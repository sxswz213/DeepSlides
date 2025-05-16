from typing import Literal
import json

from langchain.chat_models import init_chat_model
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.runnables import RunnableConfig

from langgraph.constants import Send
from langgraph.graph import START, END, StateGraph
from langgraph.types import interrupt, Command

from open_deep_research.state import (
    ReportStateInput,
    ReportStateOutput,
    Sections,
    ReportState,
    SectionState,
    SectionOutputState,
    Queries,
    Feedback
)

from open_deep_research.prompts import (
    report_planner_query_writer_instructions,
    report_planner_instructions,
    query_writer_instructions, 
    section_writer_instructions,
    final_section_writer_instructions,
    section_grader_instructions,
    section_writer_inputs
)

from open_deep_research.configuration import Configuration
from open_deep_research.utils import (
    format_sections, 
    get_config_value, 
    get_search_params, 
    select_and_execute_search,
    set_openai_api_base,
    generate_image_caption,
    generate_image_caption_v2
)

## Nodes -- 

async def process_image_input(state: ReportState, config: RunnableConfig):
    """处理图像输入并生成描述文本。
    
    这个节点：
    1. 检查状态中是否包含图像路径
    2. 如果有图像，使用Vision API生成详细描述
    3. 将描述作为额外上下文添加到后续步骤
    4. 如果未提供topic，则使用图像描述作为topic
    
    Args:
        state: 当前图状态，包含可选的图像路径
        config: 配置参数
        
    Returns:
        包含图像描述和更新topic的状态
    """
    # 检查是否提供了图像路径
    image_path = state.get("image_path")
    
    # 如果没有提供图像路径，直接返回状态
    if not image_path:
        # 确保topic存在，即使为空字符串
        if "topic" not in state:
            return {"topic": ""}
        return {}
    
    # 确保设置了正确的API基础URL
    set_openai_api_base()
    
    try:
        if not state.get("topic"):
            topic = "用户未给出主题"
        else:
            topic = state["topic"]
        # 生成图像描述
        image_result = await generate_image_caption_v2(image_path, topic)
        image_result = json.loads(image_result)
        print(image_result)
        caption, user_intent, topic = image_result["caption"], image_result["user_intent"], image_result["topic"]

        # 检查是否提供了topic
        if not state.get("topic"):
            # 如果没有提供topic，使用提取出的图像topic
            return {"caption": caption, "user_intent": user_intent, "topic": topic}
        else:
            # 如果已经提供了topic，只返回图像描述
            return {"caption": caption, "user_intent": user_intent}
    except Exception as e:
        print(f"处理图像时出错: {str(e)}")
        # 返回错误信息作为caption，并确保topic存在
        if "topic" not in state:
            return {"image_caption": f"无法处理图像: {str(e)}", "topic": ""}
        return {"image_caption": f"无法处理图像: {str(e)}"}

async def generate_report_plan(state: ReportState, config: RunnableConfig):
    """Generate the initial report plan with sections, including image caption and user intent.

    This node:
    1. Gets configuration for the report structure and search parameters
    2. Generates search queries to gather context for planning
    3. Performs web searches using those queries
    4. Uses an LLM to generate a structured plan with sections
    5. Includes image caption and user intent in planning

    Args:
        state: Current graph state containing the report topic, image caption, and user intent
        config: Configuration for models, search APIs, etc.

    Returns:
        Dict containing the generated sections
    """

    # 获取topic、图片caption和用户意图
    topic = state["topic"]
    caption = state.get("image_caption", "")
    user_intent = state.get("user_intent", "")
    feedback = state.get("feedback_on_report_plan", None)

    # Get configuration
    configurable = Configuration.from_runnable_config(config)
    report_structure = configurable.report_structure
    number_of_queries = configurable.number_of_queries
    search_api = get_config_value(configurable.search_api)
    search_api_config = configurable.search_api_config or {}
    params_to_pass = get_search_params(search_api, search_api_config)

    if isinstance(report_structure, dict):
        report_structure = str(report_structure)

    writer_provider = get_config_value(configurable.writer_provider)
    writer_model_name = get_config_value(configurable.writer_model)
    writer_model_kwargs = get_config_value(configurable.writer_model_kwargs or {})
    writer_model = init_chat_model(model=writer_model_name, model_provider=writer_provider, model_kwargs=writer_model_kwargs)
    structured_llm = writer_model.with_structured_output(Queries)

    # Format system instructions including caption and user intent
    system_instructions_query = report_planner_query_writer_instructions.format(
        topic=topic,
        caption=caption,
        user_intent=user_intent,
        report_organization=report_structure,
        number_of_queries=number_of_queries
    )

    results = await structured_llm.ainvoke([
        SystemMessage(content=system_instructions_query),
        HumanMessage(content="生成有助于规划报告各部分的搜索查询。")
    ])

    query_list = [query.search_query for query in results.queries]

    source_str = await select_and_execute_search(search_api, query_list, params_to_pass)

    image_path = state.get("image_path")
    if image_path:
        # 单独调用图片搜索API
        image_search_result = await select_and_execute_search("image_search", [image_path], params_to_pass)
        # 将图片搜索结果与之前的搜索结果合并
        source_str += f"\n\n图片搜索结果:\n{image_search_result}"

    # Include caption and user intent in report planner instructions
    system_instructions_sections = report_planner_instructions.format(
        topic=topic,
        caption=caption,
        user_intent=user_intent,
        report_organization=report_structure,
        context=source_str,
        feedback=feedback
    )

    planner_provider = get_config_value(configurable.planner_provider)
    planner_model = get_config_value(configurable.planner_model)
    planner_model_kwargs = get_config_value(configurable.planner_model_kwargs or {})

    planner_message = """生成报告的章节。您的回复必须包含一个'sections'字段，其中包含章节列表。
                        每个章节必须有：name、description、plan、research和content字段。"""

    set_openai_api_base()

    if planner_model == "claude-3-7-sonnet-latest":
        planner_llm = init_chat_model(
            model=planner_model,
            model_provider=planner_provider,
            max_tokens=20_000,
            thinking={"type": "enabled", "budget_tokens": 16_000}
        )
    else:
<<<<<<< HEAD
        planner_llm = init_chat_model(
            model=planner_model,
            model_provider=planner_provider,
            model_kwargs=planner_model_kwargs
        )

    structured_llm = planner_llm.with_structured_output(Sections)
    report_sections = await structured_llm.ainvoke([
        SystemMessage(content=system_instructions_sections),
        HumanMessage(content=planner_message)
    ])
=======
        # With other models, thinking tokens are not specifically allocated
        planner_llm = init_chat_model(model=planner_model, 
                                      model_provider=planner_provider,
                                      model_kwargs=planner_model_kwargs)

    # Generate the report sections
    structured_llm = planner_llm.with_structured_output(Sections)
    report_sections = await structured_llm.ainvoke([SystemMessage(content=system_instructions_sections),
                                             HumanMessage(content=planner_message)])
>>>>>>> dev

    sections = report_sections.sections

    return {"sections": sections}

def human_feedback(state: ReportState, config: RunnableConfig) -> Command[Literal["generate_report_plan","build_section_with_web_research"]]:
    """Get human feedback on the report plan and route to next steps.
    
    This node:
    1. Formats the current report plan for human review
    2. Gets feedback via an interrupt
    3. Routes to either:
       - Section writing if plan is approved
       - Plan regeneration if feedback is provided
    
    Args:
        state: Current graph state with sections to review
        config: Configuration for the workflow
        
    Returns:
        Command to either regenerate plan or start section writing
    """

    # Get sections
    topic = state["topic"]
    sections = state['sections']
    sections_str = "\n\n".join(
        f"Section: {section.name}\n"
        f"Description: {section.description}\n"
        f"Research needed: {'Yes' if section.research else 'No'}\n"
        for section in sections
    )

    # Get feedback on the report plan from interrupt
    interrupt_message = f"""请对以下报告计划提供反馈。
                        \n\n{sections_str}\n
                        \n此报告计划是否满足您的需求？\n传入'true'来批准报告计划。\n或者，提供反馈以重新生成报告计划："""
    
    feedback = interrupt(interrupt_message)

    # If the user approves the report plan, kick off section writing
    if isinstance(feedback, bool) and feedback is True:
        # Treat this as approve and kick off section writing
        return Command(goto=[
            Send("build_section_with_web_research", {"topic": topic, "section": s, "search_iterations": 0}) 
            for s in sections 
            if s.research
        ])
    
    # If the user provides feedback, regenerate the report plan 
    elif isinstance(feedback, str):
        # Treat this as feedback
        return Command(goto="generate_report_plan", 
                       update={"feedback_on_report_plan": feedback})
    else:
        raise TypeError(f"Interrupt value of type {type(feedback)} is not supported.")
    
async def generate_queries(state: SectionState, config: RunnableConfig):
    """Generate search queries for researching a specific section.
    
    This node uses an LLM to generate targeted search queries based on the 
    section topic and description.
    
    Args:
        state: Current state containing section details
        config: Configuration including number of queries to generate
        
    Returns:
        Dict containing the generated search queries
    """

    # Get state 
    topic = state["topic"]
    section = state["section"]

    # Get configuration
    configurable = Configuration.from_runnable_config(config)
    number_of_queries = configurable.number_of_queries

    # Set OpenAI API Base URL
    set_openai_api_base()

    # Generate queries 
    writer_provider = get_config_value(configurable.writer_provider)
    writer_model_name = get_config_value(configurable.writer_model)
    writer_model_kwargs = get_config_value(configurable.writer_model_kwargs or {})
    writer_model = init_chat_model(model=writer_model_name, model_provider=writer_provider, model_kwargs=writer_model_kwargs) 
    structured_llm = writer_model.with_structured_output(Queries)

    # Format system instructions
    system_instructions = query_writer_instructions.format(topic=topic, 
                                                           section_topic=section.description, 
                                                           number_of_queries=number_of_queries)

    # Generate queries  
    queries = await structured_llm.ainvoke([SystemMessage(content=system_instructions),
                                     HumanMessage(content="针对提供的主题生成搜索查询。")])

    return {"search_queries": queries.queries}

async def search_web(state: SectionState, config: RunnableConfig):
    """Execute web searches for the section queries.
    
    This node:
    1. Takes the generated queries
    2. Executes searches using configured search API
    3. Formats results into usable context
    
    Args:
        state: Current state with search queries
        config: Search API configuration
        
    Returns:
        Dict with search results and updated iteration count
    """

    # Get state
    search_queries = state["search_queries"]

    # Get configuration
    configurable = Configuration.from_runnable_config(config)
    search_api = get_config_value(configurable.search_api)
    search_api_config = configurable.search_api_config or {}  # Get the config dict, default to empty
    params_to_pass = get_search_params(search_api, search_api_config)  # Filter parameters

    # Web search
    query_list = [query.search_query for query in search_queries]

    # Search the web with parameters
    source_str = await select_and_execute_search(search_api, query_list, params_to_pass)

    return {"source_str": source_str, "search_iterations": state["search_iterations"] + 1}

async def write_section(state: SectionState, config: RunnableConfig) -> Command[Literal[END, "search_web"]]:
    """Write a section of the report and evaluate if more research is needed.
    
    This node:
    1. Writes section content using search results
    2. Evaluates the quality of the section
    3. Either:
       - Completes the section if quality passes
       - Triggers more research if quality fails
    
    Args:
        state: Current state with search results and section info
        config: Configuration for writing and evaluation
        
    Returns:
        Command to either complete section or do more research
    """

    # Get state 
    topic = state["topic"]
    section = state["section"]
    source_str = state["source_str"]
    
    # 提取图像信息
    images_data = []
    # dtyxs TODO: make it configurable
    max_images = 6  # 最大图像数量限制
    
    if "--- IMAGES ---" in source_str:
        try:
            # 提取图像部分
            images_section = source_str.split("--- IMAGES ---")[1].split("-" * 80)[0]
            image_blocks = images_section.strip().split("IMAGE ")[1:]  # 跳过第一个空元素
            
            for i, block in enumerate(image_blocks):
                lines = block.strip().split("\n")
                image_url = ""
                image_description = ""
                
                for line in lines:
                    if line.startswith("URL:"):
                        image_url = line.replace("URL:", "").strip()
                    elif line.startswith("DESCRIPTION:"):
                        image_description = line.replace("DESCRIPTION:", "").strip()
                
                if image_url:  # 只添加有URL的图像
                    images_data.append({
                        "index": i,
                        "url": image_url,
                        "description": image_description
                    })
        except Exception as e:
            print(f"提取图像信息时出错: {str(e)}")

    # 达到最大图像数量后停止处理
    image_num_available = len(images_data)
    if image_num_available >= max_images:
        images_data = images_data[:max_images]

    # 将图像数据格式化为JSON字符串
    images_json = json.dumps(images_data, ensure_ascii=False, indent=2) if images_data else "[]"
    
    if images_data:
        print(f"已提取 {image_num_available} 张图像（最大限制：{max_images}张）")

    # Get configuration
    configurable = Configuration.from_runnable_config(config)

    # Format system instructions
    section_writer_inputs_formatted = section_writer_inputs.format(
        topic=topic, 
        section_name=section.name, 
        section_topic=section.description, 
        context=source_str, 
        section_content=section.content,
        images_data=images_json
    )

    # Set OpenAI API Base URL
    set_openai_api_base()

    # Generate section  
    writer_provider = get_config_value(configurable.writer_provider)
    writer_model_name = get_config_value(configurable.writer_model)
    writer_model_kwargs = get_config_value(configurable.writer_model_kwargs or {})
    writer_model = init_chat_model(model=writer_model_name, model_provider=writer_provider, model_kwargs=writer_model_kwargs) 

    # dtyxs TODO: Native image input
    section_content = await writer_model.ainvoke([SystemMessage(content=section_writer_instructions),
                                           HumanMessage(content=section_writer_inputs_formatted)])
    
    # 处理返回的内容，提取图像选择信息
    content = section_content.content
    selected_image = None
    
    # 检查是否包含图像选择信息
    if "```image_selection" in content and images_data:
        try:
            # 提取图像选择JSON
            image_selection_text = content.split("```image_selection")[1].split("```")[0].strip()
            image_selection = json.loads(image_selection_text)
            
            selected_index = image_selection.get("selected_image_index", -1)
            if selected_index >= 0 and selected_index < len(images_data):
                selected_image = images_data[selected_index]
                selected_image["caption"] = image_selection.get("caption", "")
                
                # 从内容中移除图像选择部分
                content = content.split("```image_selection")[0].strip()
                
                # 在内容末尾添加选中的图像
                if selected_image:
                    content += f"\n\n![{selected_image['caption']}]({selected_image['url']})\n"
                    content += f"*{selected_image['caption']}*\n"
        except Exception as e:
            print(f"处理图像选择信息时出错: {str(e)}")
    
    # Write content to the section object  
    section.content = content

    # Grade prompt 
    section_grader_message = ("对报告进行评分并考虑针对缺失信息的后续问题。"
                              "如果评分为'pass'，则所有后续查询返回空字符串。"
                              "如果评分为'fail'，则提供具体的搜索查询以收集缺失信息。")
    
    section_grader_instructions_formatted = section_grader_instructions.format(topic=topic, 
                                                                               section_topic=section.description,
                                                                               section=section.content, 
                                                                               number_of_follow_up_queries=configurable.number_of_queries)

    # Use planner model for reflection
    planner_provider = get_config_value(configurable.planner_provider)
    planner_model = get_config_value(configurable.planner_model)
    planner_model_kwargs = get_config_value(configurable.planner_model_kwargs or {})

    # Set OpenAI API Base URL
    set_openai_api_base()

    if planner_model == "claude-3-7-sonnet-latest":
        # Allocate a thinking budget for claude-3-7-sonnet-latest as the planner model
        reflection_model = init_chat_model(model=planner_model, 
                                           model_provider=planner_provider, 
                                           max_tokens=20_000, 
                                           thinking={"type": "enabled", "budget_tokens": 16_000}).with_structured_output(Feedback)
    else:
        reflection_model = init_chat_model(model=planner_model, 
                                           model_provider=planner_provider, model_kwargs=planner_model_kwargs).with_structured_output(Feedback)
    # Generate feedback
    feedback = await reflection_model.ainvoke([SystemMessage(content=section_grader_instructions_formatted),
                                        HumanMessage(content=section_grader_message)])

    # If the section is passing or the max search depth is reached, publish the section to completed sections 
    if feedback.grade == "pass" or state["search_iterations"] >= configurable.max_search_depth:
        # Publish the section to completed sections 
        return  Command(
        update={"completed_sections": [section]},
        goto=END
    )

    # Update the existing section with new content and update search queries
    else:
        return  Command(
        update={"search_queries": feedback.follow_up_queries, "section": section},
        goto="search_web"
        )
    
async def write_final_sections(state: SectionState, config: RunnableConfig):
    """Write sections that don't require research using completed sections as context.
    
    This node handles sections like conclusions or summaries that build on
    the researched sections rather than requiring direct research.
    
    Args:
        state: Current state with completed sections as context
        config: Configuration for the writing model
        
    Returns:
        Dict containing the newly written section
    """

    # Get configuration
    configurable = Configuration.from_runnable_config(config)

    # Get state 
    topic = state["topic"]
    section = state["section"]
    completed_report_sections = state["report_sections_from_research"]
    
    # Format system instructions
    system_instructions = final_section_writer_instructions.format(topic=topic, section_name=section.name, section_topic=section.description, context=completed_report_sections)

    # Set OpenAI API Base URL
    set_openai_api_base()

    # Generate section  
    writer_provider = get_config_value(configurable.writer_provider)
    writer_model_name = get_config_value(configurable.writer_model)
    writer_model_kwargs = get_config_value(configurable.writer_model_kwargs or {})
    writer_model = init_chat_model(model=writer_model_name, model_provider=writer_provider, model_kwargs=writer_model_kwargs) 
    
    section_content = await writer_model.ainvoke([SystemMessage(content=system_instructions),
                                           HumanMessage(content="根据提供的资料生成报告章节。")])
    
    # Write content to section 
    section.content = section_content.content

    # Write the updated section to completed sections
    return {"completed_sections": [section]}

def gather_completed_sections(state: ReportState):
    """Format completed sections as context for writing final sections.
    
    This node takes all completed research sections and formats them into
    a single context string for writing summary sections.
    
    Args:
        state: Current state with completed sections
        
    Returns:
        Dict with formatted sections as context
    """

    # List of completed sections
    completed_sections = state["completed_sections"]

    # Format completed section to str to use as context for final sections
    completed_report_sections = format_sections(completed_sections)

    return {"report_sections_from_research": completed_report_sections}

def compile_final_report(state: ReportState):
    """Compile all sections into the final report.
    
    This node:
    1. Gets all completed sections
    2. Orders them according to original plan
    3. Combines them into the final report
    
    Args:
        state: Current state with all completed sections
        
    Returns:
        Dict containing the complete report
    """

    # Get sections
    sections = state["sections"]
    completed_sections = {s.name: s.content for s in state["completed_sections"]}

    # Update sections with completed content while maintaining original order
    for section in sections:
        section.content = completed_sections[section.name]

    # Compile final report
    all_sections = "\n\n".join([s.content for s in sections])

    return {"final_report": all_sections}

def initiate_final_section_writing(state: ReportState):
    """Create parallel tasks for writing non-research sections.
    
    This edge function identifies sections that don't need research and
    creates parallel writing tasks for each one.
    
    Args:
        state: Current state with all sections and research context
        
    Returns:
        List of Send commands for parallel section writing
    """

    # Kick off section writing in parallel via Send() API for any sections that do not require research
    return [
        Send("write_final_sections", {"topic": state["topic"], "section": s, "report_sections_from_research": state["report_sections_from_research"]}) 
        for s in state["sections"] 
        if not s.research
    ]

# Report section sub-graph -- 

# Add nodes 
section_builder = StateGraph(SectionState, output=SectionOutputState)
section_builder.add_node("generate_queries", generate_queries)
section_builder.add_node("search_web", search_web)
section_builder.add_node("write_section", write_section)

# Add edges
section_builder.add_edge(START, "generate_queries")
section_builder.add_edge("generate_queries", "search_web")
section_builder.add_edge("search_web", "write_section")

# Outer graph for initial report plan compiling results from each section -- 

# Add nodes
builder = StateGraph(ReportState, input=ReportStateInput, output=ReportStateOutput, config_schema=Configuration)
builder.add_node("process_image_input", process_image_input)
builder.add_node("generate_report_plan", generate_report_plan)
builder.add_node("human_feedback", human_feedback)
builder.add_node("build_section_with_web_research", section_builder.compile())
builder.add_node("gather_completed_sections", gather_completed_sections)
builder.add_node("write_final_sections", write_final_sections)
builder.add_node("compile_final_report", compile_final_report)

# Add edges
builder.add_edge(START, "process_image_input")
builder.add_edge("process_image_input", "generate_report_plan")
builder.add_edge("generate_report_plan", "human_feedback")
builder.add_edge("build_section_with_web_research", "gather_completed_sections")
builder.add_conditional_edges("gather_completed_sections", initiate_final_section_writing, ["write_final_sections"])
builder.add_edge("write_final_sections", "compile_final_report")
builder.add_edge("compile_final_report", END)

graph = builder.compile()
