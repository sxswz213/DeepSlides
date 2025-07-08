from typing import Literal
import json

from langchain.chat_models import init_chat_model
from langchain_openai import AzureChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.runnables import RunnableConfig

from langgraph.constants import Send
from langgraph.graph import START, END, StateGraph
from langgraph.types import interrupt, Command

import subprocess
import os
import datetime
import asyncio
import io
import subprocess
import httpx
from urllib.parse import urlparse

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN



import base64


from open_deep_research.state import (
    ReportStateInput,
    ReportStateOutput,
    Sections,
    ReportState,
    SectionState,
    SectionOutputState,
    Queries,
    Feedback,
    PPTSection,
    PPTOutline,
    PPTSections,
    PPTSectionState,
    PPTSlideState,
    PPTSlide,
    PPTSlideOutputState,
    PPTSectionOutputState
)

from open_deep_research.prompts import (
    report_planner_query_writer_instructions,
    report_planner_instructions,
    query_writer_instructions, 
    section_writer_instructions,
    final_section_writer_instructions,
    section_grader_instructions,
    section_writer_inputs,
    query_writer4PPT_instructions
)

from open_deep_research.configuration import Configuration
from open_deep_research.utils import (
    format_sections, 
    get_config_value, 
    get_search_params, 
    select_and_execute_search,
    set_openai_api_base,
    generate_image_caption,
    generate_image_caption_v2,
    generate_image_caption_v3
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
        image_result = await generate_image_caption_v3(image_path, topic)
        image_result = json.loads(image_result)
        print(image_result)
        caption, user_intent, topic = image_result["caption"], image_result["user_intent"], image_result["topic"]

        return {"caption": caption, "user_intent": user_intent, "topic": topic}

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
    # writer_model = init_chat_model(model=writer_model_name, model_provider=writer_provider, model_kwargs=writer_model_kwargs)
    writer_model = AzureChatOpenAI(
        model=configurable.writer_model,
        azure_endpoint=writer_model_kwargs["openai_api_base"],  # Azure's API base
        deployment_name=writer_model_kwargs["azure_deployment"],  # Azure's deployment name
        openai_api_version=writer_model_kwargs["openai_api_version"],  # Azure's API version
        temperature=0,
        max_tokens=2048
    )
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
        try:
            # 单独调用图片搜索API
            image_search_result = await select_and_execute_search("image_search", [image_path], params_to_pass)
            # 将图片搜索结果与之前的搜索结果合并
            source_str += f"\n\n图片搜索结果:\n{image_search_result}"
        except Exception as e:
            print(f"调用图片搜索API时出错: {str(e)}")

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
        # With other models, thinking tokens are not specifically allocated
        # planner_llm = init_chat_model(model=planner_model, 
        #                               model_provider=planner_provider,
        #                               model_kwargs=planner_model_kwargs)
        planner_llm = AzureChatOpenAI(
            model=configurable.writer_model,
            azure_endpoint=planner_model_kwargs["openai_api_base"],  # Azure's API base
            deployment_name=planner_model_kwargs["azure_deployment"],  # Azure's deployment name
            openai_api_version=planner_model_kwargs["openai_api_version"],  # Azure's API version
            temperature=0,
            max_tokens=2048
        )

    # Generate the report sections
    structured_llm = planner_llm.with_structured_output(Sections)
    report_sections = await structured_llm.ainvoke([SystemMessage(content=system_instructions_sections),
                                             HumanMessage(content=planner_message)])

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
    # writer_model = init_chat_model(model=writer_model_name, model_provider=writer_provider, model_kwargs=writer_model_kwargs)
    writer_model = AzureChatOpenAI(
        model=configurable.writer_model,
        azure_endpoint=writer_model_kwargs["openai_api_base"],  # Azure's API base
        deployment_name=writer_model_kwargs["azure_deployment"],  # Azure's deployment name
        openai_api_version=writer_model_kwargs["openai_api_version"],  # Azure's API version
        temperature=0,
        max_tokens=2048
    )
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
    # writer_model = init_chat_model(model=writer_model_name, model_provider=writer_provider, model_kwargs=writer_model_kwargs)
    writer_model = AzureChatOpenAI(
        model=configurable.writer_model,
        azure_endpoint=writer_model_kwargs["openai_api_base"],  # Azure's API base
        deployment_name=writer_model_kwargs["azure_deployment"],  # Azure's deployment name
        openai_api_version=writer_model_kwargs["openai_api_version"],  # Azure's API version
        temperature=0,
        max_tokens=2048
    )

    # dtyxs TODO: Native image input
    section_content = await writer_model.ainvoke([SystemMessage(content=section_writer_instructions),
                                           HumanMessage(content=section_writer_inputs_formatted)])
    
    # 处理返回的内容，提取图像选择信息
    content = section_content.content

    # 检查是否包含图像选择信息
    if "```image_selection" in content and images_data:
        try:
            # 提取图像选择JSON
            image_selection_text = content.split("```image_selection")[1].split("```")[0].strip()
            image_selection = json.loads(image_selection_text)
            
            selected_index = image_selection.get("selected_image_index", -1)
            if 0 <= selected_index < len(images_data):
                selected_image = images_data[selected_index]
                selected_image["caption"] = image_selection.get("caption", "")
                
                # 处理内容分割，保留标记前后的内容
                before_image_selection = content.split("```image_selection")[0].strip()
                after_image_selection = ""
                if "```" in content.split("```image_selection")[1]:
                    after_image_selection = content.split("```image_selection")[1].split("```", 1)[1].strip()
                
                # 重新组合内容，将图像选择部分替换为图像和标题标记
                image_content = f"\n\n![{selected_image['caption']}]({selected_image['url']})\n*{selected_image['caption']}*\n"
                if after_image_selection:
                    content = f"{before_image_selection}\n\n{image_content}\n\n{after_image_selection}"
                else:
                    content = f"{before_image_selection}\n\n{image_content}"
            else:
                # selected_index不合法，移除图像选择部分
                before_image_selection = content.split("```image_selection")[0].strip()
                after_image_selection = ""
                if "```" in content.split("```image_selection")[1]:
                    after_image_selection = content.split("```image_selection")[1].split("```", 1)[1].strip()
                
                # 重新组合内容，移除图像选择部分但保留其前后内容
                if after_image_selection:
                    content = f"{before_image_selection}\n\n{after_image_selection}"
                else:
                    content = before_image_selection
        except Exception as e:
            print(f"处理图像选择信息时出错: {str(e)}")
    
    # Write content to the section object  
    section.content = content
    section.source_str = source_str  # Store the source string in the section

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
        # reflection_model = init_chat_model(model=planner_model, 
        #                                    model_provider=planner_provider, model_kwargs=planner_model_kwargs).with_structured_output(Feedback)
        reflection_model = AzureChatOpenAI(
            model=configurable.writer_model,
            azure_endpoint=planner_model_kwargs["openai_api_base"],  # Azure's API base
            deployment_name=planner_model_kwargs["azure_deployment"],  # Azure's deployment name
            openai_api_version=planner_model_kwargs["openai_api_version"],  # Azure's API version
            temperature=0,
            max_tokens=2048
        ).with_structured_output(Feedback)
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
    # writer_model = init_chat_model(model=writer_model_name, model_provider=writer_provider, model_kwargs=writer_model_kwargs)
    writer_model = AzureChatOpenAI(
        model=configurable.writer_model,
        azure_endpoint=writer_model_kwargs["openai_api_base"],  # Azure's API base
        deployment_name=writer_model_kwargs["azure_deployment"],  # Azure's deployment name
        openai_api_version=writer_model_kwargs["openai_api_version"],  # Azure's API version
        temperature=0,
        max_tokens=2048
    )
    
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

async def generate_ppt_outline(state: ReportState, config: RunnableConfig)-> Command[Literal["generate_ppt_sections"]]:
    """
    根据演讲时长确定推荐的PPT页数，根据风格和故事线重新生成章节划分，最后基于此生成PPT大纲。

    Args:
        state: 当前状态，包含 final_report 等信息
        config: 配置参数

    Returns:
        包含PPT大纲的状态字典
    """
    configurable = Configuration.from_runnable_config(config)
    planner_model_kwargs = get_config_value(configurable.planner_model_kwargs or {})

    writer_model = AzureChatOpenAI(
        model=configurable.planner_model,
        azure_endpoint=planner_model_kwargs["openai_api_base"],
        deployment_name=planner_model_kwargs["azure_deployment"],
        openai_api_version=planner_model_kwargs["openai_api_version"],
        temperature=0,
        max_tokens=4096
    )

    topic = state.get("topic", "未指定主题")
    presentation_minutes = state.get("presentation_minutes", "15")
    style = state.get("style", "none")

    if style == "none":
        prefix = "可以参考的演讲风格：专业商务、科技现代、简约极简、创意活泼、学术严谨、故事化叙述、杂志视觉、插画卡通、复古怀旧、数据可视化。"
    else:
        prefix = f"用户期望的演讲风格：{style}"

    storyline_prompt = f"""
    你是一位经验丰富的演讲专家，要制作一份演讲PPT，现在你要确定演讲的风格、故事线和主题色（主色+辅助色）。
    演讲主题：{topic}
    演讲时长：{presentation_minutes}分钟
    {prefix}

    可以参考的故事线：
    - 问题-解决方案型：明确指出一个核心问题，并提供清晰、具体的解决方案。
    - 情境-冲突-解决-成果型：首先设定一个场景，描述面临的挑战，提供解决方法，最终呈现积极成果。
    - SCQA（背景-冲突-问题-回答）型：提供背景信息，引入冲突，明确提出关键问题并给出解决方案。
    - 时间线（过去-现在-未来）型：以时间顺序展示过去发生的事件，目前的状态，以及未来的目标。
    - 对比型（现状-未来）：清晰对比当前存在的问题与理想的未来状态，突出如何实现转变。
    - 金字塔型：从结论开始，自上而下逐层展开论据，以严谨清晰的逻辑强化核心观点。

    返回json格式：
    {{
        "style": "风格",
        "storyline": "故事线类型",
        "main_color": "推荐的主色",
        "accent_color": "推荐的辅助色"
    }}
    """
    storyline_response = await writer_model.ainvoke([
        SystemMessage(content=storyline_prompt),
        HumanMessage(content="请推荐适合的故事线")
    ])
    response_content = json.loads(storyline_response.content)
    style = response_content["style"]
    storyline = response_content["storyline"]
    main_color = response_content["main_color"]
    accent_color = response_content["accent_color"]

    ppt_length_prompt = f"""
    你是一位经验丰富的演讲专家。

    主题：{topic}
    演讲时长：{presentation_minutes}分钟

    请建议一个适合该演讲时长的PPT页数（每页内容适中，页面不拥挤，一页PPT大概对应1-2分钟的内容）。
    JSON格式：{{"recommended_slides": 10}}
    """

    ppt_length_response = await writer_model.ainvoke([
        SystemMessage(content=ppt_length_prompt),
        HumanMessage(content="请提供推荐的PPT页数。")] 
    )

    recommended_slides = json.loads(ppt_length_response.content)["recommended_slides"]

    ppt_section_distribution_prompt = f"""
    演讲主题：{topic}
    风格：{style}
    故事线：{storyline}
    推荐PPT总页数：{recommended_slides}

    注意：不要生成“提问环节”或“结束语”章节。

    根据以上信息重新规划PPT章节结构，并分配每个章节的页数。以JSON格式返回，例如：
    {{
      "section_distribution": {{
        "引言": 2,
        "方法论": 3,
        "结果": 3,
        "结论": 2
      }}
    }}
    """

    ppt_distribution_response = await writer_model.ainvoke([
        SystemMessage(content=ppt_section_distribution_prompt),
        HumanMessage(content="请规划章节结构并分配页数。")
    ])

    section_distribution = json.loads(ppt_distribution_response.content)["section_distribution"]

    ppt_outline_prompt = f"""
    你擅长设计演讲用的幻灯片大纲。
    演讲主题：{topic}
    风格：{style}
    故事线：{storyline}

    演讲内容参考资料：
    {state["final_report"]}

    每个章节的PPT页数分配如下：
    {json.dumps(section_distribution, ensure_ascii=False, indent=2)}

    请生成符合上述页数划分的PPT大纲，每页应包含：
    - 标题
    - 最多3-4个关键点

    JSON格式：
    {{
      "ppt_sections": [
        {{
          "name": "引言",
          "allocated_slides": 2,
          "slides": [{{"title": "引言", "points": ["背景介绍"]}}]
        }}
      ]
    }}
    """

    ppt_outline_response = await writer_model.ainvoke([
        SystemMessage(content=ppt_outline_prompt),
        HumanMessage(content="生成PPT大纲。")
    ])
    ppt_sections_data = json.loads(ppt_outline_response.content)["ppt_sections"]

    for section in ppt_sections_data:
        for slide in section.get("slides", []):
            slide.setdefault("codes", [])
            slide.setdefault("detail", "")
            slide.setdefault("enriched_points", "")
            slide.setdefault("path", "")

    ppt_sections = [PPTSection(**section) for section in ppt_sections_data]
    ppt_outline = PPTOutline(ppt_sections=PPTSections(sections=ppt_sections))

    return Command(
        update={
            "recommended_ppt_slides": recommended_slides,
            "section_distribution": section_distribution,
            "ppt_outline": ppt_outline,
            "ppt_sections": ppt_sections,
            "storyline": storyline,
            "style": style,
            "main_color": main_color,  
            "accent_color": accent_color
        },
        goto=[
            Send("generate_ppt_sections", {"topic": topic, "ppt_section": ppt_section, "style": style, "main_color": main_color, "accent_color": accent_color})
            for ppt_section in ppt_sections
        ]
    )

async def save_image_from_url(image_url, image_name, topic, ppt_section_name, slide_index):
    """
    异步下载图片并保存到本地，自动根据URL后缀决定图片类型。

    Args:
        image_url: 图片URL
        image_name: 保存的图片名称
        topic: 主题目录名称
        ppt_section_name: PPT章节名称
        slide_index: 幻灯片索引

    Returns:
        str: 图片保存路径或空字符串（失败时）
    """
    try:
        # 异步网络请求下载图片
        async with httpx.AsyncClient(timeout=30.0) as client:
            response = await client.get(image_url)
            response.raise_for_status()

            # 提取图片后缀
            parsed_url = urlparse(image_url)
            ext = os.path.splitext(parsed_url.path)[-1].lower()

            # 默认后缀为 .jpg，如果没有或后缀不标准时使用默认
            if ext not in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff']:
                ext = '.jpg'

            # 异步创建目录（避免阻塞）
            save_dir = os.path.join(
                ".", "saves", topic, "images", ppt_section_name, f"slide_{slide_index+1}"
            )
            await asyncio.to_thread(os.makedirs, save_dir, exist_ok=True)

            # 构建完整图片路径
            image_path = os.path.join(save_dir, f"{image_name}{ext}")

            # 异步保存文件（避免阻塞）
            await asyncio.to_thread(
                lambda: open(image_path, 'wb').write(response.content)
            )

            print(f"图片成功保存到：{image_path}")
            return image_path

    except httpx.HTTPError as e:
        print(f"HTTP请求错误: {str(e)}")
        return ""
    except Exception as e:
        print(f"其他错误: {str(e)}")
        return ""

async def truncate_by_characters(text, max_chars=250000):
    if len(text) > max_chars:
        return text[:max_chars]
    return text

async def enrich_slide_content(state: PPTSlideState, config: RunnableConfig):
    """
    第一部分：根据幻灯片的标题和要点生成详细的内容描述，并生成幻灯片布局描述。

    Args:
        state: 包含当前PPT幻灯片信息的状态。
        config: 配置参数。

    Returns:
        Tuple: 包含生成的详细内容和幻灯片布局描述。
    """
    configurable = Configuration.from_runnable_config(config)
    number_of_queries = configurable.number_of_queries
    writer_model_kwargs = configurable.writer_model_kwargs or {}
    
    # 获取当前幻灯片信息
    topic = state["topic"]
    ppt_section = state["ppt_section"]
    slide_index = state["slide_index"]
    style = state.get("style")
    main_color = state.get("main_color")
    accent_color = state.get("accent_color")
    slide = ppt_section.slides[slide_index]
    slide_title = slide.title
    slide_points = slide.points

    # 设置OpenAI API的基础URL
    set_openai_api_base()

    # 使用Azure模型生成查询语句
    writer_model = AzureChatOpenAI(
        model=configurable.writer_model,
        azure_endpoint=writer_model_kwargs["openai_api_base"],  # Azure的API基础URL
        deployment_name=writer_model_kwargs["azure_deployment"],  # Azure的部署名称
        openai_api_version=writer_model_kwargs["openai_api_version"],  # Azure的API版本
        temperature=0.7,
        max_tokens=4096
    )

    # 使用结构化输出生成查询语句
    structured_llm = writer_model.with_structured_output(Queries)

    # 格式化系统指令
    system_instructions = query_writer4PPT_instructions.format(
        topic=topic,
        section_topic=ppt_section.name,
        slide_title=slide_title,
        slide_points=", ".join(slide_points),
        number_of_queries=number_of_queries
    )

    # 生成查询语句
    query_response = await structured_llm.ainvoke([
        SystemMessage(content=system_instructions),
        HumanMessage(content="根据以上信息生成相关搜索查询")
    ])

    # 从返回的查询结果中提取查询语句
    search_queries = [query.search_query for query in query_response.queries]

    # 执行Web搜索
    search_api = get_config_value(configurable.search_api)
    search_api_config = configurable.search_api_config or {}
    params_to_pass = get_search_params(search_api, search_api_config)

    source_str = await select_and_execute_search(search_api, search_queries, params_to_pass)

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
                    # 保存图像并获取保存路径
                    image_path = await save_image_from_url(image_url, f"image_{i+1}", topic, ppt_section.name, slide_index)
                    images_data.append({
                        "index": i,
                        "url": image_url,
                        "description": image_description,
                        "local_path": image_path  # 添加本地图片路径
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

    images_data = images_data[:max_images] if len(images_data) >= max_images else images_data
    embedded_images = []
    for img in images_data:
        local_path = img.get("local_path", "")
        if local_path and os.path.exists(local_path):
            try:
                img_bytes = await asyncio.to_thread(
                    lambda path=local_path: open(path, "rb").read()
                )

                img_base64 = base64.b64encode(img_bytes).decode()
                ext = os.path.splitext(local_path)[1].lower().replace('.', '')
                mime_type = f'image/{ext if ext != "jpg" else "jpeg"}'

                embedded_images.append({
                    "index": img["index"],
                    "path": local_path,
                    "description": img["description"],
                    # "base64": f"data:{mime_type};base64,{img_base64}"
                })
            except Exception as e:
                print(f"加载图片时出错: {str(e)}")
    
    images_json_embedded = json.dumps(embedded_images, ensure_ascii=False, indent=2) if embedded_images else "[]"
    
    if embedded_images:
        print(f"成功加载 {len(embedded_images)} 张图片为Base64格式。")

    # 扩展幻灯片内容
    content_enrichment_prompt = f"""
    根据以下要点，请扩展幻灯片的内容。每个要点请写出详细描述，并确保内容与搜索结果相关联。

    幻灯片主题：{topic}
    幻灯片章节：{ppt_section.name}
    幻灯片标题：{slide_title}
    要点：{', '.join(slide_points)}

    请扩展并详细描述每个要点，适合用于演讲演示，确保语言简洁但内容详实，一般每点扩展后的内容不宜超过50字。以JSON格式返回，注意仅输出json，不要输出其他文字：

    {{
        "enriched_points": [
            {{"point_title": "要点标题1", "expanded_content": "详细扩展后的内容1"}},
            {{"point_title": "要点标题2", "expanded_content": "详细扩展后的内容2"}},
            {{"point_title": "要点标题3", "expanded_content": "详细扩展后的内容3"}},
            ...
        ]
    }}
    """

    enrichment_response = await writer_model.ainvoke([
        SystemMessage(content=content_enrichment_prompt),
        HumanMessage(content="请扩展幻灯片内容")
    ])

    enriched_points = json.loads(enrichment_response.content.replace("```json", "").replace("```", "").strip())["enriched_points"]

    # 第二阶段：使用coder_model生成幻灯片布局描述
    coder_model_kwargs = configurable.coder_model_kwargs or {}
    coder_model = AzureChatOpenAI(
        model=configurable.coder_model,
        azure_endpoint=coder_model_kwargs["openai_api_base"],
        deployment_name=coder_model_kwargs["azure_deployment"],
        openai_api_version=coder_model_kwargs["openai_api_version"],
        # temperature=0.7,
        # max_tokens=4096
        max_completion_tokens=4096
    )

    detail_prompt = f"""
    你是一位资深的幻灯片设计师，负责设计幻灯片布局。
    根据以下详细内容生成幻灯片页面布局描述，输出为JSON：

    标题: {slide_title}
    详细要点: {json.dumps(enriched_points, ensure_ascii=False)}
    幻灯片风格: {style}
    主色: {main_color}
    辅助色: {accent_color}

    一级标题字号：36
    其他级别标题字号：20
    正文字号：20
    中文字体：微软雅黑
    英文字体：Arial

    这里是一些可用的图片，你可以选择性地使用，放置在PPT中。
    <图像列表>
    {images_json_embedded}
    </图像列表>

    仅JSON格式，不要输出其他文字：
    {{
        "layout": "布局类型",
        "content_details": ["标题、详细内容、图片的位置布局、长宽、图片路径等"],
        "design_style": "设计风格"
    }}
    """

    detail_response = await coder_model.ainvoke([
        SystemMessage(content=detail_prompt),
        HumanMessage(content="请输出幻灯片布局的JSON")
    ])

    slide_detail = json.loads(detail_response.content.replace("```json", "").replace("```", "").strip())

    return {"enriched_points": enriched_points, 
            "slide_detail": slide_detail}


async def generate_slide_code_and_execute(state: PPTSlideState, config: RunnableConfig):
    enriched_points = state["enriched_points"]
    slide_detail = state["slide_detail"]

    configurable = Configuration.from_runnable_config(config)
    coder_model_kwargs = configurable.coder_model_kwargs or {}

    coder_model = AzureChatOpenAI(
        model=configurable.coder_model,
        azure_endpoint=coder_model_kwargs["openai_api_base"],
        deployment_name=coder_model_kwargs["azure_deployment"],
        openai_api_version=coder_model_kwargs["openai_api_version"],
        # temperature=0.3,
        # max_tokens=4096
        max_completion_tokens=4096
    )

    topic = state["topic"]
    ppt_section = state["ppt_section"]
    slide_index = state["slide_index"]
    main_color = state.get("main_color")
    accent_color = state.get("accent_color")
    style = state.get("style")
    slide_title = ppt_section.slides[slide_index].title
    slide_points = ppt_section.slides[slide_index].points

    save_dir = os.path.join(".", "saves", topic)
    await asyncio.to_thread(os.makedirs, save_dir, exist_ok=True)

    error_message = ""
    previous_code = ""
    python_code = None
    execution_successful = False

    async def run_script(script_path):
        def run():
            try:
                proc_result = subprocess.run(
                    ["python", script_path],
                    capture_output=True,
                    text=True,
                    timeout=60
                )
                return proc_result.returncode, proc_result.stdout, proc_result.stderr
            except subprocess.TimeoutExpired as e:
                return -1, "", f"超时异常: {str(e)}"
            except Exception as e:
                return -1, "", f"其他异常: {str(e)}"
        return await asyncio.to_thread(run)

    # 保留循环5次重试逻辑
    for attempt in range(3):
        code_prompt = f"""
        根据以下幻灯片详细描述，生成使用python-pptx库创建幻灯片的Python代码：

        标题: {slide_title}
        详细要点: {json.dumps(enriched_points, ensure_ascii=False)}
        幻灯片描述：{json.dumps(slide_detail, ensure_ascii=False, indent=2)}
        幻灯片风格: {style}
        主色: {main_color} 
        辅助色: {accent_color}


        一级标题字号：36
        其他级别标题字号：20
        正文字号：20
        中文字体：微软雅黑
        英文字体：Arial

        代码要求：
        1. 导入必要库。
        2. 创建幻灯片并确保采用宽屏标准比例: 16:9（13.33 英寸 × 7.5 英寸）。
        3. 根据详细描述在指定位置添加标题、要点和图片，设置字体和样式，确保明确设置每个元素的大小以防止重叠遮挡，注意设置文本的自动换行。
        4. 幻灯片均使用空白布局，标题、内容、图片均作为普通元素放置。
        5. 保存文件名为：\"{save_dir}/{ppt_section.name}_slide_{slide_index + 1}.pptx\"



        请根据这些信息提供完整、可执行的Python代码。注意：仅输出python代码，不要输出其他文字。
        """

        code_response = await coder_model.ainvoke([
            SystemMessage(content=code_prompt),
            HumanMessage(content="生成完整Python代码")
        ])

        python_code = code_response.content.replace("```python", "").replace("```", "").strip()

        script_path = os.path.join(save_dir, f"{ppt_section.name}_slide_{slide_index + 1}.py")

        await asyncio.to_thread(
            lambda: open(script_path, "w", encoding="utf-8").write(python_code)
        )

        returncode, stdout, stderr = await run_script(script_path)

        if returncode == 0:
            execution_successful = True
            break
        else:
            print(python_code)  # 输出用于排查错误
            error_message = stderr
            previous_code = python_code
            print(f"代码执行失败，尝试第 {attempt + 1}/3 次，错误信息：{stderr}")
    retry_count = state.get("retry_count", 0)

    if not execution_successful:
        # raise RuntimeError("代码生成执行失败，已达到最大尝试次数")
            return {
                "codes": [python_code],
                "path": "none",
                "title": slide_title,
                "points": slide_points,
            }

    return {
        "codes": [python_code],
        "path": os.path.abspath(os.path.join(save_dir, f"{ppt_section.name}_slide_{slide_index + 1}.pptx")),
        "title": slide_title,
        "points": slide_points,
    }


def ppt_to_image(slide_ppt_path, image_path):
    """
    使用 unoconv 将PPT幻灯片导出为图片
    """
    print(f"将幻灯片 {slide_ppt_path} 转换为图片 {image_path}")
    
    # 使用 unoconv 将 PPT 转换为 PNG 格式
    # try:
        # 使用 unoconv 命令行工具来将 ppt 文件转换为图片
    command = [
            "C:\\Windows\\unoconv.bat",  # 调用 unoconv 命令
            "-f", "png",  # 转换为 png 格式
            "-o", image_path,  # 输出路径
            slide_ppt_path  # 输入的 PPT 文件路径
    ]
        
        # 调用 unoconv 命令
    subprocess.run(command, check=True)
    print(f"转换成功：{slide_ppt_path} -> {image_path}")
    
    # except subprocess.CalledProcessError as e:
    #     print(f"转换失败：{str(e)}")

async def ppt_slide_to_image_and_validate(state: PPTSlideState, config: RunnableConfig):
    """
    将生成的PPT幻灯片转换为图片，并使用大模型检查布局合理性。

    Args:
        state: 当前PPT幻灯片状态。
        config: 配置参数。

    Returns:
        Command: 如果布局有效，返回Command继续执行；否则返回Command跳转到enrich_slide_content。
    """
    configurable = Configuration.from_runnable_config(config)
    planner_model_kwargs = configurable.planner_model_kwargs or {}

    planner_model = AzureChatOpenAI(
        model=configurable.planner_model,
        azure_endpoint=planner_model_kwargs["openai_api_base"],
        deployment_name=planner_model_kwargs["azure_deployment"],
        openai_api_version=planner_model_kwargs["openai_api_version"],
        temperature=0,
        max_tokens=2048
    )

    slide_ppt_path = state["path"]
    codes = state["codes"]
    title = state["title"]
    points = state["points"]
    enriched_points = state["enriched_points"]
    slide_detail = state["slide_detail"]
    max_retry_count = state.get("max_retry_count", 3)  # 默认 3 次重试
    retry_count = state.get("retry_count", 0)
    path = state.get("path")

    print("当前幻灯片:",slide_ppt_path,"当前重复次数:", retry_count, "最大重试次数:", max_retry_count)
    
    if retry_count >= max_retry_count:
        print("最大重试次数已达，停止进一步处理")
        generated_slide = PPTSlide(
            title=title,
            points=points,
            codes=codes,
            enriched_points=json.dumps(enriched_points, ensure_ascii=False, indent=2),
            detail=json.dumps(slide_detail, ensure_ascii=False, indent=2)
        )
        return Command(
            update={"completed_slides": [generated_slide]},
            goto=END
        )

    if path == "none":
        return Command(
            update={"layout_valid": False, "retry_count": retry_count + 1},
            goto="enrich_slide_content"
        )
    output_folder = os.path.dirname(slide_ppt_path)
    image_path = slide_ppt_path.replace(".pptx", ".png")
    print(f"将幻灯片 {slide_ppt_path} 转换为图片 {image_path}")

    # 使用 asyncio.to_thread 将 PowerPoint 操作移到单独的线程中
    # await asyncio.to_thread(ppt_to_image, slide_ppt_path, image_path)
    try:
        # 直接调用同步的 ppt_to_image 函数（无异步操作）
        await asyncio.to_thread(ppt_to_image, slide_ppt_path, image_path)
    except Exception as e:
        print(f"Error converting PPT to image: {e}")
        return Command(
            update={"conversion_failed": True, "retry_count": retry_count + 1},
            goto="enrich_slide_content"
        )
    print(f"幻灯片转换为图片成功，保存路径：{image_path}")
    # 将图片发送给大模型进行布局检查
    # - 幻灯片内容布局均匀，不偏重于某一侧。
    # - 文本之间没有遮挡。
    # - 所有文本内容均位于页面范围内，没有超出。
    # - 所有要点都清晰地展示。
    validation_prompt = f"""
    你是一位专业的幻灯片设计审查师，请检查幻灯片布局：

    - 文本和图片之间没有遮挡或重叠现象。
    - 所有内容均位于页面范围内，没有超出页面边界。

    请根据以上规则进行检查，如果不存在以上问题，则返回pass；如果存在以上任一问题，返回retry。
    """

    image_content = await asyncio.to_thread(
        lambda: open(image_path, "rb").read()
    )

    response = await planner_model.ainvoke([
        SystemMessage(content=validation_prompt),
        HumanMessage(content=[{"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64.b64encode(image_content).decode()}"}}])
    ])
    print(f"布局检查结果：{response.content}")
    # 如果模型返回的结果中包含有效的布局检查信息（比如返回“yes”表示有效）
    if "pass" in response.content.lower():
        # 布局有效，生成幻灯片并返回
        generated_slide = PPTSlide(
            title=title,
            points=points,
            codes=codes,
            enriched_points=json.dumps(enriched_points, ensure_ascii=False, indent=2),
            detail=json.dumps(slide_detail, ensure_ascii=False, indent=2)
        )
        return Command(
            update={"completed_slides": [generated_slide]},
            goto=END
        )
    else:
        retry_count += 1

        # 如果超过最大重试次数，结束流程
        if retry_count >= max_retry_count:
            generated_slide = PPTSlide(
                title=title,
                points=points,
                codes=codes,
                enriched_points=json.dumps(enriched_points, ensure_ascii=False, indent=2),
                detail=json.dumps(slide_detail, ensure_ascii=False, indent=2)
            )
            return Command(
                update={"completed_slides": [generated_slide]},
                goto=END
            )
        # 布局无效，返回到enrich_slide_content重新生成内容
        return Command(
            update={"layout_valid": False, "retry_count": retry_count},
            goto="enrich_slide_content"
        )





async def generate_ppt_section_start(state: PPTSectionState):
    """
    为每个PPT章节开始生成幻灯片时，初始化一个空的幻灯片列表，并为每页PPT启动一个子图。

    Args:
        state: 包含章节信息、PPT章节信息的状态

    Returns:
        Dict: 包含空的 `generated_slides` 列表，并进入 `ppt_slide_subgraph`
    """
    # 从状态中获取章节信息和PPT章节信息
    topic = state["topic"]  # 主题
    # section = state["section"]  # 章节信息
    ppt_section = state["ppt_section"]  # PPT章节信息
    style = state.get("style")
    main_color = state.get("main_color")  # 主色
    accent_color = state.get("accent_color")  # 辅助

    # 初始化幻灯片列表，准备开始生成幻灯片
    generated_slides = []

    # 打印相关信息（用于调试）
    print(f"开始生成PPT章节：{ppt_section.name}")

    # 获取该章节需要的页数
    num_slides = ppt_section.allocated_slides  # 获取分配的页数
    
    # 为每一页PPT启动一个子图
    return Command(
        update={"generated_slides": generated_slides},
        goto=[
            Send("generate_slide", {
                "topic": topic, 
                # "section": section, 
                "style": style,
                "ppt_section": ppt_section, 
                "slide_index": slide_index,  # 为每一页传递幻灯片的索引
                "main_color": main_color,
                "accent_color": accent_color,
                "max_retry_count": 3,  # 设置最大重试次数
                "retry_count": 0  # 初始化重试计数
            })
            for slide_index in range(num_slides)  # 根据分配的页数启动相应数量的子图
        ]
    )


async def generate_ppt_section_end(state: PPTSectionState):
    ppt_section = state["ppt_section"]
    ppt_section.slides = state["completed_slides"]

    return Command(
        update={"completed_ppt_sections": [ppt_section]},
        goto=END
    )

async def generate_cover_slide(state: ReportState, config: RunnableConfig):
    """生成封面幻灯片，包含布局检查步骤，最多循环3次"""
    topic = state["topic"]
    style = state.get("style", "none")
    main_color = state.get("main_color", "#FFFFFF")
    accent_color = state.get("accent_color", "#000000")

    save_dir = os.path.join(".", "saves", topic)
    cover_path = os.path.join(save_dir, "cover_slide.pptx")
    script_path = os.path.join(save_dir, "cover_slide.py")

    configurable = Configuration.from_runnable_config(config)
    coder_model_kwargs = configurable.coder_model_kwargs or {}
    coder_model = AzureChatOpenAI(
        model=configurable.coder_model,
        azure_endpoint=coder_model_kwargs["openai_api_base"],
        deployment_name=coder_model_kwargs["azure_deployment"],
        openai_api_version=coder_model_kwargs["openai_api_version"],
        # temperature=0.3,
        # max_tokens=4096
        max_completion_tokens=4096
    )

    async def run_script(script_path):
        def run():
            try:
                proc_result = subprocess.run(
                    ["python", script_path],
                    capture_output=True,
                    text=True,
                    timeout=60
                )
                return proc_result.returncode, proc_result.stdout, proc_result.stderr
            except subprocess.TimeoutExpired as e:
                return -1, "", f"超时异常: {str(e)}"
            except Exception as e:
                return -1, "", f"其他异常: {str(e)}"
        return await asyncio.to_thread(run)

    error_message = ""
    previous_code = ""

    validation_prompt = """
    你是一位专业的幻灯片设计审查师，请检查幻灯片布局：

    - 文本和图片之间没有遮挡或重叠现象。
    - 所有内容均位于页面范围内，没有超出页面边界。

    请根据以上规则进行检查，如果不存在以上问题，则返回pass；如果存在以上任一问题，返回retry。
    """

    for attempt in range(3):
        layout_prompt = f"""
        请为幻灯片封面设计一个布局，标题为：{topic}
        幻灯片风格: {style}
        主色: {main_color}
        辅助色: {accent_color}

        幻灯片封面应包括标题、演讲者姓名、日期等关键信息。

        {f"上一次的错误信息: {error_message}" if error_message else ""}
        {f"上一次生成的代码: {previous_code}" if previous_code else ""}
        """

        layout_response = await coder_model.ainvoke([
            {"role": "system", "content": layout_prompt},
            {"role": "user", "content": "生成布局描述"}
        ])

        layout_description = layout_response.content.strip()

        code_generation_prompt = f"""
        根据以下布局描述生成一个使用python-pptx库制作封面幻灯片的完整Python代码：

        布局描述：{layout_description}

        代码要求：
        1. 导入必要库。
        2. 创建幻灯片并确保采用宽屏标准比例: 16:9（13.33 英寸 × 7.5 英寸）。
        3. 页面背景采用与页面大小相同的矩形设置，不要直接设置slide.background。
        4. 保存PPT文件到路径：{cover_path}

        {f"上一次的错误信息: {error_message}" if error_message else ""}
        {f"上一次生成的代码: {previous_code}" if previous_code else ""}

        请根据这些信息提供完整、可执行的Python代码。注意：仅输出python代码，不要输出其他文字。
        """

        code_response = await coder_model.ainvoke([
            {"role": "system", "content": code_generation_prompt},
            {"role": "user", "content": "生成Python代码"}
        ])

        python_code = code_response.content.replace("```python", "").replace("```", "").strip()

        await asyncio.to_thread(
            lambda: open(script_path, "w", encoding="utf-8").write(python_code)
        )

        returncode, stdout, stderr = await run_script(script_path)

        if returncode != 0:
            error_message = stderr
            previous_code = python_code
            print(f"尝试{attempt + 1}失败，错误信息：{stderr}")
            continue

        image_path = cover_path.replace(".pptx", ".png")
        try:
            await asyncio.to_thread(ppt_to_image, cover_path, image_path)
        except Exception as e:
            print(f"幻灯片转图片失败: {e}")
            continue

        image_content = await asyncio.to_thread(lambda: open(image_path, "rb").read())

        response = await coder_model.ainvoke([
            SystemMessage(content=validation_prompt),
            HumanMessage(content=[{"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64.b64encode(image_content).decode()}"}}])
        ])

        if response.content.strip() == "pass":
            print(f"布局检查通过，路径：{cover_path}")
            return {"cover_slide_path": cover_path, "cover_layout_description": layout_description}
        else:
            error_message = "布局检查未通过"
            previous_code = python_code
            print(f"布局检查未通过，尝试重试：{attempt + 1}")

    # raise RuntimeError("生成封面幻灯片失败，已达到最大尝试次数")
    return {"cover_slide_path": cover_path, "cover_layout_description": layout_description}

async def generate_section_cover_slides(state: ReportState, config: RunnableConfig):
    """生成章节封面幻灯片，仅检查第一个章节的布局有效性，最多循环3次"""
    topic = state["topic"]
    ppt_sections = state["ppt_sections"]
    style = state.get("style", "none")
    main_color = state.get("main_color", "#FFFFFF")
    accent_color = state.get("accent_color", "#000000")

    save_dir = os.path.join(".", "saves", topic)
    script_path = os.path.join(save_dir, "section_cover_slide.py")

    configurable = Configuration.from_runnable_config(config)
    coder_model_kwargs = configurable.coder_model_kwargs or {}
    coder_model = AzureChatOpenAI(
        model=configurable.coder_model,
        azure_endpoint=coder_model_kwargs["openai_api_base"],
        deployment_name=coder_model_kwargs["azure_deployment"],
        openai_api_version=coder_model_kwargs["openai_api_version"],
        # temperature=0.3,
        # max_tokens=4096
        max_completion_tokens=4096
    )

    async def run_script(script_path):
        def run():
            try:
                proc_result = subprocess.run(
                    ["python", script_path],
                    capture_output=True,
                    text=True,
                    timeout=120
                )
                return proc_result.returncode, proc_result.stdout, proc_result.stderr
            except subprocess.TimeoutExpired as e:
                return -1, "", f"超时异常: {str(e)}"
            except Exception as e:
                return -1, "", f"其他异常: {str(e)}"
        return await asyncio.to_thread(run)

    error_message = ""
    previous_code = ""

    validation_prompt = """
    你是一位专业的幻灯片设计审查师，请检查幻灯片布局：

    - 文本和图片之间没有遮挡或重叠现象。
    - 所有内容均位于页面范围内，没有超出页面边界。

    请根据以上规则进行检查，如果不存在以上问题，则返回pass；如果存在以上任一问题，返回retry。
    """

    chapters_list = [section.name for section in ppt_sections]

    for attempt in range(3):
        layout_prompt = f"""
        请为幻灯片章节封面设计一个通用版式，适用于所有章节
        幻灯片主题: {topic}
        幻灯片风格: {style}
        主色: {main_color}
        辅助色: {accent_color}

        请返回布局描述，包含元素位置、大小及字体信息。

        {f"上一次的错误信息: {error_message}" if error_message else ""}
        {f"上一次生成的代码: {previous_code}" if previous_code else ""}
        """

        layout_response = await coder_model.ainvoke([
            {"role": "system", "content": layout_prompt},
            {"role": "user", "content": "生成布局描述"}
        ])

        layout_description = layout_response.content.strip()

        code_generation_prompt = f"""
        根据以下布局描述生成一个使用python-pptx库制作章节封面幻灯片的Python脚本：

        布局描述：{layout_description}

        脚本要求：
        1. 导入必要库。
        2. 创建幻灯片并确保采用宽屏标准比例: 16:9（13.33 英寸 × 7.5 英寸）。
        3. 代码中页面背景采用与页面大小相同的矩形设置，不要直接设置slide.background。
        4. 使用以下章节列表，为每个章节生成一个单独的pptx文件，文件保存路径为{save_dir}/section_slide_{{章节序号}}.pptx。

        章节列表：{chapters_list}

        请根据这些信息提供完整、可执行的Python代码。注意：仅输出python代码，不要输出其他文字。
        """

        code_response = await coder_model.ainvoke([
            {"role": "system", "content": code_generation_prompt},
            {"role": "user", "content": "生成Python代码"}
        ])

        python_code = code_response.content.replace("```python", "").replace("```", "").strip()

        await asyncio.to_thread(lambda: open(script_path, "w", encoding="utf-8").write(python_code))

        returncode, stdout, stderr = await run_script(script_path)

        first_slide_path = os.path.join(save_dir, "section_slide_1.pptx")

        if returncode != 0:
            error_message = stderr
            previous_code = python_code
            print(f"尝试{attempt + 1}失败，错误信息：{stderr}")
            continue

        image_path = first_slide_path.replace(".pptx", ".png")
        try:
            await asyncio.to_thread(ppt_to_image, first_slide_path, image_path)
        except Exception as e:
            print(f"幻灯片转图片失败: {e}")
            continue

        image_content = await asyncio.to_thread(lambda: open(image_path, "rb").read())

        response = await coder_model.ainvoke([
            SystemMessage(content=validation_prompt),
            HumanMessage(content=[{"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64.b64encode(image_content).decode()}"}}])
        ])

        if response.content.strip() == "pass":
            print(f"第一个章节布局检查通过，路径：{save_dir}")
            return {"section_slides_path": save_dir, "layout_description": layout_description}
        else:
            error_message = "布局检查未通过"
            previous_code = python_code
            print(f"第一个章节布局检查未通过，尝试重试：{attempt + 1}")

    # raise RuntimeError("生成章节封面幻灯片失败，已达到最大尝试次数")
    return {"section_slides_path": save_dir, "layout_description": layout_description}


async def generate_end_slide(state: ReportState, config: RunnableConfig):
    """生成封底幻灯片，包含布局检查步骤，最多循环3次"""
    topic = state["topic"]
    style = state.get("style", "专业商务")
    main_color = state.get("main_color", "#FFFFFF")
    accent_color = state.get("accent_color", "#000000")

    save_dir = os.path.join(".", "saves", topic)
    end_path = os.path.join(save_dir, "end_slide.pptx")
    script_path = os.path.join(save_dir, "end_slide.py")

    configurable = Configuration.from_runnable_config(config)
    coder_model_kwargs = configurable.coder_model_kwargs or {}
    coder_model = AzureChatOpenAI(
        model=configurable.coder_model,
        azure_endpoint=coder_model_kwargs["openai_api_base"],
        deployment_name=coder_model_kwargs["azure_deployment"],
        openai_api_version=coder_model_kwargs["openai_api_version"],
        # temperature=0.3,
        # max_tokens=4096
        max_completion_tokens=4096
    )

    async def run_script(script_path):
        def run():
            try:
                proc_result = subprocess.run(
                    ["python", script_path],
                    capture_output=True,
                    text=True,
                    timeout=60
                )
                return proc_result.returncode, proc_result.stdout, proc_result.stderr
            except subprocess.TimeoutExpired as e:
                return -1, "", f"超时异常: {str(e)}"
            except Exception as e:
                return -1, "", f"其他异常: {str(e)}"
        return await asyncio.to_thread(run)

    error_message = ""
    previous_code = ""

    validation_prompt = """
    你是一位专业的幻灯片设计审查师，请检查幻灯片布局：

    - 文本和图片之间没有遮挡或重叠现象。
    - 所有内容均位于页面范围内，没有超出页面边界。

    请根据以上规则进行检查，如果不存在以上问题，则返回pass；如果存在以上任一问题，返回retry。
    """

    for attempt in range(3):
        layout_prompt = f"""
        请为幻灯片封底设计一个布局
        幻灯片主题: {topic}
        幻灯片风格: {style}
        主色: {main_color}
        辅助色: {accent_color}

        请返回布局描述，包含元素内容、位置、大小及字体信息。

        {f"上一次的错误信息: {error_message}" if error_message else ""}
        {f"上一次生成的代码: {previous_code}" if previous_code else ""}
        """

        layout_response = await coder_model.ainvoke([
            {"role": "system", "content": layout_prompt},
            {"role": "user", "content": "生成布局描述"}
        ])

        layout_description = layout_response.content.strip()

        code_generation_prompt = f"""
        根据以下布局描述生成一个使用python-pptx库制作封底幻灯片的完整Python代码：

        布局描述：{layout_description}

        代码要求：
        1. 导入必要库。
        2. 创建幻灯片并确保采用宽屏标准比例: 16:9（13.33 英寸 × 7.5 英寸）。
        3. 页面背景采用与页面大小相同的矩形设置，不要直接设置slide.background。
        4. 保存PPT文件到路径：{end_path}

        {f"上一次的错误信息: {error_message}" if error_message else ""}
        {f"上一次生成的代码: {previous_code}" if previous_code else ""}

        请根据这些信息提供完整、可执行的Python代码。注意：仅输出python代码，不要输出其他文字。
        """

        code_response = await coder_model.ainvoke([
            {"role": "system", "content": code_generation_prompt},
            {"role": "user", "content": "生成Python代码"}
        ])

        python_code = code_response.content.replace("```python", "").replace("```", "").strip()

        await asyncio.to_thread(
            lambda: open(script_path, "w", encoding="utf-8").write(python_code)
        )

        returncode, stdout, stderr = await run_script(script_path)

        if returncode != 0:
            error_message = stderr
            previous_code = python_code
            print(f"尝试{attempt + 1}失败，错误信息：{stderr}")
            continue

        image_path = end_path.replace(".pptx", ".png")
        try:
            await asyncio.to_thread(ppt_to_image, end_path, image_path)
        except Exception as e:
            print(f"幻灯片转图片失败: {e}")
            continue

        image_content = await asyncio.to_thread(lambda: open(image_path, "rb").read())

        response = await coder_model.ainvoke([
            SystemMessage(content=validation_prompt),
            HumanMessage(content=[{"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64.b64encode(image_content).decode()}"}}])
        ])

        if response.content.strip() == "pass":
            print(f"布局检查通过，路径：{end_path}")
            return {"end_slide_path": end_path, "end_layout_description": layout_description}
        else:
            error_message = "布局检查未通过"
            previous_code = python_code
            print(f"布局检查未通过，尝试重试：{attempt + 1}")

    # raise RuntimeError("生成封底幻灯片失败，已达到最大尝试次数")
    return {"end_slide_path": end_path, "end_layout_description": layout_description}


async def compile_ppt(state: ReportState, config: RunnableConfig):
    """
    异步方式合并所有生成的pptx文件到一个pptx文件中，包含封面、章节封面和封底。

    Args:
        state: 当前状态，包含所有完成的幻灯片信息。

    Returns:
        Dict: 更新后的状态，包含最终合并的PPT路径。
    """
    topic = state["topic"]
    ppt_sections = state["ppt_sections"]

    save_dir = os.path.join(".", "saves", topic)
    final_ppt_path = os.path.join(save_dir, f"{topic}_final.pptx")

    def merge_ppts_sync():
        final_presentation = Presentation()
        final_presentation.slide_width = Inches(13.33)
        final_presentation.slide_height = Inches(7.5)

        paths = ["cover_slide.pptx"]
        for idx, section in enumerate(ppt_sections):
            # 加入章节封面
            paths.append(f"section_slide_{idx+1}.pptx")
            # 加入该章节所有幻灯片
            paths.extend([
                f"{section.name}_slide_{slide_index}.pptx"
                for slide_index, _ in enumerate(section.slides, start=1)
            ])
        paths.append("end_slide.pptx")

        for ppt_file in paths:
            ppt_path = os.path.join(save_dir, ppt_file)
            if not os.path.exists(ppt_path):
                continue

            presentation = Presentation(ppt_path)
            for slide in presentation.slides:
                new_slide = final_presentation.slides.add_slide(final_presentation.slide_layouts[6])
                for shape in slide.shapes:
                    if shape.shape_type == 13:  # 图片
                        image_stream = io.BytesIO(shape.image.blob)
                        new_slide.shapes.add_picture(
                            image_stream, shape.left, shape.top, shape.width, shape.height
                        )
                    else:
                        new_slide.shapes._spTree.insert_element_before(shape.element, 'p:extLst')

        final_presentation.save(final_ppt_path)

    await asyncio.to_thread(merge_ppts_sync)

    return {"final_ppt_path": final_ppt_path}


# async def compile_ppt(state: ReportState):
#     """
#     异步方式合并所有生成的pptx文件到一个pptx文件中，按照章节和幻灯片顺序合并。

#     Args:
#         state: 当前状态，包含所有完成的幻灯片信息。

#     Returns:
#         Dict: 更新后的状态，包含最终合并的PPT路径。
#     """
#     topic = state["topic"]
#     ppt_sections = state["ppt_sections"]

#     save_dir = os.path.join(".", "saves", topic)
#     final_ppt_path = os.path.join(save_dir, f"{topic}_final.pptx")

#     def merge_ppts_sync():
#         final_presentation = Presentation()
#         final_presentation.slide_width = Inches(20)
#         final_presentation.slide_height = Inches(11.25)

#         for section in ppt_sections:
#             for slide_index, _ in enumerate(section.slides, start=1):
#                 ppt_path = os.path.join(save_dir, f"{section.name}_slide_{slide_index}.pptx")
#                 if not os.path.exists(ppt_path):
#                     continue

#                 with open(ppt_path, "rb") as ppt_file:
#                     ppt_content = ppt_file.read()

#                 presentation = Presentation(io.BytesIO(ppt_content))

#                 for slide in presentation.slides:
#                     slide_layout = final_presentation.slide_layouts[5]
#                     new_slide = final_presentation.slides.add_slide(slide_layout)
#                     for shape in slide.shapes:
#                         new_slide.shapes._spTree.insert_element_before(shape.element, 'p:extLst')

#         final_presentation.save(final_ppt_path)

#     # 使用异步方式执行同步函数
#     await asyncio.to_thread(merge_ppts_sync)

#     return {"final_ppt_path": final_ppt_path}





# PPT Slide sub-graph -- 
ppt_slide_graph = StateGraph(PPTSlideState, output=PPTSlideOutputState)
# ppt_slide_graph.add_node("generate_slide_content", generate_slide_content)
ppt_slide_graph.add_node("enrich_slide_content", enrich_slide_content)
ppt_slide_graph.add_node("generate_slide_code_and_execute", generate_slide_code_and_execute)
ppt_slide_graph.add_node("ppt_slide_to_image_and_validate", ppt_slide_to_image_and_validate)

ppt_slide_graph.add_edge(START, "enrich_slide_content")
ppt_slide_graph.add_edge("enrich_slide_content", "generate_slide_code_and_execute")
ppt_slide_graph.add_edge("generate_slide_code_and_execute", "ppt_slide_to_image_and_validate")
# ppt_slide_graph.add_edge("generate_slide_content", END)

ppt_slide_subgraph = ppt_slide_graph.compile()

# PPT Sction sub-graph --
ppt_section_graph = StateGraph(PPTSectionState,output=PPTSectionOutputState)
ppt_section_graph.add_node("generate_ppt_section_start", generate_ppt_section_start)
ppt_section_graph.add_node("generate_slide", ppt_slide_subgraph)
ppt_section_graph.add_node("generate_ppt_section_end", generate_ppt_section_end)

ppt_section_graph.add_edge(START, "generate_ppt_section_start")
# ppt_section_graph.add_edge("generate_ppt_section_start", "generate_slide")
ppt_section_graph.add_edge("generate_slide", "generate_ppt_section_end")
# ppt_section_graph.add_edge("generate_ppt_section_end", END)

ppt_section_subgraph = ppt_section_graph.compile()

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
builder.add_node("generate_ppt_outline", generate_ppt_outline)
builder.add_node("generate_ppt_sections", ppt_section_subgraph)
builder.add_node("generate_cover_slide", generate_cover_slide)
builder.add_node("generate_section_cover_slides", generate_section_cover_slides)
builder.add_node("generate_end_slide", generate_end_slide)
builder.add_node("compile_ppt", compile_ppt)

# Add edges
builder.add_edge(START, "process_image_input")
builder.add_edge("process_image_input", "generate_report_plan")
builder.add_edge("generate_report_plan", "human_feedback")
builder.add_edge("build_section_with_web_research", "gather_completed_sections")
builder.add_conditional_edges("gather_completed_sections", initiate_final_section_writing, ["write_final_sections"])
builder.add_edge("write_final_sections", "compile_final_report")
# builder.add_edge("compile_final_report", END)
builder.add_edge("compile_final_report", "generate_ppt_outline")
# builder.add_edge("generate_ppt_outline", "generate_ppt_sections")
# builder.add_edge("generate_ppt_sections", "compile_ppt")
builder.add_edge("generate_ppt_sections", "generate_cover_slide")
builder.add_edge("generate_cover_slide", "generate_section_cover_slides")
builder.add_edge("generate_section_cover_slides", "generate_end_slide")
builder.add_edge("generate_end_slide", "compile_ppt")
builder.add_edge("compile_ppt", END)

graph = builder.compile()
