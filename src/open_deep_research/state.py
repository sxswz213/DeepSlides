from typing import Annotated, List, TypedDict, Literal, Optional, Dict, Any
from pydantic import BaseModel, Field
import operator

class Section(BaseModel):
    name: str = Field(
        description="Name for this section of the report.",
    )
    description: str = Field(
        description="Brief overview of the main topics and concepts to be covered in this section.",
    )
    research: bool = Field(
        description="Whether to perform web research for this section of the report."
    )
    content: str = Field(
        description="The content of the section."
    )   
    source_str: str = Field(
        description="Formatted string of sources used in this section, including URLs and titles."
    )

class Sections(BaseModel):
    sections: List[Section] = Field(
        description="Sections of the report.",
    )

class SearchQuery(BaseModel):
    search_query: str = Field(None, description="Query for web search.")

class Queries(BaseModel):
    queries: List[SearchQuery] = Field(
        description="List of search queries.",
    )

class Feedback(BaseModel):
    grade: Literal["pass","fail"] = Field(
        description="Evaluation result indicating whether the response meets requirements ('pass') or needs revision ('fail')."
    )
    follow_up_queries: List[SearchQuery] = Field(
        description="List of follow-up search queries.",
    )


class PPTSlide(BaseModel):
    title: str = Field(
        description="Title of the PowerPoint slide."
    )
    points: List[str] = Field(
        description="Key points or bullet points for the slide."
    )
    codes: List[str] = Field()
    detail: str = Field(
        description="Detailed slide description including layout, content positions, and design style."
    )
    enriched_points: str = Field(
        description="Enriched content for the slide, including additional details or explanations." 
    )
    layout: str = Field(
        description="Layout type for the slide."
    )


class PPTSection(BaseModel):
    name: str = Field(
        description="Name of this PPT section, corresponding to a report section."
    )
    allocated_slides: int = Field(
        description="Number of slides allocated to this PPT section."
    )
    slides: List[PPTSlide] = Field(
        description="Slides within this PPT section."
    )

class PPTSections(BaseModel):
    sections: List[PPTSection] = Field(
        description="Sections of the PowerPoint presentation."
    )

class PPTOutline(BaseModel):
    ppt_sections: PPTSections = Field(
        description="Structured outline of the entire PPT presentation, segmented by sections."
    )

class ReportStateInput(TypedDict):
    topic: str # Report topic
    image_path: Optional[str] # Optional path to input image
    presentation_minutes: Optional[str] # Optional recommended PowerPoint slides based on the report
    style: Optional[str]  # Style of the report, if applicable
    template: Optional[str]  # Template for the report, if applicable

class ReportStateOutput(TypedDict):
    final_report: str # Final report

class ReportState(TypedDict):
    topic: str # Report topic    
    image_path: Optional[str] # Optional path to input image
    caption: Optional[str] # Optional caption generated from input image
    user_intent: Optional[str] # Optional user intent generated from input image
    feedback_on_report_plan: str # Feedback on the report plan
    sections: list[Section] # List of report sections 
    completed_sections: Annotated[list, operator.add] # Send() API key
    report_sections_from_research: str # String of any completed sections from research to write final sections
    final_report: str # Final report
    ppt_outline: PPTOutline  # Outline for PowerPoint presentation
    presentation_minutes: str
    recommended_ppt_slides: int  # Recommended PowerPoint slides based on the report
    section_distribution: Dict[str, int]  # Distribution of sections in the report
    ppt_sections: List[PPTSection]  # Detailed PPT sections generated based on the outline
    completed_ppt_sections: Annotated[List[PPTSection], operator.add]  # Completed PPT sections for Send() API
    ppt_generation_codes: Annotated[List[str], operator.add]
    final_ppt_path: Optional[str]  # Path to the final generated PowerPoint presentation
    style: Optional[str]  # Style of the report, if applicable
    main_color: Optional[str]  # Main color for the PPT section, if applicable
    accent_color: Optional[str]  # Accent color for the PPT section, if applicable
    storyline: Optional[str]  # Storyline for the report, if applicable
    cover_slide_path: Optional[str]  # Path to the cover slide image, if applicable
    end_slide_path: Optional[str]  # Path to the end slide image, if applicable
    section_slides_path: Optional[str]  # Path to the section slides directory, if applicable
    cover_layout_description: Optional[str]  # Layout description for the cover slide
    layout_description: Optional[str]  # Layout description for the main slides
    end_layout_description: Optional[str]  # Layout description for the end slide
    template: Optional[str]  # Template for the report, if applicable
    background_tone: Optional[str]  # Background tone for the PPT section, if applicable
    heading_font_color: Optional[str]  # Heading font color for the PPT section, if applicable
    body_font_color: Optional[str]  # Body font color for the PPT section, if applicable
    style_summary: Optional[str]  # Summary of the style used in the presentation
    font_name: Optional[str]


class SectionState(TypedDict):
    topic: str # Report topic
    section: Section # Report section  
    search_iterations: int # Number of search iterations done
    search_queries: list[SearchQuery] # List of search queries
    source_str: str # String of formatted source content from web search
    report_sections_from_research: str # String of any completed sections from research to write final sections
    completed_sections: Annotated[List[Section], operator.add] # Final key we duplicate in outer state for Send() API

class SectionOutputState(TypedDict):
    completed_sections: list[Section] # Final key we duplicate in outer state for Send() API

class PPTSlideState(TypedDict):
    topic: str
    # section: Section
    ppt_section: PPTSection
    slide_index: int
    generated_slides: Annotated[List[PPTSlide], operator.add]
    path: Optional[str]
    enriched_points: Optional[str]
    slide_detail: Optional[str]
    codes: Optional[List[str]]
    title: Optional[str]
    points: Optional[List[str]]
    layout_valid: Optional[bool]  # Whether the slide layout is valid
    max_retry_count: Optional[int]  # Maximum retry count for layout validation (defaults to 3)
    retry_count: Optional[int]  # The current retry count
    images_json: Optional[str]  # JSON string of images available for the slide
    style: Optional[str]  # Style of the PPT section, if applicable
    main_color: Optional[str]  # Main color for the PPT section, if applicable
    accent_color: Optional[str]  # Accent color for the PPT section, if applicable
    background_tone: Optional[str]  # Background tone for the PPT section, if applicable
    heading_font_color: Optional[str]  # Heading font color for the PPT section, if applicable
    body_font_color: Optional[str]  # Body font color for the PPT section, if applicable
    style_summary: Optional[str]  # Summary of the style used in the presentation
    design_score: Optional[float]  # Design score for the slide
    aesthetics_score: Optional[float]  # Aesthetics score for the slide
    completeness_score: Optional[float]  # Completeness score for the slide
    design_suggestions: Optional[str]  # Design improvement suggestions for the slide
    aestheitcs_suggestions: Optional[str]  # Aesthetics improvement suggestions for the slide
    completeness_suggestions: Optional[str]  # Completeness improvement suggestions for the slide
    image_data: Optional[str]  # JSON string of image data embedded for the slide
    font_name: Optional[str]

class PPTSlideOutputState(TypedDict):
    completed_slides: List[PPTSlide]  # 已完成的幻灯片列表


class PPTSectionState(TypedDict):
    topic: str  # Report topic
    # section: Section
    ppt_section: PPTSection  # PPT section being generated
    completed_slides: Annotated[List[PPTSlide], operator.add]
    style: Optional[str]  # Style of the PPT section, if applicable
    # generation_iterations: int  # Number of iterations in generating PPT content
    # completed_ppt_sections: Annotated[List[PPTSection], operator.add]  # Accumulated completed PPT sections
    main_color: Optional[str]  # Main color for the PPT section, if applicable
    accent_color: Optional[str]  # Accent color for the PPT section, if applicable
    background_tone: Optional[str]  # Background tone for the PPT section, if applicable
    heading_font_color: Optional[str]  # Heading font color for the PPT section, if applicable
    body_font_color: Optional[str]  # Body font color for the PPT section, if applicable
    style_summary: Optional[str]  # Summary of the style used in the presentation
    font_name: Optional[str]

class PPTSectionOutputState(TypedDict):
    completed_ppt_sections: List[PPTSection]  # 已完成的PPT章节列表


class PPTStateInput(TypedDict):
    topic: str
    presentation_minutes: Optional[str]
    ppt_outline: Optional[str]
    section_distribution: Optional[dict]

class PPTStateOutput(TypedDict):
    final_ppt_outline: str  # Final detailed PowerPoint outline

class PPTState(TypedDict):
    topic: str
    presentation_minutes: str
    ppt_outline: str
    recommended_ppt_slides: int
    section_distribution: dict[str, int]
    ppt_sections: List[PPTSection]
    final_ppt_outline: str
    style: Optional[str]  # Style of the PPT section, if applicable
