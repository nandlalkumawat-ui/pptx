
from typing import List, Optional

from pydantic import BaseModel, ConfigDict, Field


# ============================================================
# SOURCE SCHEMA
# ============================================================

class Source(BaseModel):
    model_config = ConfigDict(extra="ignore")

    title: str
    url: str
    snippet: str = ""


# ============================================================
# TABLE SCHEMA
# ============================================================

class TableData(BaseModel):
    model_config = ConfigDict(extra="ignore")

    headers: List[str]
    rows: List[List[str]] = Field(default_factory=list)


# ============================================================
# PPTX SLIDE SCHEMA
# ============================================================

class Slide(BaseModel):
    model_config = ConfigDict(extra="ignore")

    # Slide identification
    slide_number: int
    title: str
    subtitle: str = ""

    # Slide purpose hint used for layout and image decisions:
    # title | content | process | comparison | data | conclusion | references
    layout: str = "content"

    # Slide content
    paragraphs: List[str] = Field(default_factory=list)
    bullets: List[str] = Field(default_factory=list)
    table: Optional[TableData] = None
    notes: str = ""

    # --------------------------------------------------------
    # Background design
    # --------------------------------------------------------

    background_color: str = "#FFFFFF"

    # --------------------------------------------------------
    # Typography settings (font sizes are in points: pt)
    # These values should be applied by the PPTX renderer.
    # --------------------------------------------------------

    # Header / slide title
    title_font_size: int = 32
    title_bold: bool = True
    title_font_color: str = "#0F2A44"

    # Subheader / subtitle
    subtitle_font_size: int = 20
    subtitle_bold: bool = False
    subtitle_font_color: str = "#475569"

    # Paragraph text (renderer may step down to 14pt to avoid overflow)
    paragraph_font_size: int = 18
    paragraph_bold: bool = False
    paragraph_font_color: str = "#1F2937"

    # Bullet points
    bullet_font_size: int = 18
    bullet_bold: bool = False
    bullet_font_color: str = "#1F2937"

    # Table header
    table_header_font_size: int = 15
    table_header_bold: bool = True
    table_header_font_color: str = "#FFFFFF"
    table_header_background_color: str = "#0F2A44"

    # Table body
    table_body_font_size: int = 14
    table_body_bold: bool = False
    table_body_font_color: str = "#1F2937"

    # Footer and citations
    footer_font_size: int = 10
    footer_bold: bool = False
    footer_font_color: str = "#64748B"

    # Default font family (Calibri ships with every Office install;
    # Aptos is missing on Office 2019 and older)
    font_name: str = "Calibri"

    # --------------------------------------------------------
    # Image settings
    # --------------------------------------------------------

    image_query: str = ""
    image_path: str = ""
    image_caption: str = ""
    image_caption_font_size: int = 10
    image_url: str = ""
    image_source: str = ""

    # --------------------------------------------------------
    # Research sources used on this slide
    # --------------------------------------------------------

    source_urls: List[str] = Field(default_factory=list)


# ============================================================
# PPTX PRESENTATION SCHEMA
# ============================================================

class Presentation(BaseModel):
    model_config = ConfigDict(extra="ignore")

    title: str
    subtitle: str = ""
    slides: List[Slide] = Field(default_factory=list)
    sources: List[Source] = Field(default_factory=list)


# ============================================================
# DOCX SECTION SCHEMA
# ============================================================

class Section(BaseModel):
    model_config = ConfigDict(extra="ignore")

    title: str
    paragraphs: List[str] = Field(default_factory=list)
    bullets: List[str] = Field(default_factory=list)
    table: Optional[TableData] = None


# ============================================================
# DOCX REPORT SCHEMA
# ============================================================

class Report(BaseModel):
    model_config = ConfigDict(extra="ignore")

    title: str
    subtitle: str = ""
    executive_summary: str
    sections: List[Section] = Field(default_factory=list)
    conclusion: str
    sources: List[Source] = Field(default_factory=list)