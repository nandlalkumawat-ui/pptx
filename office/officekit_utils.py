
import copy
import json
import math
import re
import shutil
import subprocess
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


# ============================================================
# CONFIGURATION (design system)
# ============================================================

SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5

MARGIN_X = 0.6
CONTENT_TOP = 1.95
CONTENT_BOTTOM = 6.75
FOOTER_Y = 7.02

HERO_TITLE_SIZE = 40
TITLE_SIZE = 32
SUBTITLE_SIZE = 20
BODY_SIZE = 18
MIN_BODY_SIZE = 14
TABLE_HEADER_SIZE = 15
TABLE_BODY_SIZE = 14
FOOTER_SIZE = 10

FONT_NAME = "Calibri"

# Palette
NAVY = "#0F2A44"
NAVY_LIGHT = "#1B3D60"
ACCENT = "#0E9F9A"
LIGHT_BG = "#FFFFFF"
PANEL_BG = "#F1F5F9"
RULE_COLOR = "#E2E8F0"

DEFAULT_BACKGROUND = LIGHT_BG
DEFAULT_TITLE_COLOR = NAVY
DEFAULT_SUBTITLE_COLOR = "#475569"
DEFAULT_BODY_COLOR = "#1F2937"
DEFAULT_FOOTER_COLOR = "#64748B"
DEFAULT_TABLE_COLOR = NAVY

ON_DARK_TITLE = "#FFFFFF"
ON_DARK_BODY = "#E2E8F0"
ON_DARK_MUTED = "#94A3B8"


# ============================================================
# OFFICEKIT HELPERS
# ============================================================

def _officekit_cli() -> str | None:
    """Find the OfficeKit CLI, including inside an un-activated venv."""

    found = shutil.which("officekit")
    if found:
        return found
    candidate = Path(__file__).resolve().parents[1] / ".venv" / "bin" / "officekit"
    return str(candidate) if candidate.is_file() else None


def run_officekit(args: list[str]) -> str:
    """Execute an OfficeKit CLI command."""

    cli = _officekit_cli()
    if not cli:
        raise RuntimeError(
            "OfficeKit CLI was not found. "
            "Install/configure OfficeKit or use python-pptx."
        )

    command = [cli, *args]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=120,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"OfficeKit command failed:\n"
            f"Command: {' '.join(command)}\n"
            f"STDOUT:\n{result.stdout}\n"
            f"STDERR:\n{result.stderr}"
        )

    return result.stdout.strip()


def create_pptx(output_file: Path, title: str) -> Path:
    """Create an initial PPTX file."""

    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH)
    prs.slide_height = Inches(SLIDE_HEIGHT)

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title_box = slide.shapes.add_textbox(
        Inches(0.7),
        Inches(2.3),
        Inches(11.9),
        Inches(1.0),
    )
    title_box.name = "Slide Title"

    paragraph = title_box.text_frame.paragraphs[0]
    paragraph.text = title
    paragraph.font.name = FONT_NAME
    paragraph.font.size = Pt(TITLE_SIZE)
    paragraph.font.bold = True

    prs.save(str(output_file))
    return output_file


def add_slide(pptx_file: Path, title: str) -> None:
    """Add a simple slide using the OfficeKit CLI."""

    run_officekit([
        "add",
        str(pptx_file),
        "slide",
        "--value",
        title,
    ])


def set_value(
    pptx_file: Path,
    target: str,
    value: str,
) -> None:
    """Set a document property through OfficeKit."""

    run_officekit([
        "set",
        str(pptx_file),
        target,
        "--value",
        value,
    ])


def inspect_document(pptx_file: Path) -> dict:
    """Inspect a document through OfficeKit (in-process, no subprocess)."""

    try:
        import officekit
        return officekit.document_info(str(pptx_file))
    except ImportError:
        output = run_officekit([
            "info",
            str(pptx_file),
            "--json",
        ])
        return json.loads(output)


def validate_document(pptx_file: Path) -> dict:
    """Run OfficeKit validation (in-process, falls back to the CLI)."""

    try:
        import officekit
        issues = officekit.validate_document(str(pptx_file))
        return {
            "valid": not issues,
            "stdout": "OK" if not issues else "\n".join(map(str, issues)),
            "stderr": "",
        }
    except ImportError:
        pass

    cli = _officekit_cli()
    if not cli:
        return {
            "valid": False,
            "stdout": "",
            "stderr": "OfficeKit CLI is not installed.",
        }

    result = subprocess.run(
        [cli, "validate", str(pptx_file)],
        capture_output=True,
        text=True,
        timeout=120,
    )

    return {
        "valid": result.returncode == 0,
        "stdout": result.stdout.strip(),
        "stderr": result.stderr.strip(),
    }


# ============================================================
# DATA CONVERSION
# ============================================================

def to_dict(value: Any) -> Any:
    """Convert Pydantic models and nested objects into dictionaries."""

    if hasattr(value, "model_dump"):
        return value.model_dump()

    if hasattr(value, "dict"):
        return value.dict()

    if isinstance(value, dict):
        return {
            key: to_dict(item)
            for key, item in value.items()
        }

    if isinstance(value, list):
        return [to_dict(item) for item in value]

    return value


# ============================================================
# COLOR HELPERS
# ============================================================

def hex_to_rgb(color: str) -> RGBColor:
    """Convert '#RRGGBB' to a PowerPoint RGBColor."""

    if not isinstance(color, str):
        return RGBColor(255, 255, 255)

    color = color.strip().lstrip("#")

    if len(color) == 3:
        color = "".join(character * 2 for character in color)

    if len(color) != 6:
        return RGBColor(255, 255, 255)

    try:
        red = int(color[0:2], 16)
        green = int(color[2:4], 16)
        blue = int(color[4:6], 16)
        return RGBColor(red, green, blue)
    except ValueError:
        return RGBColor(255, 255, 255)


def is_dark(color: str) -> bool:
    """True when a background needs light text (relative luminance)."""

    rgb = hex_to_rgb(color)

    def channel(value: int) -> float:
        value /= 255
        return value / 12.92 if value <= 0.03928 else ((value + 0.055) / 1.055) ** 2.4

    luminance = 0.2126 * channel(rgb[0]) + 0.7152 * channel(rgb[1]) + 0.0722 * channel(rgb[2])
    return luminance < 0.35


# ============================================================
# TEXT MEASUREMENT
# ============================================================

def estimate_text_height(items: list[str], width: float, font_size: float,
                         space_after: float = 8, line_spacing: float = 1.12,
                         indent: float = 0.0) -> float:
    """
    Approximate rendered height (inches) of wrapped paragraphs.

    Uses an average glyph width of ~0.5 em, slightly conservative so that
    substitute fonts (LibreOffice / non-Windows) do not overflow either.
    """

    usable = max(width - indent - 0.2, 0.5)
    chars_per_line = max(int(usable * 72 / (font_size * 0.5)), 8)
    lines = 0
    for item in items:
        words = str(item).split()
        line_length = 0
        item_lines = 1
        for word in words:
            needed = len(word) + (1 if line_length else 0)
            if line_length + needed > chars_per_line:
                item_lines += 1
                line_length = len(word)
            else:
                line_length += needed
        lines += item_lines
    height = lines * font_size * line_spacing / 72
    height += len(items) * space_after / 72
    return height + 0.15


def fit_font_size(items: list[str], width: float, height: float,
                  max_size: int = BODY_SIZE, min_size: int = MIN_BODY_SIZE,
                  indent: float = 0.0) -> int | None:
    """Largest font size (pt) at which the items fit, or None."""

    for size in range(max_size, min_size - 1, -1):
        if estimate_text_height(items, width, size, indent=indent) <= height:
            return size
    return None


# ============================================================
# TEXT FORMATTING
# ============================================================

_LEAD_TERM = re.compile(r"^\s*([^:]{2,60}?):\s+(.+)$")


def _style_run(run, font_size: float, bold: bool, color: str) -> None:
    run.font.name = FONT_NAME
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = hex_to_rgb(color)


def _write_runs(paragraph, text: str, font_size: float, bold: bool,
                color: str, emphasize_lead: bool = False,
                lead_color: str | None = None) -> None:
    """Write text into a paragraph, bolding a 'Term:' lead when present."""

    match = _LEAD_TERM.match(str(text)) if emphasize_lead else None
    if match and len(match.group(1).split()) <= 6:
        lead = paragraph.add_run()
        lead.text = match.group(1) + ": "
        _style_run(lead, font_size, True, lead_color or color)
        rest = paragraph.add_run()
        rest.text = match.group(2)
        _style_run(rest, font_size, bold, color)
    else:
        run = paragraph.add_run()
        run.text = str(text)
        _style_run(run, font_size, bold, color)


def _set_bullet(paragraph, color: str, char: str = "•", indent: float = 0.3) -> None:
    """Real PowerPoint bullet with a hanging indent (wraps cleanly)."""

    p_pr = paragraph._p.get_or_add_pPr()
    p_pr.set("marL", str(int(Inches(indent))))
    p_pr.set("indent", str(-int(Inches(indent))))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buFont"):
        for element in p_pr.findall(qn(tag)):
            p_pr.remove(element)

    bu_clr = p_pr.makeelement(qn("a:buClr"), {})
    srgb = bu_clr.makeelement(qn("a:srgbClr"), {"val": color.lstrip("#").upper()})
    bu_clr.append(srgb)
    bu_font = p_pr.makeelement(qn("a:buFont"), {"typeface": "Arial"})
    bu_char = p_pr.makeelement(qn("a:buChar"), {"char": char})
    p_pr.append(bu_clr)
    p_pr.append(bu_font)
    p_pr.append(bu_char)


def format_paragraph(
    paragraph,
    text: str,
    font_size: int,
    bold: bool = False,
    color: str = DEFAULT_BODY_COLOR,
    alignment=PP_ALIGN.LEFT,
) -> None:
    """Apply consistent font formatting to a paragraph."""

    paragraph.text = ""
    paragraph.alignment = alignment
    paragraph.space_after = Pt(6)
    paragraph.line_spacing = 1.1
    _write_runs(paragraph, text, font_size, bold, color)


def _prepare_frame(shape, anchor=MSO_ANCHOR.TOP, margin: float = 0.04):
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.NONE
    frame.vertical_anchor = anchor
    frame.margin_left = Inches(margin)
    frame.margin_right = Inches(margin)
    frame.margin_top = Inches(0.03)
    frame.margin_bottom = Inches(0.03)
    return frame


def add_text_box(
    slide,
    text: str,
    x: float,
    y: float,
    width: float,
    height: float,
    font_size: int,
    bold: bool = False,
    color: str = DEFAULT_BODY_COLOR,
    name: str = "Text Box",
    alignment=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
):
    """Add a consistently formatted text box."""

    shape = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(width),
        Inches(height),
    )

    shape.name = name
    frame = _prepare_frame(shape, anchor)

    format_paragraph(
        frame.paragraphs[0],
        text,
        font_size,
        bold,
        color,
        alignment,
    )

    return shape


def add_title_box(slide, title: str, data: dict, x: float = MARGIN_X,
                  y: float = 0.5, width: float = SLIDE_WIDTH - 2 * MARGIN_X,
                  color: str | None = None):
    """Add the slide title using the schema's typography settings."""

    size = int(data.get("title_font_size", TITLE_SIZE))
    # Long titles step down instead of wrapping into the subtitle.
    while size > 26 and estimate_text_height([title], width, size, space_after=0) > 0.95:
        size -= 2

    return add_text_box(
        slide=slide,
        text=title,
        x=x,
        y=y,
        width=width,
        height=0.9,
        font_size=size,
        bold=bool(data.get("title_bold", True)),
        color=color or data.get("title_font_color", DEFAULT_TITLE_COLOR),
        name="Slide Title",
        anchor=MSO_ANCHOR.BOTTOM,
    )


def add_subtitle_box(slide, subtitle: str, data: dict, x: float = MARGIN_X,
                     y: float = 1.42, width: float = SLIDE_WIDTH - 2 * MARGIN_X,
                     color: str | None = None):
    """Add the slide subtitle."""

    if not subtitle:
        return None

    return add_text_box(
        slide=slide,
        text=subtitle,
        x=x,
        y=y,
        width=width,
        height=0.42,
        font_size=int(data.get("subtitle_font_size", SUBTITLE_SIZE)),
        bold=bool(data.get("subtitle_bold", False)),
        color=color or data.get(
            "subtitle_font_color",
            DEFAULT_SUBTITLE_COLOR,
        ),
        name="Slide Subtitle",
    )


def add_paragraph_box(slide, text: str, x: float, y: float, width: float,
                      height: float, data: dict, font_size: int | None = None,
                      color: str | None = None):
    """Add body text at the configured paragraph font size."""

    return add_text_box(
        slide=slide,
        text=text,
        x=x,
        y=y,
        width=width,
        height=height,
        font_size=font_size or int(data.get("paragraph_font_size", BODY_SIZE)),
        bold=bool(data.get("paragraph_bold", False)),
        color=color or data.get("paragraph_font_color", DEFAULT_BODY_COLOR),
        name="Body Paragraph",
    )


def add_bullets_box(
    slide,
    bullets: list,
    x: float,
    y: float,
    width: float,
    height: float,
    data: dict,
    font_size: int | None = None,
    color: str | None = None,
    bullet_color: str = ACCENT,
):
    """Add bullet points with consistent font formatting."""

    if not bullets:
        return None

    shape = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(width),
        Inches(height),
    )

    shape.name = "Bullet List"
    frame = _prepare_frame(shape)

    size = font_size or int(data.get("bullet_font_size", BODY_SIZE))
    text_color = color or data.get("bullet_font_color", DEFAULT_BODY_COLOR)

    for index, item in enumerate(bullets):
        paragraph = (
            frame.paragraphs[0]
            if index == 0
            else frame.add_paragraph()
        )

        paragraph.level = 0
        paragraph.space_after = Pt(10)
        paragraph.line_spacing = 1.1
        _set_bullet(paragraph, bullet_color)
        _write_runs(
            paragraph,
            item,
            size,
            bool(data.get("bullet_bold", False)),
            text_color,
            emphasize_lead=True,
        )

    return shape


def _add_rect(slide, x, y, width, height, color, name="Decoration",
              shape_type=MSO_SHAPE.RECTANGLE):
    shape = slide.shapes.add_shape(
        shape_type, Inches(x), Inches(y), Inches(width), Inches(height)
    )
    shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(color)
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


# ============================================================
# TABLE RENDERING
# ============================================================

def add_table(slide, table_data: Any, data: dict,
              x: float = 0.7, y: float = 2.0,
              width: float = 11.9, height: float = 3.0):
    """Add a formatted table from the TableData schema."""

    if not table_data:
        return None

    table_data = to_dict(table_data)

    headers = table_data.get("headers", [])
    rows = table_data.get("rows", [])

    if not headers:
        return None

    rows = [
        row if isinstance(row, list) else [str(row)]
        for row in rows
    ]

    # Normalize row widths to match the number of headers.
    column_count = len(headers)
    normalized_rows = []

    for row in rows:
        row = [str(cell) for cell in row[:column_count]]
        row.extend([""] * (column_count - len(row)))
        normalized_rows.append(row)

    # Shrink the table font for big tables, never below 11pt.
    body_size = int(data.get("table_body_font_size", TABLE_BODY_SIZE))
    header_size = int(data.get("table_header_font_size", TABLE_HEADER_SIZE))
    if len(normalized_rows) > 5 or column_count > 4:
        body_size = max(11, body_size - 2)
        header_size = max(12, header_size - 2)

    # Column widths proportional to content length.
    weights = []
    for column in range(column_count):
        longest = max(
            [len(str(headers[column]))] + [len(row[column]) for row in normalized_rows]
        )
        weights.append(min(max(longest, 6), 45))
    total_weight = sum(weights)

    row_height = max(0.42, height / (len(normalized_rows) + 1))
    table_shape = slide.shapes.add_table(
        rows=len(normalized_rows) + 1,
        cols=column_count,
        left=Inches(x),
        top=Inches(y),
        width=Inches(width),
        height=Inches(row_height * (len(normalized_rows) + 1)),
    )

    table = table_shape.table
    table_shape.name = "Content Table"

    for column in range(column_count):
        table.columns[column].width = Inches(width * weights[column] / total_weight)

    header_background = hex_to_rgb(
        data.get(
            "table_header_background_color",
            DEFAULT_TABLE_COLOR,
        )
    )
    header_color = data.get(
        "table_header_font_color",
        "#FFFFFF",
    )
    body_color = data.get(
        "table_body_font_color",
        DEFAULT_BODY_COLOR,
    )

    for column_index, header in enumerate(headers):
        cell = table.cell(0, column_index)
        cell.text = str(header)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_background
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = Inches(0.1)

        for paragraph in cell.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.LEFT

            for run in paragraph.runs:
                _style_run(run, header_size, bool(data.get("table_header_bold", True)), header_color)

    for row_index, row in enumerate(normalized_rows, start=1):
        for column_index, value in enumerate(row):
            cell = table.cell(row_index, column_index)
            cell.text = value
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.1)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_to_rgb(
                PANEL_BG if row_index % 2 == 0 else "#FFFFFF"
            )

            for paragraph in cell.text_frame.paragraphs:
                paragraph.space_after = Pt(2)

                for run in paragraph.runs:
                    _style_run(
                        run,
                        body_size,
                        bool(data.get("table_body_bold", False)) or column_index == 0,
                        body_color,
                    )

    return table_shape


# ============================================================
# IMAGE RENDERING
# ============================================================

def add_image(
    slide,
    image_path: str,
    x: float = 7.4,
    y: float = 2.0,
    width: float = 5.2,
    height: float = 3.8,
    fit: str = "contain",
    max_crop: float = 0.35,
):
    """
    Insert an existing local image into a slide without distortion.

    fit="contain": scale to fit inside the box (letterboxed).
    fit="cover":   fill the box and crop the overflow equally from both
                   sides; falls back to "contain" when more than
                   ``max_crop`` of the picture would be cut away.
    """

    if not image_path:
        return None

    path = Path(image_path)

    if not path.is_file():
        print(f"Image not found; skipping: {path}")
        return None

    try:
        from PIL import Image

        with Image.open(path) as image:
            image_width, image_height = image.size

        if image_width <= 0 or image_height <= 0:
            return None

        image_ratio = image_width / image_height
        box_ratio = width / height

        if fit == "cover":
            if image_ratio > box_ratio:
                visible = box_ratio / image_ratio
                crop_x, crop_y = (1 - visible) / 2, 0.0
            else:
                visible = image_ratio / box_ratio
                crop_x, crop_y = 0.0, (1 - visible) / 2

            if 1 - visible <= max_crop:
                picture = slide.shapes.add_picture(
                    str(path),
                    Inches(x),
                    Inches(y),
                    width=Inches(width),
                    height=Inches(height),
                )
                picture.crop_left = picture.crop_right = crop_x
                picture.crop_top = picture.crop_bottom = crop_y
                picture.name = "Slide Image"
                return picture

        scale = min(
            width / image_width,
            height / image_height,
        )

        fitted_width = image_width * scale
        fitted_height = image_height * scale

        left = x + (width - fitted_width) / 2
        top = y + (height - fitted_height) / 2

        picture = slide.shapes.add_picture(
            str(path),
            Inches(left),
            Inches(top),
            width=Inches(fitted_width),
            height=Inches(fitted_height),
        )
        picture.name = "Slide Image"
        return picture

    except Exception as exc:
        print(f"Could not insert image {path}: {exc}")
        return None


def add_image_caption(slide, caption: str, x: float, y: float,
                      width: float, data: dict, color: str | None = None):
    """Add a small caption below a slide image."""

    if not caption:
        return None

    return add_text_box(
        slide=slide,
        text=caption,
        x=x,
        y=y,
        width=width,
        height=0.3,
        font_size=int(data.get("image_caption_font_size", FOOTER_SIZE)),
        color=color or data.get("footer_font_color", DEFAULT_FOOTER_COLOR),
        name="Image Caption",
        alignment=PP_ALIGN.RIGHT,
    )


# ============================================================
# FOOTER
# ============================================================

def add_footer(slide, slide_number: int, data: dict, deck_title: str = "",
               dark: bool = False):
    """Add a slide number footer (and the deck title on the left)."""

    color = ON_DARK_MUTED if dark else data.get("footer_font_color", DEFAULT_FOOTER_COLOR)
    size = int(data.get("footer_font_size", FOOTER_SIZE))

    if deck_title:
        add_text_box(
            slide=slide,
            text=deck_title,
            x=MARGIN_X,
            y=FOOTER_Y,
            width=9.0,
            height=0.3,
            font_size=size,
            color=color,
            name="Slide Footer Title",
        )

    footer = add_text_box(
        slide=slide,
        text=str(slide_number),
        x=SLIDE_WIDTH - MARGIN_X - 0.8,
        y=FOOTER_Y,
        width=0.8,
        height=0.3,
        font_size=size,
        bold=bool(data.get("footer_bold", False)),
        color=color,
        name="Slide Footer",
        alignment=PP_ALIGN.RIGHT,
    )
    return footer


# ============================================================
# SLIDE BACKGROUND
# ============================================================

def set_slide_background(slide, color: str):
    """Apply the background color to a slide."""

    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = hex_to_rgb(color)


# ============================================================
# LAYOUTS
# ============================================================

def _has_image(slide_data: dict) -> bool:
    image_path = slide_data.get("image_path", "")
    return bool(image_path) and Path(image_path).is_file()


def _paragraph_text(slide_data: dict) -> str:
    return " ".join(str(p).strip() for p in slide_data.get("paragraphs", []) or [] if str(p).strip())


def _header(slide, slide_data: dict, width: float = SLIDE_WIDTH - 2 * MARGIN_X,
            dark: bool = False) -> None:
    """Accent bar + title + subtitle at fixed, consistent positions."""

    _add_rect(slide, MARGIN_X, 0.42, 0.7, 0.07, ACCENT, name="Decoration Accent")
    add_title_box(
        slide, slide_data.get("title", ""), slide_data, width=width,
        color=ON_DARK_TITLE if dark else None,
    )
    add_subtitle_box(
        slide, slide_data.get("subtitle", ""), slide_data, width=width,
        color=ON_DARK_MUTED if dark else None,
    )


def _text_column(slide, slide_data: dict, x: float, y: float, width: float,
                 bottom: float = CONTENT_BOTTOM, color: str | None = None,
                 bullet_color: str = ACCENT) -> None:
    """Paragraph (lead) followed by bullets, sized to fit the column."""

    paragraph = _paragraph_text(slide_data)
    bullets = slide_data.get("bullets", []) or []
    available = bottom - y

    items = ([paragraph] if paragraph else []) + list(bullets)
    size = fit_font_size(items, width, available, indent=0.3) or MIN_BODY_SIZE

    if paragraph:
        paragraph_height = estimate_text_height([paragraph], width, size, space_after=0)
        add_paragraph_box(slide, paragraph, x, y, width, paragraph_height, slide_data,
                          font_size=size, color=color)
        y += paragraph_height + 0.15

    if bullets:
        add_bullets_box(slide, bullets, x, y, width, max(bottom - y, 0.5), slide_data,
                        font_size=size, color=color, bullet_color=bullet_color)


def _layout_title(slide, slide_data: dict, deck: dict) -> None:
    set_slide_background(slide, NAVY)
    title = slide_data.get("title") or deck.get("title", "")
    subtitle = slide_data.get("subtitle") or deck.get("subtitle", "")

    if _has_image(slide_data):
        image_x = 7.35
        add_image(slide, slide_data["image_path"], image_x, 0, SLIDE_WIDTH - image_x,
                  SLIDE_HEIGHT, fit="cover", max_crop=0.6)
        text_width = image_x - MARGIN_X - 0.5
        caption = add_image_caption(slide, slide_data.get("image_caption", ""), MARGIN_X,
                                    FOOTER_Y, text_width, slide_data, color=ON_DARK_MUTED)
        if caption is not None:
            caption.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    else:
        text_width = SLIDE_WIDTH - 2 * MARGIN_X - 1.5

    size = HERO_TITLE_SIZE
    while size > 30 and estimate_text_height([title], text_width, size, space_after=0) > 2.2:
        size -= 2

    _add_rect(slide, MARGIN_X, 2.35, 0.9, 0.08, ACCENT, name="Decoration Accent")
    add_text_box(slide, title, MARGIN_X, 2.55, text_width, 2.2, size, bold=True,
                 color=ON_DARK_TITLE, name="Slide Title", anchor=MSO_ANCHOR.TOP)
    if subtitle:
        add_text_box(slide, subtitle, MARGIN_X, 4.85, text_width, 1.0, SUBTITLE_SIZE,
                     color=ON_DARK_BODY, name="Slide Subtitle")
    if deck.get("author"):
        add_text_box(slide, str(deck["author"]), MARGIN_X, 5.95, text_width, 0.45, SUBTITLE_SIZE - 2,
                     bold=True, color=ON_DARK_MUTED, name="Author")


def _layout_image_text(slide, slide_data: dict, deck: dict) -> None:
    """Readable text on the left, a relevant picture on the right."""

    image_x, image_width = 7.45, SLIDE_WIDTH - 7.45 - MARGIN_X
    image_height = 4.2
    _header(slide, slide_data)
    _text_column(slide, slide_data, MARGIN_X, CONTENT_TOP, image_x - MARGIN_X - 0.45)
    picture = add_image(slide, slide_data["image_path"], image_x, CONTENT_TOP,
                        image_width, image_height, fit="cover")
    if picture is not None:
        caption_y = (picture.top + picture.height) / 914400 + 0.06
        add_image_caption(slide, slide_data.get("image_caption", ""), image_x,
                          caption_y, image_width, slide_data)


def _layout_text(slide, slide_data: dict, deck: dict) -> None:
    """No image: paragraph in a tinted lead panel, bullets alongside."""

    _header(slide, slide_data)
    paragraph = _paragraph_text(slide_data)
    bullets = slide_data.get("bullets", []) or []

    if paragraph and bullets:
        panel_width = 4.6
        inner = panel_width - 0.6
        size = fit_font_size([paragraph], inner, CONTENT_BOTTOM - CONTENT_TOP - 0.5,
                             max_size=BODY_SIZE + 2) or MIN_BODY_SIZE
        text_height = estimate_text_height([paragraph], inner, size, space_after=0)
        # Panel hugs the paragraph instead of filling the slide height.
        panel_height = min(max(text_height + 0.6, 2.0), CONTENT_BOTTOM - CONTENT_TOP)
        _add_rect(slide, MARGIN_X, CONTENT_TOP, panel_width, panel_height,
                  PANEL_BG, name="Decoration Panel")
        _add_rect(slide, MARGIN_X, CONTENT_TOP, 0.07, panel_height,
                  ACCENT, name="Decoration Accent")
        add_paragraph_box(slide, paragraph, MARGIN_X + 0.35, CONTENT_TOP + 0.3, inner,
                          panel_height - 0.5, slide_data, font_size=size,
                          color=NAVY)
        bullets_x = MARGIN_X + panel_width + 0.5
        _text_column(slide, {"bullets": bullets, **{k: v for k, v in slide_data.items() if k != "paragraphs"}},
                     bullets_x, CONTENT_TOP + 0.1, SLIDE_WIDTH - MARGIN_X - bullets_x)
    else:
        _text_column(slide, slide_data, MARGIN_X, CONTENT_TOP, SLIDE_WIDTH - 2 * MARGIN_X - 1.0)


def _cards(slide, items: list[str], x: float, y: float, width: float, height: float,
           numbered: bool, fill: str = PANEL_BG, text_color: str = DEFAULT_BODY_COLOR,
           columns: int | None = None) -> None:
    """Row/grid of rounded cards; used for process steps and comparisons."""

    count = len(items)
    if not count:
        return
    columns = columns or count
    rows = math.ceil(count / columns)
    gap = 0.25
    # Single row (process flow): number above the text. Grid: number to
    # the left of the text, which leaves more height for readable text.
    badge_inline = numbered and rows > 1
    badge_space = 0.75 if numbered and not badge_inline else 0.0
    text_offset = 0.95 if badge_inline else 0.25
    card_width = (width - gap * (columns - 1)) / columns
    card_height = (height - gap * (rows - 1)) / rows
    inner_width = card_width - text_offset - 0.25

    # One font size for every card (consistency), then shrink the cards to
    # the tallest text so they do not become large empty boxes.
    max_inner = card_height - 0.45 - badge_space
    size = min(
        fit_font_size([item], inner_width, max_inner, max_size=BODY_SIZE, min_size=13) or 13
        for item in items
    )
    needed = max(estimate_text_height([item], inner_width, size, space_after=0) for item in items)
    card_height = min(card_height, max(needed + 0.45 + badge_space, 1.2))

    for index, item in enumerate(items):
        column, row = index % columns, index // columns
        card_x = x + column * (card_width + gap)
        card_y = y + row * (card_height + gap)
        card = _add_rect(slide, card_x, card_y, card_width, card_height, fill,
                         name="Card", shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
        card.adjustments[0] = 0.06

        top = card_y + 0.25
        if numbered:
            badge_y = card_y + (card_height - 0.55) / 2 if badge_inline else top
            badge = _add_rect(slide, card_x + 0.25, badge_y, 0.55, 0.55, ACCENT,
                              name="Decoration Badge", shape_type=MSO_SHAPE.OVAL)
            frame = _prepare_frame(badge, MSO_ANCHOR.MIDDLE, margin=0)
            format_paragraph(frame.paragraphs[0], str(index + 1), 18, True, "#FFFFFF",
                             PP_ALIGN.CENTER)
            top += badge_space

        inner_height = card_y + card_height - top - 0.2
        box = slide.shapes.add_textbox(Inches(card_x + text_offset), Inches(top),
                                       Inches(inner_width), Inches(inner_height))
        box.name = "Card Text"
        frame = _prepare_frame(box, MSO_ANCHOR.MIDDLE if badge_inline else MSO_ANCHOR.TOP)
        paragraph = frame.paragraphs[0]
        paragraph.line_spacing = 1.1
        _write_runs(paragraph, item, size, False, text_color, emphasize_lead=True,
                    lead_color=NAVY if text_color == DEFAULT_BODY_COLOR else text_color)

        if numbered and index < count - 1 and rows == 1:
            arrow = _add_rect(slide, card_x + card_width + 0.03, card_y + card_height / 2 - 0.12,
                              gap - 0.06, 0.24, ACCENT, name="Decoration Arrow",
                              shape_type=MSO_SHAPE.RIGHT_ARROW)
            arrow.shadow.inherit = False


def _layout_process(slide, slide_data: dict, deck: dict) -> None:
    _header(slide, slide_data)
    paragraph = _paragraph_text(slide_data)
    y = CONTENT_TOP
    width = SLIDE_WIDTH - 2 * MARGIN_X
    if paragraph:
        size = fit_font_size([paragraph], width, 1.2) or MIN_BODY_SIZE
        height = estimate_text_height([paragraph], width, size, space_after=0)
        add_paragraph_box(slide, paragraph, MARGIN_X, y, width, height, slide_data, font_size=size)
        y += height + 0.25
    bullets = slide_data.get("bullets", []) or []
    _cards(slide, bullets, MARGIN_X, y, width, CONTENT_BOTTOM - y, numbered=True)


def _layout_comparison(slide, slide_data: dict, deck: dict) -> None:
    _header(slide, slide_data)
    paragraph = _paragraph_text(slide_data)
    y = CONTENT_TOP
    width = SLIDE_WIDTH - 2 * MARGIN_X
    if paragraph:
        size = fit_font_size([paragraph], width, 1.2) or MIN_BODY_SIZE
        height = estimate_text_height([paragraph], width, size, space_after=0)
        add_paragraph_box(slide, paragraph, MARGIN_X, y, width, height, slide_data, font_size=size)
        y += height + 0.25
    bullets = slide_data.get("bullets", []) or []
    columns = 2 if len(bullets) == 4 else min(len(bullets), 3)
    _cards(slide, bullets, MARGIN_X, y, width, CONTENT_BOTTOM - y, numbered=False,
           columns=columns)


def _layout_data(slide, slide_data: dict, deck: dict) -> None:
    _header(slide, slide_data)
    width = SLIDE_WIDTH - 2 * MARGIN_X
    paragraph = _paragraph_text(slide_data)
    bullets = slide_data.get("bullets", []) or []
    table = to_dict(slide_data.get("table")) or {}
    row_count = len(table.get("rows", [])) + 1

    y = CONTENT_TOP
    if paragraph:
        size = fit_font_size([paragraph], width, 1.0, max_size=18, min_size=14) or 14
        height = estimate_text_height([paragraph], width, size, space_after=0)
        add_paragraph_box(slide, paragraph, MARGIN_X, y, width, height, slide_data, font_size=size)
        y += height + 0.2

    table_height = min(0.48 * row_count, CONTENT_BOTTOM - y)
    if bullets:
        table_width = 7.6
        add_table(slide, table, slide_data, MARGIN_X, y, table_width, table_height)
        bullets_x = MARGIN_X + table_width + 0.45
        _text_column(slide, {"bullets": bullets, **{k: v for k, v in slide_data.items() if k != "paragraphs"}},
                     bullets_x, y, SLIDE_WIDTH - MARGIN_X - bullets_x)
    else:
        add_table(slide, table, slide_data, MARGIN_X, y, width, table_height)


def _layout_conclusion(slide, slide_data: dict, deck: dict) -> None:
    set_slide_background(slide, NAVY)
    _header(slide, slide_data, dark=True)
    paragraph = _paragraph_text(slide_data)
    bullets = slide_data.get("bullets", []) or []
    y = CONTENT_TOP
    width = SLIDE_WIDTH - 2 * MARGIN_X
    if paragraph:
        size = fit_font_size([paragraph], width, 1.2) or MIN_BODY_SIZE
        height = estimate_text_height([paragraph], width, size, space_after=0)
        add_paragraph_box(slide, paragraph, MARGIN_X, y, width, height, slide_data,
                          font_size=size, color=ON_DARK_BODY)
        y += height + 0.3
    if bullets:
        columns = 2 if len(bullets) == 4 else min(len(bullets), 3)
        _cards(slide, bullets, MARGIN_X, y, width, CONTENT_BOTTOM - y, numbered=True,
               fill=NAVY_LIGHT, text_color="#FFFFFF", columns=columns)


def _layout_references(slide, slide_data: dict, deck: dict) -> None:
    _header(slide, slide_data)
    entries = slide_data.get("_references", [])
    shape = slide.shapes.add_textbox(Inches(MARGIN_X), Inches(CONTENT_TOP),
                                     Inches(SLIDE_WIDTH - 2 * MARGIN_X),
                                     Inches(CONTENT_BOTTOM - CONTENT_TOP))
    shape.name = "Bullet List"
    frame = _prepare_frame(shape)
    size = 13 if len(entries) <= 10 else 11
    for index, (label, url) in enumerate(entries):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.space_after = Pt(5)
        _set_bullet(paragraph, ACCENT, indent=0.25)
        run = paragraph.add_run()
        run.text = label
        _style_run(run, size, False, DEFAULT_BODY_COLOR)
        if url:
            link = paragraph.add_run()
            link.text = f"  —  {_short_url(url)}"
            _style_run(link, size, False, ACCENT)
            link.hyperlink.address = url


def _short_url(url: str, limit: int = 70) -> str:
    parsed = urlparse(url)
    text = (parsed.netloc.replace("www.", "") + parsed.path).rstrip("/")
    return text if len(text) <= limit else text[: limit - 1] + "…"


LAYOUT_RENDERERS = {
    "title": _layout_title,
    "process": _layout_process,
    "comparison": _layout_comparison,
    "data": _layout_data,
    "conclusion": _layout_conclusion,
    "references": _layout_references,
}


def _choose_layout(slide_data: dict, index: int) -> str:
    layout = str(slide_data.get("layout") or "").lower()
    bullets = slide_data.get("bullets", []) or []

    if index == 1 and layout in ("", "title", "content") and not bullets:
        return "title"
    if layout == "title" and index != 1:
        layout = "content"
    if slide_data.get("table"):
        return "data"
    if layout == "process" and not 3 <= len(bullets) <= 5:
        layout = "content"
    if layout == "comparison" and not 2 <= len(bullets) <= 6:
        layout = "content"
    if layout in LAYOUT_RENDERERS:
        return layout
    return "image_text" if _has_image(slide_data) else "text"


# ============================================================
# PAGINATION (split dense slides instead of shrinking text)
# ============================================================

PAGINATE_MIN_SIZE = 16  # split a slide rather than go below this


def _paginate(slides: list[dict]) -> list[dict]:
    result = []

    def fits(items: list[str], width: float, available: float) -> bool:
        return fit_font_size(items, width, available, min_size=PAGINATE_MIN_SIZE,
                             indent=0.3) is not None

    for slide_data in slides:
        layout = str(slide_data.get("layout") or "").lower()
        bullets = list(slide_data.get("bullets", []) or [])
        if layout in ("title", "process", "comparison", "conclusion", "references") or slide_data.get("table"):
            result.append(slide_data)
            continue

        # Narrowest text column the slide will get (image+text: 6.4in,
        # lead panel + bullets: 7.0in for the bullets alone).
        width = 6.4 if _has_image(slide_data) else 7.0
        paragraph = _paragraph_text(slide_data)
        # Minus the paragraph/bullet gap and box padding _text_column adds.
        available = CONTENT_BOTTOM - CONTENT_TOP - 0.35
        if fits(([paragraph] if paragraph else []) + bullets, width, available):
            result.append(slide_data)
            continue

        # Move bullets to a continuation slide until the first part fits.
        keep = list(bullets)
        overflow: list[str] = []
        while len(keep) > 1 and not fits(([paragraph] if paragraph else []) + keep,
                                         width, available):
            overflow.insert(0, keep.pop())

        # Balance the two slides instead of leaving a near-empty one.
        while overflow and len(keep) - len(overflow) > 1:
            overflow.insert(0, keep.pop())

        first = dict(slide_data, bullets=keep)
        result.append(first)
        if overflow:
            continuation = copy.deepcopy(slide_data)
            continuation.update(
                title=f"{slide_data.get('title', '')} (continued)",
                paragraphs=[],
                bullets=overflow,
                image_path="",
                image_caption="",
            )
            result.extend(_paginate([continuation]))
    return result


def _references_slide(presentation_data: dict, slides: list[dict]) -> dict | None:
    entries = []
    # Research sources keep their [n] number, which slide footers cite.
    for number, source in enumerate(presentation_data.get("sources", []) or [], start=1):
        source = to_dict(source)
        if source.get("url") or source.get("title"):
            entries.append((f"[{number}] {source.get('title') or source.get('url')}",
                            str(source.get("url") or "")))
    for slide_data in slides:
        if slide_data.get("image_path") and slide_data.get("image_source"):
            page = next(iter(slide_data.get("source_urls") or []), "") or slide_data.get("image_url", "")
            entries.append((f"Image (slide: {slide_data.get('title', '')}) — {slide_data['image_source']}", page))
    if not entries:
        return None
    return {
        "title": "Sources & Image Credits",
        "subtitle": "Research and images used in this presentation",
        "layout": "references",
        "_references": entries[:14],
    }


def citation_label(slide_data: dict) -> str:
    """Footer citation for a slide, e.g. 'Sources [1], [4]' (or '')."""
    numbers = slide_data.get("citations") or []
    return "Sources " + ", ".join(f"[{n}]" for n in numbers) if numbers else ""


def footer_text(deck_title: str, slide_data: dict) -> str:
    return "  ·  ".join(part for part in (deck_title, citation_label(slide_data)) if part)


# ============================================================
# RICH PPTX RENDERER
# ============================================================

def render_rich_pptx(
    output_file: Path,
    presentation_data: Any,
) -> Path:
    """
    Generate a professional PPTX from Presentation or dict data.

    Supports:
    - Purpose-specific layouts (title, image + text, text, process,
      comparison, data/table, conclusion, references)
    - Paragraphs and real bullets with bold lead terms
    - Tables with proportional columns
    - Local images, cropped to fill without distortion
    - Overflow-aware font sizing and continuation slides
    - Source URLs and image credits in speaker notes and a references slide
    """

    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    presentation_data = to_dict(presentation_data)

    if not isinstance(presentation_data, dict):
        raise TypeError("Presentation data must be a dictionary or Pydantic model.")

    slides_data = [to_dict(slide) for slide in presentation_data.get("slides", [])]

    if not slides_data:
        raise ValueError("Presentation contains no slides.")

    slides_data = _paginate(slides_data)
    references = _references_slide(presentation_data, slides_data)
    if references:
        slides_data.append(references)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH)
    prs.slide_height = Inches(SLIDE_HEIGHT)

    blank_layout = prs.slide_layouts[6]
    deck_title = str(presentation_data.get("title", ""))

    for index, slide_data in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(blank_layout)
        set_slide_background(slide, DEFAULT_BACKGROUND)

        layout = _choose_layout(slide_data, index)
        renderer = LAYOUT_RENDERERS.get(layout)
        if renderer is None:
            renderer = _layout_image_text if layout == "image_text" else _layout_text
        renderer(slide, slide_data, presentation_data)

        if layout != "title":
            add_footer(slide, index, slide_data, footer_text(deck_title, slide_data),
                       dark=layout == "conclusion")

        # Add sources to speaker notes when available.
        source_urls = slide_data.get("source_urls", []) or []
        notes = slide_data.get("notes", "")

        note_parts = []

        if notes:
            note_parts.append(str(notes))

        if slide_data.get("image_url"):
            note_parts.append(f"Image source: {slide_data['image_url']}")

        if source_urls:
            note_parts.append(
                "Sources:\n" + "\n".join(str(url) for url in source_urls)
            )

        if note_parts:
            try:
                slide.notes_slide.notes_text_frame.text = "\n\n".join(
                    note_parts
                )
            except Exception as exc:
                print(f"Could not add notes for slide {index}: {exc}")

    prs.save(str(output_file))

    print(f"PPTX saved successfully: {output_file}")
    return output_file
