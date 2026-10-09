
"""
Professional PowerPoint Generator.

Pipeline:
1. Convert Pydantic presentation data into a dictionary.
2. Render formatted slides using python-pptx.
3. Inspect the generated PPTX using OfficeKit.
4. Validate the PPTX using OfficeKit.
5. Return the generated file path.
"""

import time
from pathlib import Path
from typing import Any

from pptx import Presentation as PptPresentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from office.officekit_utils import (
    render_rich_pptx,
    inspect_document,
    validate_document,
)
from utils.perf import PERF


def _to_dict(value: Any) -> dict:
    """Convert a Pydantic model or dictionary into a dictionary."""

    if hasattr(value, "model_dump"):
        data = value.model_dump()
    elif hasattr(value, "dict"):
        data = value.dict()
    elif isinstance(value, dict):
        data = value
    else:
        raise TypeError(
            "Presentation must be a Pydantic model or dictionary."
        )

    if not isinstance(data, dict):
        raise TypeError("Converted presentation data must be a dictionary.")

    return data


def _count_picture_shapes(pptx_file: Path) -> int:
    """Return the number of embedded picture shapes in a PPTX file."""
    return check_pptx_layout(pptx_file)["pictures"]


# Content shapes that must never overlap each other.
_CONTENT_SHAPES = {
    "Slide Title", "Slide Subtitle", "Body Paragraph", "Bullet List",
    "Content Table", "Slide Image", "Image Caption",
}


def check_pptx_layout(pptx_file: Path) -> dict:
    """
    Single pass over the saved PPTX: count embedded pictures and report
    off-slide shapes, overlapping content shapes and tiny fonts.
    """
    report = {"slides": 0, "pictures": 0, "picture_slides": [], "issues": []}
    try:
        document = PptPresentation(str(pptx_file))
    except Exception as exc:
        report["issues"].append(f"cannot open PPTX: {exc}")
        return report

    width, height = document.slide_width, document.slide_height
    tolerance = 9525 * 2  # 2px in EMU
    report["slides"] = len(document.slides)

    for number, slide in enumerate(document.slides, start=1):
        boxes = []
        for shape in slide.shapes:
            if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                report["pictures"] += 1
                report["picture_slides"].append(number)

            if (shape.left < -tolerance or shape.top < -tolerance
                    or shape.left + shape.width > width + tolerance
                    or shape.top + shape.height > height + tolerance):
                report["issues"].append(f"slide {number}: '{shape.name}' extends off the slide")

            if shape.name in _CONTENT_SHAPES:
                boxes.append((shape.name, shape.left, shape.top,
                              shape.left + shape.width, shape.top + shape.height))

            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.font.size is not None and run.font.size.pt < 9:
                            report["issues"].append(
                                f"slide {number}: font {run.font.size.pt}pt in '{shape.name}'"
                            )

        for i, first in enumerate(boxes):
            for second in boxes[i + 1:]:
                overlap_x = min(first[3], second[3]) - max(first[1], second[1])
                overlap_y = min(first[4], second[4]) - max(first[2], second[2])
                if overlap_x > tolerance and overlap_y > tolerance:
                    report["issues"].append(
                        f"slide {number}: '{first[0]}' overlaps '{second[0]}'"
                    )

    return report


def generate_pptx_with_officekit(
    presentation: Any,
    output_file: Path,
) -> Path:
    """
    Generate a professional PPTX and run OfficeKit checks.

    Args:
        presentation: Presentation Pydantic model or dictionary.
        output_file: Destination path for the PPTX file.

    Returns:
        Path to the generated PPTX file.

    Raises:
        TypeError: If the presentation data has an unsupported type.
        ValueError: If required presentation content is missing.
        RuntimeError: If PPTX rendering fails.
    """

    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    data = _to_dict(presentation)

    title = data.get("title", "Presentation")
    slides = data.get("slides", [])

    if not slides:
        raise ValueError(
            "The presentation contains no slides. "
            "Generate slide content before rendering."
        )

    print("\n========================================")
    print("       PROFESSIONAL PPTX GENERATOR")
    print("========================================")
    print(f"Presentation: {title}")
    print(f"Number of slides: {len(slides)}")
    print(f"Output file: {output_file}")

    # --------------------------------------------------------
    # STEP 1: Render the presentation.
    # Uses python-pptx for formatting, text, tables and images.
    # --------------------------------------------------------

    print("\n[1/3] Rendering PowerPoint slides...")

    try:
        with PERF.stage("PPTX slide building + save"):
            result = render_rich_pptx(
                output_file=output_file,
                presentation_data=data,
            )

        if result is not None:
            output_file = Path(result)

        if not output_file.is_file():
            raise RuntimeError(
                "The renderer finished without creating the PPTX file."
            )

        print(f"PPTX created: {output_file}")

    except Exception as exc:
        raise RuntimeError(
            f"PowerPoint rendering failed: {exc}"
        ) from exc

    # --------------------------------------------------------
    # STEP 2: Inspect the generated presentation with OfficeKit.
    # --------------------------------------------------------

    print("\n[2/3] Inspecting PPTX with OfficeKit...")

    validation_started = time.perf_counter()

    try:
        inspection = inspect_document(output_file)

        print("OfficeKit inspection completed.")

        if inspection:
            print(f"Inspection returned document information: {inspection}")

    except Exception as exc:
        print(f"OfficeKit inspection could not be completed: {exc}")

    # --------------------------------------------------------
    # STEP 3: Validate the generated presentation with OfficeKit.
    # --------------------------------------------------------

    print("\n[3/3] Validating PPTX with OfficeKit...")

    try:
        validation = validate_document(output_file)

        if validation.get("valid"):
            print("OfficeKit validation passed.")

            if validation.get("stdout"):
                print(validation["stdout"])
        else:
            print("OfficeKit validation reported an issue.")

            if validation.get("stdout"):
                print(validation["stdout"])

            if validation.get("stderr"):
                print(validation["stderr"])

    except Exception as exc:
        print(f"OfficeKit validation could not be completed: {exc}")

    # --------------------------------------------------------
    # FINAL IMAGE VERIFICATION
    # --------------------------------------------------------

    expected_images = 0
    for raw_slide in slides:
        slide_dict = raw_slide if isinstance(raw_slide, dict) else raw_slide.model_dump()
        image_path = slide_dict.get("image_path")
        if image_path and Path(image_path).is_file():
            expected_images += 1

    layout_report = check_pptx_layout(output_file)
    embedded_images = layout_report["pictures"]
    validation_elapsed = time.perf_counter() - validation_started
    PERF.add_time("PPTX final validation", validation_elapsed)
    print(f"[timing] PPTX final validation: {validation_elapsed:.2f}s")

    PERF.incr("slides generated", layout_report["slides"])
    PERF.incr("images inserted into PPTX", embedded_images)

    print(f"[verification] Slides in PPTX: {layout_report['slides']}")
    print(
        f"[verification] Embedded images in PPTX: {embedded_images} / expected {expected_images}"
        f" (slides {layout_report['picture_slides']})"
    )
    if expected_images and embedded_images < expected_images:
        print("[warning] Some slides were expected to contain images, but fewer pictures were embedded than expected.")
    if layout_report["issues"]:
        print("[warning] Layout check found issues:")
        for issue in layout_report["issues"]:
            print(f"  - {issue}")
    else:
        print("[verification] Layout check: no off-slide shapes, overlaps or tiny fonts.")

    # --------------------------------------------------------
    # FINAL RESULT
    # --------------------------------------------------------

    print("\n========================================")
    print("       PPTX GENERATION COMPLETED")
    print("========================================")
    print(f"File: {output_file}")

    return output_file