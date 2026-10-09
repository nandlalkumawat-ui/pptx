import argparse
import json
import threading
import time
from pathlib import Path

from research.web_search import search_web, format_search_results
from research.image_search import enrich_presentation_with_images
from generators.pptx_generator import PPTXContentGenerator
from generators.officekit_pptx_generator import (
    generate_pptx_with_officekit,
)
from utils.perf import PERF
# ---------------------------------------------------------
# Paths
# ---------------------------------------------------------

PROJECT_DIR = Path(__file__).resolve().parent

OUTPUT_DIR = PROJECT_DIR / "output"
PPTX_OUTPUT_DIR = OUTPUT_DIR / "pptx"

PPTX_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------
# Research
# ---------------------------------------------------------

def collect_research(topic: str) -> str:
    """Search the web and convert results into research text."""

    print("\nSearching web...")

    results = search_web(topic)
    PERF.incr("web search results", len(results))

    print(f"Found {len(results)} sources.")

    return format_search_results(results)


# ---------------------------------------------------------
# PPTX content generation
# ---------------------------------------------------------

def generate_pptx_content(topic: str, research: str, generator=None):
    """Generate structured presentation content using the LLM."""

    print("\nGenerating PPTX content...")

    generator = generator or PPTXContentGenerator()

    presentation = generator.generate(
        topic,
        research
    )

    return presentation


# ---------------------------------------------------------
# Save structured presentation
# ---------------------------------------------------------

def save_presentation_json(presentation, topic: str) -> Path:
    """Save generated presentation structure as JSON."""

    safe_topic = "".join(
        c.lower() if c.isalnum() else "_"
        for c in topic
    ).strip("_")

    json_file = PPTX_OUTPUT_DIR / f"{safe_topic}.json"

    # Handle Pydantic model / normal dict / object
    if hasattr(presentation, "model_dump"):
        data = presentation.model_dump()

    elif hasattr(presentation, "dict"):
        data = presentation.dict()

    elif isinstance(presentation, dict):
        data = presentation

    else:
        data = presentation.__dict__

    with open(json_file, "w", encoding="utf-8") as file:
        json.dump(
            data,
            file,
            indent=2,
            ensure_ascii=False
        )

    return json_file


# ---------------------------------------------------------
# Display presentation
# ---------------------------------------------------------

def display_presentation(presentation):
    """Print a summary of the generated presentation."""

    print("\n===== PRESENTATION GENERATED =====")

    if hasattr(presentation, "model_dump"):
        data = presentation.model_dump()

    elif hasattr(presentation, "dict"):
        data = presentation.dict()

    elif isinstance(presentation, dict):
        data = presentation

    else:
        data = presentation.__dict__

    title = data.get(
        "title",
        "Untitled Presentation"
    )

    slides = data.get(
        "slides",
        []
    )

    print(f"Title: {title}")
    print(f"Slides: {len(slides)}")

    for index, slide in enumerate(slides, start=1):

        if isinstance(slide, dict):
            slide_title = slide.get(
                "title",
                f"Slide {index}"
            )
        else:
            slide_title = getattr(
                slide,
                "title",
                f"Slide {index}"
            )

        print(f"{index}. {slide_title}")


# ---------------------------------------------------------
# OfficeKit rendering
# ---------------------------------------------------------

def render_with_officekit(presentation, topic: str):
    """
    Convert the generated presentation structure
    into an actual PPTX file using OfficeKit.
    """

    from generators.officekit_pptx_generator import (
        generate_pptx_with_officekit
    )

    safe_topic = "".join(
        c.lower() if c.isalnum() else "_"
        for c in topic
    ).strip("_")

    output_file = (
        PPTX_OUTPUT_DIR /
        f"{safe_topic}.pptx"
    )

    print("\nCreating PPTX with OfficeKit...")

    try:

        result = generate_pptx_with_officekit(
            presentation,
            output_file
        )

        print("\n===== PPTX CREATED =====")
        print(f"File: {result}")

        return result

    except Exception as exc:

        print("\nERROR: OfficeKit PPTX generation failed.")
        print(f"Reason: {exc}")

        return None


# ---------------------------------------------------------
# PPTX workflow
# ---------------------------------------------------------

def generate_pptx(topic: str):
    """Complete PPTX generation workflow."""
    PERF.reset()
    total_start = time.perf_counter()

    # Load the model and pre-process the static prompt prefix while the
    # (network-bound) web search runs, instead of one after the other.
    generator = PPTXContentGenerator()
    warm_up = threading.Thread(target=generator.warm_up, daemon=True)
    warm_up.start()

    # Step 1: Research
    with PERF.stage("Web search"):
        research = collect_research(topic)

    if not research.strip():
        print("\nWARNING: No research data was found.")

    with PERF.stage("Waiting for Ollama warm-up"):
        warm_up.join()

    # Step 2: Generate structured presentation
    try:
        with PERF.stage("Ollama content generation"):
            presentation = generate_pptx_content(
                topic,
                research,
                generator,
            )

    except Exception as exc:
        print("\nERROR: PPTX content generation failed.")
        print(f"Reason: {exc}")
        print(PERF.summary())
        return

    # Step 3: Add real slide images
    with PERF.stage("Image search + download (wall)"):
        presentation = enrich_presentation_with_images(presentation, topic)

    # Step 4: Display generated content
    display_presentation(presentation)

    # Step 5: Save JSON
    try:
        json_file = save_presentation_json(
            presentation,
            topic
        )

        print(f"\nStructured JSON saved to:")
        print(json_file)

    except Exception as exc:
        print(f"\nWARNING: Could not save JSON: {exc}")

    print("\nPPTX content generation successful.")

    # Step 6: Convert structured content to PPTX
    print("\nNext step: convert structured content into OfficeKit PPTX.")
    with PERF.stage("PPTX render + validation (wall)"):
        result = render_with_officekit(
            presentation,
            topic
        )
    PERF.add_time("TOTAL PPTX generation", time.perf_counter() - total_start)
    print(PERF.summary())
    return result


# ---------------------------------------------------------
# Main
# ---------------------------------------------------------

def main():

    # Optional non-interactive use:  python main.py --topic "..." --choice 1
    parser = argparse.ArgumentParser(description="AI document generator")
    parser.add_argument("--topic", default="")
    parser.add_argument("--choice", default="")
    args = parser.parse_args()

    print("======================================")
    print("     AI DOCUMENT GENERATOR")
    print("======================================")

    topic = args.topic.strip() or input(
        "\nEnter topic: "
    ).strip()

    if not topic:
        print("ERROR: Topic cannot be empty.")
        return

    if args.choice:
        choice = args.choice.strip()
    else:
        print("\nChoose output:")
        print("1. PPTX (8-10 slides)")
        print("2. DOCX (6-7 pages)")

        choice = input(
            "Enter choice: "
        ).strip()

    if choice == "1":

        generate_pptx(topic)

    elif choice == "2":

        print(
            "\nDOCX generation is not connected yet."
        )

        print(
            "We will connect the DOCX generator "
            "after the PPTX pipeline is working."
        )

    else:

        print(
            "\nERROR: Invalid choice."
        )

        print(
            "Please choose 1 or 2."
        )


# ---------------------------------------------------------
# Entry point
# ---------------------------------------------------------

if __name__ == "__main__":
    main()