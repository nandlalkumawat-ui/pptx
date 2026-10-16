import argparse
import json
import threading
import time
from pathlib import Path

from config import DOCX_MAX_FIGURES, DOCX_OUTPUT_DIR, PPTX_ENGINE
from research.web_search import search_web, format_search_results
from research.image_search import enrich_presentation_with_images, enrich_report_with_images
from generators.pptx_generator import (
    DEFAULT_SLIDE_COUNT,
    MAX_SLIDE_COUNT,
    MIN_SLIDE_COUNT,
    PPTXContentGenerator,
)
from generators.docx_generator import DOCXContentGenerator
from generators.officekit_pptx_generator import (
    PPTX_ENGINES,
    PptxValidationError,
    generate_pptx_with_officekit,
)
from generators.officekit_docx_generator import generate_docx_with_officekit
from office.officekit_checks import print_review, review_document
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

def generate_pptx_content(topic: str, research: str, generator=None,
                          slide_count: int = DEFAULT_SLIDE_COUNT, author: str = ""):
    """Generate structured presentation content using the LLM."""

    print("\nGenerating structured content...")

    generator = generator or PPTXContentGenerator()

    presentation = generator.generate(
        topic,
        research,
        slide_count=slide_count,
        author=author,
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

def render_with_officekit(presentation, topic: str, engine: str | None = None):
    """
    Render the generated presentation with the chosen engine and
    validate the final PPTX with Office-Kit (always).
    """

    safe_topic = "".join(
        c.lower() if c.isalnum() else "_"
        for c in topic
    ).strip("_")

    output_file = (
        PPTX_OUTPUT_DIR /
        f"{safe_topic}.pptx"
    )

    try:

        result = generate_pptx_with_officekit(
            presentation,
            output_file,
            engine,
        )

        print(f"\nFinal PPTX: {result}")

        return result

    except PptxValidationError as exc:

        print(f"\nERROR: PPTX failed at stage: {exc.stage}")
        print(f"Reason: {exc}")
        if output_file.is_file():
            print(f"The unvalidated PPTX was kept for debugging: {output_file}")

        return None

    except Exception as exc:

        print(f"\nERROR: PPTX failed at stage: {getattr(exc, 'stage', 'generation')}")
        print(f"Reason: {exc}")

        return None


# ---------------------------------------------------------
# PPTX engine menu
# ---------------------------------------------------------

PPTX_MENU = {"1": "compono", "2": "python-pptx", "3": "both"}


def choose_pptx_engine() -> str:
    """Ask for the PPTX engine until the selection is valid."""

    print("\n========================================")
    print("       PPTX GENERATION OPTIONS")
    print("========================================\n")
    print("1. Generate PPTX using Compono")
    print("2. Generate PPTX using python-pptx")
    print("3. Generate PPTX using Both\n")

    while True:
        selection = input("Select an option (1/2/3): ").strip()
        if selection in PPTX_MENU:
            return PPTX_MENU[selection]
        print(f"ERROR: Invalid option '{selection}'. Please enter 1, 2 or 3.")


def parse_slide_count(value: str) -> int:
    """'' -> default; otherwise an integer within the supported range."""

    value = str(value).strip()
    if not value:
        return DEFAULT_SLIDE_COUNT
    if not value.isdigit() or not MIN_SLIDE_COUNT <= int(value) <= MAX_SLIDE_COUNT:
        raise ValueError(
            f"Number of slides must be a whole number from {MIN_SLIDE_COUNT} to {MAX_SLIDE_COUNT}."
        )
    return int(value)


def ask_slide_count() -> int:
    """Ask for the slide count until it is valid (Enter = default)."""

    while True:
        try:
            return parse_slide_count(
                input(f"Enter number of slides (default {DEFAULT_SLIDE_COUNT}): ")
            )
        except ValueError as exc:
            print(f"ERROR: {exc}")


# ---------------------------------------------------------
# PPTX workflow
# ---------------------------------------------------------

def _wait_for_warm_up(thread: threading.Thread) -> None:
    with PERF.stage("Waiting for Ollama warm-up"):
        thread.join()


def generate_pptx(topic: str, engine: str | None = None, author: str = "",
                  slide_count: int = DEFAULT_SLIDE_COUNT):
    """Complete PPTX generation workflow."""
    PERF.reset()
    total_start = time.perf_counter()

    # Load the model and pre-process the static prompt prefix while the
    # (network-bound) web search runs, instead of one after the other.
    generator = PPTXContentGenerator()
    warm_up = threading.Thread(target=generator.warm_up, daemon=True)
    warm_up.start()

    # Step 1: Research
    print("\nResearching the topic...")
    with PERF.stage("Web search"):
        research = collect_research(topic)

    if not research.strip():
        print("\nWARNING: No research data was found.")

    # Only wait for the warm-up if the model is really called (not on a
    # cached answer).
    generator.llm.before_request = lambda: _wait_for_warm_up(warm_up)

    # Step 2: Generate structured presentation
    try:
        with PERF.stage("Ollama content generation"):
            presentation = generate_pptx_content(
                topic,
                research,
                generator,
                slide_count,
                author,
            )

    except Exception as exc:
        print("\nERROR: PPTX content generation failed.")
        print(f"Reason: {exc}")
        print(PERF.summary())
        return

    # Step 3: Add real slide images
    print("\nFinding relevant images...")
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
    print("\nNext step: render the slides and validate them.")
    with PERF.stage("PPTX render + validation (wall)"):
        result = render_with_officekit(
            presentation,
            topic,
            engine,
        )
    PERF.add_time("TOTAL PPTX generation", time.perf_counter() - total_start)
    print(PERF.summary())
    return result


# ---------------------------------------------------------
# DOCX workflow
# ---------------------------------------------------------

def generate_docx(topic: str):
    """Complete DOCX report workflow (separate from the PPTX layouts)."""
    PERF.reset()
    total_start = time.perf_counter()

    generator = DOCXContentGenerator()
    warm_up = threading.Thread(target=generator.warm_up, daemon=True)
    warm_up.start()

    with PERF.stage("Web search"):
        research = collect_research(topic)

    # Only wait for the warm-up if the model is really called (not on a
    # cached answer).
    generator.llm.before_request = lambda: _wait_for_warm_up(warm_up)

    try:
        print("\nGenerating DOCX content...")
        with PERF.stage("Ollama content generation"):
            report = generator.generate(topic, research)
    except Exception as exc:
        print("\nERROR: DOCX content generation failed.")
        print(f"Reason: {exc}")
        print(PERF.summary())
        return None

    with PERF.stage("Image search + download (wall)"):
        enrich_report_with_images(report, topic, max_figures=DOCX_MAX_FIGURES)

    safe_topic = "".join(c.lower() if c.isalnum() else "_" for c in topic).strip("_")
    json_file = DOCX_OUTPUT_DIR / f"{safe_topic}.json"
    json_file.write_text(json.dumps(report.model_dump(), indent=2, ensure_ascii=False),
                         encoding="utf-8")
    print(f"\nStructured JSON saved to:\n{json_file}")

    result = None
    with PERF.stage("DOCX render + validation (wall)"):
        try:
            result = generate_docx_with_officekit(report, DOCX_OUTPUT_DIR / f"{safe_topic}.docx")
        except Exception as exc:
            print("\nERROR: DOCX generation failed.")
            print(f"Reason: {exc}")

    PERF.add_time("TOTAL DOCX generation", time.perf_counter() - total_start)
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
    parser.add_argument("--engine", default="", choices=["", *PPTX_ENGINES],
                        help="PPTX engine; skips the PPTX options menu")
    parser.add_argument("--author", default=None,
                        help="Author shown on the PPTX title slide (optional)")
    parser.add_argument("--slides", default=None,
                        help=f"Number of PPTX slides ({MIN_SLIDE_COUNT}-{MAX_SLIDE_COUNT}, "
                             f"default {DEFAULT_SLIDE_COUNT})")
    parser.add_argument("--inspect", default="",
                        help="Inspect/validate an existing .pptx or .docx with Office-Kit and exit")
    args = parser.parse_args()

    if args.inspect:
        report = review_document(Path(args.inspect))
        print_review(report)
        print(json.dumps(report, indent=2, default=str))
        return

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
        print(f"1. PPTX (default {DEFAULT_SLIDE_COUNT} slides)")
        print("2. DOCX (6-7 pages)")

        choice = input(
            "Enter choice: "
        ).strip()

    if choice == "1":

        # Flags win; prompts only appear in an interactive run (no --choice).
        interactive = not args.choice
        if args.author is not None:
            author = args.author.strip()
        else:
            author = input("\nEnter author name (optional): ").strip() if interactive else ""

        if args.slides is not None:
            try:
                slide_count = parse_slide_count(args.slides)
            except ValueError as exc:
                print(f"ERROR: {exc}")
                raise SystemExit(2)
        else:
            slide_count = ask_slide_count() if interactive else DEFAULT_SLIDE_COUNT

        # Non-interactive --choice 1 without --engine keeps the old default
        # (config.PPTX_ENGINE with fallback) instead of prompting.
        engine = args.engine or (choose_pptx_engine() if interactive else None)
        label = PPTX_ENGINES[engine] if engine else f"default ({PPTX_ENGINES[PPTX_ENGINE]})"
        print(f"\nSelected option: {label}")
        print(f"Slides requested: {slide_count}" + (f" | Author: {author}" if author else ""))

        if generate_pptx(topic, engine, author, slide_count) is None:
            raise SystemExit(1)

    elif choice == "2":

        if generate_docx(topic) is None:
            raise SystemExit(1)

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