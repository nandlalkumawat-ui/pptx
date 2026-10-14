import re

from pydantic import ValidationError

from llm.ollama_client import OllamaClient
from research.web_search import compact_research
from schemas.document_schema import Presentation, Source


LAYOUTS = ["title", "content", "process", "comparison", "data", "conclusion"]

# JSON schema passed to Ollama's structured output. It only contains the
# fields the model must write; styling fields come from the renderer, and
# sources come from the search results, so the model does not spend
# (slow, CPU-bound) tokens repeating them.
SLIDE_SCHEMA = {
    "type": "object",
    "properties": {
        "title": {"type": "string"},
        "subtitle": {"type": "string"},
        "layout": {"type": "string", "enum": LAYOUTS},
        "paragraphs": {"type": "array", "items": {"type": "string"}, "maxItems": 2},
        # minItems/minLength are enforced by Ollama's grammar; without them
        # llama3.1:8b tends to return 2 bullets and empty image queries.
        "bullets": {"type": "array", "items": {"type": "string"}, "minItems": 3, "maxItems": 5},
        "table": {
            "anyOf": [
                {
                    "type": "object",
                    "properties": {
                        "headers": {"type": "array", "items": {"type": "string"}},
                        "rows": {
                            "type": "array",
                            "items": {"type": "array", "items": {"type": "string"}},
                        },
                    },
                    "required": ["headers", "rows"],
                },
                {"type": "null"},
            ]
        },
        "image_query": {"type": "string", "minLength": 8},
    },
    "required": ["title", "subtitle", "layout", "paragraphs", "bullets", "table", "image_query"],
}

PRESENTATION_SCHEMA = {
    "type": "object",
    "properties": {
        "title": {"type": "string"},
        "subtitle": {"type": "string"},
        "slides": {"type": "array", "items": SLIDE_SCHEMA, "minItems": 8, "maxItems": 10},
    },
    "required": ["title", "subtitle", "slides"],
}


def _sources_from_research(research: str) -> list[Source]:
    """Recover the sources from the text built by format_search_results()."""

    sources = []
    pattern = re.compile(
        r"Title:\s*\n(?P<title>.*?)\n\s*URL:\s*\n(?P<url>\S+)\s*\n\s*Snippet:\s*\n(?P<snippet>.*?)(?=\nSOURCE \d+|\Z)",
        re.S,
    )
    for match in pattern.finditer(research or ""):
        sources.append(
            Source(
                title=match.group("title").strip(),
                url=match.group("url").strip(),
                snippet=" ".join(match.group("snippet").split())[:300],
            )
        )
    return sources


# Static instructions come FIRST and never contain the topic, so Ollama can
# reuse their KV cache across runs (prompt processing is ~10-15 tok/s on
# CPU, so every cached token matters). See PPTXContentGenerator.warm_up().
INSTRUCTIONS = """You are an expert presentation writer. Write an 8-10 slide presentation as JSON, using the topic and research material given at the end.

Rules:
- Use the research material as the factual basis; do not invent facts or numbers.
- Slide 1: layout "title"; a strong presentation title and a one-line subtitle; no bullets or paragraphs.
- Slide 2: an introduction / overview of what the deck covers.
- Middle slides: one clear idea per slide, explained progressively.
- Last slide: layout "conclusion" with 3-4 key takeaways as bullets.
- Each content slide: a specific title, a short subtitle, ONE paragraph of 2-3 sentences (max 60 words) and 3-5 bullets.
- Bullets are short (max 18 words). Start a bullet with a key term followed by a colon when it helps, e.g. "Solid-state cells: higher energy density with lower fire risk".
- Do not repeat the paragraph in the bullets.
- layout values: "content" (explanation), "process" (steps or a timeline; bullets in order), "comparison" (pros/cons, options), "data" (slide built around a table), "conclusion".
- Use a "data" slide with a table only when the research contains real figures. Every table row has exactly as many cells as headers, max 6 rows. Otherwise "table": null.
- image_query: 3-7 words describing a concrete photo that illustrates THIS slide's subject (an object, place, technology or process), specific to the topic. Different for every slide. Never empty.
"""

_MARKDOWN_TABLE = re.compile(r"^\s*\|.*\|\s*$", re.M)
_SCHEMA_KEYS = set(SLIDE_SCHEMA["properties"]) | {"slides", "sources", "null"}
_REFERENCE_TITLE = re.compile(r"^\s*(sources?|references?|bibliography|further reading)\b", re.I)
_NUMBER = re.compile(r"\d+(?:[.,]\d+)*")


def _table_supported_by_research(table: dict, research: str) -> bool:
    """
    Keep a table only if most of its numbers occur in the research text;
    otherwise the model probably invented them.
    """
    numbers = [n for row in table["rows"] for cell in row for n in _NUMBER.findall(cell)]
    if not numbers:
        return True
    research_numbers = set(_NUMBER.findall(research or ""))
    found = sum(1 for n in numbers if n in research_numbers)
    return found / len(numbers) >= 0.5


class PPTXContentGenerator:

    def __init__(self):
        self.llm = OllamaClient()

    def warm_up(self) -> None:
        """
        Load the model and pre-compute the static instruction prefix.
        Safe to run in a background thread while web search is running.
        """
        self.llm.warm_up(prefix=INSTRUCTIONS)

    def generate(self, topic: str, research: str) -> Presentation:

        sources = _sources_from_research(research)
        # The model only needs titles + snippets; URLs are dropped from the
        # prompt (fewer tokens to process) and re-attached as sources below.
        prompt_research = (
            compact_research([source.model_dump() for source in sources])
            if sources else research
        )

        prompt = f"""{INSTRUCTIONS}
Topic: {topic}

Research material:
{prompt_research}
"""

        last_error = None
        for attempt in range(2):
            try:
                data = self.llm.generate_json(
                    prompt,
                    schema=PRESENTATION_SCHEMA,
                    options={"temperature": 0.3, "num_predict": 4096},
                    # A retry must not get the same (bad) cached answer.
                    use_cache=attempt == 0,
                )
                presentation = Presentation.model_validate(
                    self._normalize(data, topic, research)
                )
                if len(presentation.slides) >= 3:
                    break
                last_error = ValueError(
                    f"model returned only {len(presentation.slides)} slides"
                )
            except (ValidationError, ValueError) as exc:
                last_error = exc
            print(f"PPTX content attempt {attempt + 1} failed: {last_error}. Retrying...")
        else:
            raise RuntimeError(f"PPTX content generation failed: {last_error}")

        if not presentation.sources:
            presentation.sources = sources

        return presentation

    @staticmethod
    def _normalize(data: dict, topic: str = "", research: str = "") -> dict:
        """
        Repair structural slips of small models so validation does not fail
        and the deck has a sensible order.
        """

        if not isinstance(data, dict):
            raise ValueError("model output is not a JSON object")

        slides = [slide for slide in data.get("slides") or [] if isinstance(slide, dict)]
        data["title"] = str(data.get("title") or topic or "Presentation").strip()

        for slide in slides:
            slide["title"] = str(slide.get("title") or "").strip()
            layout = str(slide.get("layout") or "content").lower()
            slide["layout"] = layout if layout in LAYOUTS else "content"
            if "conclu" in slide["title"].lower() or "takeaway" in slide["title"].lower():
                slide["layout"] = "conclusion"

            for key in ("paragraphs", "bullets"):
                items = slide.get(key) or []
                if isinstance(items, str):
                    items = [items]
                # Drop markdown tables the model sometimes pastes into text.
                items = [_MARKDOWN_TABLE.sub("", str(item)).strip() for item in items]
                slide[key] = [item for item in items if item]

            table = slide.get("table")
            if isinstance(table, dict) and table.get("headers"):
                headers = [str(h) for h in table.get("headers") or []]
                rows = []
                for row in table.get("rows") or []:
                    if not isinstance(row, list):
                        continue
                    row = [str(cell) for cell in row[: len(headers)]]
                    row += [""] * (len(headers) - len(row))
                    rows.append(row)
                slide["table"] = {"headers": headers, "rows": rows} if rows else None
            else:
                slide["table"] = None

            if slide["table"] and research and not _table_supported_by_research(slide["table"], research):
                print(f"Dropping table on '{slide['title'] or 'untitled'}': its numbers are not in the research.")
                slide["table"] = None

            if slide["table"]:
                slide["layout"] = "data"
                slide["title"] = slide["title"] or "Key Figures"
            elif slide["layout"] == "data":
                slide["layout"] = "content"

            # Small models sometimes echo schema field names as bullets.
            slide["bullets"] = [b for b in slide["bullets"] if b.strip().lower() not in _SCHEMA_KEYS]

        # The renderer builds a real "Sources & Image Credits" slide from the
        # search results, so model-written source/reference slides (which
        # only repeat site names) are dropped.
        slides = [
            slide for slide in slides
            if not _REFERENCE_TITLE.match(slide["title"])
        ]

        # Remove slides that ended up empty (e.g. a dropped table).
        slides = [
            slide for slide in slides
            if slide["layout"] == "title" or slide["bullets"] or slide["paragraphs"] or slide["table"]
        ]

        # Slide 1 must be a real title slide for the whole deck. If the model
        # used it for content, keep that content and add a title slide.
        if not slides or slides[0]["layout"] != "title" or slides[0]["bullets"]:
            if slides and slides[0]["layout"] == "title":
                slides[0]["layout"] = "content"
            slides.insert(0, {
                "title": data["title"],
                "subtitle": str(data.get("subtitle") or ""),
                "layout": "title",
                "paragraphs": [],
                "bullets": [],
                "table": None,
                "image_query": topic or data["title"],
            })
        else:
            slides[0]["title"] = data["title"] if len(data["title"]) > 3 else slides[0]["title"]
            slides[0]["subtitle"] = slides[0].get("subtitle") or str(data.get("subtitle") or "")

        # Conclusion goes last.
        conclusions = [s for s in slides[1:] if s["layout"] == "conclusion"]
        if conclusions:
            final = conclusions[-1]
            for extra in conclusions[:-1]:
                extra["layout"] = "content"
            slides = [s for s in slides if s is not final] + [final]

        for number, slide in enumerate(slides, start=1):
            slide["slide_number"] = number
            slide["title"] = slide["title"] or f"Slide {number}"

        data["slides"] = slides
        return data


# Backward-compatible alias for older imports.
PPTXGenerator = PPTXContentGenerator
