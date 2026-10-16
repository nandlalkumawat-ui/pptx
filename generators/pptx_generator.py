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
        # Numbers [n] of the research items that support this slide.
        "sources": {"type": "array", "items": {"type": "integer"}, "maxItems": 3},
    },
    "required": ["title", "subtitle", "layout", "paragraphs", "bullets", "table",
                 "image_query", "sources"],
}

DEFAULT_SLIDE_COUNT = 8
MIN_SLIDE_COUNT = 3
MAX_SLIDE_COUNT = 15


def presentation_schema(slide_count: int = DEFAULT_SLIDE_COUNT) -> dict:
    """Ollama output schema; the grammar enforces the slide count."""
    return {
        "type": "object",
        "properties": {
            "title": {"type": "string"},
            "subtitle": {"type": "string"},
            "slides": {"type": "array", "items": SLIDE_SCHEMA,
                       "minItems": slide_count, "maxItems": slide_count},
        },
        "required": ["title", "subtitle", "slides"],
    }


PRESENTATION_SCHEMA = presentation_schema()


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
INSTRUCTIONS = """You are an expert presentation writer. Write a presentation as JSON with exactly the number of slides requested at the end, using the topic and research material given there.

Structure for 8 slides:
1. Title (layout "title"): a strong title and a one-line subtitle. Its image_query asks for a wide photograph with no text or diagram.
2. Introduction ("content"): what the topic is, its background, and why it matters in the real world.
3. Key Concepts ("comparison"): the essential concepts, components or principles, one per bullet.
4. Current Situation ("data" if the research has real figures, otherwise "content"): current trends, facts and examples from the research, with the year when the research gives one.
5. Applications ("comparison"): practical uses in the industries that really fit this topic.
6. Benefits and Challenges ("data"): a table with headers ["Benefits", "Challenges"] and 3-4 rows; EVERY row names one benefit AND one challenge, no empty cells.
7. Future Scope ("process"): emerging trends and likely developments in order; mark forecasts as forecasts ("expected", "projected").
8. Conclusion (layout "conclusion"): 3-4 key takeaways and one recommendation as bullets.
For fewer slides, merge neighbouring sections; for more, split the largest sections. The title slide is always first and the conclusion always last.

Rules:
- Use the research material as the factual basis; do not invent facts, numbers, dates or names.
- Every slide except the title: a specific title, a short subtitle, ONE paragraph of 2-3 full sentences (40-60 words) and 3 to 5 bullets. Never fewer than 3 bullets.
- Each bullet is one complete, specific point of 8-18 words. Start it with a short label and a colon, then explain, in the style of "Label: a specific fact or reason about the topic"; never copy this pattern literally or write placeholder text.
- Do not repeat the paragraph in the bullets.
- layout values: "content" (explanation), "process" (steps or a timeline; bullets in order), "comparison" (cards for concepts, options or uses), "data" (slide built around a table), "conclusion".
- Tables: only with real figures from the research, or for the Benefits and Challenges slide. Every row has exactly as many cells as headers, max 6 rows. Otherwise "table": null.
- sources: the numbers [n] of the research items that support the facts on this slide (max 3); [] when none apply. Never cite a number that is not in the research list.
- image_query: 3-7 words describing a concrete photo or diagram that illustrates THIS slide's subject, specific to the topic. Different for every slide. Never empty.
"""

_MARKDOWN_TABLE = re.compile(r"^\s*\|.*\|\s*$", re.M)
_NUMBER = re.compile(r"\d+(?:[.,]\d+)*")
# Model output can echo schema field names as bullet text; strip those and
# drop the renderer-generated source/reference slide title.
_LITERAL_LABEL = re.compile(r"^\s*label\s*:\s*", re.IGNORECASE)

_SCHEMA_KEYS = {
    "title",
    "subtitle",
    "layout",
    "paragraphs",
    "bullets",
    "table",
    "image_query",
}
_REFERENCE_TITLE = re.compile(
    r"^(?:sources?\s*(?:&|and)\s*image credits|references?|source notes?)$",
    re.I,
)


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

    def generate(self, topic: str, research: str, slide_count: int = DEFAULT_SLIDE_COUNT,
                 author: str = "") -> Presentation:

        if not MIN_SLIDE_COUNT <= slide_count <= MAX_SLIDE_COUNT:
            raise ValueError(
                f"Slide count must be between {MIN_SLIDE_COUNT} and {MAX_SLIDE_COUNT}, got {slide_count}."
            )

        sources = _sources_from_research(research)
        # The model only needs titles + snippets; URLs are dropped from the
        # prompt (fewer tokens to process) and re-attached as sources below.
        prompt_research = (
            compact_research([source.model_dump() for source in sources])
            if sources else research
        )

        # The slide count stays out of INSTRUCTIONS so the cached prefix is
        # the same for every request.
        prompt = f"""{INSTRUCTIONS}
Topic: {topic}
Number of slides: {slide_count}

Research material:
{prompt_research}
"""

        last_error = None
        for attempt in range(2):
            try:
                data = self.llm.generate_json(
                    prompt,
                    schema=presentation_schema(slide_count),
                    # ~450 tokens per slide; 4096 was sized for 8-10 slides.
                    options={"temperature": 0.3, "num_predict": max(4096, 450 * slide_count)},
                    # A retry must not get the same (bad) cached answer.
                    use_cache=attempt == 0,
                )
                presentation = Presentation.model_validate(
                    self._normalize(data, topic, research, len(sources))
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
        presentation.author = author.strip()

        # Cited URLs go into the speaker notes; the number itself is shown
        # in the slide footer and matches the references slide.
        for slide in presentation.slides:
            cited = [f"[{n}] {presentation.sources[n - 1].url}" for n in slide.citations
                     if n <= len(presentation.sources)]
            if cited:
                slide.notes = "\n\n".join(
                    part for part in (slide.notes, "Cited sources:\n" + "\n".join(cited)) if part
                )

        return presentation

    @staticmethod
    def _normalize(data: dict, topic: str = "", research: str = "", source_count: int = 0) -> dict:
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
                # Drop columns the model left empty (e.g. "Challenges").
                keep = [i for i in range(len(headers)) if any(row[i].strip() for row in rows)]
                if len(keep) < 2:
                    rows = []
                headers = [headers[i] for i in keep]
                rows = [[row[i] for i in keep] for row in rows]
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
            # ...or copy the prompt's literal "Label:" example prefix.
            slide["bullets"] = [b for b in (_LITERAL_LABEL.sub("", b).strip() for b in slide["bullets"]) if b]

            # Only citation numbers that exist in the research list.
            cited = []
            for number in slide.get("sources") or []:
                if isinstance(number, int) and 1 <= number <= source_count and number not in cited:
                    cited.append(number)
            slide["citations"] = cited

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

        # The schema grammar forces bullets on every slide, including the
        # title slide; a title slide shows none, so drop them there.
        if slides and slides[0]["layout"] == "title" and not slides[0]["table"]:
            slides[0]["bullets"] = []
            slides[0]["paragraphs"] = []
            slides[0]["citations"] = []

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
