import re

from pydantic import ValidationError

from generators.pptx_generator import _sources_from_research, _table_supported_by_research
from llm.ollama_client import OllamaClient
from research.web_search import compact_research
from schemas.document_schema import Report


_TEXT_LIST = {"type": "array", "items": {"type": "string"}}

TABLE_SCHEMA = {
    "anyOf": [
        {
            "type": "object",
            "properties": {
                "headers": _TEXT_LIST,
                "rows": {"type": "array", "items": _TEXT_LIST},
            },
            "required": ["headers", "rows"],
        },
        {"type": "null"},
    ]
}

# Structured output for Ollama: only the fields the model must write.
# Sources come from the search results, figures from image search.
REPORT_SCHEMA = {
    "type": "object",
    "properties": {
        "title": {"type": "string"},
        "subtitle": {"type": "string"},
        "executive_summary": {"type": "string", "minLength": 200},
        "sections": {
            "type": "array",
            "minItems": 4,
            "maxItems": 6,
            "items": {
                "type": "object",
                "properties": {
                    "title": {"type": "string"},
                    "paragraphs": {**_TEXT_LIST, "minItems": 1, "maxItems": 3},
                    "subsections": {
                        "type": "array",
                        "maxItems": 3,
                        "items": {
                            "type": "object",
                            "properties": {
                                "title": {"type": "string"},
                                "paragraphs": {**_TEXT_LIST, "minItems": 1, "maxItems": 2},
                                "bullets": {**_TEXT_LIST, "maxItems": 5},
                            },
                            "required": ["title", "paragraphs", "bullets"],
                        },
                    },
                    "bullets": {**_TEXT_LIST, "maxItems": 6},
                    "table": TABLE_SCHEMA,
                    "table_caption": {"type": "string"},
                    "image_query": {"type": "string", "minLength": 8},
                },
                "required": ["title", "paragraphs", "subsections", "bullets",
                             "table", "table_caption", "image_query"],
            },
        },
        "conclusion": {"type": "string", "minLength": 200},
    },
    "required": ["title", "subtitle", "executive_summary", "sections", "conclusion"],
}

# Static first so Ollama can reuse the KV cache (see warm_up()).
INSTRUCTIONS = """You are an expert technical report writer. Write a professional report as JSON, using the topic and research material given at the end.

Rules:
- Use the research material as the factual basis; do not invent facts, numbers or sources.
- executive_summary: one paragraph of 100-150 words.
- 4-6 sections, in a logical order (background, current state, applications or analysis, challenges, outlook).
- Each section: 1-2 paragraphs of 80-120 words, then 1-3 subsections. Each subsection has a specific title, one paragraph of 60-100 words and 0-5 short bullets.
- Section bullets are optional; use them for lists of key points. Start a bullet with a key term and a colon when it helps.
- Add a table only when the research contains real figures or a real comparison. Every row has exactly as many cells as headers, max 8 rows. table_caption describes the table. Otherwise "table": null and "table_caption": "".
- image_query: 3-7 words describing a concrete photo that illustrates the section's subject.
- conclusion: one paragraph of 100-150 words.
- Plain text only: no markdown, no numbering in titles.
"""


class DOCXContentGenerator:

    def __init__(self):
        self.llm = OllamaClient()

    def warm_up(self) -> None:
        """Load the model and pre-compute the static prompt prefix."""
        self.llm.warm_up(prefix=INSTRUCTIONS)

    def generate(self, topic: str, research: str) -> Report:

        sources = _sources_from_research(research)
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
                    schema=REPORT_SCHEMA,
                    options={"temperature": 0.3, "num_predict": 6144},
                    use_cache=attempt == 0,
                )
                report = Report.model_validate(self._normalize(data, topic, research))
                if report.sections:
                    break
                last_error = ValueError("model returned no sections")
            except (ValidationError, ValueError) as exc:
                last_error = exc
            print(f"DOCX content attempt {attempt + 1} failed: {last_error}. Retrying...")
        else:
            raise RuntimeError(f"DOCX content generation failed: {last_error}")

        if not report.sources:
            report.sources = sources
        return report

    @staticmethod
    def _normalize(data: dict, topic: str = "", research: str = "") -> dict:
        """Repair structural slips of small models."""

        if not isinstance(data, dict):
            raise ValueError("model output is not a JSON object")

        def clean_list(items) -> list[str]:
            if isinstance(items, str):
                items = [items]
            return [str(item).strip() for item in items or [] if str(item).strip()]

        data["title"] = str(data.get("title") or topic or "Report").strip()
        data["executive_summary"] = str(data.get("executive_summary") or "").strip()
        data["conclusion"] = str(data.get("conclusion") or "").strip()

        sections = []
        for section in data.get("sections") or []:
            if not isinstance(section, dict):
                continue
            section["title"] = re.sub(r"^\s*\d+(\.\d+)*[.)]?\s+", "", str(section.get("title") or "")).strip()
            section["paragraphs"] = clean_list(section.get("paragraphs"))
            section["bullets"] = clean_list(section.get("bullets"))
            section["subsections"] = [
                {
                    "title": str(sub.get("title") or "").strip(),
                    "paragraphs": clean_list(sub.get("paragraphs")),
                    "bullets": clean_list(sub.get("bullets")),
                }
                for sub in section.get("subsections") or []
                if isinstance(sub, dict) and (sub.get("paragraphs") or sub.get("bullets"))
            ]

            table = section.get("table")
            if isinstance(table, dict) and table.get("headers"):
                headers = [str(h) for h in table["headers"]]
                rows = []
                for row in table.get("rows") or []:
                    if isinstance(row, list):
                        row = [str(cell) for cell in row[: len(headers)]]
                        rows.append(row + [""] * (len(headers) - len(row)))
                table = {"headers": headers, "rows": rows} if rows else None
            else:
                table = None
            if table and research and not _table_supported_by_research(table, research):
                print(f"Dropping table in '{section['title']}': its numbers are not in the research.")
                table = None
            section["table"] = table
            if not table:
                section["table_caption"] = ""

            if section["title"] and (section["paragraphs"] or section["subsections"] or section["bullets"]):
                sections.append(section)

        data["sections"] = sections
        return data


# Backward-compatible alias for older imports.
DOCXGenerator = DOCXContentGenerator
