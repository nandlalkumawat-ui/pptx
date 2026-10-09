from llm.ollama_client import OllamaClient
from schemas.document_schema import Report


class DOCXContentGenerator:

    def __init__(self):
        self.llm = OllamaClient()

    def generate(self, topic: str, research: str) -> Report:

        prompt = f"""
You are an expert technical report writer.

Create a professional 6-7 page report.

Topic:
{topic}

Research:
{research}

Requirements:

1. Create content suitable for a 6-7 page document.
2. Start with an executive summary.
3. Use clear headings and subheadings.
4. Use paragraphs where explanation is required.
5. Use bullet lists for important points.
6. Include tables where useful.
7. Include a strong conclusion.
8. Include sources.
9. Do not invent facts.
10. Base factual claims on the research material.

Return ONLY valid JSON.

Structure:

{{
    "title": "...",
    "subtitle": "...",
    "executive_summary": "...",
    "sections": [
        {{
            "title": "...",
            "paragraphs": [],
            "bullets": [],
            "table": null
        }}
    ],
    "conclusion": "...",
    "sources": [
        {{
            "title": "...",
            "url": "...",
            "snippet": "..."
        }}
    ]
}}
"""

        data = self.llm.generate_json(prompt)

        return Report.model_validate(data)


# Backward-compatible alias for older imports.
DOCXGenerator = DOCXContentGenerator