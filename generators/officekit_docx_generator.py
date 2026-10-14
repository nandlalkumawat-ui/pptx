"""
DOCX Report Generator.

Pipeline:
1. Render the Report with python-docx (office/docx_renderer.py).
2. Inspect and validate it with Office-Kit: opens, stats, heading outline,
   every section heading present, unfilled placeholders.
3. Verify embedded figures and tables against the report content.
4. Return the generated file path.
"""

import time
from pathlib import Path
from typing import Any

from office.docx_renderer import render_docx_report
from office.officekit_checks import print_review, review_document
from utils.perf import PERF


def generate_docx_with_officekit(report: Any, output_file: Path) -> Path:
    output_file = Path(output_file)
    data = report.model_dump() if hasattr(report, "model_dump") else dict(report)
    sections = data.get("sections") or []
    if not sections:
        raise ValueError("The report contains no sections. Generate content before rendering.")

    print("\n========================================")
    print("          DOCX REPORT GENERATOR")
    print("========================================")
    print(f"Report: {data.get('title', 'Report')}")
    print(f"Sections: {len(sections)}")
    print(f"Output file: {output_file}")

    print("\n[1/3] Rendering DOCX with python-docx...")
    with PERF.stage("DOCX build + save (python-docx)"):
        render_docx_report(output_file, data)
    if not output_file.is_file():
        raise RuntimeError("The DOCX renderer finished without creating the file.")

    print("\n[2/3] Inspecting and validating DOCX with Office-Kit...")
    expected = [section["title"] for section in sections]
    expected += [sub["title"] for section in sections for sub in section.get("subsections") or []]
    started = time.perf_counter()
    with PERF.stage("Office-Kit inspect + validate"):
        review = review_document(output_file, expected_titles=expected,
                                 placeholder_values={"title": data.get("title", "")})
    print_review(review)
    if review["issues"] and not review["info"]:
        raise RuntimeError(f"Office-Kit cannot open the generated DOCX: {review['issues']}")

    print("\n[3/3] Checking figures and tables...")
    expected_images = sum(1 for s in sections if s.get("image_path") and Path(s["image_path"]).is_file())
    expected_tables = sum(1 for s in sections if s.get("table"))
    stats = review["stats"]
    print(f"[verification] Figures embedded: {stats.get('images')} / expected {expected_images}")
    print(f"[verification] Tables: {stats.get('tables')} / expected {expected_tables}")
    print(f"[verification] Words: {stats.get('words')}")
    if stats.get("images", 0) < expected_images or stats.get("tables", 0) < expected_tables:
        print("[warning] Fewer figures or tables in the DOCX than in the report content.")
    PERF.add_time("DOCX final validation", time.perf_counter() - started)
    PERF.incr("docx figures", stats.get("images", 0))
    PERF.incr("docx tables", stats.get("tables", 0))

    print("\n========================================")
    print("       DOCX GENERATION COMPLETED")
    print("========================================")
    print(f"File: {output_file}")
    return output_file
