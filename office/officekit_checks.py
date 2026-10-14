"""
Office-Kit inspection, validation and post-processing for PPTX and DOCX.

What the installed Office-Kit (0.1.0) actually provides, and is used here:
- validate_document(): opens the file and reports failures
- document_info() / view.view_stats(): slide, paragraph, table and image counts
- view.view_outline(): DOCX heading tree
- view.view_issues(): empty documents, unfilled {{placeholders}} (DOCX)
- extract_text(): visible text, used to check expected titles are present
- batch.run_batch() with a "merge" op: atomic {{placeholder}} filling

Office-Kit has no document-merge, layout or text-overflow checks. Layout and
overflow are covered by Compono's validator and check_pptx_layout().
"""

import re
from pathlib import Path

import officekit
from officekit import view
from officekit.batch import run_batch
from officekit.core import PLACEHOLDER


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", str(text)).strip().lower()


def review_document(
    path: Path,
    expected_titles: list[str] | None = None,
    placeholder_values: dict | None = None,
) -> dict:
    """
    Inspect and validate a .pptx/.docx with Office-Kit.

    Only edits the file when it contains unfilled {{placeholders}} and
    ``placeholder_values`` can fill them (one atomic batch).
    """
    path = Path(path)
    report = {
        "file": str(path),
        "valid": False,
        "info": {},
        "stats": {},
        "outline": [],
        "issues": [],
        "missing_titles": [],
        "edits_applied": 0,
    }

    open_issues = officekit.validate_document(path)
    if open_issues:
        report["issues"] = [str(issue) for issue in open_issues]
        return report

    report["info"] = officekit.document_info(path)
    report["stats"] = view.view_stats(path)
    if path.suffix.lower() == ".docx":
        report["outline"] = view.view_outline(path).get("headings", [])

    text = officekit.extract_text(path)
    placeholders = sorted(set(PLACEHOLDER.findall(text)))
    if placeholders and placeholder_values:
        fillable = {key: value for key, value in placeholder_values.items() if key in placeholders}
        if fillable:
            result = run_batch(path, [{"op": "merge", "data": fillable}])
            report["edits_applied"] = result["applied"]
            text = officekit.extract_text(path)

    issues = [issue["message"] for issue in view.view_issues(path)]
    leftover = sorted(set(PLACEHOLDER.findall(text)))
    if leftover and path.suffix.lower() == ".pptx":
        # view_issues() only scans DOCX text for placeholders.
        issues.extend(f"Unfilled placeholder {{{{{name}}}}}" for name in leftover)
    report["issues"] = issues

    haystack = _normalize(text)
    report["missing_titles"] = [
        title for title in (expected_titles or [])
        if title and _normalize(title) not in haystack
    ]
    report["valid"] = not report["issues"] and not report["missing_titles"]
    return report


def print_review(report: dict) -> None:
    """Human-readable summary of review_document()."""
    status = "passed" if report["valid"] else "reported problems"
    print(f"Office-Kit validation {status}: {report['file']}")
    if report["stats"]:
        print(f"  stats: {report['stats']}")
    if report["outline"]:
        levels = {}
        for heading in report["outline"]:
            levels[heading["level"]] = levels.get(heading["level"], 0) + 1
        print(f"  headings per level: {dict(sorted(levels.items()))}")
    if report["edits_applied"]:
        print(f"  Office-Kit batch edits applied: {report['edits_applied']}")
    for issue in report["issues"]:
        print(f"  [issue] {issue}")
    for title in report["missing_titles"]:
        print(f"  [missing] expected title not found: {title}")
