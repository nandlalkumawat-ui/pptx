from pathlib import Path

from officekit import create_document


def create_pptx(title: str, output_path: Path):

    create_document(
        "pptx",
        str(output_path),
        title=title,
    )

    print(f"PPTX created: {output_path}")

    return output_path