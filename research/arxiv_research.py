from pathlib import Path
import arxiv
import pymupdf


PROJECT_DIR = Path(__file__).resolve().parent.parent
ARXIV_DIR = PROJECT_DIR / "output" / "arxiv"
PDF_DIR = ARXIV_DIR / "pdfs"
IMAGE_DIR = ARXIV_DIR / "images"

PDF_DIR.mkdir(parents=True, exist_ok=True)
IMAGE_DIR.mkdir(parents=True, exist_ok=True)


def search_arxiv(topic: str, max_results: int = 3):
    client = arxiv.Client()

    search = arxiv.Search(
        query=topic,
        max_results=max_results,
        sort_by=arxiv.SortCriterion.Relevance,
    )

    papers = []

    for result in client.results(search):
        papers.append({
            "title": result.title,
            "authors": [a.name for a in result.authors],
            "abstract": result.summary,
            "url": result.entry_id,
            "pdf_url": result.pdf_url,
        })

    return papers


def extract_paper(paper: dict) -> dict:
    safe_id = paper["url"].rstrip("/").split("/")[-1]
    pdf_path = PDF_DIR / f"{safe_id}.pdf"
    paper_image_dir = IMAGE_DIR / safe_id
    paper_image_dir.mkdir(parents=True, exist_ok=True)

    # Download the PDF only when it is not already cached.
    if not pdf_path.exists():
        import requests

        response = requests.get(
            paper["pdf_url"],
            timeout=60,
            headers={"User-Agent": "ResearchPresentationGenerator/1.0"},
        )
        response.raise_for_status()
        pdf_path.write_bytes(response.content)

    document = pymupdf.open(pdf_path)
    page_text = []
    extracted_images = []

    try:
        for page_index, page in enumerate(document):
            page_text.append(page.get_text("text"))

            # Extract embedded images with their original file formats.
            for image_index, image_info in enumerate(
                page.get_images(full=True), start=1
            ):
                xref = image_info[0]
                image = document.extract_image(xref)

                # Ignore very small assets such as icons and tiny logos.
                if image["width"] < 180 or image["height"] < 120:
                    continue

                image_path = (
                    paper_image_dir
                    / f"page_{page_index + 1}_image_{image_index}."
                      f"{image['ext']}"
                )

                if not image_path.exists():
                    image_path.write_bytes(image["image"])

                extracted_images.append({
                    "path": str(image_path),
                    "page": page_index + 1,
                    "width": image["width"],
                    "height": image["height"],
                })
    finally:
        document.close()

    return {
        **paper,
        "full_text": "\n".join(page_text),
        "images": extracted_images,
        "pdf_path": str(pdf_path),
    }


def collect_arxiv_research(topic: str, max_results: int = 3):
    papers = search_arxiv(topic, max_results=max_results)
    extracted = []

    for paper in papers:
        try:
            extracted.append(extract_paper(paper))
        except Exception as exc:
            print(f"Could not process {paper['title']}: {exc}")

    return extracted