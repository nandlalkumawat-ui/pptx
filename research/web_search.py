from ddgs import DDGS

from config import MAX_SEARCH_RESULTS


def search_web(
    query: str,
    max_results: int = MAX_SEARCH_RESULTS,
):
    """
    Search the web using DDGS.

    No paid API key is required.
    """

    print(f"Searching for: {query}")

    try:
        results = DDGS().text(
            query,
            region="in-en",
            safesearch="moderate",
            max_results=max_results,
            backend="auto",
        )

    except Exception as exc:
        print(f"Search error: {exc}")
        return []

    formatted_results = []

    for item in results:

        formatted_results.append(
            {
                "title": item.get("title", ""),
                "url": item.get("href", ""),
                "snippet": item.get("body", ""),
            }
        )

    return formatted_results


def format_search_results(results):

    if not results:
        return "No web search results were found."

    chunks = []

    for index, result in enumerate(
        results,
        start=1,
    ):

        chunks.append(
            f"""
SOURCE {index}

Title:
{result["title"]}

URL:
{result["url"]}

Snippet:
{result["snippet"]}
"""
        )

    return "\n".join(chunks)