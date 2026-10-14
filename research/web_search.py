from ddgs import DDGS

from config import MAX_SEARCH_RESULTS, WEB_SEARCH_CACHE_TTL_SECONDS
from utils.cache import cache_get, cache_set
from utils.perf import PERF


def search_web(
    query: str,
    max_results: int = MAX_SEARCH_RESULTS,
):
    """
    Search the web using DDGS.

    No paid API key is required. Results are cached on disk for
    WEB_SEARCH_CACHE_TTL_SECONDS; failed or empty searches are not cached.
    """

    print(f"Searching for: {query}")

    cache_key = f"{' '.join(query.lower().split())}|{max_results}"
    cached = cache_get("web_search", cache_key, ttl=WEB_SEARCH_CACHE_TTL_SECONDS)
    if cached:
        PERF.incr("web search cache hits")
        return cached

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

    if formatted_results:
        cache_set("web_search", cache_key, formatted_results)
    return formatted_results


def compact_research(results) -> str:
    """
    Research text for the LLM prompt: numbered title + snippet, no URLs.

    URLs are long in tokens, the model never needs them (sources are taken
    from the search results directly), and on CPU every prompt token costs
    roughly 0.1-0.15 s of prompt processing.
    """

    if not results:
        return "No web search results were found."

    return "\n".join(
        f"[{index}] {result['title']}: {' '.join(str(result['snippet']).split())}"
        for index, result in enumerate(results, start=1)
    )


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