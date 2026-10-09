import hashlib
import io
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlparse

import requests
from PIL import Image, ImageOps, UnidentifiedImageError
from ddgs import DDGS
from requests.adapters import HTTPAdapter

from config import (
    IMAGE_DOWNLOAD_TIMEOUT,
    IMAGE_MAX_BYTES,
    IMAGE_MAX_EDGE_PX,
    IMAGE_MIN_EDGE_PX,
    IMAGE_SEARCH_WORKERS,
    PROJECT_ROOT,
)
from utils.cache import cache_get, cache_set
from utils.perf import PERF

IMAGE_CACHE_DIR = PROJECT_ROOT / "output" / "images"
IMAGE_CACHE_DIR.mkdir(parents=True, exist_ok=True)

# Formats python-pptx can embed. Everything else (WebP, AVIF, ...) is
# re-encoded to JPEG/PNG before insertion.
_PPTX_SAFE_SUFFIXES = (".jpg", ".png")

# Stock-photo sites serve watermarked previews and their licences forbid
# reuse, so they are never selected.
_BLOCKED_DOMAINS = (
    "shutterstock.", "istockphoto.", "gettyimages.", "alamy.", "dreamstime.",
    "123rf.", "depositphotos.", "stock.adobe.", "vectorstock.", "freepik.",
    "vecteezy.", "pngtree.", "canstockphoto.", "bigstockphoto.", "pond5.",
    "pinterest.", "pinimg.", "facebook.", "fbsbx.", "instagram.", "lookaside.",
)
_PENALTY_WORDS = {"logo", "icon", "clipart", "watermark", "meme", "banner",
                  "wallpaper", "cartoon", "vector", "stock"}
_STOPWORDS = {
    "the", "and", "for", "with", "from", "into", "that", "this", "their",
    "about", "over", "under", "between", "what", "how", "why", "are", "its",
    "of", "in", "on", "to", "a", "an", "by", "or", "as", "is", "be", "vs",
    "image", "photo", "picture", "slide", "introduction", "overview",
    "conclusion", "summary", "key", "takeaways",
}

# Slides of these kinds do not get a photo: data slides show their table,
# process/comparison slides get a native step/card diagram from the
# renderer, conclusion/reference slides stay clean.
_NO_IMAGE_LAYOUTS = {"data", "process", "comparison", "conclusion", "references", "agenda"}

_thread_local = threading.local()
_failed_urls: set[str] = set()
_failed_lock = threading.Lock()


def _normalize_query(value: str) -> str:
    """Clean a query for web/image searches."""
    return " ".join(str(value).strip().split())


def _dedupe(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        if not item:
            continue
        normalized = str(item).strip()
        if normalized and normalized not in seen:
            seen.add(normalized)
            result.append(normalized)
    return result


def _keywords(text: str) -> set[str]:
    words = re.findall(r"[a-z0-9]+", str(text).lower())
    return {word for word in words if len(word) > 2 and word not in _STOPWORDS}


def _session() -> requests.Session:
    """One pooled HTTP session per worker thread (connection reuse)."""
    session = getattr(_thread_local, "session", None)
    if session is None:
        session = requests.Session()
        session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
            ),
            "Accept": "image/avif,image/webp,image/png,image/jpeg,image/*;q=0.8",
        })
        adapter = HTTPAdapter(pool_connections=8, pool_maxsize=8, max_retries=0)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        _thread_local.session = session
    return session


def search_image_candidates(query: str, max_results: int = 5) -> list[dict[str, Any]]:
    """Search for high-quality image URLs using DuckDuckGo."""
    cleaned = _normalize_query(query)
    if not cleaned:
        return []

    cache_key = f"{cleaned}|{max_results}"
    cached = cache_get("image_search", cache_key)
    if cached is not None:
        PERF.incr("image search cache hits")
        return cached
    PERF.incr("image search cache misses")

    started = time.perf_counter()
    results = None
    # One retry with a short back-off: DuckDuckGo occasionally rate-limits
    # bursts of queries, which previously left slides silently image-less.
    for attempt in range(2):
        try:
            with DDGS(timeout=10) as ddgs:
                results = ddgs.images(
                    cleaned,
                    region="wt-wt",
                    safesearch="moderate",
                    max_results=max_results,
                )
            break
        except Exception as exc:  # pragma: no cover - defensive runtime guard
            print(f"Image search failed for '{cleaned}' (attempt {attempt + 1}): {exc}")
            time.sleep(1.5)
    PERF.add_time("image search (cumulative)", time.perf_counter() - started)

    candidates: list[dict[str, Any]] = []
    for item in results or []:
        item = item or {}
        image_url = item.get("image") or ""
        if not str(image_url).startswith(("http://", "https://")):
            continue

        candidates.append(
            {
                "title": str(item.get("title") or cleaned).strip(),
                "url": str(image_url).strip(),
                "page_url": str(item.get("url") or "").strip(),
                "thumbnail": str(item.get("thumbnail") or "").strip(),
                "source": str(item.get("source") or "").strip(),
                "width": item.get("width"),
                "height": item.get("height"),
            }
        )

    if results is not None:
        cache_set("image_search", cache_key, candidates)
    return candidates


def _domain(url: str) -> str:
    host = urlparse(url).netloc.lower()
    return host[4:] if host.startswith("www.") else host


def _rank_candidates(candidates: list[dict[str, Any]], keywords: set[str]) -> list[dict[str, Any]]:
    """Order candidates by relevance to the slide and drop unusable ones."""

    scored = []
    for candidate in candidates:
        url = candidate["url"]
        page = candidate.get("page_url", "")
        if any(blocked in _domain(url) or blocked in _domain(page) for blocked in _BLOCKED_DOMAINS):
            continue

        try:
            width = int(candidate.get("width") or 0)
            height = int(candidate.get("height") or 0)
        except (TypeError, ValueError):
            width = height = 0

        if width and height:
            if min(width, height) < IMAGE_MIN_EDGE_PX:
                continue
            ratio = width / height
            if ratio > 3.0 or ratio < 0.33:
                continue

        text = f"{candidate.get('title', '')} {urlparse(page).path} {urlparse(url).path}"
        candidate_words = _keywords(text.replace("-", " ").replace("_", " "))

        score = 2.0 * len(keywords & candidate_words)
        if width and height:
            score += 1.0 if min(width, height) >= 700 else 0.0
            score += 0.5 if 1.2 <= width / height <= 1.9 else 0.0
        score -= 2.0 * len(candidate_words & _PENALTY_WORDS)

        scored.append((score, candidate))

    scored.sort(key=lambda pair: pair[0], reverse=True)
    return [candidate for _, candidate in scored]


def _cache_path_for_url(url: str, cache_dir: Path | None = None, suffix: str = ".jpg") -> Path:
    cache_dir = cache_dir or IMAGE_CACHE_DIR
    cache_dir.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha1(url.encode("utf-8")).hexdigest()
    # The size suffix keeps files from older (unresized / unvalidated)
    # versions of this module from being reused.
    return cache_dir / f"{digest}_{IMAGE_MAX_EDGE_PX}{suffix}"


def _cached_image(url: str, cache_dir: Path | None) -> Path | None:
    for suffix in _PPTX_SAFE_SUFFIXES:
        path = _cache_path_for_url(url, cache_dir, suffix)
        if path.is_file() and path.stat().st_size > 0:
            try:
                with Image.open(path) as image:
                    image.verify()
                return path
            except Exception:
                path.unlink(missing_ok=True)
    return None


def _download_bytes(url: str, timeout=IMAGE_DOWNLOAD_TIMEOUT) -> bytes:
    with _session().get(url, timeout=timeout, stream=True) as response:
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "").lower()
        if content_type and not content_type.startswith(("image/", "application/octet-stream", "binary/")):
            raise ValueError(f"not an image (Content-Type: {content_type})")

        declared = int(response.headers.get("Content-Length") or 0)
        if declared > IMAGE_MAX_BYTES:
            raise ValueError(f"image too large ({declared} bytes)")

        chunks = []
        total = 0
        for chunk in response.iter_content(chunk_size=64 * 1024):
            total += len(chunk)
            if total > IMAGE_MAX_BYTES:
                raise ValueError("image too large")
            chunks.append(chunk)
        return b"".join(chunks)


def _convert_image(data: bytes, cache_file: Path) -> Path | None:
    """
    Validate, orient, downscale and re-encode an image as JPEG (or PNG when
    it has real transparency). Returns the written path or None if the
    payload is not a usable picture.
    """
    try:
        with Image.open(io.BytesIO(data)) as image:
            if image.format == "JPEG":
                # Decode large JPEGs at reduced scale - much faster.
                image.draft("RGB", (IMAGE_MAX_EDGE_PX, IMAGE_MAX_EDGE_PX))
            image.seek(0)  # first frame of animated GIF/WebP
            image = ImageOps.exif_transpose(image)

            if min(image.size) < IMAGE_MIN_EDGE_PX:
                return None
            ratio = image.width / image.height
            if ratio > 3.0 or ratio < 0.33:
                return None

            image.thumbnail((IMAGE_MAX_EDGE_PX, IMAGE_MAX_EDGE_PX), Image.LANCZOS)

            has_alpha = image.mode in ("RGBA", "LA") or (
                image.mode == "P" and "transparency" in image.info
            )
            if has_alpha:
                rgba = image.convert("RGBA")
                if rgba.getchannel("A").getextrema()[0] < 250:
                    target = cache_file.with_suffix(".png")
                    rgba.save(target, format="PNG", optimize=False)
                    return target
                image = rgba

            target = cache_file.with_suffix(".jpg")
            temp = target.with_suffix(".tmp")
            image.convert("RGB").save(temp, format="JPEG", quality=85, optimize=True)
            temp.replace(target)
            return target
    except (UnidentifiedImageError, OSError, ValueError, Image.DecompressionBombError):
        return None


def download_image(url: str, cache_dir: Path | None = None, timeout=IMAGE_DOWNLOAD_TIMEOUT) -> Path | None:
    """Download and normalize a single image to a local JPEG/PNG file."""
    if not url:
        return None

    cached = _cached_image(url, cache_dir)
    if cached is not None:
        PERF.incr("image file cache hits")
        return cached

    with _failed_lock:
        if url in _failed_urls:
            return None
    if cache_get("image_failed", url, ttl=3 * 24 * 3600):
        return None
    PERF.incr("image file cache misses")

    started = time.perf_counter()
    try:
        payload = _download_bytes(url, timeout=timeout)
    except Exception as exc:  # broken URL, timeout, 403, HTML page ...
        print(f"Image download failed for {url[:100]}: {exc}")
        payload = None
    finally:
        PERF.add_time("image download (cumulative)", time.perf_counter() - started)

    result = None
    if payload:
        started = time.perf_counter()
        result = _convert_image(payload, _cache_path_for_url(url, cache_dir))
        PERF.add_time("image conversion (cumulative)", time.perf_counter() - started)
        if result is None:
            print(f"Image rejected (invalid, too small or odd shape): {url[:100]}")

    if result is None:
        PERF.incr("images failed")
        with _failed_lock:
            _failed_urls.add(url)
        cache_set("image_failed", url, True)
        return None

    PERF.incr("images downloaded")
    return result


def _image_fingerprint(path: Path) -> str:
    """Tiny average-hash so the same picture from two URLs is not reused."""
    with Image.open(path) as image:
        small = image.convert("L").resize((8, 8))
        pixels = list(small.getdata())
    average = sum(pixels) / len(pixels)
    return "".join("1" if pixel > average else "0" for pixel in pixels)


def _slide_get(slide: Any, key: str, default: Any = "") -> Any:
    if isinstance(slide, dict):
        return slide.get(key, default)
    return getattr(slide, key, default)


def _slide_set(slide: Any, key: str, value: Any) -> None:
    if isinstance(slide, dict):
        slide[key] = value
    else:
        setattr(slide, key, value)


def _wants_image(slide: Any, index: int, total: int) -> bool:
    """Only add a picture where it supports the slide's message."""

    layout = str(_slide_get(slide, "layout", "") or "").lower()
    if layout in _NO_IMAGE_LAYOUTS:
        return False
    if not layout:
        # No hint: keep the title slide, skip the final (conclusion) slide.
        if total > 2 and index == total - 1:
            return False
    if _slide_get(slide, "table", None):
        return False
    return bool(_slide_get(slide, "image_query", "") or _slide_get(slide, "title", ""))


def _build_query(slide: Any, query_hint: str) -> str:
    query = _normalize_query(_slide_get(slide, "image_query", "") or _slide_get(slide, "title", ""))
    hint_words = _keywords(query_hint)
    # Anchor generic queries to the presentation topic.
    if query_hint and not (hint_words & _keywords(query)):
        query = f"{query} {query_hint}"
    return _normalize_query(query)


def enrich_presentation_with_images(presentation: Any, query_hint: str = "") -> Any:
    """Attach local image paths to each slide from its requested query or title."""
    if presentation is None:
        return presentation

    if hasattr(presentation, "slides"):
        slides = presentation.slides
    elif isinstance(presentation, dict):
        slides = presentation.get("slides", [])
    else:
        return presentation

    used_urls: set[str] = set()
    used_fingerprints: set[str] = set()
    claim_lock = threading.Lock()

    def process(index: int, slide: Any) -> None:
        image_path = _slide_get(slide, "image_path", "")
        if image_path and Path(image_path).is_file():
            return
        if not _wants_image(slide, index, len(slides)):
            return

        PERF.incr("slides requesting an image")
        query = _build_query(slide, query_hint)
        if not query:
            return

        keywords = _keywords(f"{_slide_get(slide, 'title', '')} {_slide_get(slide, 'image_query', '')}")
        candidates = _rank_candidates(search_image_candidates(query, max_results=12), keywords)
        PERF.incr("image candidates found", len(candidates))

        attempts = 0
        for item in candidates:
            if attempts >= 4:
                break
            with claim_lock:
                if item["url"] in used_urls:
                    continue
                used_urls.add(item["url"])
            attempts += 1

            downloaded = download_image(item["url"], IMAGE_CACHE_DIR)
            if downloaded is None:
                continue

            fingerprint = _image_fingerprint(downloaded)
            with claim_lock:
                if fingerprint in used_fingerprints:
                    continue
                used_fingerprints.add(fingerprint)

            page_url = item.get("page_url") or item["url"]
            _slide_set(slide, "image_path", str(downloaded))
            _slide_set(slide, "image_url", item["url"])
            _slide_set(slide, "image_source", _domain(page_url))
            _slide_set(slide, "image_caption", f"Image: {_domain(page_url)}")
            _slide_set(
                slide,
                "source_urls",
                _dedupe(list(_slide_get(slide, "source_urls", []) or []) + [page_url]),
            )
            PERF.incr("images attached to slides")
            return

        print(f"No usable image found for slide {index + 1}: '{query}'")

    # Bounded pool: enough to overlap network waits without tripping
    # DuckDuckGo's rate limiting.
    with ThreadPoolExecutor(max_workers=IMAGE_SEARCH_WORKERS) as pool:
        list(pool.map(process, range(len(slides)), slides))

    return presentation
