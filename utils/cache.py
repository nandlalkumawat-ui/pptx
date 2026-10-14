"""
Small on-disk JSON cache with a time-to-live.

Used for web and image search results so that re-running a topic does not
repeat identical DuckDuckGo requests.
"""

import hashlib
import json
import time
from pathlib import Path
from typing import Any

from config import CACHE_DIR, SEARCH_CACHE_TTL_SECONDS


def _cache_file(namespace: str, key: str) -> Path:
    digest = hashlib.sha1(key.encode("utf-8")).hexdigest()
    folder = CACHE_DIR / namespace
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{digest}.json"


def cache_get(namespace: str, key: str,
              ttl: int = SEARCH_CACHE_TTL_SECONDS) -> Any | None:
    path = _cache_file(namespace, key)
    try:
        if time.time() - path.stat().st_mtime > ttl:
            return None
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def cache_set(namespace: str, key: str, value: Any) -> None:
    path = _cache_file(namespace, key)
    temp = path.with_suffix(".tmp")
    try:
        temp.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
        temp.replace(path)
    except (OSError, TypeError, ValueError) as exc:
        # A failed cache write only costs speed on the next run.
        print(f"[warning] cache write failed ({namespace}): {exc}")
