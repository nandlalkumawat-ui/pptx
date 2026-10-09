from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent

OUTPUT_DIR = PROJECT_ROOT / "output"
PPTX_OUTPUT_DIR = OUTPUT_DIR / "pptx"
DOCX_OUTPUT_DIR = OUTPUT_DIR / "docx"

RESEARCH_DIR = PROJECT_ROOT / "data" / "research"


# Create directories
PPTX_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
DOCX_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
RESEARCH_DIR.mkdir(parents=True, exist_ok=True)


# Local LLM
OLLAMA_HOST = "http://localhost:11434"
OLLAMA_MODEL = "llama3.1:8b"


# Context window for Ollama. The PPTX prompt plus research is ~3-4k tokens
# and the JSON answer another ~2-3k, so the Ollama default (2048/4096)
# silently truncates the prompt.
OLLAMA_NUM_CTX = 8192
# Keep the model loaded between runs to avoid reload latency.
OLLAMA_KEEP_ALIVE = "30m"


# Search
MAX_SEARCH_RESULTS = 8

CACHE_DIR = OUTPUT_DIR / "cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)
SEARCH_CACHE_TTL_SECONDS = 7 * 24 * 3600


# Images
IMAGE_SEARCH_WORKERS = 3      # concurrent DuckDuckGo image queries
IMAGE_DOWNLOAD_TIMEOUT = (5, 12)  # (connect, read) seconds
IMAGE_MAX_BYTES = 15 * 1024 * 1024
IMAGE_MAX_EDGE_PX = 1600      # downscale before embedding
IMAGE_MIN_EDGE_PX = 400       # reject thumbnails / icons