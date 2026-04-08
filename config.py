import os
from dotenv import load_dotenv

load_dotenv()

# ── Gemini Settings ──
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-1.5-pro")

# ── Processing Settings ──
BATCH_SIZE = 4        # pages per Gemini API call (4 = safer for free-tier rate limits)
PAGE_DPI = 150        # resolution for page-to-PNG conversion
MAX_RETRIES = 5       # retries on Gemini API failure (handles 429 rate limit bursts)
