
import json
import re
import time
import logging
import PIL.Image
from pathlib import Path
from google import genai
from google.genai import types

logger = logging.getLogger(__name__)


class GeminiModel:
    """Thin wrapper that holds the client + model name, mimicking the old GenerativeModel interface."""

    def __init__(self, client: genai.Client, model_name: str):
        self.client = client
        self.model_name = model_name

    def generate_content(self, contents):
        """Accepts a list of [prompt_str, PIL.Image, ...] or a plain str."""
        parts = []

        if isinstance(contents, str):
            parts = [contents]
        else:
            for item in contents:
                if isinstance(item, str):
                    parts.append(item)
                elif isinstance(item, PIL.Image.Image):
                    parts.append(item)
                else:
                    parts.append(str(item))

        response = self.client.models.generate_content(
            model=self.model_name,
            contents=parts,
        )
        return response


def init_gemini(api_key: str, model_name: str) -> GeminiModel:
    """Configure and return a GeminiModel wrapper."""
    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY is empty. "
            "Add it to your .env file or pass it via --api-key."
        )
    client = genai.Client(api_key=api_key)
    model = GeminiModel(client, model_name)
    logger.info(f"Gemini initialised — model: {model_name}")
    return model


# ── JSON helpers ─────────────────────────────────────────────────────────────

def _clean_json(text: str) -> str:
    """Strip markdown code fences that Gemini sometimes wraps around JSON."""
    text = re.sub(r"```(?:json)?\s*", "", text)
    text = re.sub(r"```\s*", "", text)
    return text.strip()


def parse_json_response(text: str):
    """
    Parse JSON from a Gemini response.
    Tries direct parse first, then hunts for the outermost [ ] or { }.
    Raises json.JSONDecodeError if nothing works.
    """
    cleaned = _clean_json(text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        # Try to extract array
        m = re.search(r"(\[[\s\S]*\])", cleaned)
        if m:
            return json.loads(m.group(1))
        # Try to extract object
        m = re.search(r"(\{[\s\S]*\})", cleaned)
        if m:
            return json.loads(m.group(1))
        raise


# ── API calls ─────────────────────────────────────────────────────────────────

def send_pages_to_gemini(model: GeminiModel, page_paths: list, prompt: str, max_retries: int = 5) -> str:
    """
    Send a batch of page PNG images + a prompt to Gemini.
    Returns the raw response text.
    Waits 65 s on 429 RESOURCE_EXHAUSTED before retrying.
    """
    for attempt in range(max_retries + 1):
        try:
            images = [PIL.Image.open(str(p)) for p in page_paths]
            content = [prompt] + images
            response = model.generate_content(content)
            return response.text
        except Exception as e:
            err_str = str(e)
            is_rate_limit = "429" in err_str or "RESOURCE_EXHAUSTED" in err_str
            if attempt < max_retries:
                wait = 65 if is_rate_limit else min(2 ** attempt, 30)
                logger.warning(
                    f"Gemini image call failed (attempt {attempt + 1}/{max_retries + 1}), "
                    f"{'rate-limit — ' if is_rate_limit else ''}retrying in {wait}s: {e}"
                )
                time.sleep(wait)
            else:
                logger.error(f"Gemini image call failed after {max_retries + 1} attempts: {e}")
                raise


def send_text_to_gemini(model: GeminiModel, prompt: str, max_retries: int = 5) -> str:
    """
    Send a text-only prompt to Gemini.
    Returns the raw response text.
    Waits 65 s on 429 RESOURCE_EXHAUSTED before retrying.
    """
    for attempt in range(max_retries + 1):
        try:
            response = model.generate_content(prompt)
            return response.text
        except Exception as e:
            err_str = str(e)
            is_rate_limit = "429" in err_str or "RESOURCE_EXHAUSTED" in err_str
            if attempt < max_retries:
                wait = 65 if is_rate_limit else min(2 ** attempt, 30)
                logger.warning(
                    f"Gemini text call failed (attempt {attempt + 1}/{max_retries + 1}), "
                    f"{'rate-limit — ' if is_rate_limit else ''}retrying in {wait}s: {e}"
                )
                time.sleep(wait)
            else:
                logger.error(f"Gemini text call failed after {max_retries + 1} attempts: {e}")
                raise

