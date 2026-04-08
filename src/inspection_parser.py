
import json
import logging
import re
from pathlib import Path

from src.gemini_client import send_pages_to_gemini, parse_json_response

logger = logging.getLogger(__name__)

INSPECTION_PROMPT = """
You are a property inspection expert carefully reading page screenshots of a property inspection report.

Your job is to extract ALL area-wise findings from these pages.

For every inspection area you find, extract:
  - area_name       : exact name of the area (e.g. "Hall - Ceiling", "Bathroom", "Balcony - Wall")
  - observations    : list of defects / issues noted (dampness, cracks, leakage, efflorescence, spalling, etc.)
  - checklist       : any checklist items visible on the page and their status (Yes / No / tick)
  - photo_references: list of INTEGERS representing photo numbers mentioned (e.g. [1, 2, 3]).
                      If the page says "Photo 1", "Image 2", or just "1", include the NUMBER ONLY.
                      Do NOT include the word "Photo" or "Image" — return plain integers like [1, 2].
  - page_number     : the page number where this finding appears (integer)

Important rules:
  - Extract ONLY what is written. Do NOT invent findings.
  - If a page is a cover page, table of contents, or has no findings — return an empty array.
  - An area can span multiple pages; include one entry per page occurrence.
  - Look carefully at handwritten text, tick marks, and numbered photos.
  - photo_references MUST be a list of plain integers, e.g. [1, 2, 3]. Never strings.

Return ONLY a valid JSON array. No preamble. No explanation. No markdown fences.

Example format:
[
  {
    "area_name": "Hall - Ceiling",
    "observations": ["dampness", "efflorescence", "spalling of paint"],
    "checklist": {"Dampness": "Yes", "Leakage": "Yes", "Cracks": "No"},
    "photo_references": [1, 2, 3],
    "page_number": 5
  }
]
"""


def _normalise_photo_refs(refs) -> list:
    """
    Coerce any photo-ref format to a list of plain integers.
    Handles: int, float, "1", "Photo 1", "photo_2", "IMG_001", etc.
    """
    if not refs:
        return []
    result = []
    for ref in refs:
        digits = re.findall(r"\d+", str(ref))
        if digits:
            result.append(int(digits[0]))
    return sorted(set(result))


def parse_inspection_report(
    model,
    page_paths: list,
    batch_size: int,
    max_retries: int,
    checkpoint_dir: Path,
) -> list:
    """
    Process inspection report pages in batches.
    Returns a unified, deduplicated list of area finding dicts.
    """
    checkpoint_dir.mkdir(parents=True, exist_ok=True)
    checkpoint_path = checkpoint_dir / "inspection_data.json"

    all_findings  = []
    total_batches = (len(page_paths) + batch_size - 1) // batch_size

    for i in range(0, len(page_paths), batch_size):
        batch     = page_paths[i: i + batch_size]
        batch_num = i // batch_size + 1
        logger.info(
            f"  Inspection batch {batch_num}/{total_batches} "
            f"(pages {i + 1}–{i + len(batch)})"
        )

        try:
            raw      = send_pages_to_gemini(model, batch, INSPECTION_PROMPT, max_retries)
            findings = parse_json_response(raw)

            if not isinstance(findings, list):
                logger.warning(f"    → Unexpected response type: {type(findings)}")
                continue

            # Normalise photo_references to plain integers
            for entry in findings:
                entry["photo_references"] = _normalise_photo_refs(
                    entry.get("photo_references") or []
                )

            all_findings.extend(findings)
            logger.info(f"    → {len(findings)} findings extracted")

        except Exception as exc:
            logger.error(f"    → Batch {batch_num} failed, skipping: {exc}")

    # Deduplication: same (area_name.lower, page_number) key
    seen_keys: set = set()
    deduped   = []
    for entry in all_findings:
        key = (
            str(entry.get("area_name", "")).lower().strip(),
            str(entry.get("page_number")),
        )
        if key not in seen_keys:
            seen_keys.add(key)
            deduped.append(entry)

    removed = len(all_findings) - len(deduped)
    if removed:
        logger.info(f"Inspection dedup: {len(all_findings)} → {len(deduped)} ({removed} removed)")
    all_findings = deduped

    with open(checkpoint_path, "w", encoding="utf-8") as f:
        json.dump(all_findings, f, indent=2)
    logger.info(
        f"Inspection checkpoint saved → {checkpoint_path} "
        f"({len(all_findings)} total entries)"
    )
    return all_findings
