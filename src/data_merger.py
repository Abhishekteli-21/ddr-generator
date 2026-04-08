
import json
import logging
import re
from pathlib import Path

from src.gemini_client import send_text_to_gemini, parse_json_response

logger = logging.getLogger(__name__)


def _normalise_photo_refs(refs) -> list:
    """Coerce any photo-ref format to a sorted list of plain integers."""
    if not refs:
        return []
    result = []
    for ref in refs:
        digits = re.findall(r"\d+", str(ref))
        if digits:
            result.append(int(digits[0]))
    return sorted(set(result))


def _build_merge_prompt(inspection_data: list, thermal_data: list) -> str:
    return f"""
You are a data analyst merging two datasets from a property inspection.

── INSPECTION DATA ──
{json.dumps(inspection_data, indent=2)}

── THERMAL DATA ──
{json.dumps(thermal_data, indent=2)}

Task:
Match inspection areas with thermal images using semantic description matching.

CRITICAL NOTES:
  - The thermal data has NO area_name field. Match by reading thermal_observations
    and comparing to inspection observations + area_name.
  - Example match: thermal "cold band at bottom of wall near tiles" ↔ inspection
    area "Bathroom Wall" with "moisture at skirting".
  - inspection_photo_refs MUST be a list of PLAIN INTEGERS (e.g. [1, 2]).
    Never strings. These are the photo numbers from the inspection report.

For each merged entry produce ALL of these fields:
  - area_name               : clean standardised name (from inspection)
  - inspection_observations : list from inspection (or "Not Available")
  - inspection_checklist    : checklist dict from inspection (or "Not Available")
  - inspection_photo_refs   : list of INTEGERS (photo numbers) from inspection, or []
  - inspection_page         : integer page number from inspection, or null
  - thermal_filename        : e.g. "RB02380X.JPG" from thermal (or "Not Available")
  - thermal_observations    : from thermal (or "Not Available")
  - hotspot_temp            : from thermal (or "Not Available")
  - coldspot_temp           : from thermal (or "Not Available")
  - emissivity              : from thermal (or "Not Available")
  - thermal_date            : from thermal (or "Not Available")
  - thermal_page            : integer page number from thermal report, or null
  - match_confidence        : "high" | "medium" | "low"
  - conflicts               : description of any data conflict, or "None"
  - data_completeness       : "complete" | "inspection_only" | "thermal_only" | "partial"

Rules:
  - Do NOT invent data.
  - EVERY inspection area must appear in output, even with no thermal match.
  - If no thermal match, set all thermal fields to "Not Available" / null.
  - Avoid assigning the same thermal_page to multiple areas.
  - inspection_photo_refs: use integers ONLY, e.g. [1, 2, 3].

Return ONLY a valid JSON array. No preamble. No markdown fences.
"""


def merge_data(
    model,
    inspection_data: list,
    thermal_data: list,
    max_retries: int,
    checkpoint_dir: Path,
) -> list:
    """
    Merge inspection + thermal data via Gemini.
    Returns a list of merged area dicts with normalised photo refs.
    """
    checkpoint_dir.mkdir(parents=True, exist_ok=True)
    checkpoint_path = checkpoint_dir / "merged_data.json"

    logger.info(
        f"Merging {len(inspection_data)} inspection entries "
        f"with {len(thermal_data)} thermal entries…"
    )

    prompt = _build_merge_prompt(inspection_data, thermal_data)

    try:
        raw    = send_text_to_gemini(model, prompt, max_retries)
        merged = parse_json_response(raw)

        if not isinstance(merged, list):
            raise ValueError(f"Expected list, got {type(merged)}")

        # ── Normalise photo refs ──────────────────────────────────────────
        for item in merged:
            item["inspection_photo_refs"] = _normalise_photo_refs(
                item.get("inspection_photo_refs") or []
            )

        # ── Warn if coverage is incomplete ────────────────────────────────
        if len(merged) < len(inspection_data):
            logger.warning(
                f"Merged output has {len(merged)} entries but inspection had "
                f"{len(inspection_data)} — some areas may be missing."
            )

        # ── Deduplication by area_name (keep highest confidence) ──────────
        CONF_RANK = {"high": 3, "medium": 2, "low": 1}
        seen: dict = {}
        for item in merged:
            key = str(item.get("area_name", "")).lower().strip()
            if key in seen:
                ec = CONF_RANK.get(str(seen[key].get("match_confidence", "")).lower(), 0)
                nc = CONF_RANK.get(str(item.get("match_confidence", "")).lower(), 0)
                if nc > ec:
                    seen[key] = item
                    logger.debug(f"  Replaced duplicate '{key}' with higher-confidence entry")
            else:
                seen[key] = item

        deduped = list(seen.values())
        if len(deduped) < len(merged):
            logger.info(
                f"Dedup removed {len(merged) - len(deduped)} duplicate areas "
                f"({len(merged)} → {len(deduped)})"
            )
        merged = deduped

        with open(checkpoint_path, "w", encoding="utf-8") as f:
            json.dump(merged, f, indent=2)
        logger.info(
            f"Merged data checkpoint saved → {checkpoint_path} "
            f"({len(merged)} areas)"
        )
        return merged

    except Exception as exc:
        logger.error(f"Data merge failed: {exc}")
        raise
