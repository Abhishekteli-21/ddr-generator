

import json
import logging
from pathlib import Path

from src.gemini_client import send_pages_to_gemini, parse_json_response

logger = logging.getLogger(__name__)

THERMAL_PROMPT = """
You are a thermal imaging expert reading page screenshots of a Bosch GTC 400 C thermal inspection report.

ACTUAL PAGE LAYOUT (confirmed):
  HEADER   : "Thermal image : <FILENAME>" | "Device: GTC 400 C Professional" | "Serial Number: ..."
  LEFT SIDE: Two images stacked vertically:
               - TOP    = thermal heatmap (false-colour: red=hot, blue=cold)
               - BOTTOM = regular site photograph of the same location
  RIGHT SIDE: Data table with:
               - Date (top-right, e.g. "27/09/22")
               - Hotspot temperature (red crosshair icon)
               - Coldspot temperature (blue crosshair icon)
               - Emissivity (ε icon, always 0.94)
               - Reflected temperature

IMPORTANT: There is NO area/location name label on these pages. Do NOT invent one.

For each page in this batch, extract:
  - thermal_filename     : the filename from the header (e.g. "RB02380X.JPG")
  - hotspot_temp         : hotspot temperature shown (e.g. "28.8 °C")
  - coldspot_temp        : coldspot temperature shown (e.g. "23.4 °C")
  - emissivity           : emissivity value (e.g. "0.94")
  - reflected_temp       : reflected temperature if shown (e.g. "23 °C")
  - date                 : date shown top-right (e.g. "27/09/22")
  - thermal_observations : 2-3 sentence description of the thermal heatmap:
                           - Describe color pattern (blue=cold, green=moderate, red/yellow=hot)
                           - Location of cold patches (bottom-left, base, corner, etc.)
                           - What the site photo shows (wall, skirting, tiles, ceiling, etc.)
                           - Likely cause (moisture, dampness, cold bridge, etc.)
  - page_number          : position in THIS batch (1 = first page you see, 2 = second, etc.)

CRITICAL RULES:
  - One entry per page. Do NOT skip pages.
  - page_number is relative to this batch only (1 to batch_size).
  - Do NOT write "Not Available" for thermal_filename — it is always in the header.
  - Do NOT invent an area_name field — it does not exist in this document.

Return ONLY a valid JSON array. No preamble. No markdown. No explanation.

Example for a 2-page batch:
[
  {
    "thermal_filename": "RB02380X.JPG",
    "hotspot_temp": "28.8 °C",
    "coldspot_temp": "23.4 °C",
    "emissivity": "0.94",
    "reflected_temp": "23 °C",
    "date": "27/09/22",
    "thermal_observations": "Large blue-green cold zone covering bottom-left quadrant of the thermal heatmap, indicating significant moisture accumulation at the base of the wall. The site photo shows a corner wall junction with visible dampness staining. Consistent with water ingress at skirting level.",
    "page_number": 1
  },
  {
    "thermal_filename": "RB02386X.JPG",
    "hotspot_temp": "27.4 °C",
    "coldspot_temp": "22.4 °C",
    "emissivity": "0.94",
    "reflected_temp": "23 °C",
    "date": "27/09/22",
    "thermal_observations": "Horizontal cold band clearly visible along the bottom of the thermal image in blue tones. Site photo shows a white painted wall at skirting level with minor surface staining. Moisture ingress at the base of the wall.",
    "page_number": 2
  }
]
"""


def parse_thermal_report(
    model,
    page_paths: list,
    batch_size: int,
    max_retries: int,
    checkpoint_dir: Path
) -> list:
    """
    Process thermal report pages in batches with incremental checkpointing.
    Resumes from last saved batch if interrupted.
    Returns a unified list of thermal entries (dicts).
    """
    checkpoint_dir.mkdir(parents=True, exist_ok=True)
    checkpoint_path = checkpoint_dir / "thermal_data.json"

    # ── Resume: load any already-processed entries ────────────────────────
    all_entries = []
    processed_pages: set = set()
    if checkpoint_path.exists():
        try:
            with open(checkpoint_path, encoding="utf-8") as f:
                all_entries = json.load(f)
            processed_pages = {e.get("page_number") for e in all_entries if e.get("page_number")}
            if processed_pages:
                logger.info(f"  Resuming thermal parse — {len(all_entries)} entries already saved "
                            f"(pages {sorted(processed_pages)})")
        except Exception:
            all_entries = []

    total_batches = (len(page_paths) + batch_size - 1) // batch_size

    for i in range(0, len(page_paths), batch_size):
        batch = page_paths[i : i + batch_size]
        batch_num = i // batch_size + 1
        batch_start_page = i + 1
        batch_end_page   = i + len(batch)

        # Skip if all pages in this batch are already processed
        batch_pages = set(range(batch_start_page, batch_end_page + 1))
        if batch_pages.issubset(processed_pages):
            logger.info(f"  Thermal batch {batch_num}/{total_batches} — already done, skipping.")
            continue

        logger.info(f"  Thermal batch {batch_num}/{total_batches} "
                    f"(pages {batch_start_page}–{batch_end_page})")

        try:
            raw = send_pages_to_gemini(model, batch, THERMAL_PROMPT, max_retries)
            entries = parse_json_response(raw)

            if isinstance(entries, list):
                # Convert Gemini's relative page_number (1..8) → absolute PDF page number
                for entry in entries:
                    rel = entry.get("page_number")
                    if isinstance(rel, int) and 1 <= rel <= len(batch):
                        entry["page_number"] = batch_start_page + rel - 1
                    else:
                        entry["page_number"] = batch_start_page + entries.index(entry)
                all_entries.extend(entries)
                processed_pages.update(e.get("page_number") for e in entries)
                logger.info(f"    → {len(entries)} thermal entries extracted "
                            f"(pages {batch_start_page}–{batch_start_page+len(entries)-1})")
                # ── Incremental checkpoint after every batch ───────────────
                with open(checkpoint_path, "w", encoding="utf-8") as f:
                    json.dump(all_entries, f, indent=2)
                logger.info(f"    Checkpoint saved ({len(all_entries)} total so far)")
            else:
                logger.warning(f"    → Unexpected response type: {type(entries)}")

        except Exception as e:
            logger.error(f"    → Batch {batch_num} failed, skipping: {e}")

    # Deduplicate: same area_name + page_number from overlapping batches
    seen_keys: set = set()
    deduped = []
    for entry in all_entries:
        key = (
            str(entry.get("area_name", "")).lower().strip(),
            entry.get("page_number"),
        )
        if key not in seen_keys:
            seen_keys.add(key)
            deduped.append(entry)
    if len(deduped) < len(all_entries):
        logger.info(
            f"Thermal dedup: {len(all_entries)} → {len(deduped)} entries "
            f"({len(all_entries) - len(deduped)} duplicates removed)"
        )
    all_entries = deduped

    with open(checkpoint_path, "w", encoding="utf-8") as f:
        json.dump(all_entries, f, indent=2)
    logger.info(f"Thermal checkpoint saved → {checkpoint_path} "
                f"({len(all_entries)} total entries)")

    return all_entries
