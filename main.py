
import argparse
import json
import logging
import sys
from pathlib import Path

import config
from src.pdf_processor     import convert_pages_to_png, extract_images_from_pdf, extract_site_photos_from_pdf
from src.gemini_client     import init_gemini
from src.inspection_parser import parse_inspection_report
from src.thermal_parser    import parse_thermal_report
from src.data_merger       import merge_data
from src.report_generator  import generate_ddr_content
from src.docx_builder      import build_docx

# ── Logging setup ─────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("ddr_generator.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger("pipeline")

# ── Default paths ──────────────────────────────────────────────────────────
BASE_DIR        = Path(__file__).parent
INPUT_DIR       = BASE_DIR / "input"
OUTPUT_DIR      = BASE_DIR / "output"
TEMP_DIR        = BASE_DIR / "temp"
PAGES_DIR       = TEMP_DIR / "pages"
IMAGES_DIR      = TEMP_DIR / "extracted_images"
CHECKPOINT_DIR  = TEMP_DIR / "checkpoints"


def _load_checkpoint(path: Path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _clear_checkpoints():
    for p in CHECKPOINT_DIR.glob("*.json"):
        p.unlink()
    logger.info("All checkpoints cleared — fresh run.")


# ── Pipeline ──────────────────────────────────────────────────────────────

def run_pipeline(
    inspection_pdf: Path,
    thermal_pdf:    Path,
    output_path:    Path,
    fresh:          bool = False,
    api_key:        str  = "",
    pages_dir:      Path = PAGES_DIR,
    images_dir:     Path = IMAGES_DIR,
    checkpoint_dir: Path = CHECKPOINT_DIR,
) -> Path:
    """
    Full DDR generation pipeline.
    Returns the Path to the generated .docx file.
    """

    # ── Resolve API key ───────────────────────────────────────────────────
    api_key = api_key or config.GEMINI_API_KEY
    if not api_key:
        logger.error("No Gemini API key found. Set GEMINI_API_KEY in .env or pass --api-key.")
        sys.exit(1)

    # ── Validate inputs ───────────────────────────────────────────────────
    for label, path in [("Inspection PDF", inspection_pdf), ("Thermal PDF", thermal_pdf)]:
        if not path.exists():
            logger.error(f"{label} not found: {path}")
            sys.exit(1)

    if fresh:
        _clear_checkpoints()

    checkpoint_dir.mkdir(parents=True, exist_ok=True)

    logger.info("=" * 60)
    logger.info("DDR GENERATOR — PIPELINE START")
    logger.info("=" * 60)
    logger.info(f"Inspection : {inspection_pdf}")
    logger.info(f"Thermal    : {thermal_pdf}")
    logger.info(f"Output     : {output_path}")

    # ── Init Gemini ───────────────────────────────────────────────────────
    model = init_gemini(api_key, config.GEMINI_MODEL)

    # ════════════════════════════════════════════════════════════════
    # PHASE 1 — PDF Processing
    # ════════════════════════════════════════════════════════════════
    logger.info("\n── PHASE 1: PDF Processing ──────────────────────────────")

    logger.info("Converting inspection pages → PNG…")
    inspection_pages = convert_pages_to_png(
        inspection_pdf, "inspection", pages_dir, config.PAGE_DPI
    )

    logger.info("Converting thermal pages → PNG…")
    thermal_pages = convert_pages_to_png(
        thermal_pdf, "thermal", pages_dir, config.PAGE_DPI
    )

    logger.info("Extracting embedded images from thermal PDF…")
    thermal_images = extract_images_from_pdf(thermal_pdf, "thermal", images_dir)
    total_therm_imgs = sum(len(v) for v in thermal_images.values())
    logger.info(f"  → {total_therm_imgs} thermal images extracted across {len(thermal_images)} pages")

    logger.info("Extracting site photos from inspection PDF (logo-filtered)…")
    inspection_images = extract_site_photos_from_pdf(inspection_pdf, "inspection", images_dir)
    total_insp_imgs = sum(len(v) for v in inspection_images.values())
    logger.info(f"  → {total_insp_imgs} inspection site photos extracted")

    # ════════════════════════════════════════════════════════════════
    # PHASE 2 — AI Understanding (with checkpoint reuse)
    # ════════════════════════════════════════════════════════════════
    logger.info("\n── PHASE 2: AI Understanding ────────────────────────────")

    insp_ckpt    = checkpoint_dir / "inspection_data.json"
    thermal_ckpt = checkpoint_dir / "thermal_data.json"

    if insp_ckpt.exists():
        logger.info("Reusing inspection checkpoint (pass --fresh to reprocess)")
        inspection_data = _load_checkpoint(insp_ckpt)
    else:
        logger.info("Parsing inspection report…")
        inspection_data = parse_inspection_report(
            model, inspection_pages,
            config.BATCH_SIZE, config.MAX_RETRIES, checkpoint_dir
        )

    if thermal_ckpt.exists():
        logger.info("Reusing thermal checkpoint (pass --fresh to reprocess)")
        thermal_data = _load_checkpoint(thermal_ckpt)
    else:
        logger.info("Parsing thermal report…")
        thermal_data = parse_thermal_report(
            model, thermal_pages,
            config.BATCH_SIZE, config.MAX_RETRIES, checkpoint_dir
        )

    logger.info(f"  Inspection areas  : {len(inspection_data)}")
    logger.info(f"  Thermal entries   : {len(thermal_data)}")

    # ════════════════════════════════════════════════════════════════
    # PHASE 3 — Merge
    # ════════════════════════════════════════════════════════════════
    logger.info("\n── PHASE 3: Merging Data ────────────────────────────────")

    merged_ckpt = checkpoint_dir / "merged_data.json"
    if merged_ckpt.exists():
        logger.info("Reusing merge checkpoint")
        merged_data = _load_checkpoint(merged_ckpt)
    else:
        merged_data = merge_data(
            model, inspection_data, thermal_data,
            config.MAX_RETRIES, checkpoint_dir
        )

    logger.info(f"  Merged areas: {len(merged_data)}")

    # ════════════════════════════════════════════════════════════════
    # PHASE 4 — Generate DDR Content
    # ════════════════════════════════════════════════════════════════
    logger.info("\n── PHASE 4: Generating DDR Content ─────────────────────")

    ddr_ckpt = checkpoint_dir / "ddr_content.json"
    if ddr_ckpt.exists():
        logger.info("Reusing DDR content checkpoint")
        ddr_content = _load_checkpoint(ddr_ckpt)
    else:
        ddr_content = generate_ddr_content(
            model, merged_data,
            config.MAX_RETRIES, checkpoint_dir
        )

    # ════════════════════════════════════════════════════════════════
    # PHASE 5 — Build Word Document
    # ════════════════════════════════════════════════════════════════
    logger.info("\n── PHASE 5: Building Word Document ─────────────────────")

    result = build_docx(
        ddr_content        = ddr_content,
        merged_data        = merged_data,
        extracted_images   = thermal_images,      # thermal PDF images
        inspection_images  = inspection_images,   # ← NOW WIRED IN
        output_path        = output_path,
    )

    logger.info("\n" + "=" * 60)
    logger.info(f"✅  DDR GENERATED SUCCESSFULLY")
    logger.info(f"    Output: {result}")
    logger.info("=" * 60)
    return result


# ── CLI entry point ───────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="DDR Report Generator — converts inspection + thermal PDFs into a Word report."
    )
    parser.add_argument(
        "--inspection", type=Path,
        default=INPUT_DIR / "inspection_report.pdf",
        help="Path to the Inspection Report PDF (default: input/inspection_report.pdf)",
    )
    parser.add_argument(
        "--thermal", type=Path,
        default=INPUT_DIR / "thermal_report.pdf",
        help="Path to the Thermal Report PDF (default: input/thermal_report.pdf)",
    )
    parser.add_argument(
        "--output", type=Path,
        default=OUTPUT_DIR / "DDR_Report.docx",
        help="Output path for the DDR Word document (default: output/DDR_Report.docx)",
    )
    parser.add_argument(
        "--api-key", type=str, default="",
        help="Gemini API key (overrides .env file)",
    )
    parser.add_argument(
        "--fresh", action="store_true",
        help="Clear all checkpoints and rerun from scratch",
    )

    args = parser.parse_args()

    run_pipeline(
        inspection_pdf = args.inspection,
        thermal_pdf    = args.thermal,
        output_path    = args.output,
        fresh          = args.fresh,
        api_key        = args.api_key,
    )
