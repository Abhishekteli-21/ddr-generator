
import logging
import os
import sys
import tempfile
import time
from pathlib import Path

import streamlit as st

import config

# ── Logging ────────────────────────────────────────────────────────────────
LOG_FILE = "ddr_generator.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("app")

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DDR Report Generator",
    page_icon="🏗️",
    layout="centered",
)

st.title("🏗️ DDR Report Generator")
st.caption("AI-powered Detailed Diagnostic Report from Inspection + Thermal PDFs")
st.divider()

# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")
    user_api_key = st.text_input(
        "Gemini API Key (Optional)",
        type="password",
        help="By default, this app uses an internal API key. If you encounter Rate Limit errors, paste your own free Gemini API key here to bypass them."
    )
    st.caption("Get your free key at [Google AI Studio](https://aistudio.google.com/)")

# ── Inputs ─────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    inspection_file = st.file_uploader("📋 Inspection Report PDF", type=["pdf"])
with col2:
    thermal_file = st.file_uploader("🌡️ Thermal Report PDF", type=["pdf"])

fresh_run = st.checkbox(
    "Force fresh run (ignore cached checkpoints)",
    value=False,
    help="Check this if you uploaded different PDFs or want Gemini to re-read everything.",
)

st.divider()
generate_btn = st.button(
    "⚡ Generate DDR Report", type="primary", use_container_width=True
)

# ── Checkpoint helper ──────────────────────────────────────────────────────

CHECKPOINT_FILES = [
    "inspection_data.json",
    "thermal_data.json",
    "merged_data.json",
    "ddr_content.json",
]

def _clear_checkpoints(ckpt_dir: Path):
    """Delete all checkpoint files so the pipeline runs from scratch."""
    import shutil
    temp_dir = Path("temp")
    if temp_dir.exists():
        shutil.rmtree(temp_dir, ignore_errors=True)
    logger.info("Entire temp directory deleted.")


# ── Generation ─────────────────────────────────────────────────────────────
if generate_btn:
    errors = []
    
    # Determine which API key to use
    active_api_key = user_api_key.strip() if user_api_key.strip() else config.GEMINI_API_KEY.strip()
    
    if not active_api_key:
        errors.append("No API key available. The internal key is missing and you did not provide one.")
    if not inspection_file:
        errors.append("Please upload the Inspection Report PDF.")
    if not thermal_file:
        errors.append("Please upload the Thermal Report PDF.")

    for e in errors:
        st.error(e)
    if errors:
        st.stop()

    # ── Write uploaded files to project input directory ───────────────────
    input_dir  = Path("input")
    output_dir = Path("output")
    temp_dir   = Path("temp")
    
    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)

    if fresh_run:
        _clear_checkpoints(temp_dir)
        st.info("🔄  Fresh run: all cached temp files cleared.")
        logger.info("Fresh run requested — temp cache completely cleared.")

    insp_path   = input_dir / "inspection_report.pdf"
    therm_path  = input_dir / "thermal_report.pdf"
    output_path = output_dir / "DDR_Report.docx"
    pages_dir   = temp_dir / "pages"
    images_dir  = temp_dir / "extracted_images"
    ckpt_dir    = temp_dir / "checkpoints"

    insp_path.write_bytes(inspection_file.read())
    therm_path.write_bytes(thermal_file.read())

    if output_path.exists():
        try:
            output_path.unlink()
        except OSError:
            st.error("❌ **Cannot run pipeline!**")
            st.error(f"The file `{output_path}` is currently open in another program (like Microsoft Word). Please close the document so the AI can write to it, then click Generate again.")
            st.stop()

    status = st.status("Running pipeline…", expanded=True)
    progress_bar = st.progress(5, text="Preparing workspace...")
    t0 = time.time()

    try:
        with status:

            # ── Phase 1: PDF Processing ────────────────────────────────
            progress_bar.progress(10, text="Phase 1: Processing PDFs and Extracting Images...")
            st.write("⚙️  Phase 1 — Processing PDFs…")
            t1 = time.time()

            from src.pdf_processor import (
                convert_pages_to_png,
                extract_images_from_pdf,
                extract_site_photos_from_pdf,
            )

            insp_pages  = convert_pages_to_png(insp_path,  "inspection", pages_dir, config.PAGE_DPI)
            therm_pages = convert_pages_to_png(therm_path, "thermal",    pages_dir, config.PAGE_DPI)
            therm_imgs  = extract_images_from_pdf(therm_path, "thermal",     images_dir)
            insp_imgs   = extract_site_photos_from_pdf(insp_path, "inspection", images_dir)

            therm_img_total = sum(len(v) for v in therm_imgs.values())
            insp_img_total  = sum(len(v) for v in insp_imgs.values())

            st.write(
                f"   → {len(insp_pages)} inspection pages, "
                f"{len(therm_pages)} thermal pages | "
                f"{therm_img_total} thermal images, "
                f"{insp_img_total} inspection photos  "
                f"({time.time() - t1:.1f}s)"
            )

            # ── Phase 2: AI Understanding ──────────────────────────────
            progress_bar.progress(30, text="Phase 2: AI Analyzing Inspection & Thermal Data (This may take 2-5 minutes)...")
            st.write("🤖  Phase 2 — AI Understanding (Rate-limited, please wait)…")
            t2 = time.time()

            from src.gemini_client     import init_gemini
            from src.inspection_parser import parse_inspection_report
            from src.thermal_parser    import parse_thermal_report

            model        = init_gemini(active_api_key, config.GEMINI_MODEL)
            insp_data    = parse_inspection_report(
                model, insp_pages, config.BATCH_SIZE, config.MAX_RETRIES, ckpt_dir
            )
            thermal_data = parse_thermal_report(
                model, therm_pages, config.BATCH_SIZE, config.MAX_RETRIES, ckpt_dir
            )

            st.write(
                f"   → {len(insp_data)} inspection findings, "
                f"{len(thermal_data)} thermal entries  "
                f"({time.time() - t2:.1f}s)"
            )

            # ── Phase 3: Merge ─────────────────────────────────────────
            progress_bar.progress(70, text="Phase 3: Merging Inspection & Thermal data...")
            st.write("🔀  Phase 3 — Merging inspection + thermal data…")
            t3 = time.time()

            from src.data_merger import merge_data
            merged = merge_data(
                model, insp_data, thermal_data, config.MAX_RETRIES, ckpt_dir
            )
            st.write(
                f"   → {len(merged)} areas merged  ({time.time() - t3:.1f}s)"
            )

            # ── Phase 4: Generate DDR Content ──────────────────────────
            progress_bar.progress(85, text="Phase 4: Synthesizing Final DDR Content...")
            st.write("✍️  Phase 4 — Generating DDR content…")
            t4 = time.time()

            from src.report_generator import generate_ddr_content
            ddr_content = generate_ddr_content(
                model, merged, config.MAX_RETRIES, ckpt_dir
            )
            st.write(f"   → DDR content ready  ({time.time() - t4:.1f}s)")

            # ── Phase 5: Build Word Document ───────────────────────────
            progress_bar.progress(95, text="Phase 5: Building Microsoft Word Document...")
            st.write("📄  Phase 5 — Building Word document…")
            t5 = time.time()

            from src.docx_builder import build_docx
            build_docx(
                ddr_content       = ddr_content,
                merged_data       = merged,
                extracted_images  = therm_imgs,
                inspection_images = insp_imgs,
                output_path       = output_path,
            )
            st.write(f"   → Document built  ({time.time() - t5:.1f}s)")

        total_time = time.time() - t0
        progress_bar.progress(100, text=f"Pipeline Completed in {total_time:.0f} seconds!")
        status.update(
            label=f"✅ Done!  (total: {total_time:.0f}s)", state="complete"
        )
        logger.info(f"Pipeline completed in {total_time:.1f}s")

        st.success("DDR Report generated successfully!")

        if output_path.exists() and output_path.stat().st_size > 0:
            st.download_button(
                label="📥 Download DDR_Report.docx",
                data=output_path.read_bytes(),
                file_name="DDR_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        else:
            st.error("Output file is empty — check the log for details.")

    except Exception as exc:
        progress_bar.progress(0, text="Pipeline Failed.")
        status.update(label="❌ Pipeline failed", state="error")
        st.error(f"Error: {exc}")
        st.info(
            f"Check `{LOG_FILE}` in the project root for the full traceback."
        )
        logger.exception("Pipeline error")
