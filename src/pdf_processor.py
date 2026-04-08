
import hashlib
import io
import logging
from pathlib import Path

import fitz          # PyMuPDF
from PIL import Image

logger = logging.getLogger(__name__)

# ── Tunables ───────────────────────────────────────────────────────────────
MIN_CROP_PIXELS  = 60 * 60        # ignore tiny icon-sized crops
MIN_BYTES        = 2_048          # 2 KB minimum embedded image
DPI_PAGES        = 150            # resolution for page-to-PNG (Gemini input)
DPI_CROP         = 200            # resolution for block-crop (better quality)

# Aspect-ratio / size guard for inspection site-photo extraction
_MIN_ASPECT = 0.30    # very tall portrait allowed (narrow staircase shot)
_MAX_ASPECT = 3.20    # wide panoramic allowed
_MIN_PX     = 300     # each dimension must be ≥ this many pixels
_MAX_PX     = 4000    # skip full-page renders already saved as page PNGs

# Fraction of page width occupied by images on thermal landscape pages.
# The Bosch GTC 400 C report uses ~55 % on the left for images.
THERMAL_IMG_FRACTION = 0.58


# ── Helpers ────────────────────────────────────────────────────────────────

def _is_likely_logo(width: int, height: int) -> bool:
    """Return True when dimensions suggest a logo, banner, or icon."""
    if width < _MIN_PX or height < _MIN_PX:
        return True
    if width > _MAX_PX or height > _MAX_PX:
        return True
    aspect = width / max(height, 1)
    return aspect < _MIN_ASPECT or aspect > _MAX_ASPECT


def _pil_to_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.convert("RGB").save(buf, "JPEG", quality=85)
    return buf.getvalue()


def _is_blank(img: Image.Image, threshold: float = 0.96) -> bool:
    """True when the image is mostly white (empty region)."""
    try:
        rgb    = img.convert("RGB")
        pixels = list(rgb.getdata())
        white  = sum(1 for r, g, b in pixels if r > 238 and g > 238 and b > 238)
        return white / max(len(pixels), 1) > threshold
    except Exception:
        return False


def _save_pil(img: Image.Image, out_path: Path) -> bool:
    """Save PIL image to disk and verify it can be reopened."""
    w, h = img.size
    if w < 300 or h < 300:
        return False
    try:
        img.convert("RGB").save(str(out_path), "JPEG", quality=90)
        # quick round-trip check
        Image.open(str(out_path)).verify()
        return True
    except Exception as exc:
        logger.warning(f"Could not save crop → {out_path.name}: {exc}")
        out_path.unlink(missing_ok=True)
        return False


def _save_bytes(raw: bytes, out_path: Path) -> bool:
    """Write raw image bytes and verify PIL can open it."""
    try:
        img = Image.open(io.BytesIO(raw))
        img.verify()
        w, h = img.size
        if w < 300 or h < 300:
            return False

        out_path.write_bytes(raw)
        return True
    except Exception as exc:
        logger.warning(f"Could not save image → {out_path.name}: {exc}")
        out_path.unlink(missing_ok=True)
        return False


# ── Public API ─────────────────────────────────────────────────────────────

def convert_pages_to_png(
    pdf_path: Path,
    prefix: str,
    pages_dir: Path,
    dpi: int = DPI_PAGES,
) -> list:
    """
    Render every PDF page as a PNG for Gemini vision calls.
    Returns sorted list of Path objects.
    """
    pages_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    page_paths = []
    mat = fitz.Matrix(dpi / 72, dpi / 72)

    for page_num in range(len(doc)):
        out_path = pages_dir / f"{prefix}_page_{page_num + 1:03d}.png"
        if not out_path.exists():
            pix = doc[page_num].get_pixmap(matrix=mat)
            pix.save(str(out_path))
        page_paths.append(out_path)
        logger.debug(f"  Page {page_num + 1}/{len(doc)} → {out_path.name}")

    doc.close()
    logger.info(f"Page PNGs ready: {len(page_paths)} from {pdf_path.name}")
    return page_paths


def extract_images_from_pdf(pdf_path: Path, prefix: str, images_dir: Path) -> dict:
    """
    Extract thermal-report images, one entry per page.

    Strategy (tried in order):
      Tier 1 – get_image_info(hashes=True) : best, per-page, content-dedup
      Tier 2 – block-based bbox crop        : works even for shared-resource PDFs
      Tier 3 – legacy get_images() + MD5   : last resort

    Returns
    -------
    {1-indexed page number: [Path, ...]}
    For each thermal page: [0] = thermal heatmap, [1] = regular site photo
    (top-to-bottom order within the image region).
    """
    images_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    page_images: dict = {i + 1: [] for i in range(len(doc))}

    total = _extract_with_image_info(doc, prefix, images_dir, page_images)

    if total == 0:
        logger.warning("Tier-1 (get_image_info) returned nothing → trying block-crop.")
        total = _extract_by_page_crop(doc, prefix, images_dir, page_images)

    if total == 0:
        logger.warning("Tier-2 (block-crop) returned nothing → legacy fallback.")
        total = _extract_legacy(doc, prefix, images_dir, page_images)

    page1_count = len(page_images.get(1, []))
    other_count = sum(len(v) for k, v in page_images.items() if k != 1)
    if page1_count > 4 and other_count == 0:
        logger.warning(
            f"All {page1_count} images still on page 1 — PDF uses shared "
            "resource dictionary. Per-area image mapping will be limited."
        )

    doc.close()
    pages_with = sum(1 for v in page_images.values() if v)
    logger.info(f"Extracted {total} images from {pdf_path.name} across {pages_with} pages.")
    return page_images


def extract_site_photos_from_pdf(pdf_path: Path, prefix: str, images_dir: Path) -> dict:
    """
    Extract site photos from inspection PDFs.

    Uses get_images(full=True) to get original high-res blocks, then filters:
      • 150 ≤ w,h ≤ 4000 px
      • 0.30 ≤ aspect ≤ 3.20
      • MD5 dedup across the PDF

    Images per page are sorted by their bbox Y-position so the global flat
    list preserves the "Photo 1, Photo 2 …" numbering from the report.

    Returns {1-indexed page: [Path, ...]}
    """
    images_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    page_images: dict = {i + 1: [] for i in range(len(doc))}
    seen_hashes: set  = set()
    total = 0

    for page_num in range(len(doc)):
        page     = doc[page_num]
        page_key = page_num + 1

        # Collect (y_position, xref) pairs so we can sort by Y
        img_with_pos = []
        for img in page.get_images(full=True):
            xref = img[0]
            # Get bounding box on page for sort key
            try:
                rects = page.get_image_rects(xref)
                y0    = rects[0].y0 if rects else 9999
            except Exception:
                y0 = 9999
            img_with_pos.append((y0, xref))

        img_with_pos.sort(key=lambda t: t[0])  # top-to-bottom

        img_idx = 0
        for _y, xref in img_with_pos:
            try:
                base = doc.extract_image(xref)
                if not base:
                    continue
                w, h      = base["width"], base["height"]
                img_bytes = base["image"]
                ext       = base.get("ext", "jpg")

                if _is_likely_logo(w, h):
                    logger.debug(f"  Skip logo/icon page={page_key} {w}×{h}")
                    continue
                if len(img_bytes) < MIN_BYTES:
                    continue

                md5 = hashlib.md5(img_bytes).hexdigest()
                if md5 in seen_hashes:
                    logger.debug(f"  Duplicate skipped page={page_key}")
                    continue
                seen_hashes.add(md5)

                out = images_dir / f"{prefix}_page{page_key:03d}_img{img_idx}.{ext}"
                if _save_bytes(img_bytes, out):
                    page_images[page_key].append(out)
                    total   += 1
                    img_idx += 1
                    logger.debug(f"  Saved inspection photo page={page_key} {w}×{h} → {out.name}")

            except Exception as exc:
                logger.debug(f"  xref={xref} page={page_key}: {exc}")

    doc.close()
    pages_with = sum(1 for v in page_images.values() if v)
    logger.info(
        f"Extracted {total} inspection site photos from {pdf_path.name} "
        f"across {pages_with} pages (logos/banners filtered)."
    )
    return page_images


# ── Private helpers ────────────────────────────────────────────────────────

def _extract_with_image_info(
    doc, prefix: str, images_dir: Path, page_images: dict
) -> int:
    """
    Tier 1: get_image_info(hashes=True).
    Dedup is PER-PAGE so different thermal pages that share an xref still
    each get their own saved copy.
    """
    total = 0

    for page_num in range(len(doc)):
        page     = doc[page_num]
        page_key = page_num + 1
        seen_this_page: set = set()

        try:
            infos = page.get_image_info(hashes=True)
        except (AttributeError, Exception) as exc:
            logger.debug(f"get_image_info unavailable: {exc}")
            return 0  # trigger Tier 2

        # Sort top-to-bottom by Y of bbox
        infos_sorted = sorted(infos, key=lambda x: x.get("bbox", (0, 0, 0, 0))[1])

        for img_idx, info in enumerate(infos_sorted):
            w = info.get("width", 0)
            h = info.get("height", 0)
            if w < 100 or h < 100:
                continue

            digest = info.get("digest")
            if digest:
                if digest in seen_this_page:
                    continue
                seen_this_page.add(digest)

            xref = info.get("xref", 0)
            if not xref:
                continue

            saved = _save_image_xref(doc, xref, page_key, img_idx, prefix, images_dir)
            if saved:
                page_images[page_key].append(saved)
                total += 1

    return total


def _extract_by_page_crop(
    doc, prefix: str, images_dir: Path, page_images: dict,
    dpi: int = DPI_CROP,
) -> int:
    """
    Tier 2: Render each page and crop image bounding boxes.

    KEY FIX (v5): For LANDSCAPE pages (thermal reports), when no image blocks
    are found via rawdict, we crop the LEFT portion (image side) of the page,
    then split THAT vertically into thermal heatmap (top) and site photo (bottom).

    Old code cropped the full page width top/bottom → heatmap & site photo were
    mixed with the right-side data table, producing garbled crops.
    """
    seen_hashes: set = set()
    total = 0
    mat   = fitz.Matrix(dpi / 72, dpi / 72)

    for page_num in range(len(doc)):
        page     = doc[page_num]
        page_key = page_num + 1

        pix      = page.get_pixmap(matrix=mat)
        pil_page = Image.open(io.BytesIO(pix.tobytes("png")))
        pw, ph   = pil_page.size
        is_landscape = pw > ph  # thermal reports are landscape

        # Try block-based detection first
        try:
            blocks = page.get_text(
                "rawdict", flags=fitz.TEXT_PRESERVE_IMAGES
            ).get("blocks", [])
        except Exception:
            blocks = []

        image_blocks = sorted(
            [b for b in blocks if b.get("type") == 1],
            key=lambda b: b["bbox"][1],
        )

        crops_saved = []

        if image_blocks:
            px_w = page.rect.width
            px_h = page.rect.height
            sx   = pw / px_w
            sy   = ph / px_h

            for img_idx, block in enumerate(image_blocks):
                x0, y0, x1, y1 = block["bbox"]
                px0 = max(0, int(x0 * sx))
                py0 = max(0, int(y0 * sy))
                px1 = min(pw, int(x1 * sx))
                py1 = min(ph, int(y1 * sy))

                if (px1 - px0) * (py1 - py0) < MIN_CROP_PIXELS:
                    continue

                crop      = pil_page.crop((px0, py0, px1, py1))
                img_bytes = _pil_to_bytes(crop)
                h         = hashlib.md5(img_bytes).hexdigest()
                if h in seen_hashes:
                    continue
                seen_hashes.add(h)

                out = images_dir / f"{prefix}_page{page_key:03d}_img{img_idx}.jpg"
                if _save_pil(crop, out):
                    crops_saved.append(out)
                    total += 1

        if not crops_saved:
            # ── Fallback split ────────────────────────────────────────────
            if is_landscape:
                # Thermal report layout: left ~58 % = images, right = data table
                img_right = int(pw * THERMAL_IMG_FRACTION)
                left_half = pil_page.crop((0, 0, img_right, ph))
                lw, lh    = left_half.size
                # Give a small top-margin so we skip any caption/header text
                top_margin = int(lh * 0.05)
                halves = [
                    ("img0", left_half.crop((0, top_margin, lw, lh // 2))),  # thermal heatmap
                    ("img1", left_half.crop((0, lh // 2,   lw, lh))),        # site photo
                ]
            else:
                # Portrait: straight top / bottom split
                halves = [
                    ("img0", pil_page.crop((0, 0,      pw, ph // 2))),
                    ("img1", pil_page.crop((0, ph // 2, pw, ph))),
                ]

            for tag, crop in halves:
                if _is_blank(crop):
                    continue
                img_bytes = _pil_to_bytes(crop)
                h         = hashlib.md5(img_bytes).hexdigest()
                if h in seen_hashes:
                    continue
                seen_hashes.add(h)

                out = images_dir / f"{prefix}_page{page_key:03d}_{tag}.jpg"
                if _save_pil(crop, out):
                    crops_saved.append(out)
                    total += 1

        page_images[page_key].extend(crops_saved)
        if crops_saved:
            logger.debug(f"  Page {page_key}: {len(crops_saved)} crop(s) saved")

    return total


def _extract_legacy(
    doc, prefix: str, images_dir: Path, page_images: dict
) -> int:
    """Tier 3: get_images() + MD5 dedup — last resort."""
    seen_hashes: set = set()
    total = 0

    for page_num in range(len(doc)):
        page     = doc[page_num]
        page_key = page_num + 1

        for img_idx, img_info in enumerate(page.get_images(full=True)):
            xref = img_info[0]
            try:
                base = doc.extract_image(xref)
                if not base:
                    continue
                img_bytes = base["image"]
                ext       = base.get("ext", "jpg")

                if len(img_bytes) < MIN_BYTES:
                    continue

                h = hashlib.md5(img_bytes).hexdigest()
                if h in seen_hashes:
                    continue
                seen_hashes.add(h)

                out = images_dir / f"{prefix}_page{page_key:03d}_img{img_idx}.{ext}"
                if _save_bytes(img_bytes, out):
                    page_images[page_key].append(out)
                    total += 1

            except Exception as exc:
                logger.warning(f"  Legacy page={page_key} img={img_idx}: {exc}")

    return total


def _save_image_xref(
    doc, xref: int, page_key: int, img_idx: int, prefix: str, images_dir: Path
) -> "Path | None":
    """Extract one image by xref, validate, and write to disk."""
    try:
        base = doc.extract_image(xref)
        if not base:
            return None
        img_bytes = base["image"]
        ext       = base.get("ext", "jpg")

        if len(img_bytes) < MIN_BYTES:
            return None

        out = images_dir / f"{prefix}_page{page_key:03d}_img{img_idx}.{ext}"
        return out if _save_bytes(img_bytes, out) else None

    except Exception as exc:
        logger.warning(f"  xref={xref} page={page_key}: {exc}")
        return None
