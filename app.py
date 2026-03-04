import os
import shutil
import tempfile
import zipfile
from io import BytesIO

import imagehash
import streamlit as st
from PIL import Image, ImageDraw

# ============================================================
# BASIC IMAGE HELPERS
# ============================================================

def get_image_hash(img_bytes: bytes):
    return imagehash.phash(Image.open(BytesIO(img_bytes)))


def to_png(img_bytes: bytes) -> bytes:
    buf = BytesIO()
    Image.open(BytesIO(img_bytes)).convert("RGBA").save(buf, format="PNG")
    return buf.getvalue()


def thumbnail_bytes(img_bytes: bytes, size: tuple = (220, 220)) -> bytes:
    img = Image.open(BytesIO(img_bytes)).convert("RGBA")
    img.thumbnail(size, Image.LANCZOS)
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def annotate_bbox(
    img_bytes: bytes,
    x1: int, y1: int, x2: int, y2: int,
    color: str = "#ff3300",
    line_width: int = 4,
) -> bytes:
    """Draw a coloured rectangle over the detected region."""
    img = Image.open(BytesIO(img_bytes)).convert("RGBA")
    draw = ImageDraw.Draw(img)
    for i in range(line_width):
        draw.rectangle([x1 - i, y1 - i, x2 + i, y2 + i], outline=color)
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ============================================================
# OCR  (EasyOCR, cached so model loads once)
# ============================================================

@st.cache_resource(show_spinner="Loading OCR engine…")
def _ocr_reader():
    import easyocr  # noqa: lazy import
    return easyocr.Reader(["en"], gpu=False)


def detect_by_ocr(
    img_bytes: bytes,
    search_text: str = "LTIMindtree",
    expand_left_ratio: float = 0.9,
    v_padding_ratio: float = 0.15,
) -> dict | None:
    """
    Find `search_text` in `img_bytes` using EasyOCR.
    Because the LTIMindtree logo has a circular icon immediately to the *left*
    of the text, we expand the bounding-box leftward by `expand_left_ratio × text_width`.

    Returns {"x1","y1","x2","y2","text","confidence"} or None.
    """
    import numpy as np  # noqa

    reader = _ocr_reader()
    img = Image.open(BytesIO(img_bytes)).convert("RGB")
    results = reader.readtext(np.array(img))

    best = None
    for bbox, text, conf in results:
        if search_text.lower() in text.lower():
            xs = [p[0] for p in bbox]
            ys = [p[1] for p in bbox]
            rx1, ry1 = int(min(xs)), int(min(ys))
            rx2, ry2 = int(max(xs)), int(max(ys))
            if best is None or conf > best["confidence"]:
                best = {"x1": rx1, "y1": ry1, "x2": rx2, "y2": ry2,
                        "text": text, "confidence": conf}

    if best is None:
        return None

    w = best["x2"] - best["x1"]
    h = best["y2"] - best["y1"]
    pad_v = int(h * v_padding_ratio)
    pad_left = int(w * expand_left_ratio)

    img_w, img_h = img.size
    best["x1"] = max(0, best["x1"] - pad_left)
    best["y1"] = max(0, best["y1"] - pad_v)
    best["y2"] = min(img_h, best["y2"] + pad_v)
    return best


# ============================================================
# SCREENSHOT PREPROCESSING
# ============================================================

def autocrop_screenshot(img_bytes: bytes, tolerance: int = 30) -> bytes:
    """
    Remove uniform solid-colour borders/padding from a screenshot.
    Works by repeatedly checking whether the outermost row/column is
    within `tolerance` of the image corner pixel colour.
    Helps when the user pastes a screenshot that has desktop background
    or application chrome around the actual logo.
    """
    import numpy as np

    img  = Image.open(BytesIO(img_bytes)).convert("RGB")
    arr  = np.array(img)
    h, w = arr.shape[:2]

    # Sample the corner colour as the background guess
    bg = arr[0, 0].astype(int)

    def _is_border_row(row):  return np.all(np.abs(arr[row].astype(int) - bg) <= tolerance)
    def _is_border_col(col):  return np.all(np.abs(arr[:, col].astype(int) - bg) <= tolerance)

    top    = 0
    while top < h and _is_border_row(top):          top    += 1
    bottom = h - 1
    while bottom > top and _is_border_row(bottom):  bottom -= 1
    left   = 0
    while left < w and _is_border_col(left):        left   += 1
    right  = w - 1
    while right > left and _is_border_col(right):   right  -= 1

    if top >= bottom or left >= right:   # nothing to crop
        return img_bytes

    cropped = img.crop((left, top, right + 1, bottom + 1))
    buf = BytesIO()
    cropped.save(buf, format="PNG")
    return buf.getvalue()


def _normalise_for_matching(gray_img):  # noqa: numpy ndarray → ndarray
    """Apply CLAHE contrast normalisation so screenshots and embedded images match better."""
    import cv2
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    return clahe.apply(gray_img)


# ============================================================
# TEMPLATE MATCHING  (OpenCV multi-scale)
# ============================================================

def detect_by_template(
    img_bytes: bytes,
    template_bytes: bytes,
    min_score: float = 0.45,
) -> dict | None:
    """
    Multi-scale OpenCV template matching.
    Automatically crops and normalises the template so screenshots
    with padding/background work correctly.
    Tries scales 40 %–160 % so size differences are handled.
    Returns {"x1","y1","x2","y2","score"} or None.
    """
    import cv2
    import numpy as np

    # Preprocess template: autocrop then normalise
    template_bytes = autocrop_screenshot(template_bytes)

    img  = _normalise_for_matching(
        cv2.cvtColor(np.array(Image.open(BytesIO(img_bytes)).convert("RGB")),      cv2.COLOR_RGB2GRAY)
    )
    tmpl = _normalise_for_matching(
        cv2.cvtColor(np.array(Image.open(BytesIO(template_bytes)).convert("RGB")), cv2.COLOR_RGB2GRAY)
    )
    ih, iw = img.shape
    th, tw = tmpl.shape

    best_score, best_loc, best_scale = -1.0, None, 1.0

    for scale in [s / 100 for s in range(40, 165, 5)]:
        new_w = max(1, int(tw * scale))
        new_h = max(1, int(th * scale))
        if new_w > iw or new_h > ih:
            continue
        resized = cv2.resize(tmpl, (new_w, new_h), interpolation=cv2.INTER_AREA)
        res = cv2.matchTemplate(img, resized, cv2.TM_CCOEFF_NORMED)
        _, score, _, loc = cv2.minMaxLoc(res)
        if score > best_score:
            best_score, best_loc, best_scale = score, loc, scale

    if best_score < min_score or best_loc is None:
        return None

    bw = int(tw * best_scale)
    bh = int(th * best_scale)
    x1, y1 = best_loc
    return {"x1": x1, "y1": y1, "x2": x1 + bw, "y2": y1 + bh, "score": float(best_score)}


# ============================================================
# REGION REPLACE  (paste new logo into detected bbox)
# ============================================================

def replace_region_in_image(
    img_bytes: bytes,
    x1: int, y1: int, x2: int, y2: int,
    new_logo_bytes: bytes,
) -> bytes:
    """
    Paste `new_logo_bytes` (resized to the bbox dimensions) into `img_bytes`.
    The background under the new logo is first cleared to the surrounding
    average colour so there is no bleed-through from the old logo.
    """
    img     = Image.open(BytesIO(img_bytes)).convert("RGBA")
    new_w   = max(1, x2 - x1)
    new_h   = max(1, y2 - y1)
    new_logo = Image.open(BytesIO(new_logo_bytes)).convert("RGBA")
    new_logo = new_logo.resize((new_w, new_h), Image.LANCZOS)

    # Clear old-logo area to white (or transparent) before pasting
    clear = Image.new("RGBA", (new_w, new_h), (255, 255, 255, 255))
    img.paste(clear, (x1, y1))
    img.paste(new_logo, (x1, y1), new_logo)

    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ============================================================
# DOCUMENT EXTRACTION  (returns zip_path / xref for re-embedding)
# ============================================================

def _collect_pics(shapes):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    pics = []
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            pics.append(shape)
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            pics.extend(_collect_pics(shape.shapes))
    return pics


def extract_images_pptx(pptx_bytes: bytes) -> list[dict]:
    """
    Scans ppt/media/* directly (gives zip_path), then uses python-pptx
    to annotate where each image appears (master / layout / slide N).
    """
    from pptx import Presentation

    # ── step 1: collect all media files with their zip paths ──
    media: dict[str, dict] = {}   # hash_str → entry
    with zipfile.ZipFile(BytesIO(pptx_bytes)) as z:
        for name in z.namelist():
            if name.lower().startswith("ppt/media/"):
                try:
                    data = z.read(name)
                    key  = str(get_image_hash(data))
                    if key not in media:
                        media[key] = {
                            "bytes":    data,
                            "zip_path": name,
                            "locations": [],
                            "hash_str": key,
                        }
                except Exception:
                    pass

    # ── step 2: annotate locations via python-pptx ──
    prs = Presentation(BytesIO(pptx_bytes))

    def annotate(shapes, location: str):
        for pic in _collect_pics(shapes):
            try:
                key = str(get_image_hash(pic.image.blob))
                if key in media and location not in media[key]["locations"]:
                    media[key]["locations"].append(location)
            except Exception:
                pass

    annotate(prs.slide_master.shapes, "Slide Master")
    for layout in prs.slide_master.slide_layouts:
        annotate(layout.shapes, f"Layout — {layout.name}")
    for i, slide in enumerate(prs.slides):
        annotate(slide.shapes, f"Slide {i + 1}")

    # Fill in location for any image not referenced via shapes
    for entry in media.values():
        if not entry["locations"]:
            entry["locations"].append("Media file (not linked to a shape)")

    return list(media.values())


def extract_images_docx(docx_bytes: bytes) -> list[dict]:
    seen: dict[str, dict] = {}
    with zipfile.ZipFile(BytesIO(docx_bytes)) as z:
        for name in z.namelist():
            if name.lower().startswith("word/media/"):
                try:
                    data = z.read(name)
                    key  = str(get_image_hash(data))
                    if key not in seen:
                        seen[key] = {
                            "bytes":    data,
                            "zip_path": name,
                            "locations": ["Document body"],
                            "hash_str": key,
                        }
                except Exception:
                    pass
    return list(seen.values())


def extract_images_pdf(pdf_bytes: bytes) -> list[dict]:
    import fitz
    seen: dict[str, dict] = {}
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for page_num, page in enumerate(doc):
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                try:
                    data = doc.extract_image(xref)["image"]
                    key  = str(get_image_hash(data))
                    loc  = f"Page {page_num + 1}"
                    if key not in seen:
                        seen[key] = {
                            "bytes":    data,
                            "xref":     xref,
                            "locations": [loc],
                            "hash_str": key,
                        }
                    elif loc not in seen[key]["locations"]:
                        seen[key]["locations"].append(loc)
                except Exception:
                    pass
    finally:
        doc.close()
    return list(seen.values())


def extract_images_any(file_bytes: bytes, filename: str) -> list[dict]:
    ext = os.path.splitext(filename.lower())[1]
    if ext == ".pptx":
        return extract_images_pptx(file_bytes)
    if ext in (".docx", ".doc"):
        return extract_images_docx(file_bytes)
    if ext == ".pdf":
        return extract_images_pdf(file_bytes)
    raise ValueError(f"Unsupported: {ext}")


# ============================================================
# RE-EMBED A (modified) IMAGE BACK INTO THE DOCUMENT
# ============================================================

def _zip_replace(zip_bytes: bytes, replacements: dict[str, bytes]) -> bytes:
    """Replace specific paths inside a zip with new bytes."""
    buf = BytesIO()
    with zipfile.ZipFile(BytesIO(zip_bytes)) as zin, \
         zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in replacements:
                zout.writestr(item.filename, replacements[item.filename])
            else:
                zout.writestr(item.filename, zin.read(item.filename))
    return buf.getvalue()


def embed_image_in_document(
    file_bytes: bytes,
    filename: str,
    img_meta: dict,
    new_img_bytes: bytes,
) -> bytes:
    """
    Put a (possibly region-modified) image back into the document.
    Uses zip_path for pptx/docx, xref for pdf.
    """
    ext = os.path.splitext(filename.lower())[1]

    if ext in (".pptx", ".docx", ".doc"):
        return _zip_replace(file_bytes, {img_meta["zip_path"]: new_img_bytes})

    if ext == ".pdf":
        import fitz
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        try:
            doc.replace_image(img_meta["xref"], stream=new_img_bytes)
            return doc.tobytes(garbage=4, deflate=True)
        finally:
            doc.close()

    raise ValueError(f"Unsupported: {ext}")


# ============================================================
# STANDARD HASH-BASED FULL-IMAGE REPLACEMENT (existing flow)
# ============================================================

def replace_logo_pptx(pptx_bytes, old_logo_bytes, new_logo_bytes, threshold):
    from pptx import Presentation

    old_hash = get_image_hash(old_logo_bytes)
    new_png  = to_png(new_logo_bytes)
    temp_dir = tempfile.mkdtemp()
    try:
        inp = os.path.join(temp_dir, "in.pptx")
        out = os.path.join(temp_dir, "out.pptx")
        with open(inp, "wb") as f:
            f.write(pptx_bytes)

        prs   = Presentation(inp)
        count = 0

        def process(tree):
            nonlocal count
            for pic in _collect_pics(tree):
                try:
                    blob = pic.image.blob
                except Exception:
                    continue
                if old_hash - get_image_hash(blob) <= threshold:
                    l, t, w, h = pic.left, pic.top, pic.width, pic.height
                    pic._element.getparent().remove(pic._element)
                    tree.add_picture(BytesIO(new_png), l, t, w, h)
                    count += 1

        process(prs.slide_master.shapes)
        for layout in prs.slide_master.slide_layouts:
            process(layout.shapes)
        for slide in prs.slides:
            process(slide.shapes)

        prs.save(out)
        return open(out, "rb").read(), count
    finally:
        shutil.rmtree(temp_dir)


def replace_logo_docx(docx_bytes, old_logo_bytes, new_logo_bytes, threshold):
    old_hash = get_image_hash(old_logo_bytes)
    new_png  = to_png(new_logo_bytes)
    buf, count = BytesIO(), 0
    with zipfile.ZipFile(BytesIO(docx_bytes)) as zin, \
         zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.lower().startswith("word/media/"):
                try:
                    if old_hash - get_image_hash(data) <= threshold:
                        zout.writestr(item.filename, new_png)
                        count += 1
                        continue
                except Exception:
                    pass
            zout.writestr(item.filename, data)
    return buf.getvalue(), count


def replace_logo_pdf(pdf_bytes, old_logo_bytes, new_logo_bytes, threshold):
    import fitz
    old_hash = get_image_hash(old_logo_bytes)
    new_png  = to_png(new_logo_bytes)
    count, done = 0, set()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for page in doc:
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                if xref in done:
                    continue
                try:
                    data = doc.extract_image(xref)["image"]
                    if old_hash - get_image_hash(data) <= threshold:
                        doc.replace_image(xref, stream=new_png)
                        done.add(xref)
                        count += 1
                except Exception:
                    pass
        return doc.tobytes(garbage=4, deflate=True), count
    finally:
        doc.close()


def replace_logo_any(file_bytes, filename, old_logo_bytes, new_logo_bytes, threshold):
    ext = os.path.splitext(filename.lower())[1]
    if ext == ".pptx":
        return replace_logo_pptx(file_bytes, old_logo_bytes, new_logo_bytes, threshold)
    if ext in (".docx", ".doc"):
        return replace_logo_docx(file_bytes, old_logo_bytes, new_logo_bytes, threshold)
    if ext == ".pdf":
        return replace_logo_pdf(file_bytes, old_logo_bytes, new_logo_bytes, threshold)
    raise ValueError(f"Unsupported: {ext}")


def replace_multiple_logos_any(
    file_bytes: bytes,
    filename: str,
    old_logos_bytes: list[bytes],   # one entry per selected logo
    new_logo_bytes: bytes,
    threshold: int = 8,
) -> tuple[bytes, int]:
    """
    Replace all logos in `old_logos_bytes` inside a single document in one go.
    Iterates through each reference logo and applies hash-based replacement in turn,
    accumulating changes so the output contains all replacements.
    Returns (updated_file_bytes, total_replacements).
    """
    data  = file_bytes
    total = 0
    for old_bytes in old_logos_bytes:
        data, n = replace_logo_any(data, filename, old_bytes, new_logo_bytes, threshold)
        total += n
    return data, total


def process_zip(zip_bytes, old_logo_bytes, new_logo_bytes, threshold):
    """Hash-based whole-image replacement across all files in a ZIP."""
    buf, total, files = BytesIO(), 0, 0
    with zipfile.ZipFile(BytesIO(zip_bytes)) as zin, \
         zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            ext  = os.path.splitext(item.filename.lower())[1]
            if ext in (".pptx", ".docx", ".doc", ".pdf"):
                try:
                    updated, n = replace_logo_any(data, item.filename, old_logo_bytes, new_logo_bytes, threshold)
                    zout.writestr(item.filename, updated)
                    total += n
                    files += 1
                    continue
                except Exception:
                    pass
            zout.writestr(item.filename, data)
    return buf.getvalue(), total, files


def process_zip_region(
    zip_outer_bytes: bytes,
    detect_fn,          # callable: (img_bytes: bytes) -> dict|None  {x1,y1,x2,y2}
    new_logo_bytes: bytes,
    progress_cb=None,   # optional callable(msg: str) for live status updates
) -> tuple[bytes, int, int, int]:
    """
    Fully automated region-replacement across every supported document in a ZIP.

    For each file:
      1. Extract all embedded images.
      2. Run detect_fn on each image.
      3. For every image where a region is found, paint the new logo over it
         and re-embed it back into that document.
      4. If multiple distinct images in the same file all match, all are handled.

    Returns (output_zip_bytes, files_processed, regions_replaced, files_with_no_match).
    """
    DOC_EXTS = {".pptx", ".docx", ".doc", ".pdf"}
    out_buf          = BytesIO()
    files_processed  = 0
    regions_replaced = 0
    files_no_match   = 0

    with zipfile.ZipFile(BytesIO(zip_outer_bytes)) as zin, \
         zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:

        all_items = zin.infolist()
        for idx, item in enumerate(all_items):
            data = zin.read(item.filename)
            ext  = os.path.splitext(item.filename.lower())[1]

            if ext not in DOC_EXTS:
                zout.writestr(item.filename, data)
                continue

            if progress_cb:
                short = os.path.basename(item.filename)
                progress_cb(f"[{idx+1}/{len(all_items)}] Processing {short}…")

            try:
                images    = extract_images_any(data, item.filename)
                file_data = data
                found_in_file = 0

                for img_meta in images:
                    try:
                        bbox = detect_fn(img_meta["bytes"])
                    except Exception:
                        bbox = None

                    if bbox is None:
                        continue

                    modified_img       = replace_region_in_image(
                        img_meta["bytes"],
                        bbox["x1"], bbox["y1"], bbox["x2"], bbox["y2"],
                        new_logo_bytes,
                    )
                    img_meta["bytes"]  = modified_img   # update for chained calls
                    file_data          = embed_image_in_document(
                        file_data, item.filename, img_meta, modified_img
                    )
                    found_in_file     += 1
                    regions_replaced  += 1

                if found_in_file == 0:
                    files_no_match += 1
                    if progress_cb:
                        progress_cb(f"  ↳ No logo found in {os.path.basename(item.filename)}")
                else:
                    if progress_cb:
                        progress_cb(
                            f"  ↳ Replaced {found_in_file} region(s) "
                            f"in {os.path.basename(item.filename)}"
                        )

                zout.writestr(item.filename, file_data)
                files_processed += 1

            except Exception as e:
                if progress_cb:
                    progress_cb(f"  ↳ ERROR on {os.path.basename(item.filename)}: {e}")
                zout.writestr(item.filename, data)   # pass through unchanged

    return out_buf.getvalue(), files_processed, regions_replaced, files_no_match


def process_zip_by_refs(
    zip_bytes: bytes,
    ref_images: list[bytes],   # reference logo images selected from the sample scan
    new_logo_bytes: bytes,
    threshold: int = 8,        # perceptual hash distance tolerance
    progress_cb=None,
) -> tuple[bytes, int, int]:
    """
    Replace every embedded image across all documents in a ZIP that perceptually
    matches ANY of the supplied reference images.

    ref_images  – list of raw image bytes extracted from the user-selected logos
    threshold   – max pHash distance to count as a match (0=exact, 20=very loose)

    Returns (out_zip_bytes, files_processed, total_replacements).
    """
    ref_hashes = [get_image_hash(b) for b in ref_images]
    out_buf      = BytesIO()
    files_proc   = 0
    total_repl   = 0
    DOC_EXTS     = {".pptx", ".docx", ".doc", ".pdf"}

    with zipfile.ZipFile(BytesIO(zip_bytes)) as zin, \
         zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:

        all_items = zin.infolist()
        for idx, item in enumerate(all_items):
            data = zin.read(item.filename)
            ext  = os.path.splitext(item.filename.lower())[1]

            if ext not in DOC_EXTS:
                zout.writestr(item.filename, data)
                continue

            short = os.path.basename(item.filename)
            if progress_cb:
                progress_cb(f"[{idx+1}/{len(all_items)}] Processing {short}…")

            try:
                images    = extract_images_any(data, item.filename)
                file_data = data
                count     = 0

                for img_meta in images:
                    try:
                        h = get_image_hash(img_meta["bytes"])
                    except Exception:
                        continue

                    if any(abs(h - rh) <= threshold for rh in ref_hashes):
                        file_data = embed_image_in_document(
                            file_data, item.filename, img_meta,
                            to_png(new_logo_bytes),
                        )
                        count      += 1
                        total_repl += 1

                if progress_cb:
                    if count:
                        progress_cb(f"  ↳ Replaced {count} logo image(s)")
                    else:
                        progress_cb(f"  ↳ No matching logos found")

                zout.writestr(item.filename, file_data)
                files_proc += 1

            except Exception as e:
                if progress_cb:
                    progress_cb(f"  ↳ ERROR: {e}")
                zout.writestr(item.filename, data)

    return out_buf.getvalue(), files_proc, total_repl


# ============================================================
# SESSION STATE
# ============================================================

_defaults = {
    "images":          [],
    "selected":        None,
    "doc_cache":       None,
    "bbox":            None,    # detected region dict for intra-image mode
    "drill_active":    False,   # whether the "Find inside image" panel is open
    # ZIP guided flow
    "zip_sample_imgs": [],      # images scanned from the chosen sample file
    "zip_ref_indices": [],      # indices of logo images selected by the user
    # single-file multi-select
    "selected_set":    set(),   # indices of logo images marked for bulk replacement
}
for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ============================================================
# UI HELPERS
# ============================================================

MIME = {
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".doc":  "application/msword",
    ".pdf":  "application/pdf",
}


def _download_doc(label: str, data: bytes, stem: str, ext: str):
    st.download_button(
        label, data,
        file_name=f"{stem}_updated{ext}",
        mime=MIME.get(ext, "application/octet-stream"),
    )


# ============================================================
# PAGE
# ============================================================

st.set_page_config(page_title="Document Logo Replacer", page_icon="🖼️", layout="wide")
st.title("🖼️ Document Logo Replacer")
st.caption(
    "Supports PPTX · DOCX · PDF · ZIP batch. "
    "Finds logos on slide masters and layouts. "
    "Can also detect logos **baked into** a template background image using OCR or template matching."
)
st.divider()

# ── Upload ────────────────────────────────────────────────────────────────────

doc_file = st.file_uploader(
    "Upload document",
    type=["pptx", "docx", "doc", "pdf", "zip"],
    help="Single file or a ZIP containing multiple documents.",
)
is_zip = doc_file is not None and doc_file.name.lower().endswith(".zip")

if doc_file:
    cache = st.session_state.doc_cache
    if cache is None or cache[0] != doc_file.name:
        doc_file.seek(0)
        st.session_state.doc_cache        = (doc_file.name, doc_file.read())
        st.session_state.images           = []
        st.session_state.selected         = None
        st.session_state.selected_set     = set()
        st.session_state.bbox             = None
        st.session_state.drill_active     = False
        st.session_state.zip_sample_imgs  = []
        st.session_state.zip_ref_indices  = []

# ── ZIP flow ──────────────────────────────────────────────────────────────────

if is_zip:
    name_zip, zbytes_zip = st.session_state.doc_cache
    stem_zip = os.path.splitext(name_zip)[0]

    # Count supported files in the ZIP for display
    with zipfile.ZipFile(BytesIO(zbytes_zip)) as _z:
        _supported = [
            i.filename for i in _z.infolist()
            if os.path.splitext(i.filename.lower())[1] in (".pptx",".docx",".doc",".pdf")
        ]
    st.info(
        f"ZIP contains **{len(_supported)}** supported file(s): "
        + ", ".join(os.path.basename(f) for f in _supported[:8])
        + (" …" if len(_supported) > 8 else "")
    )

    zip_tab_hash, zip_tab_guided, zip_tab_region = st.tabs([
        "🔄 Quick — upload old logo file",
        "🎯 Guided — scan sample & select logos",
        "🔬 Region detect — logos baked into backgrounds",
    ])

    # ── ZIP Tab A: hash-based ─────────────────────────────────────────────────
    with zip_tab_hash:
        st.markdown(
            "Upload the **old logo** as a reference image. "
            "Every embedded image across all files that matches it will be swapped out wholesale."
        )
        zh_c1, zh_c2 = st.columns(2)
        with zh_c1:
            old_logo_hash = st.file_uploader(
                "Old logo (reference)", type=["png","jpg","jpeg","bmp","webp"], key="zip_hash_old"
            )
            if old_logo_hash:
                st.image(old_logo_hash, use_container_width=True)
        with zh_c2:
            new_logo_hash = st.file_uploader(
                "New logo", type=["png","jpg","jpeg","bmp","webp"], key="zip_hash_new"
            )
            if new_logo_hash:
                st.image(new_logo_hash, use_container_width=True)

        thr_hash = st.slider("Match sensitivity", 0, 20, 5, key="zip_hash_thresh")

        if st.button("Replace Logos (whole-image) in all files", type="primary",
                     disabled=not (old_logo_hash and new_logo_hash), key="btn_zip_hash"):
            with st.spinner("Processing ZIP…"):
                out_zip, total, files = process_zip(
                    zbytes_zip, old_logo_hash.read(), new_logo_hash.read(), thr_hash
                )
            if files == 0:
                st.warning("No supported files found in the ZIP.")
            elif total == 0:
                st.warning(f"Processed {files} file(s) but no matching logos found. Try raising sensitivity.")
            else:
                st.success(f"Replaced **{total}** logo(s) across **{files}** file(s).")
            st.download_button(
                "⬇️ Download updated ZIP", out_zip,
                file_name=stem_zip + "_updated.zip", mime="application/zip"
            )

    # ── ZIP Tab B: guided scan + multi-select ────────────────────────────────
    with zip_tab_guided:
        st.markdown(
            "Scan one file from the ZIP to see all embedded images, "
            "then **select which images are logos** you want replaced. "
            "The tool will find those logos in **every file in the ZIP** and swap them "
            "for the new logo automatically. You can select multiple logos "
            "(e.g. a title-slide logo **and** a footer logo)."
        )
        st.divider()

        # Step 1 ──────────────────────────────────────────────────────────────
        st.markdown("#### Step 1 — Scan a sample file from the ZIP")
        _sample_names  = [os.path.basename(f) for f in _supported]
        _sample_choice = st.selectbox(
            "Pick which file to use as a reference scan",
            _sample_names, key="zip_sample_choice",
        )

        if st.button("🔍 Scan sample file", key="btn_zip_scan"):
            with zipfile.ZipFile(BytesIO(zbytes_zip)) as _zs:
                _full_paths = [f for f in _supported if os.path.basename(f) == _sample_choice]
                if _full_paths:
                    _raw = _zs.read(_full_paths[0])
                    with st.spinner(f"Scanning {_sample_choice}…"):
                        _scanned = extract_images_any(_raw, _full_paths[0])
                    st.session_state.zip_sample_imgs  = _scanned
                    st.session_state.zip_ref_indices  = []
                    if not _scanned:
                        st.warning("No images found in this file. Try picking a different sample.")
                    st.rerun()

        # Step 2 ──────────────────────────────────────────────────────────────
        _sample_imgs = st.session_state.zip_sample_imgs
        if _sample_imgs:
            st.divider()
            st.markdown(
                "#### Step 2 — Select the logo(s) to replace\n"
                "Click **Select as logo** to toggle an image. "
                "Select as many as needed — e.g. one header logo and one footer logo."
            )
            _COLS = 5
            for _ri, _row in enumerate(
                [_sample_imgs[i : i + _COLS] for i in range(0, len(_sample_imgs), _COLS)]
            ):
                _cols = st.columns(_COLS)
                for _ci, (_col, _img) in enumerate(zip(_cols, _row)):
                    _idx    = _ri * _COLS + _ci
                    _is_sel = _idx in st.session_state.zip_ref_indices
                    _border = "2px solid #00c853" if _is_sel else "2px solid #ddd"
                    with _col:
                        st.markdown(
                            f'<div style="border:{_border};border-radius:6px;padding:4px">',
                            unsafe_allow_html=True,
                        )
                        st.image(thumbnail_bytes(_img["bytes"]), use_container_width=True)
                        _locs  = _img.get("locations", [])
                        _label = (
                            _locs[0] if len(_locs) == 1
                            else (f"{_locs[0]} +{len(_locs)-1} more" if _locs else "—")
                        )
                        st.caption(_label)
                        _btn_lbl = "✅ Selected" if _is_sel else "Select as logo"
                        if st.button(_btn_lbl, key=f"zg_{_idx}", use_container_width=True):
                            _refs = list(st.session_state.zip_ref_indices)
                            if _is_sel:
                                _refs.remove(_idx)
                            else:
                                _refs.append(_idx)
                            st.session_state.zip_ref_indices = _refs
                            st.rerun()
                        st.markdown("</div>", unsafe_allow_html=True)

            _n_sel = len(st.session_state.zip_ref_indices)
            if _n_sel:
                st.success(f"{_n_sel} logo image(s) selected for replacement.")
            else:
                st.info("Click images above to mark them as logos to be replaced.")

        # Step 3 ──────────────────────────────────────────────────────────────
        _refs_idx = st.session_state.zip_ref_indices
        if _refs_idx and _sample_imgs:
            st.divider()
            st.markdown("#### Step 3 — Upload new logo & run")
            _gc1, _gc2 = st.columns(2)
            with _gc1:
                st.markdown("**Logo(s) that will be replaced (your selection):**")
                _sel_thumbs = [_sample_imgs[i] for i in _refs_idx]
                _tcols      = st.columns(min(len(_sel_thumbs), 4))
                for _tc, _si in zip(_tcols, _sel_thumbs):
                    with _tc:
                        st.image(thumbnail_bytes(_si["bytes"]), use_container_width=True)
            with _gc2:
                st.markdown("**New logo to apply to all selected logos:**")
                zg_new_logo = st.file_uploader(
                    "New logo",
                    type=["png", "jpg", "jpeg", "bmp", "webp"],
                    key="zip_guided_new",
                    label_visibility="collapsed",
                )
                if zg_new_logo:
                    st.image(zg_new_logo, use_container_width=True)

            zg_thresh = st.slider(
                "Match sensitivity — hash distance (higher = more tolerant of minor variations between files)",
                0, 20, 8, key="zip_guided_thresh",
            )

            if st.button(
                "🚀 Replace logos across all files in ZIP",
                type="primary",
                disabled=zg_new_logo is None,
                key="btn_zip_guided",
            ):
                _ref_bytes = [_sample_imgs[i]["bytes"] for i in _refs_idx]
                _new_bytes = zg_new_logo.read()

                _log_ph    = st.empty()
                _log_lines_g: list[str] = []

                def _glog(msg: str):
                    _log_lines_g.append(msg)
                    _log_ph.code("\n".join(_log_lines_g[-30:]))

                _glog("Starting guided batch replacement…")
                with st.spinner("Processing ZIP — this may take a moment…"):
                    _out_zip, _f_done, _r_done = process_zip_by_refs(
                        zbytes_zip, _ref_bytes, _new_bytes, zg_thresh, progress_cb=_glog
                    )
                _glog("")
                _glog(f"✅ Done — {_f_done} file(s) processed, {_r_done} replacement(s) made.")

                if _r_done == 0:
                    st.warning(
                        "No matching logos were found. "
                        "Try raising the sensitivity slider or re-selecting the logo images in Step 2."
                    )
                else:
                    st.success(
                        f"Replaced **{_r_done}** logo instance(s) across **{_f_done}** file(s)."
                    )
                st.download_button(
                    "⬇️ Download updated ZIP", _out_zip,
                    file_name=stem_zip + "_updated.zip", mime="application/zip",
                    key="dl_guided_zip",
                )

    # ── ZIP Tab C: automated region detection (baked-in logos) ───────────────
    with zip_tab_region:
        st.markdown(
            "Use this when the logo is **baked into a full-slide background image** "
            "and cannot be selected as a separate shape. "
            "The tool scans every image in every file and detects the logo region using "
            "OCR (text recognition) or visual template matching, then pastes the new logo over it."
        )
        st.divider()

        zr_method = st.radio(
            "Detection method",
            ["OCR — find by text", "Template match — find by visual crop"],
            horizontal=True, key="zip_region_method",
        )

        zr_c1, zr_c2 = st.columns(2)

        with zr_c1:
            st.markdown("**Detection settings**")
            if zr_method.startswith("OCR"):
                zr_ocr_text   = st.text_input("Text to search for", value="LTIMindtree", key="zr_ocr_text")
                zr_expand_l   = st.slider("Expand left (% of text width, to include icon)",
                                          0, 200, 90, key="zr_expand_l")
                zr_vpad       = st.slider("Vertical padding (%)", 0, 50, 15, key="zr_vpad")
                zr_tmpl_bytes = None
            else:
                st.markdown(
                    "📸 **A screenshot of the logo is fine.** "
                    "The tool will automatically remove any surrounding padding / desktop background "
                    "and normalise contrast before searching. "
                    "Try to capture the logo area as closely as possible for best results."
                )
                zr_tmpl_file  = st.file_uploader(
                    "Screenshot or crop of old logo",
                    type=["png","jpg","jpeg","bmp","webp"], key="zr_tmpl"
                )
                zr_tmpl_bytes = zr_tmpl_file.read() if zr_tmpl_file else None
                if zr_tmpl_bytes:
                    col_raw, col_crop = st.columns(2)
                    with col_raw:
                        st.caption("Your upload")
                        st.image(zr_tmpl_bytes, use_container_width=True)
                    with col_crop:
                        st.caption("After auto-crop (what will be matched)")
                        st.image(autocrop_screenshot(zr_tmpl_bytes), use_container_width=True)
                zr_min_score  = st.slider(
                    "Minimum match score (0=loose, 1=exact)", 0.0, 1.0, 0.40, step=0.05, key="zr_score"
                )

        with zr_c2:
            st.markdown("**New logo**")
            zr_new_logo = st.file_uploader(
                "New logo", type=["png","jpg","jpeg","bmp","webp"], key="zr_new_logo"
            )
            if zr_new_logo:
                st.image(zr_new_logo, use_container_width=True)

        # Readiness check
        zr_ocr_ready  = zr_method.startswith("OCR") and zr_new_logo is not None
        zr_tmpl_ready = zr_method.startswith("Template") and zr_tmpl_bytes is not None and zr_new_logo is not None
        zr_ready = zr_ocr_ready or zr_tmpl_ready

        if st.button(
            "🚀 Run automated replacement across all files in ZIP",
            type="primary", disabled=not zr_ready, key="btn_zip_region"
        ):
            zr_new_bytes = zr_new_logo.read()

            # Build the detect_fn from current settings
            if zr_method.startswith("OCR"):
                _text   = zr_ocr_text
                _expl   = zr_expand_l / 100
                _vpad   = zr_vpad / 100
                def _detect(img_bytes, _t=_text, _e=_expl, _v=_vpad):
                    return detect_by_ocr(img_bytes, _t, _e, _v)
            else:
                _tmpl   = zr_tmpl_bytes
                _score  = zr_min_score
                def _detect(img_bytes, _tmpl=_tmpl, _s=_score):
                    return detect_by_template(img_bytes, _tmpl, _s)

            log_placeholder = st.empty()
            log_lines: list[str] = []

            def _log(msg: str):
                log_lines.append(msg)
                log_placeholder.code("\n".join(log_lines[-30:]))  # show last 30 lines

            _log("Starting automated region replacement…")

            with st.spinner("Running detection across all files — this may take a few minutes…"):
                out_zip, files_done, regions_done, no_match = process_zip_region(
                    zbytes_zip, _detect, zr_new_bytes, progress_cb=_log
                )

            _log("")
            _log(f"✅ Done — {files_done} file(s) processed, "
                 f"{regions_done} region(s) replaced, "
                 f"{no_match} file(s) had no match.")

            if regions_done == 0:
                st.warning(
                    "No logo regions were detected in any file. "
                    "Try adjusting detection settings or switching methods."
                )
            else:
                st.success(
                    f"Replaced **{regions_done}** logo region(s) across "
                    f"**{files_done - no_match}** file(s)."
                )

            st.download_button(
                "⬇️ Download updated ZIP", out_zip,
                file_name=stem_zip + "_updated.zip", mime="application/zip"
            )

    st.stop()

# ── Scan ──────────────────────────────────────────────────────────────────────

if doc_file:
    if st.button("🔍 Scan document for all embedded images"):
        name, fbytes = st.session_state.doc_cache
        with st.spinner("Scanning…"):
            imgs = extract_images_any(fbytes, name)
        st.session_state.images     = imgs
        st.session_state.selected   = None
        st.session_state.selected_set = set()
        st.session_state.bbox       = None
        st.session_state.drill_active = False
        if not imgs:
            st.warning("No images found in this document.")

# ── Image grid ────────────────────────────────────────────────────────────────

if st.session_state.images:
    st.divider()
    n = len(st.session_state.images)
    _n_sel = len(st.session_state.selected_set)
    st.subheader(
        f"Step 2 \u2014 Mark logo(s) to replace  ({n} unique image{'s' if n != 1 else ''} found)"
    )
    st.caption(
        "**Mark as logo** to add an image to the replacement list (supports multi-select). "
        "**Inspect** to open the region-detect panel for a single image."
    )

    COLS = 5
    imgs = st.session_state.images
    for row_i, row in enumerate([imgs[i:i+COLS] for i in range(0, len(imgs), COLS)]):
        cols = st.columns(COLS)
        for col_i, (col, img) in enumerate(zip(cols, row)):
            idx        = row_i * COLS + col_i
            is_marked  = idx in st.session_state.selected_set
            is_inspect = st.session_state.selected == idx
            if is_marked:
                border = "2px solid #00c853"
            elif is_inspect:
                border = "2px solid #2196f3"
            else:
                border = "2px solid #ddd"
            with col:
                st.markdown(
                    f'<div style="border:{border};border-radius:6px;padding:4px">',
                    unsafe_allow_html=True,
                )
                st.image(thumbnail_bytes(img["bytes"]), use_container_width=True)
                locs  = img["locations"]
                label = locs[0] if len(locs) == 1 else f"{locs[0]} +{len(locs)-1} more"
                st.caption(label)
                bc1, bc2 = st.columns(2)
                with bc1:
                    mark_lbl = "\u2705 Marked" if is_marked else "Mark as logo"
                    if st.button(mark_lbl, key=f"mark_{idx}", use_container_width=True):
                        _s = set(st.session_state.selected_set)
                        if is_marked:
                            _s.discard(idx)
                        else:
                            _s.add(idx)
                        st.session_state.selected_set = _s
                        st.rerun()
                with bc2:
                    insp_lbl = "\U0001f535 Inspecting" if is_inspect else "\U0001f50d Inspect"
                    if st.button(insp_lbl, key=f"insp_{idx}", use_container_width=True):
                        st.session_state.selected     = idx
                        st.session_state.bbox         = None
                        st.session_state.drill_active = False
                        st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

    if _n_sel:
        st.success(f"{_n_sel} logo image{'s' if _n_sel != 1 else ''} marked for replacement.")
    else:
        st.info("Click **Mark as logo** on one or more images above, then upload the new logo below.")

# ── Step 3: Bulk replace all marked logos ───────────────────────────────────────

_marked_indices = list(st.session_state.selected_set)
if _marked_indices and st.session_state.images:
    _marked_imgs = [st.session_state.images[i] for i in sorted(_marked_indices)]
    name_sf, fbytes_sf = st.session_state.doc_cache
    ext_sf  = os.path.splitext(name_sf.lower())[1]
    stem_sf = os.path.splitext(name_sf)[0]

    st.divider()
    st.subheader("Step 3 \u2014 Replace selected logo(s)")

    _mc1, _mc2 = st.columns([2, 1])
    with _mc1:
        st.markdown("**Logo(s) that will be replaced:**")
        _tcols = st.columns(min(len(_marked_imgs), 5))
        for _tc, _mi in zip(_tcols, _marked_imgs):
            with _tc:
                st.image(thumbnail_bytes(_mi["bytes"]), use_container_width=True)
                _ml = _mi.get("locations", [])
                st.caption(_ml[0] if len(_ml) == 1 else (f"{_ml[0]} +{len(_ml)-1} more" if _ml else "—"))
    with _mc2:
        st.markdown("**New logo to apply:**")
        _new_logo_multi = st.file_uploader(
            "New logo",
            type=["png", "jpg", "jpeg", "bmp", "webp"],
            key="multi_new_logo",
            label_visibility="collapsed",
        )
        if _new_logo_multi:
            st.image(_new_logo_multi, use_container_width=True)

    _thr_multi = st.slider(
        "Match sensitivity (hash distance — higher = more tolerant)",
        0, 20, 8, key="multi_thresh",
    )

    if st.button(
        f"\U0001f504 Replace {len(_marked_imgs)} logo image{'s' if len(_marked_imgs) != 1 else ''} in document",
        type="primary",
        disabled=_new_logo_multi is None,
        key="btn_multi_replace",
    ):
        _old_refs  = [_mi["bytes"] for _mi in _marked_imgs]
        _new_bytes = _new_logo_multi.read()
        with st.spinner("Replacing logos\u2026"):
            _out_bytes, _count = replace_multiple_logos_any(
                fbytes_sf, name_sf, _old_refs, _new_bytes, _thr_multi
            )
        if _count == 0:
            st.warning("No replacements made. Try raising the sensitivity slider.")
        else:
            st.success(f"Replaced **{_count}** logo instance(s) across all locations.")
        _download_doc("\u2b07\ufe0f Download updated document", _out_bytes, stem_sf, ext_sf)

# ── Step 4: Inspect / region-detect for a single selected image ─────────────

if st.session_state.selected is not None:
    sel  = st.session_state.images[st.session_state.selected]
    name, fbytes = st.session_state.doc_cache
    ext  = os.path.splitext(name.lower())[1]
    stem = os.path.splitext(name)[0]

    st.divider()
    st.subheader("Step 4 — Inspect / region-detect inside one image")

    tab_normal, tab_drill = st.tabs([
        "🔄 Replace this whole image",
        "🔬 Find & replace logo **inside** this image (OCR / template match)",
    ])

    # ── Tab A: Normal whole-image replacement ────────────────────────────────
    with tab_normal:
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            st.markdown("**Selected image**")
            st.image(sel["bytes"], use_container_width=True)
            st.caption("Found in: " + ", ".join(sel["locations"]))
            st.download_button("⬇️ Save this image", to_png(sel["bytes"]),
                               file_name="extracted_image.png", mime="image/png")
        with c2:
            st.markdown("**New logo to replace with**")
            new_logo_whole = st.file_uploader(
                "New logo", type=["png","jpg","jpeg","bmp","webp"],
                label_visibility="collapsed", key="whole_new",
            )
            if new_logo_whole:
                st.image(new_logo_whole, use_container_width=True)
        with c3:
            st.markdown("**Settings**")
            thr_whole = st.slider("Match sensitivity", 0, 20, 5, key="whole_thresh")
            st.caption("Lower = exact match only.")

        if st.button("🔄 Replace whole image", type="primary",
                     disabled=new_logo_whole is None, key="btn_whole"):
            old_bytes = to_png(sel["bytes"])
            new_bytes = new_logo_whole.read()
            with st.spinner("Replacing…"):
                out_bytes, count = replace_logo_any(fbytes, name, old_bytes, new_bytes, thr_whole)
            if count == 0:
                st.warning("No replacements made. Try raising match sensitivity.")
            else:
                st.success(f"Replaced {count} instance(s).")
            _download_doc("⬇️ Download updated document", out_bytes, stem, ext)

    # ── Tab B: Intra-image logo detection + region replace ───────────────────
    with tab_drill:
        st.markdown(
            "Use this when the logo is **baked into a larger background image** "
            "and can't be selected separately in the document editor."
        )
        st.image(sel["bytes"], caption="Selected image — logo will be hunted inside this", use_container_width=True)
        st.divider()

        detect_method = st.radio(
            "Detection method",
            ["OCR — find by text  (recommended for LTIMindtree)", "Template match — find by visual crop"],
            horizontal=True,
            key="detect_method",
        )

        # --- OCR branch ---
        if detect_method.startswith("OCR"):
            st.markdown("**OCR settings**")
            search_text     = st.text_input("Text to search for", value="LTIMindtree", key="ocr_text")
            expand_left     = st.slider(
                "Expand left (to include icon beside text, as % of text width)",
                0, 200, 90, key="ocr_expand",
            )
            v_pad           = st.slider("Vertical padding (%)", 0, 50, 15, key="ocr_vpad")

            if st.button("🔍 Detect logo via OCR", key="btn_ocr"):
                with st.spinner("Running OCR… (first run downloads language model ~50 MB)"):
                    result = detect_by_ocr(
                        sel["bytes"], search_text,
                        expand_left_ratio=expand_left / 100,
                        v_padding_ratio=v_pad / 100,
                    )
                if result is None:
                    st.warning(f'Could not find "{search_text}" in the image. '
                               "Try a partial string (e.g. 'LTI' or 'Mindtree') or switch to Template match.")
                else:
                    st.session_state.bbox = result
                    st.success(f'Found "{result["text"]}" — confidence {result["confidence"]:.0%}')

        # --- Template match branch ---
        else:
            st.markdown(
                "📸 **A screenshot is fine** — you do not need the original logo file. "
                "The tool automatically strips surrounding padding and normalises contrast "
                "so screenshots match reliably. Try to capture just the logo area."
            )
            tmpl_file = st.file_uploader(
                "Screenshot or crop of the old logo",
                type=["png","jpg","jpeg","bmp","webp"], key="tmpl_crop",
            )
            if tmpl_file:
                raw_bytes  = tmpl_file.read()
                crop_bytes = autocrop_screenshot(raw_bytes)
                tc1, tc2 = st.columns(2)
                with tc1:
                    st.caption("Your upload")
                    st.image(raw_bytes, use_container_width=True)
                with tc2:
                    st.caption("After auto-crop (what will be matched)")
                    st.image(crop_bytes, use_container_width=True)
            else:
                raw_bytes = None

            min_score = st.slider("Minimum match score (0 = very loose, 1 = exact)", 0.0, 1.0, 0.40,
                                  step=0.05, key="tmpl_score")

            if raw_bytes and st.button("🔍 Detect logo via template matching", key="btn_tmpl"):
                with st.spinner("Running multi-scale template matching…"):
                    result = detect_by_template(sel["bytes"], raw_bytes, min_score=min_score)
                if result is None:
                    st.warning(
                        f"No match found above score {min_score:.2f}. "
                        "Try lowering the minimum score, or switch to OCR mode."
                    )
                else:
                    st.session_state.bbox = result
                    st.success(f'Match found — score {result["score"]:.0%}')

        # --- Show detected region & replace ---
        if st.session_state.bbox:
            bbox = st.session_state.bbox
            x1, y1, x2, y2 = bbox["x1"], bbox["y1"], bbox["x2"], bbox["y2"]

            st.divider()
            st.markdown("**Detected region** (red box)")
            preview = annotate_bbox(sel["bytes"], x1, y1, x2, y2)
            st.image(preview, use_container_width=True)
            st.caption(f"Region: ({x1}, {y1}) → ({x2}, {y2})  |  size: {x2-x1} × {y2-y1} px")

            # Manual fine-tune
            with st.expander("Fine-tune bounding box"):
                fc1, fc2 = st.columns(2)
                with fc1:
                    x1 = st.number_input("x1 (left)",  value=x1, step=1, key="bb_x1")
                    y1 = st.number_input("y1 (top)",   value=y1, step=1, key="bb_y1")
                with fc2:
                    x2 = st.number_input("x2 (right)", value=x2, step=1, key="bb_x2")
                    y2 = st.number_input("y2 (bottom)", value=y2, step=1, key="bb_y2")
                if st.button("↺ Update preview", key="btn_preview"):
                    preview2 = annotate_bbox(sel["bytes"], x1, y1, x2, y2)
                    st.image(preview2, use_container_width=True)

            st.divider()
            new_logo_drill = st.file_uploader(
                "Upload new logo to place in this region",
                type=["png","jpg","jpeg","bmp","webp"], key="drill_new",
            )
            if new_logo_drill:
                st.image(new_logo_drill, caption="New logo (will be scaled to fit red box)", width=200)

            if st.button("✂️ Replace region & rebuild document", type="primary",
                         disabled=new_logo_drill is None, key="btn_drill_replace"):
                with st.spinner("Replacing region…"):
                    # 1. paint new logo into the selected image
                    modified_img = replace_region_in_image(
                        sel["bytes"], x1, y1, x2, y2, new_logo_drill.read()
                    )
                    # 2. put modified image back into the document
                    out_bytes = embed_image_in_document(fbytes, name, sel, modified_img)

                st.success("Done! The logo region has been replaced in the document.")
                _download_doc("⬇️ Download updated document", out_bytes, stem, ext)

elif doc_file and not is_zip and not st.session_state.images:
    st.info("Click **Scan document for all embedded images** above to get started.")