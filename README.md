# Document Logo Replacer

A Streamlit web application that automates logo and text replacement inside **PPTX**, **DOCX**, **PDF**, and **ZIP batch** files — without corrupting any other content in the document.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [How It Works — Core Architecture](#how-it-works--core-architecture)
- [Supported File Formats](#supported-file-formats)
- [Installation](#installation)
- [Running the App](#running-the-app)
- [Usage Guide](#usage-guide)
  - [Single File — Logo Replacement](#single-file--logo-replacement)
  - [Single File — Text Replacement](#single-file--text-replacement)
  - [Single File — Logo Baked into Background](#single-file--logo-baked-into-background)
  - [ZIP Batch — Quick Replace](#zip-batch--quick-replace)
  - [ZIP Batch — Guided Replace](#zip-batch--guided-replace)
  - [ZIP Batch — Region Detection](#zip-batch--region-detection)
- [Detection Methods Explained](#detection-methods-explained)
  - [Perceptual Hashing](#perceptual-hashing)
  - [OCR (Text Detection)](#ocr-text-detection)
  - [Template Matching](#template-matching)
- [Understanding Match Sensitivity](#understanding-match-sensitivity)
- [Technical Deep-Dive](#technical-deep-dive)
  - [Why ZIP-Level Byte Swapping?](#why-zip-level-byte-swapping)
  - [Split-Run Text Replacement](#split-run-text-replacement)
  - [PDF Alpha Channel Fix](#pdf-alpha-channel-fix)
- [Project Structure](#project-structure)
- [Dependencies](#dependencies)
- [Troubleshooting](#troubleshooting)

---

## Overview

This tool was built to solve a real-world rebranding problem: when a company renames or rebrands, hundreds of PPTX/DOCX/PDF documents need their old logos replaced with the new one, and footer text like `"LTIMindtree | Privileged and Confidential"` updated to `"LTM | Privileged and Confidential"`. Manual editing is error-prone and slow. This app automates the entire process.

---

## Features

| Feature | Details |
|---|---|
| **Logo replacement** | Hash-based perceptual matching. Finds logos even if slightly resized or re-compressed |
| **Multi-logo replacement** | Select multiple logos (e.g. header + footer logo) in one pass |
| **Text replacement** | Handles text split across formatting runs — a common Office XML quirk |
| **Combined logo + text** | One button, one download — do both operations simultaneously |
| **Baked-in logo detection** | OCR or template matching to find logos inside background images |
| **ZIP batch processing** | Process an entire folder of documents at once |
| **Safe operation** | Only the targeted image bytes or text nodes are ever modified; all other XML, animations, shapes, and relationships are byte-copied unchanged |
| **Supports all major formats** | PPTX, DOCX, DOC, PDF |

---

## How It Works — Core Architecture

Office documents (PPTX, DOCX) are ZIP archives. Instead of parsing and re-serialising the XML — which risks changing z-order, breaking relationship references, or corrupting animations — this app treats the file as a ZIP and **swaps only the specific bytes that need to change**:

```
Input file (ZIP)
    ppt/slides/slide1.xml       ← copied verbatim
    ppt/slides/slide2.xml       ← copied verbatim
    ppt/slideMasters/master1.xml ← copied verbatim
    ppt/media/image1.png        ← OLD LOGO → replaced with new logo bytes
    ppt/media/image2.png        ← copied verbatim
    [Content_Types].xml         ← copied verbatim
    ...
Output file (ZIP)               ← structurally identical, only image1.png changed
```

For PDF files, PyMuPDF's `replace_image(xref, ...)` API is used, which replaces the image XObject at the given cross-reference ID while leaving the page content streams untouched.

---

## Supported File Formats

| Format | Logo Replace | Text Replace | Baked-in Detection |
|--------|:---:|:---:|:---:|
| PPTX | ✅ | ✅ | ✅ |
| DOCX / DOC | ✅ | ✅ | ✅ |
| PDF | ✅ | ❌ | ✅ |
| ZIP (batch) | ✅ | ❌ | ✅ |

> Text replacement in PDFs requires editing page content streams which is out of scope for this tool. Use the OCR/template "region replace" approach to overwrite baked-in text in PDFs.

---

## Installation

### Prerequisites

- Python 3.10 or higher
- pip

### Steps

1. **Clone or download** this project folder.

2. **Create and activate a virtual environment** (recommended):

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

3. **Install dependencies**:

   ```powershell
   pip install -r requirements.txt
   ```

   > The first run of OCR mode will download the EasyOCR English model (~300 MB). This is cached locally and only happens once.

---

## Running the App

```powershell
streamlit run app.py
```

The app opens automatically in your browser at `http://localhost:8501`.

---

## Usage Guide

### Single File — Logo Replacement

This is the standard workflow for replacing a logo that exists as a distinct embedded image.

1. **Upload** your PPTX, DOCX, or PDF.
2. Click **🔍 Scan document for all embedded images** — a gallery of every unique embedded image appears.
3. Click **Mark as logo** on each image you want to replace. You can select multiple (e.g. a header logo and a footer logo).
4. In the **Step 3** section that appears, upload your **new logo**.
5. Optionally, expand **"🔤 Also replace text"** and fill in the find/replace fields to do both in one step.
6. Click **Replace logos** → download the updated file.

**Match sensitivity slider:**  
Controls how strictly the old logo must match. `0` = exact pixel match only. `8` (default) tolerates minor JPEG re-compression. `20` = very loose (may accidentally replace similar-looking images).

---

### Single File — Text Replacement

Use this to update footer text, copyright notices, or any text string in the document.

1. Upload your PPTX or DOCX.
2. Scroll to the **🔤 Text Replacement** section at the bottom.
3. Enter the **exact text to find** and the **replacement text**.
4. Click **Replace Text** → download the updated file.

> If logos are already marked in Step 2 and a new logo is uploaded in Step 3, a checkbox **"Also apply the logo replacement"** appears. Ticking it produces a single combined file with both changes applied.

**Important:** The search is case-sensitive and must match the visible string exactly. If "0 occurrences found", the text may be split across formatting runs or stored with different casing.

---

### Single File — Logo Baked into Background

Sometimes the logo is embedded as part of a full-slide background image (a common PowerPoint template technique). It cannot be replaced by swapping a media file because it's not a separate shape.

1. Upload your document and scan it.
2. Click **🔍 Inspect** on the background image that contains the logo.
3. Switch to the **"🔬 Find & replace logo inside this image"** tab.
4. Choose a detection method:
   - **OCR** — searches for text (e.g. "LTIMindtree") within the image. Best when the logo includes readable text. Expands the bounding box leftward to include any icon next to the text.
   - **Template match** — you upload a screenshot of the old logo. The app auto-crops desktop padding and uses multi-scale OpenCV matching to locate it.
5. Click the **Detect** button. A red bounding box preview shows what was found.
6. Fine-tune the box if needed using the coordinate inputs.
7. Upload the new logo and click **Replace region**.

---

### ZIP Batch — Quick Replace

Upload a ZIP of documents and provide the old + new logo as reference images. Every document in the ZIP is processed using hash matching.

1. Upload a `.zip` file.
2. Go to the **🔄 Quick** tab.
3. Upload the old logo and new logo reference images.
4. Click **Replace Logos**.

---

### ZIP Batch — Guided Replace

Useful when different files in the ZIP may have slightly different versions of the logo (different compression, minor resizing).

1. Upload a `.zip` file → go to **🎯 Guided** tab.
2. **Step 1:** Pick a sample file from the ZIP and click **Scan** to see its embedded images.
3. **Step 2:** Click **Select as logo** on each image that should be replaced across all files.
4. **Step 3:** Upload the new logo and click **Replace logos across all files**.

The selected images become reference hashes. Any embedded image in any file that perceptually matches any reference is replaced.

---

### ZIP Batch — Region Detection

For ZIPs where the logo is baked into background images across all files.

1. Upload a `.zip` → go to **🔬 Region detect** tab.
2. Choose **OCR** or **Template match** and configure detection settings.
3. Upload the new logo.
4. Click **Run automated replacement** — the app iterates every document, every embedded image, detects the logo region, and replaces it.

---

## Detection Methods Explained

### Perceptual Hashing

Perceptual hashing (pHash) converts an image into a compact 64-bit fingerprint based on its visual content (not pixel values). Two images with the same logo at different resolutions or compression levels will have similar hashes. The **Hamming distance** between two hashes measures visual difference:

| Distance | Interpretation |
|----------|----------------|
| 0 | Pixel-identical |
| 1–5 | Nearly identical (minor compression artefacts) |
| 6–10 | Very similar (slight resize or re-save) |
| 11–20 | Somewhat similar |
| >20 | Likely different images |

The sensitivity slider maps directly to the maximum allowed Hamming distance.

### OCR (Text Detection)

Uses **EasyOCR** with an English language model. The engine detects all text regions in the image, then filters for the target substring (case-insensitive). Since many logos (like LTIMindtree) have a graphical icon to the *left* of the text, the bounding box is expanded leftward by a configurable ratio of the text width to capture the full logo including the icon.

### Template Matching

Uses **OpenCV's `matchTemplate`** with `TM_CCOEFF_NORMED`. This measures normalised cross-correlation between the template (your screenshot) and sliding windows across the target image.

Key preprocessing steps that make this reliable even with screenshots:
- **Auto-crop:** Removes uniform background/desktop padding from screenshots using corner-pixel colour sampling.
- **CLAHE normalisation:** Applies contrast-limited adaptive histogram equalisation so that brightness differences between a screen screenshot and a document rendering don't hurt matching quality.
- **Multi-scale search:** Tries scales from 40% to 160% of the template size in 5% steps to handle size differences between your screenshot and the embedded image.

---

## Understanding Match Sensitivity

The sensitivity slider (0–20) controls the maximum perceptual hash distance:

- **Too low (0–3):** Only exact or near-exact matches. May miss logos that were resaved or slightly resized across different documents.
- **Recommended (5–10):** Catches the same logo across typical document variations.
- **Too high (15–20):** May accidentally replace unrelated images that look superficially similar to the logo (same general colour scheme, similar size).

Start at the default (5–8) and raise it only if expected logos are not being found.

---

## Technical Deep-Dive

### Why ZIP-Level Byte Swapping?

Office Open XML formats (PPTX, DOCX) are ZIP archives containing XML files and media assets. A naïve approach using python-pptx to remove a shape and re-add it with a new picture:

1. Strips the shape's XML element from the slide tree
2. Appends a brand-new `<p:pic>` element at the end
3. This **changes the z-order** (the new shape is on top of everything)
4. It also fires python-pptx's shape-ID auto-increment logic which can break cross-references in animations and hyperlinks
5. Master/layout shapes don't support `add_picture` natively, requiring fragile monkey-patching

By swapping only the bytes of the media file inside the ZIP, the XML (including the `<p:blipFill r:embed="rId2"/>` reference) is never touched. The document reopens referencing the same `rId2`, which still points to the same `ppt/media/image1.png` path — which now contains the new logo.

### Split-Run Text Replacement

Microsoft Office frequently stores a single visible string as multiple XML run (`<a:r>`) elements for internal reasons (spell-check boundaries, formatting history, mixed character styles). For example:

```xml
<a:p>
  <a:r><a:t>LTIMind</a:t></a:r>
  <a:r><a:t>tree | Privileged</a:t></a:r>
  <a:r><a:t> and Confidential</a:t></a:r>
</a:p>
```

A simple string search for `"LTIMindtree | Privileged and Confidential"` on raw XML bytes would find nothing. This app's `_replace_text_in_xml` function:

1. Concatenates all text node values within a paragraph into one string
2. Checks if the target text exists in the concatenation
3. Writes the replacement into the **first** run's text node
4. Clears all subsequent run text nodes (while preserving their `<a:rPr>` formatting properties)

This ensures all formatting runs are preserved — only the text content changes.

### PDF Alpha Channel Fix

PDFs represent image transparency via a separate `/SMask` (soft-mask) XObject. When PyMuPDF's `replace_image` is given an RGBA PNG:
- It stores all 4 bytes per pixel in the image stream
- But it does **not** automatically create a `/SMask` XObject to tell the renderer which byte is alpha
- PDF readers interpret the raw data as a 4-channel image, producing a corrupted magenta/tinted appearance

The fix: composite the replacement logo onto a white background before embedding, producing a standard 3-channel RGB PNG that any PDF renderer handles correctly without needing an `/SMask`.

---

## Project Structure

```
Powerpoint Automation/
├── app.py              # Main application (all logic + Streamlit UI)
├── requirements.txt    # Python dependencies
├── Dockerfile          # Container deployment configuration
└── README.md           # This file
```

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `streamlit` | Web UI framework |
| `python-pptx` | PPTX introspection (location annotation only — not used for editing) |
| `Pillow` | Image loading, conversion, compositing, thumbnails |
| `ImageHash` | Perceptual hashing (pHash) for image similarity |
| `PyMuPDF (fitz)` | PDF image extraction and replacement |
| `easyocr` | Text detection inside images for baked-in logo finding |
| `opencv-python-headless` | Multi-scale template matching |
| `numpy` | Array operations for OpenCV and EasyOCR |
| `lxml` | Fast XML parsing for split-run text replacement |

---

## Troubleshooting

### "No replacements made" on logo replacement
- Raise the **Match sensitivity** slider. The old logo in the document may have been resaved at a different compression level.
- Ensure you uploaded the correct old logo. Scan the document first and visually compare the extracted image with your reference.

### "Text not found" on text replacement
- Check the exact casing and spacing of the target string. The search is case-sensitive.
- The text may exist only in a baked-in background image (not as XML text). Use the OCR region detection instead.
- Look for hidden characters: copy the text directly from the document (not from memory) and paste it into the find field.

### OCR finds nothing
- Try a shorter search string (e.g. just `"LTIMindtree"` instead of the full footer).
- The logo may be too small or blurry in the extracted image.
- Switch to Template Match mode.

### Template matching score is always low
- Use Auto-crop preview to verify your screenshot doesn't have excessive desktop background padding.
- Lower the **Minimum match score** slider (try 0.25–0.35).
- Ensure the logo in the template and in the document face the same direction and have similar aspect ratios.

### PDF replacement looks wrong (tinted/corrupted)
- This was a known bug (RGBA alpha channel mishandled by PyMuPDF). It is fixed in the current version via `_to_pdf_safe_png`, which composites to RGB before embedding.

### ZIP processing is slow
- Region detection (OCR/template matching) is compute-intensive and runs on CPU. A ZIP with 50 slides × 3 background images × multi-scale template search will take several minutes. Hash-based replacement is fast (typically under 1 second per document).
