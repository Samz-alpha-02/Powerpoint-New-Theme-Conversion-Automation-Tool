"""
app.py  –  PowerPoint Rebranding Tool
Streamlit web UI that accepts:
  • a ZIP archive of old .pptx files
  • the new company template .pptx
and returns a single ZIP containing all presentations rebranded with
the new template (slide master, theme colours, fonts, layouts).
"""

import time
import traceback
from io import BytesIO

import streamlit as st

from template_applier import process_zip

# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="PowerPoint Rebranding Tool",
    page_icon="🎨",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Minimal, clean CSS – no background overrides, pure Streamlit dark/light
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* Upload zones – subtle border only */
    [data-testid="stFileUploader"] {
        border: 1.5px dashed #555;
        border-radius: 8px;
        padding: 0.6rem 0.9rem;
        transition: border-color 0.2s;
    }
    [data-testid="stFileUploader"]:hover { border-color: #00b4d8; }

    /* Primary action button – vivid cyan accent */
    div.stButton > button {
        background: #00b4d8;
        color: #000;
        font-size: 1rem;
        font-weight: 700;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        width: 100%;
        letter-spacing: 0.3px;
        transition: background 0.2s, opacity 0.2s;
    }
    div.stButton > button:hover  { background: #0096c7; }
    div.stButton > button:active { opacity: 0.8; }

    /* Download button – emerald green */
    div[data-testid="stDownloadButton"] > button {
        background: #2dc653;
        color: #000;
        font-size: 1rem;
        font-weight: 700;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 2rem;
        width: 100%;
        margin-top: 0.75rem;
        transition: background 0.2s, opacity 0.2s;
    }
    div[data-testid="stDownloadButton"] > button:hover  { background: #22a244; }
    div[data-testid="stDownloadButton"] > button:active { opacity: 0.8; }

    /* Terminal-style log box */
    .log-box {
        background: #0d1117;
        color: #3fb950;
        font-family: "Cascadia Code", "Fira Mono", "Courier New", monospace;
        font-size: 0.8rem;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        max-height: 320px;
        overflow-y: auto;
        white-space: pre-wrap;
        word-break: break-all;
        line-height: 1.65;
    }

    /* Thin section divider */
    hr { border: none; border-top: 1px solid #30363d; margin: 1.25rem 0; }

    /* Step labels */
    .step-label {
        font-size: 0.78rem;
        font-weight: 700;
        color: #00b4d8;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 0.2rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.title("🎨 PowerPoint Rebranding Tool")
st.markdown(
    "Upload your old presentations and the new company template "
    "to automatically apply the new design to every slide."
)
st.markdown("<hr>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# How it works
# ---------------------------------------------------------------------------
with st.expander("ℹ️  How it works", expanded=False):
    st.markdown(
        """
**What this tool does**

1. You upload a **ZIP file** containing one or more `.pptx` presentations.
2. You upload your **new company template** (`.pptx`).
3. Click **Rebrand Presentations**.
4. The tool will, for every presentation:
   - Replace the **slide master** (theme colours, fonts, logo placeholders)  
     with the one from your template.
   - **Remap each slide layout** to the closest matching layout in the new master.
   - Apply the **correct heading / body fonts** from the new theme to all text.
   - Enable **word-wrap / auto-shrink** so no text overflows.
   - Preserve all **images, shapes, tables, and charts**.
5. Download the output ZIP containing the rebranded files.

> **Tip:** The new template must be a real PowerPoint file whose *Design* tab shows
> the theme you want. A presentation you saved as "Save As > PowerPoint Presentation"
> after applying a theme works perfectly.
        """
    )

# ---------------------------------------------------------------------------
# Upload section
# ---------------------------------------------------------------------------
col_left, col_right = st.columns(2, gap="large")

with col_left:
    st.markdown('<p class="step-label">Step 1 – Old Presentations</p>', unsafe_allow_html=True)
    zip_file = st.file_uploader(
        "Upload ZIP of .pptx files",
        type=["zip"],
        help="ZIP archive containing one or more .pptx files to rebrand.",
        key="zip_uploader",
    )
    if zip_file:
        st.success(f"📦 `{zip_file.name}` uploaded ({zip_file.size / 1024:.1f} KB)")

with col_right:
    st.markdown('<p class="step-label">Step 2 – New Template</p>', unsafe_allow_html=True)
    template_file = st.file_uploader(
        "Upload new template .pptx",
        type=["pptx"],
        help="The new company template – the theme/design from this file will be applied.",
        key="template_uploader",
    )
    if template_file:
        st.success(f"🎨 `{template_file.name}` uploaded ({template_file.size / 1024:.1f} KB)")

st.markdown("<hr>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Submit button  +  processing
# ---------------------------------------------------------------------------
submit_disabled = (zip_file is None) or (template_file is None)

if submit_disabled:
    st.info("⬆️  Please upload both files above to enable rebranding.", icon="ℹ️")

submit = st.button(
    "🚀  Rebrand Presentations",
    disabled=submit_disabled,
    use_container_width=True,
)

if submit:
    # ------------------------------------------------------------------ #
    # Read uploaded bytes
    # ------------------------------------------------------------------ #
    zip_bytes      = zip_file.read()
    template_bytes = template_file.read()

    # ------------------------------------------------------------------ #
    # Live log area
    # ------------------------------------------------------------------ #
    st.markdown("### 📋 Processing Log")
    log_placeholder   = st.empty()
    prog_placeholder  = st.empty()
    status_placeholder = st.empty()

    log_lines: list[str] = []

    def update_log(msg: str):
        log_lines.append(msg)
        log_html = "\n".join(log_lines[-120:])   # cap at 120 lines for UI perf
        log_placeholder.markdown(
            f'<div class="log-box">{log_html}</div>',
            unsafe_allow_html=True,
        )

    # ------------------------------------------------------------------ #
    # Progress bar
    # ------------------------------------------------------------------ #
    progress_bar = prog_placeholder.progress(0, text="Starting …")
    total_steps  = [0]   # mutable container for closure

    def progress_hook(msg: str):
        update_log(msg)
        # Rough progress estimation based on log volume
        total_steps[0] += 1
        pct = min(int(total_steps[0] * 2), 95)
        progress_bar.progress(pct, text=msg[:80])

    # ------------------------------------------------------------------ #
    # Run the rebranding
    # ------------------------------------------------------------------ #
    start_time = time.time()
    output_zip_bytes: bytes | None = None
    error_msg: str | None = None

    with st.spinner("Rebranding in progress – please wait …"):
        try:
            output_zip_bytes = process_zip(
                zip_bytes,
                template_bytes,
                progress_callback=progress_hook,
            )
        except ValueError as exc:
            error_msg = str(exc)
        except Exception as exc:
            error_msg = (
                f"Unexpected error: {exc}\n\n"
                + traceback.format_exc()
            )

    elapsed = time.time() - start_time

    # ------------------------------------------------------------------ #
    # Result
    # ------------------------------------------------------------------ #
    if error_msg:
        progress_bar.empty()
        status_placeholder.error(f"❌ Processing failed:\n{error_msg}")
        update_log(f"\n❌ Error: {error_msg}")
    else:
        progress_bar.progress(100, text="Complete!")
        status_placeholder.success(
            f"✅ All presentations rebranded successfully in {elapsed:.1f}s!"
        )
        update_log(f"\n✅ Finished in {elapsed:.1f}s")

        # Derive output filename from input ZIP name
        stem = zip_file.name.rsplit(".", 1)[0]
        out_filename = f"{stem}_rebranded.zip"

        st.download_button(
            label="⬇️  Download Rebranded ZIP",
            data=output_zip_bytes,
            file_name=out_filename,
            mime="application/zip",
            use_container_width=True,
        )

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown(
    "<p style='text-align:center; color:#6e7681; font-size:0.78rem;'>"
    "PowerPoint Rebranding Tool &nbsp;|&nbsp; "
    "Built with python-pptx &amp; Streamlit"
    "</p>",
    unsafe_allow_html=True,
)
