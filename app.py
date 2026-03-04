import os
import tempfile
import shutil
import zipfile
from io import BytesIO

import streamlit as st
from PIL import Image
import imagehash


# ----------------------------------------------------------
# Image helpers
# ----------------------------------------------------------

def get_image_hash(img_bytes):
    img = Image.open(BytesIO(img_bytes))
    return imagehash.phash(img)


def convert_new_logo(new_logo_bytes):
    img = Image.open(BytesIO(new_logo_bytes)).convert("RGBA")
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ----------------------------------------------------------
# PowerPoint processing
# ----------------------------------------------------------

def _extract_pictures(shape):
    """
    Recursively find picture shapes even inside grouped shapes.
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    pictures = []

    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        pictures.append(shape)

    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for s in shape.shapes:
            pictures.extend(_extract_pictures(s))

    return pictures


def replace_logo_pptx(input_bytes, old_logo_bytes, new_logo_bytes, threshold=5):

    from pptx import Presentation

    old_hash = get_image_hash(old_logo_bytes)
    new_logo_png = convert_new_logo(new_logo_bytes)

    temp_dir = tempfile.mkdtemp()

    try:

        input_path = os.path.join(temp_dir, "input.pptx")
        output_path = os.path.join(temp_dir, "output.pptx")

        with open(input_path, "wb") as f:
            f.write(input_bytes)

        prs = Presentation(input_path)

        replaced_count = 0

        for slide in prs.slides:

            pictures = []

            for shape in slide.shapes:
                pictures.extend(_extract_pictures(shape))

            for pic in pictures:

                try:
                    blob = pic.image.blob
                except Exception:
                    continue

                img_hash = get_image_hash(blob)

                if old_hash - img_hash <= threshold:

                    left = pic.left
                    top = pic.top
                    width = pic.width
                    height = pic.height

                    pic._element.getparent().remove(pic._element)

                    slide.shapes.add_picture(
                        BytesIO(new_logo_png),
                        left,
                        top,
                        width,
                        height
                    )

                    replaced_count += 1

        prs.save(output_path)

        with open(output_path, "rb") as f:
            result = f.read()

        return result, replaced_count

    finally:
        shutil.rmtree(temp_dir)


# ----------------------------------------------------------
# Word processing
# ----------------------------------------------------------

def replace_logo_docx(input_bytes, old_logo_bytes, new_logo_bytes, threshold=5):

    old_hash = get_image_hash(old_logo_bytes)
    new_logo_png = convert_new_logo(new_logo_bytes)

    output_buffer = BytesIO()
    replaced_count = 0

    with zipfile.ZipFile(BytesIO(input_bytes), "r") as zin, \
         zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():

            data = zin.read(item.filename)

            if item.filename.lower().startswith("word/media/"):

                try:
                    img_hash = get_image_hash(data)

                    if old_hash - img_hash <= threshold:

                        zout.writestr(item.filename, new_logo_png)
                        replaced_count += 1
                        continue

                except Exception:
                    pass

            zout.writestr(item.filename, data)

    return output_buffer.getvalue(), replaced_count


# ----------------------------------------------------------
# PDF processing
# ----------------------------------------------------------

def replace_logo_pdf(input_bytes, old_logo_bytes, new_logo_bytes, threshold=5):

    import fitz

    old_hash = get_image_hash(old_logo_bytes)
    new_logo_png = convert_new_logo(new_logo_bytes)

    replaced_count = 0

    doc = fitz.open(stream=input_bytes, filetype="pdf")

    try:

        replaced_xrefs = set()

        for page in doc:

            for img in page.get_images(full=True):

                xref = img[0]

                if xref in replaced_xrefs:
                    continue

                try:

                    extracted = doc.extract_image(xref)
                    img_hash = get_image_hash(extracted["image"])

                    if old_hash - img_hash <= threshold:

                        doc.replace_image(xref, stream=new_logo_png)

                        replaced_xrefs.add(xref)
                        replaced_count += 1

                except Exception:
                    continue

        return doc.tobytes(garbage=4, deflate=True), replaced_count

    finally:
        doc.close()


# ----------------------------------------------------------
# Dispatcher
# ----------------------------------------------------------

def replace_logo_any(input_bytes, filename, old_logo_bytes, new_logo_bytes, threshold):

    ext = os.path.splitext(filename.lower())[1]

    if ext == ".pptx":
        return replace_logo_pptx(input_bytes, old_logo_bytes, new_logo_bytes, threshold)

    if ext in [".docx", ".doc"]:
        return replace_logo_docx(input_bytes, old_logo_bytes, new_logo_bytes, threshold)

    if ext == ".pdf":
        return replace_logo_pdf(input_bytes, old_logo_bytes, new_logo_bytes, threshold)

    raise ValueError("Unsupported file type")


# ----------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------

st.set_page_config(page_title="Document Logo Replacer", page_icon="🖼️")

st.title("🖼️ Document Logo Replacer")

doc_file = st.file_uploader("Upload document", type=["pptx", "docx", "doc", "pdf"])
old_logo = st.file_uploader("Upload OLD logo", type=["png", "jpg", "jpeg"])
new_logo = st.file_uploader("Upload NEW logo", type=["png", "jpg", "jpeg"])

threshold = st.slider("Match sensitivity", 0, 20, 5)

if st.button("Replace Logo"):

    if not doc_file or not old_logo or not new_logo:
        st.warning("Please upload all required files.")
        st.stop()

    input_bytes = doc_file.read()
    old_logo_bytes = old_logo.read()
    new_logo_bytes = new_logo.read()

    with st.spinner("Processing..."):

        output_bytes, count = replace_logo_any(
            input_bytes,
            doc_file.name,
            old_logo_bytes,
            new_logo_bytes,
            threshold
        )

    if count == 0:
        st.warning("No matching logos were found.")
    else:
        st.success(f"Replaced {count} logos successfully.")

    st.download_button(
        "Download Updated Document",
        output_bytes,
        file_name="updated_" + doc_file.name
    )