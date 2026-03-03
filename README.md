# PowerPoint Rebranding Tool

A Streamlit web app that applies a new company template to a batch of
legacy `.pptx` presentations in one click.

---

## Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the app

```bash
streamlit run app.py
```

The app opens automatically at **http://localhost:8501**

---

## Usage

| Step | Action |
|------|--------|
| 1 | Upload a **ZIP archive** containing your old `.pptx` files |
| 2 | Upload the **new company template** `.pptx` |
| 3 | Click **Rebrand Presentations** |
| 4 | Download the output ZIP with all rebranded files |

---

## What gets changed

- **Slide master** replaced with the new template's master (theme colours, fonts, placeholder positions, logo, background).
- **Slide layouts** remapped to the closest-named layout in the new master.
- **Heading font** (major font) applied to title/heading placeholders.
- **Body font** (minor font) applied to all other text runs.
- **Word-wrap / auto-shrink** enabled on every text box to prevent overflow.
- Slide size synced to the template's dimensions.

## What is preserved

- All text content
- All images and pictures
- All shapes, lines, and SmartArt containers
- All tables and charts

---

## File structure

```
Powerpoint Automation/
├── app.py               # Streamlit UI
├── template_applier.py  # Core rebranding logic
├── requirements.txt     # Python dependencies
└── README.md            # This file
```
