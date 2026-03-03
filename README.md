# 🎨 PowerPoint Rebranding Tool

> Batch-apply a new company template to every `.pptx` in a ZIP — in seconds.

![Python](https://img.shields.io/badge/Python-3.11-blue?logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.35%2B-FF4B4B?logo=streamlit&logoColor=white)
![python-pptx](https://img.shields.io/badge/python--pptx-0.6.23%2B-green)
![Docker](https://img.shields.io/badge/Docker-ready-2496ED?logo=docker&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow)

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [How It Works](#how-it-works)
- [Tech Stack](#tech-stack)
- [Project Structure](#project-structure)
- [Getting Started](#getting-started)
  - [Local Development](#local-development)
  - [Docker](#docker)
- [Deploying to Render](#deploying-to-render)
- [Usage Guide](#usage-guide)
- [Configuration](#configuration)
- [License](#license)

---

## Overview

The **PowerPoint Rebranding Tool** is a self-hosted web application that automates the process of applying a new corporate identity to existing PowerPoint presentations.

Upload a ZIP of legacy `.pptx` files and your new company template — the tool re-skins every slide with the template's decorative shapes, colour scheme, and fonts, while preserving all original content (titles, body text, bullet points) exactly as written.

---

## Features

| Feature | Description |
|---|---|
| **Batch processing** | Rebrand an entire folder of `.pptx` files in a single upload |
| **Template-faithful design** | Copies decorative shapes (rectangles, ovals, logos) placed directly on template slides |
| **Smart content injection** | Matches source placeholders and TextBoxes to the template's content slots by index, type, and position |
| **No repair dialogs** | Works entirely within a single OPC package — PowerPoint never prompts to repair the file |
| **Theme colour preservation** | Strips inline colour overrides so the new theme's palette shows through |
| **Font enforcement** | Reads heading/body fonts from the template's theme and applies them to all text |
| **Overflow protection** | Enables word-wrap and `normAutofit` on every text frame |
| **Live progress log** | Terminal-style log streamed directly in the browser during processing |
| **Zero data retention** | All files processed in memory — nothing is written to disk |

---

## How It Works

```
 ┌─────────────────────┐     ┌──────────────────────┐
 │  old_presentations  │     │  company_template     │
 │       .zip          │     │       .pptx           │
 └────────┬────────────┘     └──────────┬───────────┘
          │                             │
          ▼                             ▼
 ┌─────────────────────────────────────────────────┐
 │              template_applier.py                │
 │                                                 │
 │  1. Load template as dst_prs (all decorative    │
 │     shapes & master remain intact)              │
 │                                                 │
 │  2. Adjust slide count:                         │
 │     • Too few  → duplicate last template slide  │
 │     • Too many → remove excess slides           │
 │                                                 │
 │  3. For each slide pair (template ↔ source):    │
 │     a. Clear template placeholder content       │
 │     b. Inject source title / body text          │
 │     c. Copy source media (images, charts)       │
 │     d. Strip old inline colour overrides        │
 │                                                 │
 │  4. Apply theme fonts & word-wrap               │
 └─────────────────────────────────────────────────┘
          │
          ▼
 ┌─────────────────────┐
 │  rebranded_files    │
 │       .zip          │
 └─────────────────────┘
```

The key design principle is **single-package operation**: the template itself is the destination. Its slides already contain every decorative shape — no cross-package OPC copies are made, which is why output files open cleanly in PowerPoint with no repair prompts.

---

## Tech Stack

| Layer | Library | Purpose |
|---|---|---|
| Presentation manipulation | [`python-pptx`](https://python-pptx.readthedocs.io/) | OPC package access, slide/shape/XML editing |
| XML processing | [`lxml`](https://lxml.de/) | Low-level element tree operations |
| Web UI | [`Streamlit`](https://streamlit.io/) | Browser-based upload / download interface |
| Image handling | [`Pillow`](https://python-pillow.org/) | Used internally by python-pptx |
| Containerisation | [Docker](https://www.docker.com/) | Consistent deployment environment |

---

## Project Structure

```
Powerpoint Automation/
│
├── app.py                   # Streamlit web UI
├── template_applier.py      # Core rebranding engine
│
├── requirements.txt         # Python dependencies
├── Dockerfile               # Container definition (Render-ready)
│
├── .streamlit/
│   └── config.toml          # Streamlit server config (headless, CORS off)
│
└── README.md
```

---

## Getting Started

### Prerequisites

- Python **3.11+**
- `pip`

### Local Development

```bash
# 1. Clone the repository
git clone https://github.com/your-org/pptx-rebranding-tool.git
cd pptx-rebranding-tool

# 2. Create and activate a virtual environment
python -m venv .venv

# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Launch the app
streamlit run app.py
```

The app will open at **http://localhost:8501**.

---

### Docker

#### Build the image

```bash
docker build -t pptx-rebrand .
```

#### Run locally

```bash
docker run -p 8501:10000 pptx-rebrand
```

Open **http://localhost:8501** in your browser.

#### Environment variables

| Variable | Default | Description |
|---|---|---|
| `PORT` | `10000` | Port Streamlit listens on (injected automatically by Render) |

---

## Deploying to Render

The `Dockerfile` is pre-configured for Render's Docker runtime.

### Step-by-step

1. **Push your code** to a GitHub (or GitLab) repository.

2. **Create a new Web Service** on [render.com](https://render.com):
   - Go to **Dashboard → New → Web Service**
   - Connect your repository

3. **Configure the service**:

   | Setting | Value |
   |---|---|
   | **Runtime** | Docker |
   | **Branch** | `main` (or your default branch) |
   | **Instance type** | Starter (512 MB RAM) or higher |
   | **Health Check Path** | `/_stcore/health` |

4. **Environment variables** — Render injects `PORT` automatically; no manual setup needed.

5. Click **Create Web Service**. Render builds the Docker image and deploys.

> **Note on memory:** For ZIP archives larger than ~50 MB, upgrade to a **Standard instance (2 GB RAM)** to avoid out-of-memory errors during processing.

---

## Usage Guide

1. **Prepare your inputs**
   - Create a `.zip` archive containing all `.pptx` files you want to rebrand.
   - Have your new company template `.pptx` ready (the file whose design/theme you want applied).

2. **Upload**
   - **Step 1 — Old Presentations:** Upload the ZIP file.
   - **Step 2 — New Template:** Upload the template `.pptx`.

3. **Rebrand**
   - Click **🚀 Rebrand Presentations**.
   - Watch the live log as each file is processed.

4. **Download**
   - Click **⬇️ Download Rebranded ZIP** to save the output archive.
   - The output ZIP preserves the original filenames.

### Tips for best results

- Place your company's decorative shapes (logos, branded rectangles, coloured ovals) **directly on the template's slides**, not only in the Slide Master — the tool uses the slides themselves as the visual base.
- The tool maps content top-to-bottom by vertical position. Ensure your template's TextBoxes are ordered: **title at the top**, body/content below.
- If your source presentations have more slides than the template, the tool **duplicates the last template slide** to fill the gap — no blank, un-styled slides.

---

## Configuration

### `.streamlit/config.toml`

```toml
[server]
headless = true              # Required for Docker / Render
enableCORS = false           # Safe behind Render's reverse proxy
enableXsrfProtection = false

[browser]
gatherUsageStats = false     # Opt out of Streamlit telemetry
```

---

## License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

---

<p align="center">
  Built with <a href="https://python-pptx.readthedocs.io/">python-pptx</a> &amp;
  <a href="https://streamlit.io/">Streamlit</a>
</p>
