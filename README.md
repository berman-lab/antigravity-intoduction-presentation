# Antigravity Lab Presentation

This repository contains the automated pipeline for generating the slide presentation for the Berman Lab meeting on **Antigravity and Agentic AI in the Wet Lab**.

The presentation is formatted for 16:9 widescreen, with a clean white background, high-contrast typography, and strict 1- or 2-column layouts designed for maximum readability.

---

## Output File

- **PowerPoint Presentation:** [`antigravity_lab_presentation.pptx`](file:///Users/bermanlab/dev/agy_lab_presentation/antigravity_lab_presentation.pptx)
  - 18 slides formatted in widescreen (16:9)
  - Includes 2 dedicated breaks for live demonstrations
  - Compatible with Microsoft PowerPoint, Apple Keynote, and Google Slides (upload directly)
- **Interactive Context Visualizer:** [`context_visualizer.html`](file:///Users/bermanlab/dev/agy_lab_presentation/context_visualizer.html)
  - Interactive simulator demonstrating context window accumulation, token hygiene, and compaction.

---

## Environment & Reproduction

The presentation generation script uses a local Conda environment with `python-pptx` and `pillow`.

### 1. Conda Environment
To recreate or activate the local environment:
```bash
# Created via Miniconda3:
/Users/bermanlab/miniconda3/bin/conda create -p ./env -c conda-forge python=3.11 python-pptx pillow -y
```

### 2. Regenerate Slides
Whenever you edit slide contents in `build_deck.py`, run:
```bash
./env/bin/python build_deck.py
```

---

## Presentation Structure (18 Slides)

### Part I: The Paradigm Shift (Chatbots vs. Agents)
1. **Title Slide:** *Antigravity in the Wet Lab — Your On-Call Junior Bioinformatician*
2. **The Wet-Lab Reality & Computational Bottleneck:** Data generation vs. copy-paste chatbot friction.
3. **Web Chatbots vs. Antigravity Platform:** Contrasting disconnected text generation with autonomous execution loops.
4. **The Mental Model:** Senior Scientist (biology & QC) + Junior Bioinformatician (code & debugging).

### Part II: What Did I Do With AGY? (Real Scientific Wins)
5. **Project 1: The Autonomous Paper Reproducer:** From genomics paper PDF/URL to executed reproduction pipeline.
6. **[LIVE DEMO 1] Paper Reproducer:** Live walkthrough of the genomics reproduction workspace.
7. **Project 2: The Petite Paper Figure Generator:** Consolidating 5 bloated Jupyter notebooks into 1 script with Google Sheets palette sync.
8. **The 20-Minute Mystery:** How AGY independently diagnosed and resolved Anastasia's silent Windows Matplotlib crash.
9. **[LIVE DEMO 2] Petite Paper Figure Generator:** Live demonstration of single-command figure generation and palette syncing.

### Part III: The Scientist in the Loop (Mindset & Quality Control)
10. **Domain Expertise & Ground Truth:** The Johannes proteomics normalization case and keeping the agent honest.
11. **The Golden Rule (Code-First Figures):** Why Illustrator is un-reproducible and code-defined figures win.

### Part IV: Best Practices & The AGY Toolkit
12. **Context Windows & Token Hygiene:** Working memory, attention dilution, and the interactive visualizer (`context_visualizer.html`).
13. **Institutional Memory (`AGENTS.md`):** Defining persistent Candida strain aliases, palettes, and folder rules.
14. **Power Steering (Slash Commands & Skills):** Using `/plan`, `/grill-me`, `/learn`, and custom skills.
15. **Safety, Sandboxing & The Git Safety Net:** Folder boundaries, immutable raw data, and parallel agents via Git Worktrees.

### Part V: The Path Forward (Berman Lab Action Plan)
16. **Immediate Opportunities (Lab Tools):** Revamping OD Normalizer, DDA panel creator (Ze'ev's challenge!), Scanalyzer, and diskImageR.
17. **Personal Callouts & Model Economics:** Tailored advice for Shani, Sofya, Alona, and TAU academic licensing.
18. **Getting Started & Summary Checklist:** 3 starter projects to build confidence this week.
