# Antigravity Lab Presentation

This repository contains the automated pipeline for generating the slide presentation for the Berman Lab meeting on **Antigravity and Agentic AI in the Wet Lab**.

The presentation is formatted for 16:9 widescreen, with a clean white background, high-contrast typography, and strict 1- or 2-column layouts designed for maximum readability.

---

## Output File

- **PowerPoint Presentation:** [`antigravity_lab_presentation.pptx`](file:///Users/bermanlab/dev/agy_lab_presentation/antigravity_lab_presentation.pptx)
  - 17 slides formatted in widescreen (16:9)
  - Includes 2 dedicated breaks for live demonstrations
  - Compatible with Microsoft PowerPoint, Apple Keynote, and Google Slides (upload directly)

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

## Presentation Structure (17 Slides)

### Part I: The Paradigm Shift (Chatbots vs. Agents)
1. **Title & Hook:** *Beyond the Chatbot: Antigravity as Your On-Call Bioinformatician*
2. **The Wet-Lab Bottleneck:** The gap between biological hypotheses and computational implementation.
3. **Chatbots vs. Agents:** Contrasting disconnected web chatbots with local autonomous coding agents.
4. **The Mental Model:** The Senior Scientist (biology, hypothesis, QC) + Junior Bioinformatician (code, syntax, debugging).

### Part II: Steering the Agent (Concepts & Best Practices)
5. **Context Windows & Token Hygiene:** Understanding working memory, avoiding marathon chats, using `/compact`.
6. **Choosing the Right Model:** Lightweight/fast models (Flash) for tweaks vs. deep thinking models (Pro) for architecture.
7. **Power Steering (Slash Commands):** Using `/plan` (think first), `/grill-me` (interrogate assumptions), and `/learn` (institutional habits).
8. **Institutional Memory (`AGENTS.md`):** Setting persistent project rules, palettes, and folder structures.
9. **The Golden Rule (Code-First Figures):** Why manual tweaking in Photoshop/Illustrator is un-reproducible; why code-generated figures win.

### Part III: Live Demo Break 1
10. **[LIVE DEMO 1] The Petite Mutant Pipeline:** Real-world walkthrough of `~/dev/petite_paper` (growth curves, microscopy, proteomics).

### Part IV: Rigor, Safety & Scaling
11. **Git & GitHub:** The scientist's universal undo button and auditable provenance ledger.
12. **Data Hygiene & Documentation:** Read-only `data/` folders, the power of `data/README.md`, and centralized paths.
13. **Advanced Scale (Git Worktrees):** Parallelizing multiple agents simultaneously across figures without branch conflicts.
14. **Prompting Like a Lab PI:** Moving from vague chatbot prompts to structured scientific directives.

### Part V: The Path Forward
15. **[LIVE DEMO 2] From Scratch in 5 Minutes:** Live interactive demo from raw unformatted CSV to published figure.
16. **What This Unlocks for Your Science:** Speed of iteration, experimental ambition/scale, publication rigor.
17. **Getting Started:** Three low-hanging fruit projects for lab members this week.
