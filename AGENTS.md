# AGENTS.md — Presentation Project Rules

## Slide Design Standards
- **Background:** Clean solid white (`#FFFFFF`).
- **Aspect Ratio:** 16:9 widescreen (`13.333` inches x `7.5` inches).
- **Typography:**
  - Font family: Calibri / Arial.
  - Colors: High-contrast charcoal text (`#111827`), secondary text (`#4B5563`), blue accent (`#2563EB`).
  - Slide Titles: 26pt bold.
  - Section Tags: 11pt bold uppercase accent.
- **Layout:** Strictly 1-column, 2-column, or 3-column card containers with subtle borders (`#E2E8F0`).
- **Demo Breaks:** Slides 10 and 15 use amber highlight cards to clearly signal live demonstration interludes.

## Animation Standards
- **Sequential Reveal Strategy:**
  - **Slides with > 2 Bounding Boxes (3 or 4 cards):** Every bounding box appears on click as an indivisible unit (card shape + all text inside).
  - **Slides with 1 or 2 Bounding Boxes:** Every bounding box appears on click (container + title/subtitle), and each bullet point within it appears on its own sequential click.
  - **Non-bullet cards (code blocks, image/diagram placeholders):** The entire container and internal content appear together on click.
  - **Title Slide:** Unanimated (visible immediately on load).
- **Implementation via `pptx_animations`:**
  - Do not write raw XML in slide generators. Use the high-level helper module:
    `from pptx_animations import CardInfo, apply_card_animations`
  - Attach animations to every slide using `apply_card_animations(slide, [CardInfo(...)])`.

## Environment & Tooling
- Python environment: `./env/bin/python` (Conda environment built with Python 3.11, `python-pptx`, `pillow`).
- Regeneration script: `build_deck.py`. Output is saved to `antigravity_lab_presentation.pptx`.
- Animation module: `pptx_animations.py`.
