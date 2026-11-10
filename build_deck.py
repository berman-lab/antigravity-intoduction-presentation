"""
Generate Antigravity Lab Presentation (.pptx)
Synthesized from Presentation notes.md and build_deck.py.
Widescreen 16:9, clean white background, high contrast, 1-2-3 column layout.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx_animations import CardInfo, apply_card_animations

# --- Presentation Setup ---
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_slide_layout = prs.slide_layouts[6]  # Blank layout

# --- Color Palette ---
COLOR_BG = RGBColor(255, 255, 255)
COLOR_TEXT_PRIMARY = RGBColor(17, 24, 39)      # Charcoal / near-black
COLOR_TEXT_SECONDARY = RGBColor(75, 85, 99)    # Medium gray
COLOR_TEXT_MUTED = RGBColor(107, 114, 128)    # Lighter gray
COLOR_ACCENT = RGBColor(37, 99, 235)           # Blue accent
COLOR_ACCENT_BG = RGBColor(239, 246, 255)      # Light blue tint
COLOR_CARD_BG = RGBColor(248, 250, 252)        # Off-white / light slate
COLOR_CARD_BORDER = RGBColor(226, 232, 240)    # Slate border
COLOR_ALERT_BG = RGBColor(254, 243, 199)       # Warm amber for demo breaks
COLOR_ALERT_BORDER = RGBColor(245, 158, 11)    # Amber border
COLOR_CODE_BG = RGBColor(241, 245, 249)        # Code background
COLOR_SUCCESS = RGBColor(16, 185, 129)         # Green accent
COLOR_DANGER = RGBColor(239, 68, 68)           # Red accent

FONT_FAMILY = "Calibri"


def add_header(slide, section_tag, title, subtitle=None):
    """Adds a standard clean header to a slide."""
    header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.733), Inches(1.3))
    tf = header_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    # Section Tag
    p_tag = tf.paragraphs[0]
    p_tag.text = section_tag.upper()
    p_tag.font.name = FONT_FAMILY
    p_tag.font.size = Pt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_ACCENT
    p_tag.space_after = Pt(4)

    # Title
    p_title = tf.add_paragraph()
    p_title.text = title
    p_title.font.name = FONT_FAMILY
    p_title.font.size = Pt(26)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_TEXT_PRIMARY

    # Subtitle
    if subtitle:
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle
        p_sub.font.name = FONT_FAMILY
        p_sub.font.size = Pt(14)
        p_sub.font.color.rgb = COLOR_TEXT_SECONDARY
        p_sub.space_before = Pt(4)


def add_card(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER):
    """Creates a subtle rounded rectangle container card."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


# ==============================================================================
# SLIDE 1: Title Slide
# ==============================================================================
s1 = prs.slides.add_slide(blank_slide_layout)
tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(4.5))
tf1 = tb1.text_frame
tf1.word_wrap = True

p = tf1.paragraphs[0]
p.text = "ANTIGRAVITY IN THE WET LAB"
p.font.name = FONT_FAMILY
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(12)

p = tf1.add_paragraph()
p.text = "From Chatbot to Colleague:"
p.font.name = FONT_FAMILY
p.font.size = Pt(40)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_PRIMARY

p = tf1.add_paragraph()
p.text = "Your On-Call Junior Bioinformatician"
p.font.name = FONT_FAMILY
p.font.size = Pt(36)
p.font.bold = True
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(24)

p = tf1.add_paragraph()
p.text = "Accelerating Candida research from raw assays to reproducible publication figures"
p.font.name = FONT_FAMILY
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_SECONDARY
p.space_after = Pt(40)

p = tf1.add_paragraph()
p.text = "Berman Lab • Yeast Biology Meeting"
p.font.name = FONT_FAMILY
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_MUTED


# ==============================================================================
# SLIDE 2: The Wet-Lab Reality & The Computational Bottleneck
# ==============================================================================
s2 = prs.slides.add_slide(blank_slide_layout)
add_header(s2, "Part I: The Paradigm Shift", "The Wet-Lab Reality & Computational Bottleneck",
           "Everyone in the lab generates high-throughput data; turning it into figures is the real grind.")

card_s2_left = add_card(s2, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s2_left = s2.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s2_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Scientific Reality in Berman Lab"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

bullets_s2_left = [
    ("Rich Assay Streams:", "Growth curves, disk diffusion, spot assays, colony growth, proteomics, and whole-genome sequencing in Candida albicans & C. glabrata."),
    ("Custom Scientific Needs:", "No 'one-size-fits-all' script exists. Everyone has distinct experimental conditions, strain sets, and visual requirements."),
    ("The Bioinformatics Chokepoint:", "Access to dedicated bioinformaticians is limited and impractical for daily figure tweaks and exploratory plotting."),
    ("Industry Contrast:", "Software engineers and bioinformaticians haven't written manual boilerplate in over a year—AI agents do the typing.")
]
for title, desc in bullets_s2_left:
    p = tf.add_paragraph()
    p.text = f"• {title} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = desc
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s2_right = add_card(s2, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s2_right = s2.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s2_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Web Chatbot Bottleneck"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

bullets_s2_right = [
    ("The Copy-Paste Carousel:", "Prompt in browser -> paste code into R/Python script -> execute -> hit traceback -> paste error back to chatbot."),
    ("Creeping Technical Debt:", "Chatbots generate isolated snippets without seeing prior architecture, leading to bloated, unmaintainable code."),
    ("Palette Inconsistencies:", "Figures across panels have mismatched fonts, conflicting color schemes, and fragile manual tweaks."),
    ("The Real Hidden Cost:", "Researchers run smaller, fewer experiments simply to avoid the painful computational backlog.")
]
for title, desc in bullets_s2_right:
    p = tf.add_paragraph()
    p.text = f"• {title} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = desc
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s2, [
    CardInfo(card_s2_left, tb_s2_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s2_right, tb_s2_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 3: Web Chatbots vs. Antigravity Platform
# ==============================================================================
s3 = prs.slides.add_slide(blank_slide_layout)
add_header(s3, "Part I: The Paradigm Shift", "Web Chatbots vs. Autonomous Coding Agents",
           "A fundamental architectural leap from answering questions to achieving goals.")

card_s3_left = add_card(s3, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s3_left = s3.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s3_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Web Chatbots (ChatGPT / Claude.ai)"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

cb_points = [
    ("Disconnected from Workspace:", "Cannot see your raw CSVs, plate layouts, image directories, or existing scripts."),
    ("Copy-Paste Friction:", "You act as a manual courier carrying code back and forth between terminal and browser."),
    ("Execution in a Vacuum:", "Generates code without testing whether packages exist, inputs parse, or arrays align."),
    ("Mental Model:", "Like texting a friend for advice who cannot see your computer screen.")
]
for title, desc in cb_points:
    p = tf.add_paragraph()
    p.text = f"• {title} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = desc
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s3_right = add_card(s3, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb_s3_right = s3.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s3_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Antigravity Development Environment"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

ag_points = [
    ("Native Workspace Access:", "Directly reads CSVs, FASTA files, plate reader exports, and project scripts."),
    ("Autonomous Local Execution:", "Executes scripts directly inside your local Conda environment and inspects stdout/stderr."),
    ("Self-Correction Looping:", "If a script throws a traceback or bad shape, AGY reads the error, modifies code, and re-runs. Goal-oriented!"),
    ("Mental Model:", "Like having a dedicated junior bioinformatician working right at your desk.")
]
for title, desc in ag_points:
    p = tf.add_paragraph()
    p.text = f"• {title} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = desc
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s3, [
    CardInfo(card_s3_left, tb_s3_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s3_right, tb_s3_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 4: The Mental Model: Senior Scientist + Junior Bioinformatician
# ==============================================================================
s4 = prs.slides.add_slide(blank_slide_layout)
add_header(s4, "Part I: The Paradigm Shift", "Mental Model: Division of Scientific Labor",
           "You provide the biological hypothesis; the agent provides computational muscle.")

card_w = Inches(3.64)
card_h = Inches(4.8)

# Card 1: Your Role
card_s4_1 = add_card(s4, Inches(0.8), Inches(2.0), card_w, card_h)
tb_s4_1 = s4.shapes.add_textbox(Inches(1.0), Inches(2.2), card_w - Inches(0.4), card_h - Inches(0.4))
tf = tb_s4_1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "1. You (Senior Scientist)"
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(12)
items1 = [
    ("Biological Hypothesis", "What question are we testing with this mutant or drug combination?"),
    ("Experimental Context", "Which strains, drug concentrations, blank wells, replicates?"),
    ("Mathematical Logic", "Doubling times, MIC50, baseline normalization, log2 fold-changes."),
    ("Quality Control", "Reviewing figures, catching biological anomalies, auditing logic.")
]
for t, d in items1:
    p = tf.add_paragraph()
    p.text = f"• {t}:\n  {d}"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_SECONDARY
    p.space_after = Pt(8)

# Card 2: The Agent's Role
card_s4_2 = add_card(s4, Inches(4.84), Inches(2.0), card_w, card_h, bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb_s4_2 = s4.shapes.add_textbox(Inches(5.04), Inches(2.2), card_w - Inches(0.4), card_h - Inches(0.4))
tf = tb_s4_2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "2. Agent (Junior Bioinformatician)"
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(12)
items2 = [
    ("Syntax & Libraries", "Pandas, Matplotlib, Seaborn, Scipy, Statannotations, Biopython."),
    ("Data Wrangling", "Parsing messy plate reader CSVs, melting & transposing tables."),
    ("Execution & Debugging", "Running code, reading tracebacks, fixing cross-platform bugs."),
    ("Styling & Formatting", "300-DPI publication PNGs, vector SVGs, font scaling, palette sync.")
]
for t, d in items2:
    p = tf.add_paragraph()
    p.text = f"• {t}:\n  {d}"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_SECONDARY
    p.space_after = Pt(8)

# Card 3: The Result
card_s4_3 = add_card(s4, Inches(8.88), Inches(2.0), card_w, card_h)
tb_s4_3 = s4.shapes.add_textbox(Inches(9.08), Inches(2.2), card_w - Inches(0.4), card_h - Inches(0.4))
tf = tb_s4_3.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "3. The Scientific Result"
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_SUCCESS
p.space_after = Pt(12)
items3 = [
    ("10x Faster Iteration", "Change a palette or statistical test in 15 seconds across all figures."),
    ("Auditable Provenance", "Every figure backed by committed, deterministic Git code."),
    ("Democratized Analysis", "Every bench scientist runs sophisticated pipelines effortlessly."),
    ("Zero 'Vibe Coding'", "Code is tested, structured, reproducible, and verifiable.")
]
for t, d in items3:
    p = tf.add_paragraph()
    p.text = f"• {t}:\n  {d}"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_SECONDARY
    p.space_after = Pt(8)

apply_card_animations(s4, [
    CardInfo(card_s4_1, tb_s4_1),
    CardInfo(card_s4_2, tb_s4_2),
    CardInfo(card_s4_3, tb_s4_3)
])


# ==============================================================================
# SLIDE 5: Project 1: The Paper Reproducer
# ==============================================================================
s5 = prs.slides.add_slide(blank_slide_layout)
add_header(s5, "Part II: What Did I Do With AGY?", "Project 1: The Autonomous Paper Reproducer",
           "Taking a genomics paper (PDF or URL) and autonomously generating a pipeline to reproduce its findings.")

card_s5_left = add_card(s5, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s5_left = s5.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s5_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Autonomous Pipeline Orchestration"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

repro_steps = [
    ("1. Tool Extraction & Environment Setup:", "Extracted legacy bioinformatic tools and versions from the text; patched old C/C++ code to compile on modern OS."),
    ("2. High-Speed Data Ingestion:", "Identified accession numbers, queried repositories (SRA/ENA), and optimized parallel download speeds."),
    ("3. Branching Pipeline Assembly:", "Determined execution order, input/output piping, and parameter flags without human coding."),
    ("4. Figure Comparison:", "Generated final figures and benchmarked them against the published paper's figures."),
    ("5. Hardware Adaptation:", "Optimized memory and storage footprints for personal Mac vs. HPC SLURM cluster.")
]
for title, desc in repro_steps:
    p = tf.add_paragraph()
    p.text = f"• {title} "
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(6)
    run = p.add_run()
    run.text = desc
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

# Right Card: Placeholder for Branching Pathway & Figure
card_s5_right = add_card(s5, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb_s5_right = s5.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s5_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Pipeline Pathway & Figure Result"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

# Placeholder box
card_s5_ph = add_card(s5, Inches(7.2), Inches(3.0), Inches(5.0), Inches(3.4), bg_color=RGBColor(255, 255, 255), border_color=COLOR_ACCENT)
tb_s5_ph = s5.shapes.add_textbox(Inches(7.4), Inches(3.3), Inches(4.6), Inches(2.8))
tf_ph = tb_s5_ph.text_frame
tf_ph.word_wrap = True
p = tf_ph.paragraphs[0]
p.text = "[IMAGE PLACEHOLDER: Branching Pipeline & Output Figure]"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(8)

p = tf_ph.add_paragraph()
p.text = "• Flowchart of the branching execution pathway inferred by AGY.\n• Preliminary reproduced figure comparing output to the paper's original figure.\n\n(Drop your exported PNG or SVG screenshot directly into this frame)."
p.font.size = Pt(11)
p.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s5, [
    CardInfo(card_s5_left, tb_s5_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s5_right, tb_s5_right, header_paras=1, is_bullet_card=False, extra_shapes=[card_s5_ph, tb_s5_ph])
])


# ==============================================================================
# SLIDE 6: LIVE DEMO BREAK: Paper Reproducer
# ==============================================================================
s6 = prs.slides.add_slide(blank_slide_layout)
card_s6 = add_card(s6, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9), bg_color=COLOR_ALERT_BG, border_color=COLOR_ALERT_BORDER)
tb_s6 = s6.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(10.9), Inches(5.1))
tf = tb_s6.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "LIVE DEMONSTRATION 1"
p.font.bold = True
p.font.size = Pt(14)
p.font.color.rgb = COLOR_ALERT_BORDER
p.space_after = Pt(8)

p = tf.add_paragraph()
p.text = "LIVE DEMO: The Autonomous Paper Reproducer"
p.font.bold = True
p.font.size = Pt(32)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(16)

p = tf.add_paragraph()
p.text = "Walking through the genomics reproduction workspace, pipeline DAG, and preliminary figures."
p.font.size = Pt(16)
p.font.color.rgb = COLOR_TEXT_SECONDARY
p.space_after = Pt(28)

demo1_items = [
    ("The Paper Directive:", "Showing how AGY parsed the paper PDF and identified requisite toolchains."),
    ("Legacy Code Patching:", "Inspecting source code patches applied by the agent to get legacy bioinformatic tools compiling."),
    ("Branching Execution Graph:", "Viewing the pipeline dependency graph connecting raw FASTQ reads to final alignments."),
    ("Figure Output:", "Displaying the preliminary reproduced figure generated without manual coding.")
]
for t, d in demo1_items:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(10)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s6, [
    CardInfo(card_s6, tb_s6, header_paras=3, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 7: Project 2: The Petite Paper Figure Generator
# ==============================================================================
s7 = prs.slides.add_slide(blank_slide_layout)
add_header(s7, "Part II: What Did I Do With AGY?", "Project 2: The Petite Paper Figure Generator",
           "Rescuing years of Candida data dispersed across 5 bloated Jupyter notebooks.")

card_s7_left = add_card(s7, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s7_left = s7.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s7_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Real-World Chaos"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

chaos_points = [
    ("5 Fragmented Jupyter Notebooks:", "Years of Candida petite experiments, bloated with dead code, experimental plots, and un-tracked data paths."),
    ("Uncoordinated Color Palettes:", "Strains and conditions colored differently across notebooks. Unifying them manually meant weeks of refactoring."),
    ("The Archival Nightmare:", "Providing runnable code and data for journal publication or peer review would have been almost impossible."),
    ("Manual Refactoring Estimate:", "At least 3-4 weeks of full-time tedious coding to clean, consolidate, and harmonize.")
]
for t, d in chaos_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s7_right = add_card(s7, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb_s7_right = s7.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s7_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Antigravity Transformation"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

sol_points = [
    ("Single Point of Execution:", "Consolidated into one master pipeline: `python main.py --figure 2` produces publication-ready panels on demand."),
    ("Live Google Sheets Palette Sync:", "Color configuration is hosted on Google Sheets! Collaborators can tweak colors live, and all figures sync automatically."),
    ("Automated Data Consolidation:", "Gathered, organized, and documented all raw data into an immutable, structured `data/` directory."),
    ("Automated Aesthetic Polish:", "Legends, tick labels, font hierarchies, and visual artifacts cleaned up effortlessly. Not a single line written manually.")
]
for t, d in sol_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s7, [
    CardInfo(card_s7_left, tb_s7_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s7_right, tb_s7_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 8: The 20-Minute Mystery: Cross-Platform Self-Correction
# ==============================================================================
s8 = prs.slides.add_slide(blank_slide_layout)
add_header(s8, "Part II: What Did I Do With AGY?", "The 20-Minute Mystery: Cross-Platform Debugging",
           "A silent failure on Windows that would have taken hours to troubleshoot manually.")

card_s8_left = add_card(s8, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s8_left = s8.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s8_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Silent Crash on Windows"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

bug_points = [
    ("Flawless on macOS:", "The entire Petite pipeline ran smoothly on my Mac laptop from terminal with zero warnings."),
    ("Sharing with Anastasia:", "I sent the repository to Anastasia, who runs Windows. The script failed immediately."),
    ("The Worst Kind of Bug:", "Complete, total silence. No error message, no traceback, no crash dump. The process just exited silently."),
    ("The Human Diagnostic Time Sink:", "Debugging silent C-level or graphic backend crashes on an OS you don't use typically consumes hours or days of frustrating forum hunting.")
]
for t, d in bug_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s8_right = add_card(s8, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb_s8_right = s8.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s8_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "AGY Solved It in 20 Minutes"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_SUCCESS
p.space_after = Pt(14)

fix_points = [
    ("Tasking the Agent:", "I described the silent exit symptom on Windows and gave AGY access to the codebase."),
    ("Autonomous Investigation:", "AGY reasoned through potential culprits: font cache corruption, GUI thread contention, Matplotlib backend."),
    ("The Root Cause:", "Identified a known silent crash within the default Matplotlib Windows rendering backend under specific Python versions."),
    ("The One-Line Fix:", "Configured an environmental variable override in the startup loader. Resolved in 20 minutes with zero human frustration.")
]
for t, d in fix_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s8, [
    CardInfo(card_s8_left, tb_s8_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s8_right, tb_s8_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 9: LIVE DEMO BREAK: Petite Paper Pipeline
# ==============================================================================
s9 = prs.slides.add_slide(blank_slide_layout)
card_s9 = add_card(s9, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9), bg_color=COLOR_ALERT_BG, border_color=COLOR_ALERT_BORDER)
tb_s9 = s9.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(10.9), Inches(5.1))
tf = tb_s9.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "LIVE DEMONSTRATION 2"
p.font.bold = True
p.font.size = Pt(14)
p.font.color.rgb = COLOR_ALERT_BORDER
p.space_after = Pt(8)

p = tf.add_paragraph()
p.text = "LIVE DEMO: The Petite Paper Figure Generator"
p.font.bold = True
p.font.size = Pt(32)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(16)

p = tf.add_paragraph()
p.text = "Live walkthrough of the Petite Paper repository (~/dev/petite_paper): One script, multi-modal figures."
p.font.size = Pt(16)
p.font.color.rgb = COLOR_TEXT_SECONDARY
p.space_after = Pt(28)

demo2_items = [
    ("Single-Command Execution:", "Running `python main.py --figure 2` live to output 300-DPI publication plots."),
    ("Google Sheets Palette Sync:", "Opening Google Sheets, changing a mutant's color, re-running the command, and watching the figure update live."),
    ("Multi-Modal Figures:", "Demonstrating growth curves + MIC50 (Fig 2), MitoTracker microscopy montages (Fig 3), Proteomics PCA & GO (Fig 4)."),
    ("Clean Repository Structure:", "Showing how raw data, utility modules, and figure generators stay modular and version-controlled.")
]
for t, d in demo2_items:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(10)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s9, [
    CardInfo(card_s9, tb_s9, header_paras=3, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 10: What You Need: Domain Knowledge & Ground Truth
# ==============================================================================
s10 = prs.slides.add_slide(blank_slide_layout)
add_header(s10, "Part III: The Scientist in the Loop", "Domain Expertise & Ground Truth: Human in the Loop",
           "AI is an accelerator, not an oracle. Responsibility for scientific validity remains 100% yours.")

card_s10_left = add_card(s10, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s10_left = s10.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s10_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Johannes Proteomics Story"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

story_points = [
    ("The AI's Silent Mistake:", "During analysis of my proteomics dataset, AGY implemented an incorrect normalization formula."),
    ("Why Did It Happen?", "I hadn't explicitly explained what the raw data represented and how it had already been processed upstream."),
    ("Domain Expertise Saved the Day:", "Johannes audited the data and caught the anomaly immediately because he understood the biological context."),
    ("The Critical Lesson:", "Without domain oversight, subtle mathematical or biological bugs will slip through and corrupt conclusions. You MUST audit the logic!")
]
for t, d in story_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s10_right = add_card(s10, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s10_right = s10.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s10_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "How to Keep the Agent Honest"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_SUCCESS
p.space_after = Pt(14)

honest_points = [
    ("Demand Explanations:", "Instruct AGY: 'Explain your reasoning and document every mathematical step directly in the code comments.'"),
    ("Ground Truth Benchmarks:", "Run pipelines on known control datasets (e.g. wild-type CBS138 growth curves) to verify expected values before analyzing mutants."),
    ("Architectural Hygiene:", "Insist that parameters, strain names, and thresholds are kept in textual configs, never hardcoded in scripts."),
    ("Catching Code Bloat:", "Untended agents write redundant, duplicated code. Use occasional refactoring prompts to keep pipelines lean and token-efficient.")
]
for t, d in honest_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s10, [
    CardInfo(card_s10_left, tb_s10_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s10_right, tb_s10_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 11: The Golden Rule: Code-First, Text-First Figures
# ==============================================================================
s11 = prs.slides.add_slide(blank_slide_layout)
add_header(s11, "Part III: The Scientist in the Loop", "The Golden Rule: Code-First, Text-First Figures",
           "Never ask AI or humans to manually edit publication figures in Photoshop or Illustrator.")

card_s11_left = add_card(s11, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s11_left = s11.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s11_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Pitfalls of Manual Figure Tweaks"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

manual_points = [
    ("Un-reproducible Tweaks:", "Resizing an axis, nudging a label, or coloring a bar in Illustrator leaves zero record of how the figure was assembled."),
    ("The 'New Replicate' Disaster:", "When you perform 3 additional biological replicates, all manual adjustments must be laboriously repeated from scratch."),
    ("Visual Distortions:", "Accidentally misaligning an error bar or significance star during manual layout assembly risks data distortion."),
    ("Image Hallucinations:", "GenAI image editors 'paint pixels'—they cannot be trusted to edit scientific charts. Always work through code!")
]
for t, d in manual_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s11_right = add_card(s11, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb_s11_right = s11.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s11_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Antigravity Standard: 100% Code"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

code_points = [
    ("Every Pixel is Code-Defined:", "Axis limits, typography, significance brackets (***), error bands, and multi-panel composites are explicitly scripted."),
    ("10-Second Regeneration:", "New raw data arrives? Simply re-run `python main.py --figure 2` and the entire 300-DPI composite figure updates instantly."),
    ("Centralized Palette Sync:", "Change a strain's color hex code in your config or Google Sheet, and all 15 figures update synchronously."),
    ("Reviewer-Proof Provenance:", "When peer reviewers or editors ask for raw code and data, your Git repository provides a complete, auditable ledger.")
]
for t, d in code_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s11, [
    CardInfo(card_s11_left, tb_s11_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s11_right, tb_s11_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 12: Context Windows & Token Hygiene (with Visualizer Link)
# ==============================================================================
s12 = prs.slides.add_slide(blank_slide_layout)
add_header(s12, "Part IV: Best Practices & Toolkit", "Context Windows & Token Hygiene: Working Memory",
           "An agent's working memory is finite. Good token hygiene makes or breaks your results.")

card_s12_left = add_card(s12, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s12_left = s12.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s12_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Working Memory & Attention Dilution"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

ctx_points = [
    ("For LLMs, Only Context Exists:", "If something isn't in the active context, it does not exist. Each session starts fresh, seeded only by AGENTS.md."),
    ("The 'Marathon Chat' Trap:", "A chat running for 50 turns accumulates thousands of tokens of old outputs, error logs, and file dumps. Attention gets diluted."),
    ("Symptom: Confusion & Drift:", "The agent becomes 'lazy', forgets strain aliases or styling rules, and takes longer to respond."),
    ("Proactive /compact:", "Use `/compact` when a chunk of work finishes to summarize progress and clear token bloat. Note: usage stats exist in CLI; in UI, keep chats focused to avoid sudden auto-compaction.")
]
for t, d in ctx_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(6)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s12_right = add_card(s12, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb_s12_right = s12.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s12_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Interactive Context Visualizer"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(12)

p = tf.add_paragraph()
p.text = "Open in browser: context_visualizer.html"
p.font.bold = True
p.font.size = Pt(13)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(10)

vis_points = [
    ("Visual Token Meter:", "See system prompt (~5k), AGENTS.md (~4k), chat turns, and raw file dumps filling the context bucket."),
    ("Attention Degradation:", "Watch the attention indicator drop from 99% (Optimal) to <50% (Diluted / Hallucination Risk)."),
    ("Heavy Dump Simulation:", "Test what happens when dumping a 10,000-line raw CSV (+70k tokens) vs. running `/compact` to restore laser focus."),
    ("Rule of Thumb:", "Many short, modular conversations (one figure or task per chat) beats one giant marathon conversation every time.")
]
for t, d in vis_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(6)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s12, [
    CardInfo(card_s12_left, tb_s12_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s12_right, tb_s12_right, header_paras=2, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 13: Institutional Memory: AGENTS.md
# ==============================================================================
s13 = prs.slides.add_slide(blank_slide_layout)
add_header(s13, "Part IV: Best Practices & Toolkit", "Institutional Memory: AGENTS.md",
           "Teach your agent once; benefit across every future conversation in the project.")

card_s13_left = add_card(s13, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s13_left = s13.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s13_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "What is AGENTS.md?"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(12)

md_points = [
    ("A Single Source of Truth:", "A plain Markdown file placed in your project root. Antigravity automatically reads it at the start of every session."),
    ("Permanent Lab Invariants:", "Define your preferred color palettes (Okabe-Ito), strain naming conventions, directory paths, and figure resolutions."),
    ("Zero Repetitive Prompting:", "No need to re-type: 'Remember our WT strain is CBS138, font is Arial, and all figures must be 300 DPI' every morning."),
    ("Lab-Wide Standardization:", "Commit AGENTS.md to Git so every lab member's agent follows the exact same lab standards.")
]
for t, d in md_points:
    p = tf.add_paragraph()
    p.text = f"• {t}\n  "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

# Right Box: Sample AGENTS.md
card_s13_right = add_card(s13, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_CODE_BG)
tb_s13_right = s13.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s13_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Example: Berman Lab AGENTS.md"
p.font.bold = True
p.font.size = Pt(15)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(10)

sample_code = (
    "# Lab Project Guidelines\n\n"
    "## Visual Standards\n"
    "- Colors: Use Okabe-Ito colorblind-safe palette\n"
    "- Fonts: Arial, min 10pt for tick labels\n"
    "- Resolution: Save high-res PNG (300 DPI) & vector SVG\n\n"
    "## Candida Nomenclature\n"
    "- C. glabrata genes: *CAGL0...* format or standard alias\n"
    "- WT strain: CBS138 (ATCC 2001)\n\n"
    "## Project Layout\n"
    "- Raw data in `data/` (READ-ONLY, never overwrite)\n"
    "- Reusable utilities in `utils/`\n"
    "- Master figure outputs in `output/`"
)
p = tf.add_paragraph()
p.text = sample_code
p.font.name = "Courier New"
p.font.size = Pt(10)
p.font.color.rgb = COLOR_TEXT_PRIMARY

apply_card_animations(s13, [
    CardInfo(card_s13_left, tb_s13_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s13_right, tb_s13_right, header_paras=1, is_bullet_card=False)
])


# ==============================================================================
# SLIDE 14: Power Steering: Slash Commands & Skills
# ==============================================================================
s14 = prs.slides.add_slide(blank_slide_layout)
add_header(s14, "Part IV: Best Practices & Toolkit", "Power Steering: Slash Commands & Skills",
           "Built-in shortcuts to automate complex workflows and enforce scientific rigor.")

commands = [
    ("/plan", "Architect Before Coding",
     "Forces the agent to produce a structured, step-by-step roadmap before touching code.\n• Prevents hasty or broken implementations\n• Checkpoint: read, edit, and approve the plan before execution begins.",
     COLOR_TEXT_PRIMARY),
    ("/grill-me", "Interactive Alignment Interview",
     "Instructs the agent to interrogate YOU with targeted questions to clarify ambiguities.\n• Uncovers missing parameters (e.g., 'What units are in column 3? Which well is media blank?')\n• Ensures the agent never assumes incorrect biological logic.",
     COLOR_ACCENT),
    ("/learn", "Institutional Habit Learning",
     "Saves a successful workflow or correction into permanent agent memory / AGENTS.md.\n• 'Remember that fluconazole MIC50 is measured at 24h, not 48h'\n• Turns hard-won fixes into automatic future habits.",
     COLOR_SUCCESS),
    ("Custom Skills", "Reusable Agent Functions",
     "Skills describe specialized tasks, guardrails, and tools.\n• Example: A Palette Skill that generates master hex palettes from websites and populates configuration lists.\n• Import community skills or write your own.",
     COLOR_TEXT_SECONDARY)
]

cards_s14 = []
for i, (cmd, sub, desc, col) in enumerate(commands):
    x = Inches(0.8 + (i % 2) * 6.1)
    y = Inches(2.0 + (i // 2) * 2.5)
    card = add_card(s14, x, y, Inches(5.6), Inches(2.2))
    tb = s14.shapes.add_textbox(x + Inches(0.25), y + Inches(0.2), Inches(5.1), Inches(1.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = cmd + " "
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = col
    run = p.add_run()
    run.text = f"— {sub}"
    run.font.bold = False
    run.font.size = Pt(13)
    run.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(6)

    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.size = Pt(11)
    p2.font.color.rgb = COLOR_TEXT_SECONDARY
    cards_s14.append(CardInfo(card, tb))

apply_card_animations(s14, cards_s14)


# ==============================================================================
# SLIDE 15: Safety, Sandboxing & The Git Safety Net
# ==============================================================================
s15 = prs.slides.add_slide(blank_slide_layout)
add_header(s15, "Part IV: Best Practices & Toolkit", "Safety, Sandboxing & The Git Safety Net",
           "Work fearlessly: Protect raw data, verify execution, and track every line.")

card_s15_left = add_card(s15, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s15_left = s15.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s15_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Folder Sandboxing & Data Hygiene"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

safety_points = [
    ("Project Folder Boundary:", "AGY considers everything inside the project folder safe. It prompts you before accessing files outside it."),
    ("Immutable Raw Data:", "Always keep original plate reader and sequencing files in `data/` as read-only. Store an archival copy on Dropbox or the lab NAS."),
    ("Beware Code Execution:", "The AI runs real bash and python. It could theoretically access outside directories through scripts—always audit code before executing."),
    ("Switching Brainpower:", "Use cheap, fast models (Flash) for quick cosmetic tweaks; use deep reasoning models (Pro/Thinking) for architecture and multi-file debugging.")
]
for t, d in safety_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(6)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s15_right = add_card(s15, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s15_right = s15.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s15_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Git & Worktrees: Total Fearlessness"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_SUCCESS
p.space_after = Pt(14)

git_points = [
    ("The Universal Undo Button:", "Every change by AGY is tracked. If an experiment doesn't work, revert with a single click. No fear of broken code."),
    ("Visual Diffs (Line-by-Line):", "Review exact green/red code diffs before accepting changes."),
    ("Safe Collaboration:", "Shared repository with Anastasia: she can run figure scripts, add tests, and push updates without code collisions."),
    ("Git Worktrees (Parallel Agents):", "Run 2 or 3 independent agents on different figures simultaneously in parallel workspaces without branch conflicts.")
]
for t, d in git_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(6)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s15, [
    CardInfo(card_s15_left, tb_s15_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s15_right, tb_s15_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 16: Immediate Opportunities: Revamping Berman Lab Tools
# ==============================================================================
s16 = prs.slides.add_slide(blank_slide_layout)
add_header(s16, "Part V: The Path Forward", "Immediate Opportunities: Revamping Lab Tools",
           "Un-maintained lab tools can now be revived, modernized, and extended with AGY.")

card_s16_left = add_card(s16, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s16_left = s16.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s16_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Lab Tools Waiting for an Upgrade"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

tools_points = [
    ("OD Normalizer:", "Plate reader normalization and growth rate metrics. Can be refactored into a modern, robust CLI tool in an afternoon."),
    ("DDA Panel Creator:", "Disk diffusion assay image tiling, zone diameter measurements, and rad measurement pipelines."),
    ("Legacy Classics:", "Scanalyzer and diskImageR—vital historical tools that nobody maintains because older dependencies are broken."),
    ("The Agent Advantage:", "AGY excels at updating legacy Python/R environments, resolving broken dependencies, and writing modern tests.")
]
for t, d in tools_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s16_right = add_card(s16, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb_s16_right = s16.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s16_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Ze'ev's Homework Challenge"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

zeev_points = [
    ("The Mission:", "Feed the GitHub issues of the DDA panel creator directly to Antigravity."),
    ("The Workflow:", "Point the agent to each issue: 'Read issue #4, find where the bug occurs in the script, write a test to reproduce it, and fix it.'"),
    ("The Outcome:", "You can resolve months of accumulated bug reports and feature requests in less than a week."),
    ("Auditable Commits:", "Submit clean pull requests with green passing tests that the whole lab can trust.")
]
for t, d in zeev_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s16, [
    CardInfo(card_s16_left, tb_s16_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s16_right, tb_s16_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 17: Personal Callouts & Subscriptions
# ==============================================================================
s17 = prs.slides.add_slide(blank_slide_layout)
add_header(s17, "Part V: The Path Forward", "Personal Callouts & Model Economics",
           "Tailored recommendations for lab members and navigating subscriptions at TAU.")

card_s17_left = add_card(s17, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s17_left = s17.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s17_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Berman Lab Member Callouts"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

callouts = [
    ("Shani:", "Consolidate your web Claude analysis chats into a single dedicated AGY project. Enjoy unified figures and zero copy-paste friction."),
    ("Sofya:", "Use AGY for exploratory quality assessment on transposon insertion libraries. Let the agent parse mapping files and check insertion distributions."),
    ("Alona:", "Leverage your Claude subscription! You can plug Anthropic models (e.g. Claude 3.5 Sonnet) directly into AGY via API keys, or explore the Claude Code CLI."),
    ("The Entire Lab:", "Start treating analysis code as a permanent, shared lab asset rather than disposable one-off scripts.")
]
for t, d in callouts:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s17_right = add_card(s17, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s17_right = s17.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s17_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Model Subscriptions & TAU Access"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

sub_points = [
    ("Antigravity is Free:", "The development environment platform itself is free to download and use."),
    ("Gemini Free Tier vs. API:", "Google AI Studio provides a free tier for Gemini models (good for testing, but heavy use hits rate limits)."),
    ("TAU Academic Pricing:", "TAU offers a ~550 NIS/year Gemini Pro license (verify whether it covers developer API quotas or just Google Workspace web chat). Private developer billing runs ~75 NIS/month (~900 NIS/year)."),
    ("Model Flexibility:", "You can plug in Gemini, Claude, or OpenAI models. For our wet-lab figure pipelines, Gemini Pro is cost-effective and fully capable.")
]
for t, d in sub_points:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s17, [
    CardInfo(card_s17_left, tb_s17_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s17_right, tb_s17_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 18: Summary Checklist & Getting Started This Week
# ==============================================================================
s18 = prs.slides.add_slide(blank_slide_layout)
add_header(s18, "Part V: The Path Forward", "Getting Started: Low-Hanging Fruit for the Lab",
           "Three simple projects you can test this week to build confidence.")

card_s18_left = add_card(s18, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb_s18_left = s18.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s18_left.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "3 Starter Projects to Try This Week"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

starter_projects = [
    ("Project 1: The 'Excel Figure Makeover':", "Take an old Excel chart from a past lab meeting. Feed the raw table to Antigravity and ask it to generate a publication-grade Seaborn/Matplotlib plot with 300 DPI."),
    ("Project 2: Machine Export Parser:", "Automate parsing an annoying machine export (plate reader, qPCR, NanoDrop). Turn 30 minutes of manual formatting into a 2-second script."),
    ("Project 3: A Quick Sequence Exploration:", "Ask the agent to pull a Candida gene sequence from CGD/NCBI and check its domain architecture or homolog in S. cerevisiae.")
]
for t, d in starter_projects:
    p = tf.add_paragraph()
    p.text = f"• {t}\n  "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

card_s18_right = add_card(s18, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb_s18_right = s18.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb_s18_right.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The 5 Core Agent Habits"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

checklist = [
    ("1. Isolate Raw Data:", "Put raw files in `data/` and treat them as read-only."),
    ("2. Seed with AGENTS.md:", "Give your agent your lab rules, strain aliases, and palettes."),
    ("3. Start with /plan or /grill-me:", "Align on the scientific plan before writing any code."),
    ("4. Keep Chats Focused:", "One task/figure per session to keep working memory clean."),
    ("5. Commit to Git:", "Treat code as an auditable scientific asset."),
    ("Questions & Discussion:", "Let's open the floor and explore your current bottlenecks!")
]
for t, d in checklist:
    p = tf.add_paragraph()
    p.text = f"• {t} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(8)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s18, [
    CardInfo(card_s18_left, tb_s18_left, header_paras=1, is_bullet_card=True),
    CardInfo(card_s18_right, tb_s18_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# Save Presentation
# ==============================================================================
output_path = "antigravity_lab_presentation.pptx"
prs.save(output_path)
print(f"Presentation successfully saved to {output_path} (18 slides, 16:9 widescreen).")
