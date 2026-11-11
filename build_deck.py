"""
Generate Antigravity Lab Presentation (.pptx) — Refined 28-Slide Edition
Synthesized from Presentation notes.md and updated user feedback.
Widescreen 16:9, clean white background, high contrast, 1-2-3 column layouts.
Includes native PowerPoint OpenXML entrance animations on every content slide.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx_animations import CardInfo, apply_card_animations

# --- Presentation Setup ---
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]  # Blank layout

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
COLOR_SUCCESS_BG = RGBColor(236, 253, 245)     # Light green tint
COLOR_DANGER = RGBColor(239, 68, 68)           # Red accent
COLOR_PURPLE = RGBColor(147, 51, 234)          # Purple accent
COLOR_PURPLE_BG = RGBColor(250, 245, 255)      # Light purple tint

FONT_FAMILY = "Calibri"


def add_header(slide, section_tag, title, subtitle=None):
    """Adds a standard clean header to a slide."""
    header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.733), Inches(1.35))
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
    p_tag.space_after = Pt(3)

    # Title
    p_title = tf.add_paragraph()
    p_title.text = title
    p_title.font.name = FONT_FAMILY
    p_title.font.size = Pt(25)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_TEXT_PRIMARY

    # Subtitle
    if subtitle:
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle
        p_sub.font.name = FONT_FAMILY
        p_sub.font.size = Pt(13)
        p_sub.font.color.rgb = COLOR_TEXT_SECONDARY
        p_sub.space_before = Pt(3)


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


def add_bullet_list(tf, card_title, bullets, title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=7):
    """Fills a text frame with a card title and structured bullet points."""
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = card_title
    p.font.name = FONT_FAMILY
    p.font.bold = True
    p.font.size = Pt(17)
    p.font.color.rgb = title_color
    p.space_after = Pt(12)

    for prefix, body in bullets:
        p_b = tf.add_paragraph()
        p_b.text = f"• {prefix} "
        p_b.font.name = FONT_FAMILY
        p_b.font.bold = True
        p_b.font.size = Pt(bullet_pt)
        p_b.font.color.rgb = COLOR_TEXT_PRIMARY
        p_b.space_after = Pt(space_after)

        run = p_b.add_run()
        run.text = body
        run.font.name = FONT_FAMILY
        run.font.bold = False
        run.font.color.rgb = COLOR_TEXT_SECONDARY


# ==============================================================================
# SLIDE 1: Title Slide (Unanimated)
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(11.333), Inches(4.8))
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
# SLIDE 2: The Wet-Lab Reality & The Programming Bottleneck
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "Part I: The Paradigm Shift", "The Wet-Lab Reality & The Programming Bottleneck",
           "Modern wet labs generate rich high-throughput data; the programming learning curve chokes discovery.")

c2_left = add_card(s2, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb2_left = s2.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb2_left.text_frame, "The High-Throughput Wet-Lab Reality", [
    ("High-Throughput Automation:", "Rich phenotypic and molecular data enabled by automation (e.g. TECAN EvoWare, automated plate readers, high-content imagers)."),
    ("Diverse Assay Streams:", "Growth curves, disk diffusion assays (DDA), spot assays, colony growth, proteomics, and whole-genome sequencing in Candida albicans & C. glabrata."),
    ("Custom Scientific Logic:", "No 'one-size-fits-all' script exists. Every project has distinct strain sets, drug combinations, and non-standard plate layouts."),
    ("The Bioinformatic Bottleneck:", "Access to dedicated bioinformaticians is limited and impractical for daily exploratory plotting and iterative figure tweaks."),
    ("Exploratory Imperative:", "Lab members must analyze and visualize their own data for lab meetings, grant applications, and publication.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=11.5, space_after=6)

c2_right = add_card(s2, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb2_right = s2.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb2_right.text_frame, "The Steep Cost of Programming", [
    ("Months of Dedicated Study:", "Mastering Python/R, pandas dataframes, matplotlib/ggplot2, and conda environments requires months to years of dedicated effort."),
    ("Not What Biologists Were Hired For:", "Hours spent debugging syntax errors, import tracebacks, and array indexing issues are hours stolen from wet-lab benchwork."),
    ("The 'Fragile Snippet' Trap:", "Picking up scripting piecemeal produces fragile, copy-pasted scripts that crash whenever column names or plate formats shift."),
    ("Constrained 'Search Space':", "Because coding is slow, researchers test fewer hypotheses, explore fewer normalizations, and settle for the first working plot."),
    ("The Industry Contrast:", "Meanwhile, industry software engineers and bioinformaticians haven't written manual boilerplate in over a year—AI agents do the typing.")
], title_color=COLOR_DANGER, bullet_pt=11.5, space_after=6)

apply_card_animations(s2, [
    CardInfo(c2_left, tb2_left, header_paras=1, is_bullet_card=True),
    CardInfo(c2_right, tb2_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 3: Web Chatbots vs. Antigravity Platform
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "Part I: The Paradigm Shift", "Web Chatbots vs. Autonomous Coding Agents",
           "A fundamental architectural leap from answering isolated questions to autonomously achieving goals.")

c3_left = add_card(s3, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb3_left = s3.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb3_left.text_frame, "Web Chatbots (ChatGPT / Claude.ai)", [
    ("Disconnected from Workspace:", "Cannot see your raw CSVs, plate reader exports, FASTA files, or existing project scripts."),
    ("The Copy-Paste Carousel:", "You act as a manual courier carrying code back and forth between terminal and browser."),
    ("Execution in a Vacuum:", "Generates code without testing whether packages exist, inputs parse cleanly, or array dimensions match."),
    ("Compounding Technical Debt:", "Generates isolated snippets without seeing prior architecture, creating fragile, unmaintainable spaghetti scripts."),
    ("Mental Model:", "Like texting a friend for coding advice who cannot see your computer screen or inspect your data.")
], title_color=COLOR_DANGER, bullet_pt=11.5, space_after=6)

c3_right = add_card(s3, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb3_right = s3.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb3_right.text_frame, "Antigravity Development Environment", [
    ("Native Workspace Access:", "Directly inspects local CSVs, FASTQ directories, plate reader files, and existing scripts on your disk."),
    ("Autonomous Local Execution:", "Executes scripts directly inside your local Conda environment and inspects real terminal stdout/stderr."),
    ("Self-Correction Looping:", "If code throws a traceback or bad shape, AGY reads the error, modifies code, and re-runs. Goal-oriented!"),
    ("Multi-Model Flexibility:", "Powered natively by Google Gemini, with full flexibility to plug in Anthropic Claude or OpenAI models."),
    ("Mental Model:", "Like having a dedicated junior bioinformatician sitting right at your desk, taking assignments and reporting results.")
], title_color=COLOR_ACCENT, bullet_pt=11.5, space_after=6)

apply_card_animations(s3, [
    CardInfo(c3_left, tb3_left, header_paras=1, is_bullet_card=True),
    CardInfo(c3_right, tb3_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 4: Division of Scientific Labor (Reworked 2-Column Layout)
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "Part I: The Paradigm Shift", "Division of Labor: Scientific Specification & Implementation",
           "You provide the precise scientific boundaries; the agent builds the computational logic and executes.")

c4_left = add_card(s4, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb4_left = s4.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb4_left.text_frame, "1. You (Senior Scientist): Specification", [
    ("Input Data Geometry & Quirks:", "Describe raw instrument exports (Tecan 96/384-well format, OD600 units, time intervals, blanks, irregular headers)."),
    ("Target Visual Hierarchy:", "Define exact plot type (growth curve, disk diffusion bar plot, volcano plot), axis ranges, labels, and 300-DPI publication format."),
    ("Mathematical & Normalization Logic:", "Specify baseline subtractions (media blanks), reference strain normalizations (WT CBS138 at t=0), and doubling time definitions."),
    ("Statistical Testing Requirements:", "Prescribe specific hypothesis tests (e.g., two-way ANOVA, Tukey HSD, Mann-Whitney), significance thresholds, and replicate groupings."),
    ("Quality Control & Ground Truth:", "Inspect outputs against biological plausibility and audit calculated metrics against known wet-lab benchmarks.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=11.5, space_after=6)

c4_right = add_card(s4, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb4_right = s4.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb4_right.text_frame, "2. Agent (Junior Bioinformatician): Execution", [
    ("Builds Ingestion & Cleaning Logic:", "Writes robust parsers to clean messy instrument exports, melt tables, transpose matrices, and validate missing values."),
    ("Selects & Chains Toolchains:", "Integrates standard scientific libraries (pandas, scipy.stats, statannotations, seaborn, matplotlib)."),
    ("Implements Mathematical Algorithms:", "Codifies normalization formulas, spline curve fitting, numerical derivatives (Vmax), and statistical tests."),
    ("Autonomous Execution & Self-Correction:", "Executes scripts in your local Conda environment, reads tracebacks, self-corrects bugs, and verifies output shapes."),
    ("Aesthetic Polish & Vector Exports:", "Formats publication-grade aesthetics, styling legends, font hierarchies, colorblind palettes, and exporting SVG/PNG.")
], title_color=COLOR_ACCENT, bullet_pt=11.5, space_after=6)

apply_card_animations(s4, [
    CardInfo(c4_left, tb4_left, header_paras=1, is_bullet_card=True),
    CardInfo(c4_right, tb4_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 5: Project 1: The Autonomous Paper Reproducer
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "Part II: What Did I Do With AGY?", "Project 1: The Autonomous Paper Reproducer",
           "Feeding a genomics paper PDF or URL to AGY and autonomously reproducing its computational findings.")

c5_left = add_card(s5, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb5_left = s5.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb5_left.text_frame, "Autonomous Pipeline Orchestration", [
    ("The Ambition:", "Take a published genomics paper and task the agent to reproduce its pipeline and figures from scratch."),
    ("Tool Extraction & OS Patching:", "Extracted legacy bioinformatic tools and specific versions; patched obsolete C/C++ code to compile on modern macOS."),
    ("High-Speed Data Ingestion:", "Extracted accession numbers, queried SRA/ENA repositories, and optimized parallel downloads."),
    ("Branching Pipeline DAG:", "Dynamically resolved tool dependencies, piping inputs into outputs across multi-step alignment workflows."),
    ("Zero Manual Code:", "The entire reproduction pipeline was designed, orchestrated, and executed without writing manual code.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=7)

c5_right = add_card(s5, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb5_right = s5.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb5_right.text_frame, "Technical Challenges Overcome", [
    ("Hardware Adaptability:", "Automatically tuned memory and thread footprints for a personal laptop vs. a high-performance SLURM cluster."),
    ("Error Recovery:", "Handled interrupted downloads, missing reference genome indices, and deprecated command-line arguments autonomously."),
    ("Benchmarking Output:", "Reconstructed published figures and benchmarked them against the original publication figures."),
    ("Proof of Principle:", "Proved that Antigravity can navigate complex, multi-tool bioinformatic workflows from raw data to plots.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=7)

apply_card_animations(s5, [
    CardInfo(c5_left, tb5_left, header_paras=1, is_bullet_card=True),
    CardInfo(c5_right, tb5_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 6: Case Study 1: The Branching Pipeline Architecture & Visual Flowchart
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
add_header(s6, "Part II: What Did I Do With AGY?", "Case Study 1: The Branching Pipeline Architecture",
           "How Antigravity mapped the paper's computational dependency graph and produced preliminary figures.")

c6_left = add_card(s6, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb6_left = s6.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb6_left.text_frame, "Pipeline Architecture & Execution DAG", [
    ("Stage 1: Toolchain Discovery:", "Parsed PDF methods, extracted exact versions, and built a custom conda environment."),
    ("Stage 2: Legacy Source Patching:", "Modified obsolete C/C++ compiler flags and headers to build cleanly on modern ARM/POSIX systems."),
    ("Stage 3: Parallel Data Streaming:", "Queried ENA mirrors with aria2c to stream raw sequencing fastqs at maximum throughput."),
    ("Stage 4: Branching Alignment:", "Chained Trimmomatic -> STAR/Bowtie2 -> featureCounts without human shell scripts."),
    ("Stage 5: Differential Expression:", "Calculated log2 fold-changes and generated preliminary validation figures.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=6)

# Right Card: Visual Flowchart Illustration
c6_right = add_card(s6, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb6_right = s6.shapes.add_textbox(Inches(7.2), Inches(2.15), Inches(5.0), Inches(0.6))
tf6_r = tb6_right.text_frame
tf6_r.word_wrap = True
p = tf6_r.paragraphs[0]
p.text = "Autonomous Pipeline Execution Flowchart"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(16)
p.font.color.rgb = COLOR_ACCENT

flow_steps = [
    ("Input: Genomics Paper (PDF / URL)", "Extracts tool versions, datasets, accession IDs", COLOR_TEXT_PRIMARY, COLOR_CARD_BG, COLOR_CARD_BORDER),
    ("Environment Setup & Tool Patching", "Patches legacy C/C++ source code for modern OS", COLOR_TEXT_PRIMARY, COLOR_CARD_BG, COLOR_CARD_BORDER),
    ("Parallel Data Ingestion (SRA / ENA)", "High-speed multi-threaded raw FASTQ streaming", COLOR_TEXT_PRIMARY, COLOR_CARD_BG, COLOR_CARD_BORDER),
    ("Branching Analysis DAG (QC -> Alignment)", "Trimmomatic -> STAR / Bowtie2 -> featureCounts", COLOR_TEXT_PRIMARY, COLOR_CARD_BG, COLOR_CARD_BORDER),
    ("Figure Reproduction & Comparison", "Generates publication plots matching original paper", COLOR_SUCCESS, COLOR_SUCCESS_BG, COLOR_SUCCESS)
]

extra_shapes_s6 = []
step_y = Inches(2.7)
box_h = Inches(0.62)
box_w = Inches(5.0)

for idx, (title_text, sub_text, text_col, fill_col, border_col) in enumerate(flow_steps):
    step_shape = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.2), step_y, box_w, box_h)
    step_shape.fill.solid()
    step_shape.fill.fore_color.rgb = fill_col
    step_shape.line.color.rgb = border_col
    step_shape.line.width = Pt(1)

    tf_s = step_shape.text_frame
    tf_s.word_wrap = True
    tf_s.margin_left = tf_s.margin_right = Inches(0.15)
    tf_s.margin_top = Inches(0.06)
    tf_s.margin_bottom = 0

    p1 = tf_s.paragraphs[0]
    p1.text = f"{idx + 1}. {title_text}"
    p1.font.name = FONT_FAMILY
    p1.font.bold = True
    p1.font.size = Pt(11)
    p1.font.color.rgb = text_col

    p2 = tf_s.add_paragraph()
    p2.text = sub_text
    p2.font.name = FONT_FAMILY
    p2.font.size = Pt(9.5)
    p2.font.color.rgb = COLOR_TEXT_SECONDARY

    extra_shapes_s6.append(step_shape)

    if idx < len(flow_steps) - 1:
        arrow = s6.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(9.55), step_y + box_h + Inches(0.02), Inches(0.3), Inches(0.16))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = COLOR_ACCENT
        arrow.line.fill.background()
        extra_shapes_s6.append(arrow)

    step_y += Inches(0.82)

apply_card_animations(s6, [
    CardInfo(c6_left, tb6_left, header_paras=1, is_bullet_card=True),
    CardInfo(c6_right, tb6_right, header_paras=1, is_bullet_card=False, extra_shapes=extra_shapes_s6)
])


# ==============================================================================
# SLIDE 7: LIVE DEMO BREAK: The Autonomous Paper Reproducer
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
c7 = add_card(s7, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9), bg_color=COLOR_ALERT_BG, border_color=COLOR_ALERT_BORDER)
tb7 = s7.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(10.9), Inches(5.1))
tf7 = tb7.text_frame
tf7.word_wrap = True

p = tf7.paragraphs[0]
p.text = "LIVE DEMONSTRATION 1"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(14)
p.font.color.rgb = COLOR_ALERT_BORDER
p.space_after = Pt(8)

p = tf7.add_paragraph()
p.text = "LIVE DEMO: The Autonomous Paper Reproducer"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(32)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

p = tf7.add_paragraph()
p.text = "Walking through the genomics reproduction workspace, pipeline DAG, and preliminary figures."
p.font.name = FONT_FAMILY
p.font.size = Pt(16)
p.font.color.rgb = COLOR_TEXT_SECONDARY
p.space_after = Pt(26)

demo1_items = [
    ("The Paper Directive:", "Showing how AGY parsed the paper PDF and identified requisite toolchains."),
    ("Legacy Code Patching:", "Inspecting source code patches applied by the agent to get legacy bioinformatic tools compiling."),
    ("Branching Execution Graph:", "Viewing the pipeline dependency graph connecting raw FASTQ reads to final alignments."),
    ("Figure Output:", "Displaying the preliminary reproduced figure generated without manual coding.")
]
for t, d in demo1_items:
    p = tf7.add_paragraph()
    p.text = f"• {t} "
    p.font.name = FONT_FAMILY
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(10)
    run = p.add_run()
    run.text = d
    run.font.name = FONT_FAMILY
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s7, [
    CardInfo(c7, tb7, header_paras=3, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 8: Project 2: Rescuing the Petite Paper Figures
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
add_header(s8, "Part II: What Did I Do With AGY?", "Project 2: Rescuing the Petite Paper Figures",
           "Overcoming years of accumulated experimental chaos across 5 bloated Jupyter notebooks.")

c8_left = add_card(s8, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb8_left = s8.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb8_left.text_frame, "The Real-World Notebook Chaos", [
    ("5 Fragmented Jupyter Notebooks:", "Years of Candida petite experiments, bloated with dead code, scratch plots, and un-tracked file paths."),
    ("Buried Production Figures:", "Publishable figures were buried among dozens of exploratory plots; re-running required hunting through hundreds of messy cells."),
    ("Uncoordinated Color Palettes:", "Mutant strains and conditions were colored differently across notebooks. Unifying them manually meant endless tedious edits."),
    ("The Archival Nightmare:", "Providing runnable code and data for journal publication or peer review would have been almost impossible.")
], title_color=COLOR_DANGER, bullet_pt=12, space_after=8)

c8_right = add_card(s8, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb8_right = s8.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb8_right.text_frame, "The Refactoring Dilemma", [
    ("3-4 Weeks of Manual Toil:", "Consolidating, cleaning, and harmonizing all 5 notebooks by hand would have consumed nearly a month of full-time work."),
    ("Cognitive Overload:", "Remembering which data files corresponded to which notebook cells created massive friction."),
    ("Risk of Error:", "Manual copy-pasting of plotting logic risks introducing subtle data transposition or mislabeling errors."),
    ("The Decision:", "Use Antigravity to autonomously consolidate the entire project into a clean, unified repository.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

apply_card_animations(s8, [
    CardInfo(c8_left, tb8_left, header_paras=1, is_bullet_card=True),
    CardInfo(c8_right, tb8_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 9: Case Study 2: The Unified Figure Generation Pipeline
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "Part II: What Did I Do With AGY?", "Case Study 2: The Unified Figure Generation Pipeline",
           "A single command produces all publication figures with live-synced Google Sheets palettes.")

c9_left = add_card(s9, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb9_left = s9.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb9_left.text_frame, "The Consolidated Architecture", [
    ("Single Point of Execution:", "Consolidated into one master pipeline: `python main.py --figure 2` produces publication-ready panels on demand."),
    ("Immutable Data Vault:", "Gathered, organized, and documented all raw data files into an immutable, structured `data/` directory."),
    ("Automated Aesthetic Polish:", "Legends, tick labels, font hierarchies, and visual artifacts aligned effortlessly without manual code."),
    ("Reviewer-Proof Archival Quality:", "Generates reproducible, auditable figure sets that can be shared with journals and collaborators immediately.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=10)

c9_right = add_card(s9, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb9_right = s9.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb9_right.text_frame, "Collaborative Google Sheets Palette Sync", [
    ("Live Cloud Configuration:", "Color configuration is hosted directly on Google Sheets, accessible to all collaborators."),
    ("Non-Technical Collaboration:", "Collaborators test and update color palettes directly in the spreadsheet without touching code."),
    ("Instant Synchronization:", "Re-running the script fetches the latest online palette configuration and applies it across all figures."),
    ("Separation of Concerns:", "Visual design is cleanly decoupled from mathematical data processing.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=10)

apply_card_animations(s9, [
    CardInfo(c9_left, tb9_left, header_paras=1, is_bullet_card=True),
    CardInfo(c9_right, tb9_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 10: The 20-Minute Mystery: Cross-Platform Self-Correction
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "Part II: What Did I Do With AGY?", "The 20-Minute Mystery: Cross-Platform Debugging",
           "Resolving a silent Windows failure that would have taken hours to troubleshoot manually.")

c10_left = add_card(s10, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb10_left = s10.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb10_left.text_frame, "The Silent Crash on Windows", [
    ("Flawless on macOS:", "The entire Petite pipeline ran smoothly on my Mac laptop from terminal with zero warnings."),
    ("Sharing with Anastasia:", "I sent the repository to Anastasia, who runs Windows. The script failed immediately."),
    ("The Worst Kind of Bug:", "Complete, total silence. No error message, no traceback, no crash dump. The process just exited silently."),
    ("The Human Diagnostic Time Sink:", "Debugging silent C-level or graphic backend crashes on an OS you don't use typically consumes hours of frustrating forum hunting.")
], title_color=COLOR_DANGER, bullet_pt=12, space_after=8)

c10_right = add_card(s10, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb10_right = s10.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb10_right.text_frame, "AGY Solved It in 20 Minutes", [
    ("Tasking the Agent:", "I described the silent exit symptom on Windows and gave AGY access to the codebase."),
    ("Autonomous Investigation:", "AGY reasoned through potential culprits: font cache corruption, GUI thread contention, Matplotlib backend."),
    ("The Root Cause:", "Identified a known silent crash within the default Matplotlib Windows rendering backend under specific Python versions."),
    ("The One-Line Fix:", "Configured an environmental variable override in the startup loader. Resolved in 20 minutes with zero human frustration.")
], title_color=COLOR_SUCCESS, bullet_pt=12, space_after=8)

apply_card_animations(s10, [
    CardInfo(c10_left, tb10_left, header_paras=1, is_bullet_card=True),
    CardInfo(c10_right, tb10_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 11: LIVE DEMO BREAK: The Petite Paper Figure Generator
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
c11 = add_card(s11, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9), bg_color=COLOR_ALERT_BG, border_color=COLOR_ALERT_BORDER)
tb11 = s11.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(10.9), Inches(5.1))
tf11 = tb11.text_frame
tf11.word_wrap = True

p = tf11.paragraphs[0]
p.text = "LIVE DEMONSTRATION 2"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(14)
p.font.color.rgb = COLOR_ALERT_BORDER
p.space_after = Pt(8)

p = tf11.add_paragraph()
p.text = "LIVE DEMO: The Petite Paper Figure Generator"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(32)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

p = tf11.add_paragraph()
p.text = "Live walkthrough of the Petite Paper repository (~/dev/petite_paper): One script, publication figures."
p.font.name = FONT_FAMILY
p.font.size = Pt(16)
p.font.color.rgb = COLOR_TEXT_SECONDARY
p.space_after = Pt(26)

demo2_items = [
    ("Single-Command Execution:", "Running `python main.py --figure 2` live to output 300-DPI publication plots."),
    ("Google Sheets Palette Sync:", "Opening Google Sheets, changing a mutant's color, re-running the command, and watching the figure update live."),
    ("Clean Repository Structure:", "Showing how raw data, utility modules, and figure generators stay modular and version-controlled."),
    ("On-Demand Figure Iteration:", "Modifying plot styling or strain groupings in code and watching panels re-render in seconds.")
]
for t, d in demo2_items:
    p = tf11.add_paragraph()
    p.text = f"• {t} "
    p.font.name = FONT_FAMILY
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(10)
    run = p.add_run()
    run.text = d
    run.font.name = FONT_FAMILY
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s11, [
    CardInfo(c11, tb11, header_paras=3, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 12: Domain Expertise: The Scientist Remains the Ground Truth
# ==============================================================================
s12 = prs.slides.add_slide(blank_layout)
add_header(s12, "Part III: The Scientist in the Loop", "Domain Expertise: The Scientist Remains the Ground Truth",
           "AI accelerates computation, but responsibility for biological accuracy is 100% yours.")

c12_left = add_card(s12, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb12_left = s12.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb12_left.text_frame, "The Johannes Proteomics Story", [
    ("The AI's Silent Mistake:", "During automated analysis of my proteomics dataset, AGY implemented an incorrect normalization formula."),
    ("Why Did It Happen?", "I hadn't explicitly explained what the raw data represented and how it had already been processed upstream."),
    ("Domain Expertise Saved the Day:", "Johannes audited the data and caught the anomaly immediately because he understood the biological context."),
    ("The Critical Lesson:", "Without domain oversight, subtle mathematical or biological bugs will slip through and corrupt conclusions. You MUST audit the logic!")
], title_color=COLOR_DANGER, bullet_pt=12, space_after=8)

c12_right = add_card(s12, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb12_right = s12.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb12_right.text_frame, "Verification & Ground Truth Benchmarks", [
    ("Control Datasets First:", "Benchmark pipelines against known wild-type controls (e.g. CBS138 standard growth curves) before analyzing novel mutants."),
    ("Sanity Checking Numbers:", "Always verify that calculated doubling times, MICs, and fold-changes align with wet-lab biological reality."),
    ("Demand Explicit Explanations:", "Instruct AGY: 'Explain your mathematical logic step-by-step and document assumptions directly in code comments.'"),
    ("The Scientist's Shield:", "You are the scientific filter—treat agent output as a draft from a junior colleague that requires peer review.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

apply_card_animations(s12, [
    CardInfo(c12_left, tb12_left, header_paras=1, is_bullet_card=True),
    CardInfo(c12_right, tb12_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 13: Technical & Architectural Discipline: Clean Systems
# ==============================================================================
s13 = prs.slides.add_slide(blank_layout)
add_header(s13, "Part III: The Scientist in the Loop", "Technical & Architectural Discipline: Clean Systems",
           "Good architecture prevents code bloat, runaway token costs, and fragile pipelines.")

c13_left = add_card(s13, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb13_left = s13.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb13_left.text_frame, "Why Technical Knowledge Still Matters", [
    ("Programming Isn't Obsolete:", "Having an intuitive understanding of code structure helps you spot bad logic and fragile implementations early."),
    ("Catching Flawed AI Logic:", "When an agent formulates rules or code, human technical insight is required to spot flaws and steer it right."),
    ("The Presentation Animation Story:", "When teaching AGY PowerPoint animations in this project, it formulated the XML rules incorrectly on the first try. Technical insight caught the bug and guided it to the right abstraction."),
    ("The Code Bloat Trap:", "Left unguided, agents write redundant helper functions, duplicate parsing logic, and generate unmaintainable script bloat."),
    ("Token Burn & Sluggish Performance:", "Bloated repositories flood the agent's context window, burning tokens and degrading agent reasoning.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=11.5, space_after=6)

c13_right = add_card(s13, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb13_right = s13.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb13_right.text_frame, "Core Architectural Invariants & Refactoring", [
    ("Zero Hard-Coded Values:", "Never allow strain names, file paths, drug concentrations, or color codes to be hardcoded in scripts."),
    ("External Text Configurations:", "Keep parameters in clean YAML, JSON, or CSV configs that humans can easily edit and inspect."),
    ("Modular Pipeline Design:", "Separate data loading, mathematical transformation, and figure rendering into independent modules."),
    ("Regular Refactoring Passes:", "Use focused sessions with high-capability models (Pro/Thinking) specifically to clean, modularize, and simplify code."),
    ("Defensive Code:", "Insist that scripts validate data shapes, check for missing values, and fail loudly with informative error messages.")
], title_color=COLOR_ACCENT, bullet_pt=11.5, space_after=6)

apply_card_animations(s13, [
    CardInfo(c13_left, tb13_left, header_paras=1, is_bullet_card=True),
    CardInfo(c13_right, tb13_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 14: The Golden Rule: Code-First, Text-First Figures
# ==============================================================================
s14 = prs.slides.add_slide(blank_layout)
add_header(s14, "Part III: The Scientist in the Loop", "The Golden Rule: Code-First, Text-First Figures",
           "Never ask AI or humans to manually edit publication figures in Photoshop or Illustrator.")

c14_left = add_card(s14, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb14_left = s14.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb14_left.text_frame, "The Pitfalls of Manual Figure Tweaks", [
    ("Un-reproducible Tweaks:", "Resizing an axis, nudging a label, or coloring a bar in Illustrator leaves zero record of how the figure was assembled."),
    ("The 'New Replicate' Disaster:", "When you perform 3 additional biological replicates, all manual adjustments must be laboriously repeated from scratch."),
    ("Visual Distortions:", "Accidentally misaligning an error bar or significance star during manual layout assembly risks data distortion."),
    ("Image Hallucinations:", "GenAI image editors 'paint pixels'—they cannot be trusted to edit scientific charts. Always work through code!")
], title_color=COLOR_DANGER, bullet_pt=12, space_after=8)

c14_right = add_card(s14, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb14_right = s14.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb14_right.text_frame, "The Antigravity Standard: 100% Code", [
    ("Every Pixel is Code-Defined:", "Axis limits, typography, significance brackets (***), error bands, and multi-panel composites are explicitly scripted."),
    ("10-Second Regeneration:", "New raw data arrives? Simply re-run `python main.py --figure 2` and the entire 300-DPI composite figure updates instantly."),
    ("Centralized Palette Sync:", "Change a strain's color hex code in your config or Google Sheet, and all 15 figures update synchronously."),
    ("Reviewer-Proof Provenance:", "When peer reviewers or editors ask for raw code and data, your Git repository provides a complete, auditable ledger.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=8)

apply_card_animations(s14, [
    CardInfo(c14_left, tb14_left, header_paras=1, is_bullet_card=True),
    CardInfo(c14_right, tb14_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 15: How AGY Sessions Work: The Clean Slate Principle
# ==============================================================================
s15 = prs.slides.add_slide(blank_layout)
add_header(s15, "Part IV: Working with Antigravity", "How AGY Sessions Work: The Clean Slate Principle",
           "Understanding project boundaries, stateless sessions, and conversational steering.")

c15_left = add_card(s15, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb15_left = s15.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb15_left.text_frame, "The Project Sandbox & Fresh Sessions", [
    ("The Project Root Boundary:", "Everything inside your designated project directory is considered local and safe to inspect and run."),
    ("Fresh Conversational Memory:", "Every new conversation starts completely fresh with zero memory of prior chat turns."),
    ("Seeding via AGENTS.md:", "Automatic institutional memory—AGY loads your lab guidelines and file index at the start of every session."),
    ("The Power of Mentions (@):", "Use `@filename` to inject specific scripts, data tables, or past conversation summaries into the active context.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

c15_right = add_card(s15, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb15_right = s15.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb15_right.text_frame, "Live Observability & Tactical Steering", [
    ("Transparent Execution:", "Watch the agent work in real time—view terminal commands, file edits, and tool arguments as they execute."),
    ("Early Intervention:", "If the agent begins pursuing an incorrect biological assumption, stop it immediately and redirect its approach."),
    ("Step-by-Step Approvals:", "Choose whether to auto-run safe commands or review edits before they hit disk."),
    ("Human in Command:", "The agent proposes and executes; you maintain tactical oversight throughout the process.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=8)

apply_card_animations(s15, [
    CardInfo(c15_left, tb15_left, header_paras=1, is_bullet_card=True),
    CardInfo(c15_right, tb15_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 16: Context Windows & Token Hygiene: Working Memory
# ==============================================================================
s16 = prs.slides.add_slide(blank_layout)
add_header(s16, "Part IV: Working with Antigravity", "Context Windows & Token Hygiene: Working Memory",
           "An agent's working memory is finite. Token hygiene makes or breaks your results.")

c16_left = add_card(s16, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb16_left = s16.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb16_left.text_frame, "Working Memory & Attention Dilution", [
    ("For LLMs, Only Context Exists:", "If something isn't in the active context, it does not exist. Each session starts fresh, seeded only by AGENTS.md."),
    ("The 'Marathon Chat' Trap:", "A chat running for 50 turns accumulates thousands of tokens of old outputs, error logs, and file dumps. Attention gets diluted."),
    ("Symptom: Confusion & Drift:", "The agent becomes 'lazy', forgets strain aliases or styling rules, and takes longer to respond."),
    ("Proactive /compact:", "Use `/compact` when a chunk of work finishes to summarize progress and clear token bloat before context overflows.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=7)

c16_right = add_card(s16, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb16_right = s16.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
tf16_r = tb16_right.text_frame
tf16_r.word_wrap = True

p = tf16_r.paragraphs[0]
p.text = "Interactive Context Visualizer"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(8)

p = tf16_r.add_paragraph()
p.text = "Open in browser: context_visualizer.html"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(12)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(10)

vis_points = [
    ("Visual Token Meter:", "See system prompt (~5k), AGENTS.md (~4k), chat turns, and raw file dumps filling the context bucket."),
    ("Attention Degradation:", "Watch the attention indicator drop from 99% (Optimal) to <50% (Diluted / Hallucination Risk)."),
    ("Heavy Dump Simulation:", "Test what happens when dumping a 10,000-line raw CSV (+70k tokens) vs. running `/compact` to restore laser focus."),
    ("Rule of Thumb:", "Many short, modular conversations (one figure or task per chat) beats one giant marathon conversation every time.")
]
for title_text, desc_text in vis_points:
    p_v = tf16_r.add_paragraph()
    p_v.text = f"• {title_text} "
    p_v.font.name = FONT_FAMILY
    p_v.font.bold = True
    p_v.font.size = Pt(11.5)
    p_v.font.color.rgb = COLOR_TEXT_PRIMARY
    p_v.space_after = Pt(6)
    r_v = p_v.add_run()
    r_v.text = desc_text
    r_v.font.name = FONT_FAMILY
    r_v.font.bold = False
    r_v.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s16, [
    CardInfo(c16_left, tb16_left, header_paras=1, is_bullet_card=True),
    CardInfo(c16_right, tb16_right, header_paras=2, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 17: Model Selection: Matching Brainpower to the Task
# ==============================================================================
s17 = prs.slides.add_slide(blank_layout)
add_header(s17, "Part IV: Working with Antigravity", "Model Selection: Matching Brainpower to the Task",
           "Strategic model switching balances speed, token economics, and reasoning depth.")

c17_left = add_card(s17, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb17_left = s17.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb17_left.text_frame, "Lightweight vs. Heavyweight Models", [
    ("Fast & Lightweight (Flash / Flash-Lite):", "Lightning-fast responses with low token cost. Ideal for finding files, renaming variables, text edits, and minor formatting."),
    ("Deep Reasoning (Pro / Claude 3.5 Sonnet / Opus):", "Massive context windows with rigorous analytical capabilities. Ideal for complex multi-file debugging, mathematical logic, and architecture."),
    ("Avoid Token Waste:", "Don't burn expensive reasoning quota on trivial file lookups or simple label changes."),
    ("Tactical Model Switching:", "You can switch models mid-project or per conversation depending on the immediate technical challenge.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

c17_right = add_card(s17, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb17_right = s17.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb17_right.text_frame, "Thinking Effort & Multi-Model Flexibility", [
    ("Adjustable Thinking Levels:", "Dial thinking effort between 'Shallow' (rapid direct answers) and 'Deep' (extended internal chain-of-thought analysis)."),
    ("Deep Thinking for Tough Bugs:", "Engage deep thinking when unraveling multi-file race conditions, silent statistical anomalies, or complex pipeline errors."),
    ("Model-Agnostic Freedom:", "AGY works natively with Google Gemini, but also connects seamlessly to Anthropic Claude or OpenAI models via API keys."),
    ("Economic Optimization:", "Use cost-effective models for daily tasks, reserving flagship models for architectural heavy lifting.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=8)

apply_card_animations(s17, [
    CardInfo(c17_left, tb17_left, header_paras=1, is_bullet_card=True),
    CardInfo(c17_right, tb17_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 18: Prompt Engineering & Conversational Steering
# ==============================================================================
s18 = prs.slides.add_slide(blank_layout)
add_header(s18, "Part IV: Working with Antigravity", "Prompt Engineering & Conversational Steering",
           "How to communicate with coding agents to get precise, robust scientific results.")

c18_left = add_card(s18, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb18_left = s18.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb18_left.text_frame, "Principles of Effective Bioinfo Prompting", [
    ("Avoid Vague Directives:", "Never say 'make a nice plot.' Specify inputs, axes, units, color palette, statistical tests, and output resolution explicitly."),
    ("Provide Data Shapes & Schemas:", "Paste 3-5 sample rows or describe column types so the agent understands table geometry upfront."),
    ("State Invariants & Constraints:", "'Do not modify raw data in `data/`', 'Use Okabe-Ito palette', 'Normalize to WT strain CBS138 at t=0'."),
    ("Request Verification Steps:", "Ask the agent to print intermediate summary statistics (mean, SD, n replicates) before rendering the final plot.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

c18_right = add_card(s18, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb18_right = s18.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb18_right.text_frame, "Meta-Prompting & Active Steering", [
    ("Meta-Prompting with Web Chat:", "Ask Gemini or Claude web chat: 'Write the optimal prompt for Antigravity to build a growth curve pipeline with slash commands.'"),
    ("Incremental Prompting:", "Break large goals into sequential steps: first verify data parsing, then compute metrics, then assemble the figure."),
    ("Active Steering:", "If the agent starts writing complex custom math instead of using standard scipy/statannotations, stop it and specify the package."),
    ("Reinforce Success:", "When the agent hits the exact desired aesthetic, instruct it to save the pattern for future figures.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=8)

apply_card_animations(s18, [
    CardInfo(c18_left, tb18_left, header_paras=1, is_bullet_card=True),
    CardInfo(c18_right, tb18_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 19: Institutional Memory: AGENTS.md
# ==============================================================================
s19 = prs.slides.add_slide(blank_layout)
add_header(s19, "Part V: The Best Practices Toolkit", "Institutional Memory: AGENTS.md",
           "Teach your agent once; benefit across every future session in the project.")

c19_left = add_card(s19, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb19_left = s19.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb19_left.text_frame, "What is AGENTS.md?", [
    ("Single Source of Truth:", "A Markdown file placed in your project root. Antigravity automatically reads it at the start of every session."),
    ("Permanent Lab Invariants:", "Define your preferred color palettes (Okabe-Ito), strain naming conventions, directory paths, and figure resolutions."),
    ("Zero Repetitive Prompting:", "No need to re-type: 'Remember our WT strain is CBS138, font is Arial, and all figures must be 300 DPI' every morning."),
    ("Lab-Wide Standardization:", "Commit AGENTS.md to Git so every lab member's agent follows the exact same lab standards.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

# Right Box: Sample AGENTS.md
c19_right = add_card(s19, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_CODE_BG)
tb19_right = s19.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
tf19_r = tb19_right.text_frame
tf19_r.word_wrap = True

p = tf19_r.paragraphs[0]
p.text = "Example: Berman Lab AGENTS.md"
p.font.name = FONT_FAMILY
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
p_code = tf19_r.add_paragraph()
p_code.text = sample_code
p_code.font.name = "Courier New"
p_code.font.size = Pt(10)
p_code.font.color.rgb = COLOR_TEXT_PRIMARY

apply_card_animations(s19, [
    CardInfo(c19_left, tb19_left, header_paras=1, is_bullet_card=True),
    CardInfo(c19_right, tb19_right, header_paras=1, is_bullet_card=False)
])


# ==============================================================================
# SLIDE 20: Command 1: /plan — Architect Before Execution
# ==============================================================================
s20 = prs.slides.add_slide(blank_layout)
add_header(s20, "Part V: The Best Practices Toolkit", "Power Tool: /plan — Architect Before Coding",
           "Enforce a structured architectural roadmap before touching a single line of code.")

c20_left = add_card(s20, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb20_left = s20.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb20_left.text_frame, "How /plan Works & Why It Matters", [
    ("Halts Hasty Coding:", "Prevents the agent from rushing into writing code before understanding the full problem scope."),
    ("Comprehensive Roadmap:", "Produces a structured plan: files to inspect, data transformations, required packages, and potential edge cases."),
    ("The Scientist's Checkpoint:", "You read, edit, and approve the plan before execution begins. You stay in the driver's seat."),
    ("Interactive Refinement:", "Say 'In step 2, use median instead of mean'—the agent updates the roadmap before touching files.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

c20_right = add_card(s20, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb20_right = s20.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb20_right.text_frame, "Example Wet-Lab Workflow: /plan", [
    ("Prompt:", "/plan 'Build a parser for Tecan plate reader exports and plot 48h Candida growth curves'"),
    ("Step 1 (Data Audit):", "Inspect raw ASCII header rows and extract time-stamped optical density readings."),
    ("Step 2 (Quality Control):", "Subtract media-only blank wells and flag wells exhibiting sudden evaporation spikes."),
    ("Step 3 (Metric Calculation):", "Compute maximum growth rate (Vmax) and lag time using spline fitting."),
    ("Step 4 (Visualization):", "Render publication-ready facet grid grouping strains by drug concentration.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=7)

apply_card_animations(s20, [
    CardInfo(c20_left, tb20_left, header_paras=1, is_bullet_card=True),
    CardInfo(c20_right, tb20_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 21: Command 2: /grill-me — Interactive Alignment Interview
# ==============================================================================
s21 = prs.slides.add_slide(blank_layout)
add_header(s21, "Part V: The Best Practices Toolkit", "Power Tool: /grill-me — The Alignment Interview",
           "Instruct the agent to cross-examine you, uncovering ambiguities before execution.")

c21_left = add_card(s21, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb21_left = s21.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb21_left.text_frame, "How /grill-me Eliminates Assumptions", [
    ("Flips the Interrogation:", "Instead of you asking questions, the agent interviews YOU about your experimental setup."),
    ("Targets Missing Wet-Lab Context:", "Pinpoints ambiguous details that biologists take for granted: blank wells, replicates, units, and drug dilutions."),
    ("Prevents Silent Errors:", "Stops the agent from guessing normalization baselines or misinterpreting strain aliases."),
    ("The Perfect Partner for /plan:", "Run `/plan` followed by `/grill-me` for airtight, bug-free project execution.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=8)

c21_right = add_card(s21, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb21_right = s21.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb21_right.text_frame, "Example Wet-Lab Dialogue: /grill-me", [
    ("Directive:", "/grill-me 'I want to calculate MIC50 for our fluconazole disk diffusion assays'"),
    ("Agent Question 1:", "'How do you define the baseline? Is it 50% inhibition relative to the untreated growth control well, or media background?'"),
    ("Agent Question 2:", "'Some strains show trailing growth at 48 hours. Should the endpoint be evaluated at 24 hours or 48 hours?'"),
    ("Agent Question 3:", "'How should we handle non-inhibited resistant isolates where OD never drops below 50%?'"),
    ("Scientist Impact:", "2 minutes of alignment prevents hours of erroneous re-analysis.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=11.5, space_after=7)

apply_card_animations(s21, [
    CardInfo(c21_left, tb21_left, header_paras=1, is_bullet_card=True),
    CardInfo(c21_right, tb21_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 22: Command 3: /learn — Automating Knowledge Capture
# ==============================================================================
s22 = prs.slides.add_slide(blank_layout)
add_header(s22, "Part V: The Best Practices Toolkit", "Power Tool: /learn — Automating Knowledge Capture",
           "Convert hard-won bug fixes and workflow insights into permanent institutional memory.")

c22_left = add_card(s22, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb22_left = s22.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb22_left.text_frame, "How /learn Captures Experience", [
    ("Triggered After Success:", "Execute `/learn` immediately after solving a tricky bug, refining a workflow, or establishing project-wide standards."),
    ("Autonomous Knowledge Extraction:", "The agent analyzes what was changed, extracts generalizable patterns, and records them."),
    ("Permanent Institutional Memory:", "Automatically updates `AGENTS.md` so the new rule or invariant persists across all future conversations."),
    ("Preventing Repetitive Teaching:", "Teaches the agent your conventions once, eliminating the need to re-prompt in future chats."),
    ("Zero Overhead:", "Can be invoked as a single standalone slash command without complex arguments.")
], title_color=COLOR_SUCCESS, bullet_pt=12, space_after=7)

c22_right = add_card(s22, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_SUCCESS_BG, border_color=COLOR_SUCCESS)
tb22_right = s22.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb22_right.text_frame, "Case Study: Animation Rules in AGENTS.md", [
    ("Real Project Scenario:", "In this presentation project, we wanted all slides to follow a strict sequential entrance animation strategy."),
    ("Invoking /learn:", "After testing animation timing, the user instructed AGY to learn and record the architectural pattern permanently."),
    ("Extracted Rule 1 (Module Extraction):", "Never write raw XML in slide generators; always use the high-level helper `from pptx_animations import CardInfo, apply_card_animations`."),
    ("Extracted Rule 2 (Sequential Reveal):", "For 1-2 bounding boxes, reveal cards then bullet points sequentially; for >2 cards, reveal each bounding box as an indivisible unit."),
    ("Long-Term Impact:", "Every subsequent conversation (including this expansion!) automatically adheres to the animation rules without prompting.")
], title_color=COLOR_SUCCESS, bullet_pt=11, space_after=6)

apply_card_animations(s22, [
    CardInfo(c22_left, tb22_left, header_paras=1, is_bullet_card=True),
    CardInfo(c22_right, tb22_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 23: Custom Skills: Modular Functions for Coding Agents
# ==============================================================================
s23 = prs.slides.add_slide(blank_layout)
add_header(s23, "Part V: The Best Practices Toolkit", "Custom Skills: Modular Functions for Coding Agents",
           "Package complex multi-step workflows into reusable, shareable agent capabilities.")

c23_left = add_card(s23, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb23_left = s23.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb23_left.text_frame, "What Are Agent Skills?", [
    ("Functions for AI Agents:", "A skill is a self-contained bundle of instructions, reference scripts, and guardrails tailored to a specific task."),
    ("On-Demand Invocation:", "The agent automatically activates the skill whenever you request the corresponding task."),
    ("Shareable Across the Lab:", "Skills live in your repository or global config—commit a skill to Git so all lab members benefit."),
    ("Built-in vs. Custom:", "Antigravity comes with built-in skills (data analysis, web browsing), and you can write custom wet-lab skills.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

c23_right = add_card(s23, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb23_right = s23.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb23_right.text_frame, "Berman Lab Case Study & Future Ideas", [
    ("The Figure Palette Skill (Case Study):", "Input: A URL or list of color names from a journal. Function: Autonomously extracts hex codes, validates colorblind contrast, and updates Google Sheets config."),
    ("Manuscript Standards Checker (Future):", "Incorporate Judy's scientific writing guidelines into an automated skill that audits manuscript drafts for clarity and conciseness."),
    ("Plate Layout Converter (Future):", "Translates 96/384-well plate layouts from hand-written lab notebook scans into tidy analysis dataframes.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=8)

apply_card_animations(s23, [
    CardInfo(c23_left, tb23_left, header_paras=1, is_bullet_card=True),
    CardInfo(c23_right, tb23_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 24: Safety, Sandboxing & Data Hygiene: Work Fearlessly
# ==============================================================================
s24 = prs.slides.add_slide(blank_layout)
add_header(s24, "Part V: The Best Practices Toolkit", "Safety, Sandboxing & Data Hygiene: Work Fearlessly",
           "How Antigravity protects your system and keeps raw scientific data immutable.")

c24_left = add_card(s24, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb24_left = s24.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb24_left.text_frame, "Folder Sandboxing & Execution Boundaries", [
    ("The Workspace Boundary:", "Antigravity considers files inside the project folder safe to modify, but prompts you before accessing outside directories."),
    ("Permission Controls:", "You can set commands to run in sandboxed mode, or require explicit manual approval for system-level actions."),
    ("Beware Indirect Execution:", "Agents run real Python and Bash. They could theoretically modify outside files through code—audit scripts before running!"),
    ("Network Isolation:", "Safe execution mode prevents accidental external data exfiltration or unintended API calls.")
], title_color=COLOR_DANGER, bullet_pt=12, space_after=8)

c24_right = add_card(s24, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb24_right = s24.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb24_right.text_frame, "Data Hygiene: The Immutable Raw Vault", [
    ("Read-Only `data/` Rule:", "Keep original plate reader exports, sequencing FASTQs, and microscopy TIFFs strictly read-only."),
    ("Write Code, Never Overwrite Data:", "All pipeline outputs must be written to dedicated `output/` or `processed/` folders."),
    ("Off-Site Archival Copies:", "Always maintain the primary raw data backup on Dropbox, Google Drive, or the lab NAS."),
    ("Fearless Experimentation:", "When raw data is guaranteed immutable, you can let the agent refactor and test code with zero anxiety.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=8)

apply_card_animations(s24, [
    CardInfo(c24_left, tb24_left, header_paras=1, is_bullet_card=True),
    CardInfo(c24_right, tb24_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 25: Version Control: Git & The Version Control Tree (Schematic for Newcomers)
# ==============================================================================
s25 = prs.slides.add_slide(blank_layout)
add_header(s25, "Part V: The Best Practices Toolkit", "Version Control: Git & The Version Control Tree",
           "The ultimate scientific safety net: review visual diffs, collaborate safely, and branch work.")

c25_left = add_card(s25, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb25_left = s25.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb25_left.text_frame, "Why Version Control is Essential", [
    ("Visual Green/Red Diffs:", "Inspect exact line-by-line code changes before accepting agent modifications."),
    ("The Universal Undo Button:", "If an experimental refactor or plot doesn't pan out, roll back with a single Git command. No fear of broken code."),
    ("Collaboration Without Collision:", "Team members can work on separate branches or operating systems (macOS + Windows) without overwriting each other's scripts."),
    ("Reviewer-Proof Provenance:", "Every figure is permanently tied to an exact Git commit hash, ensuring 100% reproducible results for publication.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

# Right Card: Visual Git Tree Schematic for Newcomers
c25_right = add_card(s25, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb25_right = s25.shapes.add_textbox(Inches(7.1), Inches(2.1), Inches(5.2), Inches(0.65))
tf25_r = tb25_right.text_frame
tf25_r.word_wrap = True
p = tf25_r.paragraphs[0]
p.text = "Schematic: How Version Control Works"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(16)
p.font.color.rgb = COLOR_SUCCESS

p2 = tf25_r.add_paragraph()
p2.text = "A visual mental model of commits, branches, and merges"
p2.font.name = FONT_FAMILY
p2.font.size = Pt(11)
p2.font.color.rgb = COLOR_TEXT_SECONDARY

extra_shapes_s25 = []

# Main branch line (horizontal connecting bar)
main_line = s25.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.0), Inches(3.14), Inches(3.8), Inches(0.04))
main_line.fill.solid()
main_line.fill.fore_color.rgb = COLOR_ACCENT
main_line.line.fill.background()
extra_shapes_s25.append(main_line)

# Feature branch line (horizontal connecting bar)
feat_line = s25.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.8), Inches(4.14), Inches(1.2), Inches(0.04))
feat_line.fill.solid()
feat_line.fill.fore_color.rgb = COLOR_SUCCESS
feat_line.line.fill.background()
extra_shapes_s25.append(feat_line)

# Branch down connector (diagonal branch out)
branch_out = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.1), Inches(3.3), Inches(0.4), Inches(0.8))
branch_out.fill.solid()
branch_out.fill.fore_color.rgb = COLOR_SUCCESS_BG
branch_out.line.color.rgb = COLOR_SUCCESS
branch_out.line.width = Pt(1)
extra_shapes_s25.append(branch_out)

# Branch up connector (merge back into main)
merge_in = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.1), Inches(3.3), Inches(0.4), Inches(0.8))
merge_in.fill.solid()
merge_in.fill.fore_color.rgb = COLOR_PURPLE_BG
merge_in.line.color.rgb = COLOR_PURPLE
merge_in.line.width = Pt(1)
extra_shapes_s25.append(merge_in)

# Branch label pills
pill_main = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.15), Inches(2.98), Inches(0.8), Inches(0.35))
pill_main.fill.solid()
pill_main.fill.fore_color.rgb = COLOR_ACCENT
pill_main.line.fill.background()
pill_main.text_frame.margin_left = Inches(0.05)
p_pm = pill_main.text_frame.paragraphs[0]
p_pm.text = "main"
p_pm.font.name = FONT_FAMILY
p_pm.font.bold = True
p_pm.font.size = Pt(10)
p_pm.font.color.rgb = RGBColor(255, 255, 255)
extra_shapes_s25.append(pill_main)

pill_feat = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.15), Inches(3.98), Inches(1.5), Inches(0.35))
pill_feat.fill.solid()
pill_feat.fill.fore_color.rgb = COLOR_SUCCESS
pill_feat.line.fill.background()
pill_feat.text_frame.margin_left = Inches(0.05)
p_pf = pill_feat.text_frame.paragraphs[0]
p_pf.text = "feature/palette"
p_pf.font.name = FONT_FAMILY
p_pf.font.bold = True
p_pf.font.size = Pt(9.5)
p_pf.font.color.rgb = RGBColor(255, 255, 255)
extra_shapes_s25.append(pill_feat)

# Commit Nodes (Circles)
commits_data = [
    # (x, y, label, sublabel, color, text_above)
    (8.15, 2.98, "C1", "C1: Setup", COLOR_ACCENT, True),
    (9.05, 2.98, "C2", "C2: Baseline", COLOR_ACCENT, True),
    (10.1, 2.98, "C3", "C3: Plot v1", COLOR_ACCENT, True),
    (11.55, 2.98, "C6", "C6: Merged", COLOR_PURPLE, True),
    (9.65, 3.98, "C4", "C4: New Colors", COLOR_SUCCESS, False),
    (10.75, 3.98, "C5", "C5: Polish", COLOR_SUCCESS, False)
]

for cx, cy, node_txt, sub_txt, col, above in commits_data:
    node = s25.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(cy), Inches(0.36), Inches(0.36))
    node.fill.solid()
    node.fill.fore_color.rgb = col
    node.line.fill.background()
    p_n = node.text_frame.paragraphs[0]
    p_n.text = node_txt
    p_n.font.name = FONT_FAMILY
    p_n.font.bold = True
    p_n.font.size = Pt(9)
    p_n.font.color.rgb = RGBColor(255, 255, 255)
    extra_shapes_s25.append(node)

    # Label box
    lbl_y = Inches(cy - 0.28) if above else Inches(cy + 0.38)
    lbl_box = s25.shapes.add_textbox(Inches(cx - 0.3), lbl_y, Inches(1.0), Inches(0.28))
    tf_lbl = lbl_box.text_frame
    tf_lbl.margin_left = tf_lbl.margin_right = tf_lbl.margin_top = tf_lbl.margin_bottom = 0
    p_lbl = tf_lbl.paragraphs[0]
    p_lbl.text = sub_txt
    p_lbl.font.name = FONT_FAMILY
    p_lbl.font.size = Pt(8.5)
    p_lbl.font.color.rgb = COLOR_TEXT_MUTED
    extra_shapes_s25.append(lbl_box)

# Bottom Legend Card explaining concepts for newcomers
leg_card = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.1), Inches(4.85), Inches(5.2), Inches(1.8))
leg_card.fill.solid()
leg_card.fill.fore_color.rgb = RGBColor(255, 255, 255)
leg_card.line.color.rgb = COLOR_CARD_BORDER
leg_card.line.width = Pt(1)
tf_leg = leg_card.text_frame
tf_leg.word_wrap = True
tf_leg.margin_left = tf_leg.margin_right = tf_leg.margin_top = Inches(0.12)
tf_leg.margin_bottom = 0

legend_items = [
    ("● Commit:", "An immutable checkpoint of all files at that moment. Jump back anytime."),
    ("↗ Branch:", "A safe sandbox to test new code or palettes without risking working scripts."),
    ("↘ Merge:", "Safely incorporates tested, verified changes back into the production pipeline."),
    ("↺ Revert:", "If an agent experiment fails, discard the branch and roll back with zero risk.")
]
for idx, (term, desc) in enumerate(legend_items):
    p_l = tf_leg.paragraphs[0] if idx == 0 else tf_leg.add_paragraph()
    p_l.text = f"{term} "
    p_l.font.name = FONT_FAMILY
    p_l.font.size = Pt(9.5)
    p_l.font.bold = True
    p_l.font.color.rgb = COLOR_TEXT_PRIMARY
    p_l.space_after = Pt(2.5)
    r_l = p_l.add_run()
    r_l.text = desc
    r_l.font.name = FONT_FAMILY
    r_l.font.bold = False
    r_l.font.color.rgb = COLOR_TEXT_SECONDARY
extra_shapes_s25.append(leg_card)

apply_card_animations(s25, [
    CardInfo(c25_left, tb25_left, header_paras=1, is_bullet_card=True),
    CardInfo(c25_right, tb25_right, header_paras=1, is_bullet_card=False, extra_shapes=extra_shapes_s25)
])


# ==============================================================================
# SLIDE 26: Berman Lab Member Missions & Shared Infrastructure (Merged 26-27)
# ==============================================================================
s26 = prs.slides.add_slide(blank_layout)
add_header(s26, "Part VI: The Path Forward", "Berman Lab Member Missions & Shared Infrastructure",
           "Practical next steps for lab members and moving toward a shared, agentic lab codebase.")

c26_left = add_card(s26, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb26_left = s26.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb26_left.text_frame, "Lab Member Missions: Immediate Entry Points", [
    ("Ze'ev (DDA Panel Creator):", "Feed the GitHub issues of the DDA panel creator directly to AGY to reproduce bugs, write automated tests, and resolve issues within days."),
    ("Shani (Analysis Consolidation):", "Consolidate scattered Claude web chats into a single dedicated AGY project with unified figures and zero copy-paste friction."),
    ("Sofya (Transposon Libraries):", "Use AGY for exploratory quality assessment on Candida transposon insertion libraries, analyzing mapping files and insertion distributions."),
    ("Alona (Model Integration):", "Leverage your Claude subscription! Plug Anthropic models (Claude 3.5/3.7 Sonnet) into AGY via API keys, or explore the Claude Code CLI."),
    ("Judy (Scientific Writing Skills):", "Extract lessons and guidelines from your scientific writing course into automated skills that audit manuscript drafts for clarity and conciseness.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=11.5, space_after=6)

c26_right = add_card(s26, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb26_right = s26.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb26_right.text_frame, "Shared Lab Tools & The Berman Codebase", [
    ("Reviving & Maintaining Legacy Tools:", "Tools like OD Normalizer, Scanalyzer, diskImageR, and Ymap can now be updated, maintained, and run much easier without needing the original authors or professional programmers."),
    ("Towards a Lab Skill Repository (Simple RAG):", "Create a central repository of custom skills and guidelines that is aware of our lab's existing scripts, assay schemas, and figure conventions—so any lab member's agent can instantly retrieve and reuse lab knowledge."),
    ("Permanent Lab Intellectual Property:", "Stop treating analysis code as disposable, one-off scripts. Commit pipelines to Git so they become permanent, shared lab assets."),
    ("Culture of Collaborative Automation:", "Lab members build upon each other's figure generators and parsers rather than reinventing the wheel.")
], title_color=COLOR_ACCENT, bullet_pt=11.5, space_after=6)

apply_card_animations(s26, [
    CardInfo(c26_left, tb26_left, header_paras=1, is_bullet_card=True),
    CardInfo(c26_right, tb26_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 27: Platform Economics & Navigating Subscriptions (Split from old 28)
# ==============================================================================
s27 = prs.slides.add_slide(blank_layout)
add_header(s27, "Part VI: The Path Forward", "Platform Economics & Navigating Subscriptions",
           "Understanding Antigravity costs, Google AI Studio tiers, and academic subscriptions at TAU.")

c27_left = add_card(s27, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb27_left = s27.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb27_left.text_frame, "Platform Tiers & TAU Academic Licensing", [
    ("Antigravity Platform is 100% Free:", "The local agentic development environment itself is completely free to download and use."),
    ("Google AI Studio Free Tier:", "Generous free quota for Gemini models—great for getting started, but heavy multi-turn sessions can hit daily rate limits."),
    ("TAU Academic Gemini License:", "TAU offers an academic Gemini Pro license (~550 NIS/year). Note: verify whether it covers developer API quotas or just the web chat interface."),
    ("Standard Private Developer API:", "Private Google AI Studio billing runs on pay-as-you-go token consumption (~75 NIS/month or ~900 NIS/year for active daily coding).")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

c27_right = add_card(s27, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb27_right = s27.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb27_right.text_frame, "Multi-Model Freedom & Cost Optimization", [
    ("Model-Agnostic Setup:", "Connect Gemini, Anthropic Claude (e.g. Claude 3.5/3.7 Sonnet), or OpenAI via standard API keys in your configuration."),
    ("Tiered Brainpower Strategy:", "Use cheap, fast models (Gemini Flash) for scanning files and formatting text; reserve high-end reasoning models (Pro / Sonnet) for architecture and debugging."),
    ("Biology Figure Sufficiency:", "For our wet-lab plate reader, microscopy, and sequencing figure pipelines, Gemini Pro is cost-effective and fully capable."),
    ("Collaborative Efficiency:", "A single lab subscription or shared project API key can power figure generation across the entire team.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=8)

apply_card_animations(s27, [
    CardInfo(c27_left, tb27_left, header_paras=1, is_bullet_card=True),
    CardInfo(c27_right, tb27_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 28: Bonus Cheatsheet: The 8 Core Agent Habits for Reliable Science
# ==============================================================================
s28 = prs.slides.add_slide(blank_layout)
add_header(s28, "Bonus Cheatsheet", "The 8 Core Agent Habits for Reliable Science",
           "Keep this handy: essential rules of thumb for wet-lab scientists working with coding agents.")

c28_left = add_card(s28, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb28_left = s28.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb28_left.text_frame, "Project Setup & Working Memory (Habits 1–4)", [
    ("1. Isolate Raw Data in `data/`:", "Always treat original plate reader exports, FASTQs, and images as read-only. Keep gold-standard archival copies on Dropbox/NAS."),
    ("2. Anchor Invariants in `AGENTS.md`:", "Record strain aliases (WT=CBS138), color palettes (Okabe-Ito), and file structures once so every conversation follows them."),
    ("3. Align Before Coding (`/plan` & `/grill-me`):", "Never let the agent write code blind. Review the architectural plan and let it interview you about wet-lab assumptions."),
    ("4. Keep Chats Focused & Modular:", "One figure, one panel, or one bug per chat session. Many short chats beat one giant marathon conversation every time.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=11.5, space_after=7)

c28_right = add_card(s28, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb28_right = s28.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb28_right.text_frame, "Execution, Learning & Hygiene (Habits 5–8)", [
    ("5. Proactive Token Hygiene (`/compact`):", "When a task is complete, run `/compact` to clear execution logs and reset attention before starting the next step."),
    ("6. Codify Hard-Won Fixes (`/learn`):", "When you solve a tricky bug or establish a new standard, run `/learn` to save the rule into `AGENTS.md` permanently."),
    ("7. Code-First, Text-First Figures:", "Never manually edit charts in Illustrator or Photoshop. If an axis or color needs changing, change the code that generated it."),
    ("8. Git as the Universal Undo Button:", "Commit working milestones to Git. Review visual green/red diffs before saving, and rollback mistakes fearlessly.")
], title_color=COLOR_ACCENT, bullet_pt=11.5, space_after=7)

apply_card_animations(s28, [
    CardInfo(c28_left, tb28_left, header_paras=1, is_bullet_card=True),
    CardInfo(c28_right, tb28_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# Save Presentation
# ==============================================================================
output_path = "antigravity_lab_presentation.pptx"
prs.save(output_path)
print(f"Presentation successfully saved to {output_path} (28 slides, 16:9 widescreen).")
