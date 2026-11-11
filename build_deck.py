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
from render_assets import render_paper_repro_svg, crop_agy_figure, render_git_tree_schematic

# Ensure intermediate visual assets are up-to-date
render_paper_repro_svg()
crop_agy_figure()
render_git_tree_schematic()

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
# SLIDE 2 (NEW): If You Remember One Thing From This Presentation
# ==============================================================================
s1b = prs.slides.add_slide(blank_layout)
add_header(s1b, "Core Takeaway", "If You Remember One Thing From This Presentation",
           "You can write bespoke software solutions for your scientific domain without knowing how to program.")

c1b_left = add_card(s1b, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb1b_left = s1b.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb1b_left.text_frame, "The 3 Ingredients for Domain Software", [
    ("1. You're a Domain Expert:", "You understand the biology, experimental assays, controls, and scientific questions that matter."),
    ("2. A Little Technical Ability:", "You have basic comfort navigating files, running tools, and thinking logically about data workflows."),
    ("3. Time to Learn & Play Around:", "You take the time to experiment interactively with agentic tools and learn how to steer them."),
    ("The Core Formula:", "With these three ingredients, you can write software solutions for your domain without knowing how to program.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=8)

c1b_right = add_card(s1b, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb1b_right = s1b.shapes.add_textbox(Inches(7.2), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb1b_right.text_frame, "Why This Is a Game Changer for Our Lab", [
    ("Programming Language Optional:", "For us, learning a programming language from scratch may no longer be mandatory."),
    ("Bespoke Wet-Lab Solutions:", "Wet-lab researchers can create custom analysis and visualization solutions by themselves."),
    ("Tools Are Improving Fast:", "As these AI coding platforms continue getting better fast, this is a true game changer."),
    ("Accelerate & Unlock Science:", "Dramatically accelerates your current work and unlocks possibilities that were previously unavailable.")
], title_color=COLOR_ACCENT, bullet_pt=12, space_after=8)

apply_card_animations(s1b, [
    CardInfo(c1b_left, tb1b_left, header_paras=1, is_bullet_card=True),
    CardInfo(c1b_right, tb1b_right, header_paras=1, is_bullet_card=True)
])


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
           "How Antigravity inferred the paper's branching computational dependency graph and produced preliminary figures.")

c6_left = add_card(s6, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb6_left = s6.shapes.add_textbox(Inches(1.1), Inches(2.2), Inches(5.0), Inches(4.4))
add_bullet_list(tb6_left.text_frame, "Pipeline Architecture & Execution DAG", [
    ("Stage 1: Toolchain Discovery:", "Parsed PDF methods, extracted exact versions, and built a custom conda environment."),
    ("Stage 2: Legacy Source Patching:", "Modified obsolete C/C++ compiler flags and headers to build cleanly on modern ARM/POSIX systems."),
    ("Stage 3: Parallel Data Streaming:", "Queried ENA mirrors with aria2c to stream raw sequencing fastqs at maximum throughput."),
    ("Stage 4: Inferred Branching DAG:", "Split trimmed reads into two parallel branches: Reference-Guided De Novo Pan-Genome vs. GATK SNP Phylogeny."),
    ("Stage 5: Pan-GWAS Convergence:", "Merged gene presence/absence and core phylogeny into Scoary to identify resistance-associated loci.")
], title_color=COLOR_TEXT_PRIMARY, bullet_pt=12, space_after=6)

# Right Card: paper-repro-flow.svg rendered diagram
c6_right = add_card(s6, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb6_right = s6.shapes.add_textbox(Inches(7.1), Inches(2.05), Inches(5.2), Inches(0.42))
tf6_r = tb6_right.text_frame
tf6_r.word_wrap = True
p = tf6_r.paragraphs[0]
p.text = "Inferred Branching Pipeline (paper-repro-flow.svg)"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(15)
p.font.color.rgb = COLOR_ACCENT

# Insert paper-repro-flow.png (exact high-res render of paper-repro-flow.svg, aspect ratio 0.7597)
pic6_h = Inches(4.25)
pic6_w = Inches(4.25 * (982.08 / 1292.6))
pic6_left = Inches(6.9 + (5.6 - 4.25 * (982.08 / 1292.6)) / 2)
pic6_flow = s6.shapes.add_picture("paper-repro-flow.png", pic6_left, Inches(2.48), width=pic6_w, height=pic6_h)

apply_card_animations(s6, [
    CardInfo(c6_left, tb6_left, header_paras=1, is_bullet_card=True),
    CardInfo(c6_right, tb6_right, header_paras=1, is_bullet_card=False, extra_shapes=[pic6_flow])
])


# ==============================================================================
# SLIDE 7: Case Study 2: Pipeline Figures vs. Published Paper
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
add_header(s7, "Part II: What Did I Do With AGY?", "Case Study 2: Pipeline Figures vs. Published Paper",
           "Side-by-side comparison of the original published figure (Barber et al., 2021) and the AGY pipeline's reproduction.")

c7_left = add_card(s7, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=RGBColor(255, 255, 255), border_color=COLOR_CARD_BORDER)
tb7_left = s7.shapes.add_textbox(Inches(1.0), Inches(2.08), Inches(5.2), Inches(0.55))
tf7_l = tb7_left.text_frame
tf7_l.word_wrap = True
p7_l1 = tf7_l.paragraphs[0]
p7_l1.text = "Published Paper: Barber et al. (2021) Figure 1"
p7_l1.font.name = FONT_FAMILY
p7_l1.font.bold = True
p7_l1.font.size = Pt(15)
p7_l1.font.color.rgb = COLOR_TEXT_PRIMARY
p7_l2 = tf7_l.add_paragraph()
p7_l2.text = "Original pan-genome matrix, rarefaction curves, and gene length comparisons"
p7_l2.font.name = FONT_FAMILY
p7_l2.font.size = Pt(11)
p7_l2.font.color.rgb = COLOR_TEXT_SECONDARY

# Paper figure_1.png (1200x688, aspect ratio 1.744)
pic7_l_w = Inches(5.3)
pic7_l_h = Inches(5.3 * (688 / 1200))
pic7_paper = s7.shapes.add_picture("figure_1_paper.png", Inches(0.95), Inches(3.15), width=pic7_l_w, height=pic7_l_h)

c7_right = add_card(s7, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb7_right = s7.shapes.add_textbox(Inches(7.1), Inches(2.08), Inches(5.2), Inches(0.55))
tf7_r = tb7_right.text_frame
tf7_r.word_wrap = True
p7_r1 = tf7_r.paragraphs[0]
p7_r1.text = "AGY Pipeline Output (Figure_1_pangenome_analysis-AGY.png)"
p7_r1.font.name = FONT_FAMILY
p7_r1.font.bold = True
p7_r1.font.size = Pt(15)
p7_r1.font.color.rgb = COLOR_ACCENT
p7_r2 = tf7_r.add_paragraph()
p7_r2.text = "Generated autonomously from raw reads with zero manual coding"
p7_r2.font.name = FONT_FAMILY
p7_r2.font.size = Pt(11)
p7_r2.font.color.rgb = COLOR_TEXT_SECONDARY

# AGY cropped figure (4520x3114, aspect ratio 1.4515)
pic7_r_w = Inches(5.3)
pic7_r_h = Inches(5.3 * (3114 / 4520))
pic7_agy = s7.shapes.add_picture("Figure_1_pangenome_analysis-AGY_cropped.png", Inches(7.05), Inches(2.90), width=pic7_r_w, height=pic7_r_h)

apply_card_animations(s7, [
    CardInfo(c7_left, tb7_left, header_paras=2, is_bullet_card=False, extra_shapes=[pic7_paper]),
    CardInfo(c7_right, tb7_right, header_paras=2, is_bullet_card=False, extra_shapes=[pic7_agy])
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

# Right Card: Visual Git Tree Schematic for Newcomers (High-res rendered figure)
c25_right = add_card(s25, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb25_right = s25.shapes.add_textbox(Inches(7.1), Inches(2.06), Inches(5.2), Inches(0.55))
tf25_r = tb25_right.text_frame
tf25_r.word_wrap = True
p = tf25_r.paragraphs[0]
p.text = "Schematic: How Version Control Works"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(16)
p.font.color.rgb = COLOR_SUCCESS

p2 = tf25_r.add_paragraph()
p2.text = "A visual mental model of commits, branches, merges, and rollbacks"
p2.font.name = FONT_FAMILY
p2.font.size = Pt(11)
p2.font.color.rgb = COLOR_TEXT_SECONDARY

# Insert git_tree_schematic.png (1560x1200, aspect ratio 1.30)
pic25_w = Inches(5.33)
pic25_h = Inches(5.33 * (1200 / 1560))
pic25_left = Inches(6.9 + (5.6 - 5.33) / 2)
pic25_git = s25.shapes.add_picture("git_tree_schematic.png", pic25_left, Inches(2.64), width=pic25_w, height=pic25_h)

apply_card_animations(s25, [
    CardInfo(c25_left, tb25_left, header_paras=1, is_bullet_card=True),
    CardInfo(c25_right, tb25_right, header_paras=2, is_bullet_card=False, extra_shapes=[pic25_git])
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
print(f"Presentation successfully saved to {output_path} ({len(prs.slides)} slides, 16:9 widescreen).")
