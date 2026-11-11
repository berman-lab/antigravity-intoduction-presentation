"""
Generate Antigravity Lab Presentation (.pptx) — Refined 29-Slide Edition
Synthesized from Presentation notes.md and updated user feedback.
Widescreen 16:9, clean white background, high contrast, 1-2-3 column layouts.
Large, readable typography (26pt titles, 19pt card headers, 14.5–15pt punchy bullets).
Includes native PowerPoint OpenXML entrance animations on every content slide.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx_animations import CardInfo, apply_card_animations
from render_assets import render_paper_repro_svg, crop_agy_figure, crop_paper_figure, render_git_tree_schematic

# Ensure intermediate visual assets are up-to-date
render_paper_repro_svg()
crop_agy_figure()
crop_paper_figure()
render_git_tree_schematic()

# --- Presentation Setup ---
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]  # Blank layout

# --- Color Palette ---
COLOR_BG = RGBColor(255, 255, 255)
COLOR_TEXT_PRIMARY = RGBColor(17, 24, 39)      # Charcoal / near-black (#111827)
COLOR_TEXT_SECONDARY = RGBColor(75, 85, 99)    # Medium gray (#4B5563)
COLOR_TEXT_MUTED = RGBColor(107, 114, 128)     # Lighter gray (#6B7280)
COLOR_ACCENT = RGBColor(37, 99, 235)           # Blue accent (#2563EB)
COLOR_ACCENT_BG = RGBColor(239, 246, 255)      # Light blue tint
COLOR_CARD_BG = RGBColor(248, 250, 252)        # Off-white / light slate
COLOR_CARD_BORDER = RGBColor(226, 232, 240)    # Slate border (#E2E8F0)
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
    header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.42), Inches(11.733), Inches(1.4))
    tf = header_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    # Section Tag (11pt bold uppercase accent per AGENTS.md)
    p_tag = tf.paragraphs[0]
    p_tag.text = section_tag.upper()
    p_tag.font.name = FONT_FAMILY
    p_tag.font.size = Pt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_ACCENT
    p_tag.space_after = Pt(3)

    # Title (26pt bold per AGENTS.md)
    p_title = tf.add_paragraph()
    p_title.text = title
    p_title.font.name = FONT_FAMILY
    p_title.font.size = Pt(26)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_TEXT_PRIMARY

    # Subtitle (15pt for clear legibility)
    if subtitle:
        p_sub = tf.add_paragraph()
        p_sub.text = subtitle
        p_sub.font.name = FONT_FAMILY
        p_sub.font.size = Pt(15)
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


def add_bullet_list(tf, card_title, bullets, title_color=COLOR_TEXT_PRIMARY, bullet_pt=None, space_after=None):
    """Fills a text frame with a card title and concise, large-font bullet points."""
    if bullet_pt is None:
        bullet_pt = 15 if len(bullets) <= 4 else 14.5
    if space_after is None:
        space_after = 10 if len(bullets) <= 4 else 8

    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = card_title
    p.font.name = FONT_FAMILY
    p.font.bold = True
    p.font.size = Pt(19)
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
        run.font.size = Pt(bullet_pt)
        run.font.color.rgb = COLOR_TEXT_SECONDARY


# ==============================================================================
# SLIDE 1: Title Slide (Unanimated)
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.55), Inches(11.333), Inches(4.9))
tf1 = tb1.text_frame
tf1.word_wrap = True

p = tf1.paragraphs[0]
p.text = "ANTIGRAVITY IN THE WET LAB"
p.font.name = FONT_FAMILY
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(12)

p = tf1.add_paragraph()
p.text = "From Chatbot to Colleague:"
p.font.name = FONT_FAMILY
p.font.size = Pt(44)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_PRIMARY

p = tf1.add_paragraph()
p.text = "Your On-Call Junior Bioinformatician"
p.font.name = FONT_FAMILY
p.font.size = Pt(38)
p.font.bold = True
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(24)

p = tf1.add_paragraph()
p.text = "Accelerating Candida research from raw assays to reproducible publication figures"
p.font.name = FONT_FAMILY
p.font.size = Pt(20)
p.font.color.rgb = COLOR_TEXT_SECONDARY
p.space_after = Pt(40)

p = tf1.add_paragraph()
p.text = "Berman Lab • Yeast Biology Meeting"
p.font.name = FONT_FAMILY
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = COLOR_TEXT_MUTED


# ==============================================================================
# SLIDE 2: If You Remember One Thing From This Presentation
# ==============================================================================
s1b = prs.slides.add_slide(blank_layout)
add_header(s1b, "Core Takeaway", "If You Remember One Thing From This Presentation",
           "Write bespoke scientific software for your domain without knowing how to program.")

c1b_left = add_card(s1b, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb1b_left = s1b.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb1b_left.text_frame, "The 3 Ingredients You Need", [
    ("1. Domain Expertise:", "You know the biology, assays, controls, and key questions."),
    ("2. A Little Tech Ability:", "Basic comfort with files, tools, and logical workflows."),
    ("3. Time to Play Around:", "Willingness to experiment interactively and steer the agent."),
    ("The Core Formula:", "With these three, you can build custom software without coding.")
], title_color=COLOR_TEXT_PRIMARY)

c1b_right = add_card(s1b, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb1b_right = s1b.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb1b_right.text_frame, "Why This Is a Game Changer", [
    ("Coding Optional:", "Learning Python or R from scratch may no longer be mandatory."),
    ("Bespoke Lab Solutions:", "Wet-lab researchers can build custom analyses independently."),
    ("Improving Fast:", "Agentic tools are advancing rapidly—human language is the UI."),
    ("Unlock New Science:", "Accelerates current work and opens possibilities once out of reach.")
], title_color=COLOR_ACCENT)

apply_card_animations(s1b, [
    CardInfo(c1b_left, tb1b_left, header_paras=1, is_bullet_card=True),
    CardInfo(c1b_right, tb1b_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 3: The Wet-Lab Reality & The Programming Bottleneck
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "Part I: The Paradigm Shift", "The Wet-Lab Reality & The Programming Bottleneck",
           "Our lab generates rich, complex data—but manual coding and web chatbots slow us down.")

c2_left = add_card(s2, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb2_left = s2.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb2_left.text_frame, "The High-Throughput Wet-Lab Reality", [
    ("Routine & Fancy Data:", "Growth curves, DDAs, spot assays, plus WGS and proteomics."),
    ("No 'One Shoe Fits All':", "Every project has unique strains, layouts, and analysis needs."),
    ("Limited Tech Access:", "Bioinformaticians aren't available for day-to-day plotting tasks."),
    ("DIY Requirement:", "We must analyze and visualize our own data for papers and talks."),
    ("The Web Chatbot Habit:", "Most of us use ChatGPT or Claude web UIs to write R/Python.")
], title_color=COLOR_TEXT_PRIMARY)

c2_right = add_card(s2, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb2_right = s2.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb2_right.text_frame, "The Bottleneck & Industry Shift", [
    ("Slow to Write & Manage:", "Chatbot copy-pasting is tedious and builds up technical debt."),
    ("Fragile Code:", "Ad-hoc scripts break easily when plate layouts or columns change."),
    ("Stolen Bench Time:", "Hours spent fixing syntax errors are hours lost from biology."),
    ("The Industry Secret:", "My programmer & bioinfo friends haven't typed code in a year."),
    ("Interactive AI Dev:", "They build projects in human language with fast agent feedback.")
], title_color=COLOR_DANGER)

apply_card_animations(s2, [
    CardInfo(c2_left, tb2_left, header_paras=1, is_bullet_card=True),
    CardInfo(c2_right, tb2_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 4: Web Chatbots vs. Antigravity Platform
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "Part I: The Paradigm Shift", "Web Chatbots vs. Antigravity Environment",
           "Moving from question-oriented browser chats to a goal-oriented local platform.")

c3_left = add_card(s3, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb3_left = s3.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb3_left.text_frame, "Web Chatbots (Local or Web UI)", [
    ("Disconnected from Files:", "Cannot see your workspace, raw data files, or existing scripts."),
    ("Copy-Paste Friction:", "Copy prompt -> paste code -> run in terminal -> copy error back."),
    ("No Execution Power:", "Writes code in a vacuum without testing if it actually runs."),
    ("Technical Debt:", "Produces isolated snippets that grow hard to maintain."),
    ("Mental Model:", "Like texting a friend who can't see your computer screen.")
], title_color=COLOR_DANGER)

c3_right = add_card(s3, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb3_right = s3.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb3_right.text_frame, "Antigravity Development Environment", [
    ("Native Workspace Access:", "Directly reads your CSVs, image folders, configs, and scripts."),
    ("Autonomous Execution:", "Runs scripts in your local Conda environment and checks output."),
    ("Self-Correction Loop:", "Reads errors, edits code, and re-runs—making it goal-oriented!"),
    ("Local AI Platform:", "Google Gemini ecosystem, with support for Anthropic Claude too."),
    ("Mental Model:", "A dedicated junior bioinformatician you can task and teach fast.")
], title_color=COLOR_ACCENT)

apply_card_animations(s3, [
    CardInfo(c3_left, tb3_left, header_paras=1, is_bullet_card=True),
    CardInfo(c3_right, tb3_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 5: Division of Scientific Labor
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "Part I: The Paradigm Shift", "Division of Labor: Scientist vs. Coding Agent",
           "You define the biological rules and ground truth; the agent handles implementation.")

c4_left = add_card(s4, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb4_left = s4.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb4_left.text_frame, "1. You (Senior Scientist): Specification", [
    ("Input Data Context:", "Explain plate formats, OD600 units, time steps, and blanks."),
    ("Visual Target:", "Specify plot types, axes, groupings, and 300-DPI export format."),
    ("Math & Normalization:", "Define blank subtraction, WT baselines (CBS138), and metrics."),
    ("Statistical Rules:", "Choose hypothesis tests (ANOVA, Mann-Whitney) and replicates."),
    ("Ground-Truth Audit:", "Verify results against biological reality and control datasets.")
], title_color=COLOR_TEXT_PRIMARY)

c4_right = add_card(s4, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb4_right = s4.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb4_right.text_frame, "2. Agent (Junior Bioinformatician): Execution", [
    ("Data Parsing:", "Writes parsers to clean, reshape, and validate raw exports."),
    ("Library Chaining:", "Connects pandas, scipy.stats, seaborn, and matplotlib."),
    ("Math Implementation:", "Codes normalizations, curve fits, derivatives, and stat tests."),
    ("Autonomous Debugging:", "Runs code locally, reads tracebacks, and fixes bugs automatically."),
    ("Publication Styling:", "Formats legends, fonts, palettes, and vector SVG/PNG outputs.")
], title_color=COLOR_ACCENT)

apply_card_animations(s4, [
    CardInfo(c4_left, tb4_left, header_paras=1, is_bullet_card=True),
    CardInfo(c4_right, tb4_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 6: Project 1: The Autonomous Paper Reproducer
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "Part II: What Did I Do With AGY?", "Project 1: The Autonomous Paper Reproducer",
           "Taking a genomics paper (PDF or URL) and reproducing its results from scratch.")

c5_left = add_card(s5, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb5_left = s5.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb5_left.text_frame, "5 Steps the AI Performed Autonomously", [
    ("1. Extract & Install Tools:", "Found exact versions; patched old source code for modern OSes."),
    ("2. Fast Data Downloads:", "Fetched raw repository reads and optimized slow downloads."),
    ("3. Infer Execution Order:", "Figured out how tools connect and in what sequence to run."),
    ("4. Reproduce Figures:", "Generated plots from tool outputs and compared to the paper."),
    ("5. Runtime Adaptation:", "Optimized space and runtime for laptop vs. SLURM cluster.")
], title_color=COLOR_TEXT_PRIMARY)

c5_right = add_card(s5, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb5_right = s5.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb5_right.text_frame, "Project Status & Key Takeaway", [
    ("Learning Project:", "Built to test how far AGY can push complex genomics workflows."),
    ("Preliminary Figures Ready:", "Still in development, but already produces working figures."),
    ("End-to-End Proof:", "Proved the autonomous reproduction pipeline works in practice."),
    ("Zero Manual Code:", "I didn't write a single line of code across the entire project!")
], title_color=COLOR_ACCENT)

apply_card_animations(s5, [
    CardInfo(c5_left, tb5_left, header_paras=1, is_bullet_card=True),
    CardInfo(c5_right, tb5_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 7: Case Study 1: The Branching Pipeline Architecture
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
add_header(s6, "Part II: What Did I Do With AGY?", "Case Study 1: The Branching Pipeline Architecture",
           "The branching execution pathway AGY inferred automatically from the paper.")

c6_left = add_card(s6, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb6_left = s6.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb6_left.text_frame, "Inferred Execution Pathway", [
    ("Stage 1 (Environment):", "Parsed methods text and built a dedicated Conda environment."),
    ("Stage 2 (OS Patching):", "Patched legacy C/C++ tools from source to run on macOS."),
    ("Stage 3 (Fast Streaming):", "Streamed raw FASTQ reads in parallel from ENA via aria2c."),
    ("Stage 4 (Branching DAG):", "Split reads into De Novo Pan-Genome vs. GATK SNP Phylogeny."),
    ("Stage 5 (Convergence):", "Merged gene matrix and tree in Scoary for resistance GWAS.")
], title_color=COLOR_TEXT_PRIMARY)

# Right Card: paper-repro-flow.svg rendered diagram
c6_right = add_card(s6, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb6_right = s6.shapes.add_textbox(Inches(7.1), Inches(2.05), Inches(5.2), Inches(0.45))
tf6_r = tb6_right.text_frame
tf6_r.word_wrap = True
p = tf6_r.paragraphs[0]
p.text = "Inferred Branching Pipeline (paper-repro-flow.svg)"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_ACCENT

# Insert paper-repro-flow.png (exact high-res render of paper-repro-flow.svg, aspect ratio 0.7597)
pic6_h = Inches(4.22)
pic6_w = Inches(4.22 * (982.08 / 1292.6))
pic6_left = Inches(6.9 + (5.6 - 4.22 * (982.08 / 1292.6)) / 2)
pic6_flow = s6.shapes.add_picture("paper-repro-flow.png", pic6_left, Inches(2.50), width=pic6_w, height=pic6_h)

apply_card_animations(s6, [
    CardInfo(c6_left, tb6_left, header_paras=1, is_bullet_card=True),
    CardInfo(c6_right, tb6_right, header_paras=1, is_bullet_card=False, extra_shapes=[pic6_flow])
])


# ==============================================================================
# SLIDE 8: Case Study 2: Pipeline Figures vs. Published Paper
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
add_header(s7, "Part II: What Did I Do With AGY?", "Case Study 2: Pipeline Figures vs. Published Paper",
           "Side-by-side comparison of Barber et al. (2021) Figure 1 and AGY's autonomous reproduction.")

c7_left = add_card(s7, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=RGBColor(255, 255, 255), border_color=COLOR_CARD_BORDER)
tb7_left = s7.shapes.add_textbox(Inches(1.0), Inches(2.06), Inches(5.2), Inches(0.65))
tf7_l = tb7_left.text_frame
tf7_l.word_wrap = True
p7_l1 = tf7_l.paragraphs[0]
p7_l1.text = "Published Paper: Barber et al. (2021) Figure 1"
p7_l1.font.name = FONT_FAMILY
p7_l1.font.bold = True
p7_l1.font.size = Pt(17)
p7_l1.font.color.rgb = COLOR_TEXT_PRIMARY
p7_l2 = tf7_l.add_paragraph()
p7_l2.text = "Original pan-genome matrix, rarefaction curves, and gene lengths"
p7_l2.font.name = FONT_FAMILY
p7_l2.font.size = Pt(13)
p7_l2.font.color.rgb = COLOR_TEXT_SECONDARY

# Paper figure_1_paper_cropped.png (882x688, panels c & f removed, aspect ratio 1.282)
pic7_l_w = Inches(4.75)
pic7_l_h = Inches(4.75 * (688 / 882))
pic7_paper = s7.shapes.add_picture("figure_1_paper_cropped.png", Inches(1.225), Inches(2.88), width=pic7_l_w, height=pic7_l_h)

c7_right = add_card(s7, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb7_right = s7.shapes.add_textbox(Inches(7.1), Inches(2.06), Inches(5.2), Inches(0.65))
tf7_r = tb7_right.text_frame
tf7_r.word_wrap = True
p7_r1 = tf7_r.paragraphs[0]
p7_r1.text = "AGY Pipeline Output (Preliminary Reproduction)"
p7_r1.font.name = FONT_FAMILY
p7_r1.font.bold = True
p7_r1.font.size = Pt(17)
p7_r1.font.color.rgb = COLOR_ACCENT
p7_r2 = tf7_r.add_paragraph()
p7_r2.text = "Generated from raw reads with zero manual coding"
p7_r2.font.name = FONT_FAMILY
p7_r2.font.size = Pt(13)
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
# SLIDE 9: Project 2: Rescuing the Petite Paper Figures
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
add_header(s8, "Part II: What Did I Do With AGY?", "Project 2: Rescuing the Petite Paper Figures",
           "Turning years of dispersed data and 5 bloated Jupyter notebooks into one clean project.")

c8_left = add_card(s8, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb8_left = s8.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb8_left.text_frame, "The Real-World Notebook Chaos", [
    ("5 Bloated Notebooks:", "Years of dispersed data and exploratory plots in messy cells."),
    ("Hidden Final Figures:", "Hard to remember where publishable figures lived or how to run them."),
    ("Uncoordinated Colors:", "Color palettes didn't match across figures or notebooks."),
    ("Sharing Nightmare:", "Providing working code and data to journals or peers was impractical.")
], title_color=COLOR_DANGER)

c8_right = add_card(s8, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb8_right = s8.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb8_right.text_frame, "Refactoring With Antigravity", [
    ("Weeks of Manual Work:", "Refactoring all 5 notebooks by hand would have taken me weeks."),
    ("Marking Target Cells:", "I helped AGY by marking the cells containing real production figures."),
    ("Prompt-Driven Build:", "With careful prompts, I tasked AGY to consolidate the whole project."),
    ("Zero Manual Code:", "Transformed into a clean, modular repository without typing code.")
], title_color=COLOR_TEXT_PRIMARY)

apply_card_animations(s8, [
    CardInfo(c8_left, tb8_left, header_paras=1, is_bullet_card=True),
    CardInfo(c8_right, tb8_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 10: Case Study 2: The Unified Figure Generation Pipeline
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "Part II: What Did I Do With AGY?", "Case Study 2: The Unified Figure Pipeline",
           "One script generates all figures on demand, hooked up to a live Google Sheet.")

c9_left = add_card(s9, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb9_left = s9.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb9_left.text_frame, "The Consolidated Architecture", [
    ("Single Execution Point:", "One script (`main.py`) generates all or specific figures on demand."),
    ("Documented Data Folder:", "Collected and organized all raw data into a single documented folder."),
    ("Easy Figure Polish:", "Refined legends, labels, sizes, and visual artifacts via prompts."),
    ("Ready for Typesetting:", "Reproducible code and data ready for publication and peer review.")
], title_color=COLOR_TEXT_PRIMARY)

c9_right = add_card(s9, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb9_right = s9.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb9_right.text_frame, "Live Google Sheets Palette Sync", [
    ("Online Color Config:", "Figure generation connects directly to an online Google Sheet."),
    ("Collaborator Friendly:", "Co-authors can edit and test color palettes in the spreadsheet."),
    ("Instant Updates:", "Re-running the script applies the new sheet colors across all figures."),
    ("Clean Separation:", "Visual styling is decoupled from the core analysis code.")
], title_color=COLOR_ACCENT)

apply_card_animations(s9, [
    CardInfo(c9_left, tb9_left, header_paras=1, is_bullet_card=True),
    CardInfo(c9_right, tb9_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 11: The 15-Minute Mystery: Cross-Platform Debugging
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "Part II: What Did I Do With AGY?", "The 15-Minute Mystery: Cross-Platform Debugging",
           "How AGY autonomously solved a silent Windows crash that left me flabbergasted.")

c10_left = add_card(s10, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb10_left = s10.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb10_left.text_frame, "The Silent Crash on Windows", [
    ("Worked on macOS:", "The Petite figure pipeline ran flawlessly on my Mac laptop."),
    ("Failed for Anastasia:", "When shared with Anastasia on Windows, it failed immediately."),
    ("Completely Silent:", "No warning, no error, no crash report—it just exited silently."),
    ("Finding the Line:", "I isolated the crash to a matplotlib plotting call and passed it to AGY.")
], title_color=COLOR_DANGER)

c10_right = add_card(s10, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb10_right = s10.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb10_right.text_frame, "Solved Autonomously in 15 Minutes", [
    ("Autonomous Investigation:", "AGY analyzed the matplotlib call and Windows runtime behavior."),
    ("Root Cause Identified:", "Realized the silent crash was a Windows-specific matplotlib issue."),
    ("Clean Environment Fix:", "Solved it automatically by setting an environment variable."),
    ("Massive Time Saver:", "Would have taken me hours to debug—AGY did it in 15 minutes!")
], title_color=COLOR_SUCCESS)

apply_card_animations(s10, [
    CardInfo(c10_left, tb10_left, header_paras=1, is_bullet_card=True),
    CardInfo(c10_right, tb10_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 12: LIVE DEMO BREAK: The Petite Paper Figure Generator
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
c11 = add_card(s11, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9), bg_color=COLOR_ALERT_BG, border_color=COLOR_ALERT_BORDER)
tb11 = s11.shapes.add_textbox(Inches(1.2), Inches(1.15), Inches(10.9), Inches(5.2))
tf11 = tb11.text_frame
tf11.word_wrap = True

p = tf11.paragraphs[0]
p.text = "LIVE DEMONSTRATION"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(15)
p.font.color.rgb = COLOR_ALERT_BORDER
p.space_after = Pt(8)

p = tf11.add_paragraph()
p.text = "LIVE DEMO: The Petite Paper Figure Generator"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(32)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(12)

p = tf11.add_paragraph()
p.text = "Live walkthrough of ~/dev/petite_paper: one script, cloud palette sync, publication figures."
p.font.name = FONT_FAMILY
p.font.size = Pt(17)
p.font.color.rgb = COLOR_TEXT_SECONDARY
p.space_after = Pt(24)

demo2_items = [
    ("Single-Command Run:", "Run `python main.py --figure 2` live to generate 300-DPI figures."),
    ("Google Sheets Sync:", "Change a color in Google Sheets, re-run, and watch the plot update."),
    ("Organized Data & Code:", "Tour the unified `data/` folder, documentation, and modular scripts."),
    ("Rapid Visual Polish:", "Tweak labels, legends, or layout via prompt and re-render in seconds.")
]
for t, d in demo2_items:
    p = tf11.add_paragraph()
    p.text = f"• {t} "
    p.font.name = FONT_FAMILY
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(12)
    run = p.add_run()
    run.text = d
    run.font.name = FONT_FAMILY
    run.font.bold = False
    run.font.size = Pt(16)
    run.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s11, [
    CardInfo(c11, tb11, header_paras=3, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 13: Domain Expertise: The Scientist Remains the Ground Truth
# ==============================================================================
s12 = prs.slides.add_slide(blank_layout)
add_header(s12, "Part III: What Do You Need?", "Domain Expertise: Responsibility Is Still Yours!",
           "You need strong domain knowledge to spot errors and explain what to fix.")

c12_left = add_card(s12, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb12_left = s12.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb12_left.text_frame, "The Johannes Proteomics Story", [
    ("Wrong Normalization:", "In my proteomics work, AGY normalized the dataset incorrectly."),
    ("Why It Happened:", "I hadn't explained what the data was or how it was pre-processed."),
    ("Domain Expert Catch:", "Johannes spotted the error—otherwise it would have altered results."),
    ("The Core Lesson:", "Never trust AI blindly. Responsibility for the results is still yours!")
], title_color=COLOR_DANGER)

c12_right = add_card(s12, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb12_right = s12.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb12_right.text_frame, "How to Audit Core Analysis Code", [
    ("Demand Explanations:", "Make AGY explain its math in detail and document it in comments."),
    ("Ask for Tests:", "Have the agent write automated tests for core calculations."),
    ("Use Ground Truths:", "Run pipelines on known control datasets first to verify accuracy."),
    ("Inspect & Guide:", "Be ready to explain in detail what went wrong and how to fix it.")
], title_color=COLOR_TEXT_PRIMARY)

apply_card_animations(s12, [
    CardInfo(c12_left, tb12_left, header_paras=1, is_bullet_card=True),
    CardInfo(c12_right, tb12_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 14: Technical & Architectural Discipline: Clean Systems
# ==============================================================================
s13 = prs.slides.add_slide(blank_layout)
add_header(s13, "Part III: What Do You Need?", "Technical & Architectural Knowledge",
           "Good architecture prevents code bloat, slow runs, and wasted tokens as projects grow.")

c13_left = add_card(s13, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb13_left = s13.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb13_left.text_frame, "Why Technical Intuition Helps", [
    ("Coding Isn't Obsolete:", "Intuitive understanding helps you catch mistakes early."),
    ("Animation Bug Story:", "AGY first wrote flawed slide XML rules; human intuition fixed it."),
    ("Working != Well-Built:", "Working AI code can still hide bloat and poor architecture."),
    ("The Cost of Bloat:", "Messy code burns tokens, runs slowly, and degrades as projects grow.")
], title_color=COLOR_TEXT_PRIMARY)

c13_right = add_card(s13, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb13_right = s13.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb13_right.text_frame, "Key Architectural Rules", [
    ("No Hard-Coded Values:", "Never hard-code names or numbers—compute them or use text configs."),
    ("No Code Duplication:", "Keep shared data loading and plotting logic in reusable functions."),
    ("Readable Config Files:", "Store parameters in simple text files you can edit directly."),
    ("Occasional Refactors:", "Use stronger reasoning models periodically to clean up the codebase.")
], title_color=COLOR_ACCENT)

apply_card_animations(s13, [
    CardInfo(c13_left, tb13_left, header_paras=1, is_bullet_card=True),
    CardInfo(c13_right, tb13_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 15: The Golden Rule: Code-First, Text-First Figures
# ==============================================================================
s14 = prs.slides.add_slide(blank_layout)
add_header(s14, "Part III: What Do You Need?", "The Golden Rule: Work With Text & Code, Not Images",
           "Code is text! Never ask the AI to edit scientific figure images directly.")

c14_left = add_card(s14, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb14_left = s14.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb14_left.text_frame, "Why Direct Image Editing Fails", [
    ("AIs Aren't Magicians:", "Image models paint pixels—they distort data and hallucinate labels."),
    ("Un-reproducible Tweaks:", "Manual edits in Illustrator leave no record of how a plot was made."),
    ("New Data Pain:", "When new replicates arrive, manual image edits must be redone."),
    ("Hard to Manage:", "Direct image edits can't be tracked, audited, or batch-updated.")
], title_color=COLOR_DANGER)

c14_right = add_card(s14, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb14_right = s14.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb14_right.text_frame, "The Deterministic Code Workflow", [
    ("Raw Data -> Code -> Plot:", "Give AGY raw data and have it write code to analyze and plot it."),
    ("Edit Code for Tweaks:", "Need a label moved or color changed? Have AGY edit the script."),
    ("100% Deterministic:", "Re-running the script always produces the exact same figure."),
    ("Easy to Manage:", "Text code is clean, searchable, and easy to maintain over time.")
], title_color=COLOR_ACCENT)

apply_card_animations(s14, [
    CardInfo(c14_left, tb14_left, header_paras=1, is_bullet_card=True),
    CardInfo(c14_right, tb14_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 16: How AGY Sessions Work: The Clean Slate Principle
# ==============================================================================
s15 = prs.slides.add_slide(blank_layout)
add_header(s15, "Part IV: Working with Antigravity", "How AGY Works: Projects & Conversations",
           "Local project folders, fresh conversation sessions, and real-time supervision.")

c15_left = add_card(s15, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb15_left = s15.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb15_left.text_frame, "Project Folders & Fresh Chats", [
    ("Local Project Folder:", "Everything inside your project folder is considered local and safe."),
    ("Copy Needed Files In:", "AGY asks before touching outside files—copy what you need inside."),
    ("Chats Start Fresh:", "Every new conversation starts with zero memory of prior chats."),
    ("Seed via AGENTS.md & @:", "Auto-load rules via `AGENTS.md` and reference files/chats with `@`.")
], title_color=COLOR_TEXT_PRIMARY)

c15_right = add_card(s15, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb15_right = s15.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb15_right.text_frame, "Live Monitoring & Course Correction", [
    ("See Tools Running Live:", "Watch terminal commands and file edits happen in real time."),
    ("Thought Process Note:", "Shows active steps (though recent versions hide raw chain-of-thought)."),
    ("Stop & Correct Course:", "Interrupt the agent anytime if it heads in the wrong direction."),
    ("Per-Chat Controls:", "Pick the model and thinking effort for each conversation.")
], title_color=COLOR_ACCENT)

apply_card_animations(s15, [
    CardInfo(c15_left, tb15_left, header_paras=1, is_bullet_card=True),
    CardInfo(c15_right, tb15_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 17: Context Windows & Token Hygiene: Working Memory
# ==============================================================================
s16 = prs.slides.add_slide(blank_layout)
add_header(s16, "Part IV: Working with Antigravity", "Managing the Context Window",
           "For LLMs, only the context exists—keep conversations short and focused.")

c16_left = add_card(s16, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb16_left = s16.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb16_left.text_frame, "Context Limits & Compaction", [
    ("Only Context Exists:", "If it isn't in the active context window, the LLM doesn't know it."),
    ("Auto-Compaction Drop:", "When context fills up, old turns get summarized and performance drops."),
    ("Many Short Chats Win:", "One short chat per feature, panel, or bug beats one giant chat."),
    ("Use /compact Wisely:", "Run `/compact` when context fills up AND the last task finished cleanly.")
], title_color=COLOR_TEXT_PRIMARY)

c16_right = add_card(s16, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb16_right = s16.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
tf16_r = tb16_right.text_frame
tf16_r.word_wrap = True

p = tf16_r.paragraphs[0]
p.text = "Interactive Context Visualizer"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(19)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(6)

p = tf16_r.add_paragraph()
p.text = "Open in browser: context_visualizer.html"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(13.5)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(10)

vis_points = [
    ("Visual Token Bucket:", "See how `AGENTS.md`, chat turns, and file scans fill memory."),
    ("Attention Dilution:", "Watch focus drop as old logs and large CSV dumps pile up."),
    ("CLI vs. App Stats:", "Usage stats show in CLI (harder to spot in App)—watch chat length!"),
    ("Golden Rule:", "Start a fresh conversation whenever you move to a new task.")
]
for title_text, desc_text in vis_points:
    p_v = tf16_r.add_paragraph()
    p_v.text = f"• {title_text} "
    p_v.font.name = FONT_FAMILY
    p_v.font.bold = True
    p_v.font.size = Pt(14.5)
    p_v.font.color.rgb = COLOR_TEXT_PRIMARY
    p_v.space_after = Pt(8)
    r_v = p_v.add_run()
    r_v.text = desc_text
    r_v.font.name = FONT_FAMILY
    r_v.font.bold = False
    r_v.font.size = Pt(14.5)
    r_v.font.color.rgb = COLOR_TEXT_SECONDARY

apply_card_animations(s16, [
    CardInfo(c16_left, tb16_left, header_paras=1, is_bullet_card=True),
    CardInfo(c16_right, tb16_right, header_paras=2, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 18: Model Selection: Matching Brainpower to the Task
# ==============================================================================
s17 = prs.slides.add_slide(blank_layout)
add_header(s17, "Part IV: Working with Antigravity", "Models & Thinking Effort: Matching the Task",
           "Switch between fast and deep-reasoning models to balance speed and token cost.")

c17_left = add_card(s17, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb17_left = s17.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb17_left.text_frame, "Choosing the Right Model", [
    ("Bigger = Pricier:", "Stronger models have larger contexts, but burn token quotas faster."),
    ("Cheap Models for Routine:", "Use fast models (Flash) to find files or make quick text edits."),
    ("Strong Models for Hard Tasks:", "Use Pro or Claude Sonnet/Opus for debugging and refactoring."),
    ("Switch Anytime:", "Change models between tasks so you don't waste quota.")
], title_color=COLOR_TEXT_PRIMARY)

c17_right = add_card(s17, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb17_right = s17.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb17_right.text_frame, "Thinking Levels & Chatbot Coaching", [
    ("Deep vs. Shallow Thinking:", "Dial from 'shallow' (fast, cheap) to 'deep' (slower, thorough)."),
    ("When to Think Deep:", "Use deep thinking for tough bugs and multi-file architecture."),
    ("Ask Gemini for Advice:", "Ask a chatbot to engineer your prompt and recommend a model."),
    ("Command Tips:", "Chatbots can also suggest which AGY slash commands fit your task.")
], title_color=COLOR_ACCENT)

apply_card_animations(s17, [
    CardInfo(c17_left, tb17_left, header_paras=1, is_bullet_card=True),
    CardInfo(c17_right, tb17_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 19: Prompt Engineering & Conversational Steering
# ==============================================================================
s18 = prs.slides.add_slide(blank_layout)
add_header(s18, "Part IV: Working with Antigravity", "Effective Prompting & Interactive Steering",
           "How to communicate clearly with coding agents for reliable scientific code.")

c18_left = add_card(s18, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb18_left = s18.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb18_left.text_frame, "Writing Clear Scientific Prompts", [
    ("Be Specific:", "State exact inputs, axes, units, stat tests, and figure formats."),
    ("Explain Data History:", "Tell AGY what the numbers mean and any prior normalizations."),
    ("Set Clear Guardrails:", "'Keep `data/` read-only', 'No hard-coded names', 'Use CBS138 WT'."),
    ("Request Sanity Checks:", "Ask AGY to print summary stats and explain its math first.")
], title_color=COLOR_TEXT_PRIMARY)

c18_right = add_card(s18, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb18_right = s18.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb18_right.text_frame, "Meta-Prompting & Steering", [
    ("Prompt Engineering Help:", "Ask Gemini or Claude web chat to draft your AGY prompt."),
    ("Work Incrementally:", "Break work into steps: parse files -> check stats -> plot figure."),
    ("Steer in Real Time:", "Stop the agent early if you spot a wrong biological assumption."),
    ("Point With @ Mentions:", "Use `@filename` to focus the agent on exact scripts and tables.")
], title_color=COLOR_ACCENT)

apply_card_animations(s18, [
    CardInfo(c18_left, tb18_left, header_paras=1, is_bullet_card=True),
    CardInfo(c18_right, tb18_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 20: Institutional Memory: AGENTS.md
# ==============================================================================
s19 = prs.slides.add_slide(blank_layout)
add_header(s19, "Part V: The Best Practices Toolkit", "Institutional Memory: AGENTS.md",
           "The key document auto-loaded into every conversation in your project.")

c19_left = add_card(s19, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb19_left = s19.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb19_left.text_frame, "What Goes Into AGENTS.md?", [
    ("Auto-Loaded Every Chat:", "Seeds every new conversation so AGY works fast without scanning."),
    ("Goal & File Overview:", "Explains the project goal and maps out key folders and scripts."),
    ("Tool Permissions:", "Lists which tools and commands the AI can (and cannot!) run."),
    ("Rules & Invariants:", "Defines writing style, visual standards, and coding guardrails.")
], title_color=COLOR_TEXT_PRIMARY)

# Right Box: Sample AGENTS.md
c19_right = add_card(s19, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_CODE_BG)
tb19_right = s19.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
tf19_r = tb19_right.text_frame
tf19_r.word_wrap = True

p = tf19_r.paragraphs[0]
p.text = "Example: Berman Lab AGENTS.md"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(10)

sample_code = (
    "# Lab Project Guidelines\n\n"
    "## Visual Standards\n"
    "- Colors: Sync from Google Sheet palette\n"
    "- Fonts: Arial, min 10pt tick labels\n"
    "- Output: Save 300-DPI PNG & vector SVG\n\n"
    "## Candida Conventions\n"
    "- WT reference: CBS138 (ATCC 2001)\n"
    "- No hard-coded strain names in code\n\n"
    "## Project Rules\n"
    "- Raw data in `data/` (STRICTLY READ-ONLY)\n"
    "- Run scripts via `./env/bin/python`"
)
p_code = tf19_r.add_paragraph()
p_code.text = sample_code
p_code.font.name = "Courier New"
p_code.font.size = Pt(12.5)
p_code.font.color.rgb = COLOR_TEXT_PRIMARY

apply_card_animations(s19, [
    CardInfo(c19_left, tb19_left, header_paras=1, is_bullet_card=True),
    CardInfo(c19_right, tb19_right, header_paras=1, is_bullet_card=False)
])


# ==============================================================================
# SLIDE 21: Command 1: /plan — Architect Before Execution
# ==============================================================================
s20 = prs.slides.add_slide(blank_layout)
add_header(s20, "Part V: The Best Practices Toolkit", "Slash Command: /plan — Architect First",
           "Ask AGY to draft an editable plan before starting any big chunk of work.")

c20_left = add_card(s20, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb20_left = s20.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb20_left.text_frame, "Why Use /plan?", [
    ("For Big Tasks:", "Use when starting a major feature or parsing a new file format."),
    ("Read & Edit Iteratively:", "Review the proposed plan and refine steps before coding starts."),
    ("Precise Instructions:", "Gives the agent a clear, approved blueprint to follow."),
    ("Peace of Mind:", "Ensures the agent is on the right track before touching files.")
], title_color=COLOR_TEXT_PRIMARY)

c20_right = add_card(s20, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb20_right = s20.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb20_right.text_frame, "Example Wet-Lab /plan Workflow", [
    ("Prompt:", "/plan 'Parse Tecan exports and plot 48h Candida growth curves'"),
    ("Step 1 (Inspect):", "Check raw ASCII headers and extract time-stamped OD600 values."),
    ("Step 2 (Clean):", "Subtract media blanks and flag wells with evaporation spikes."),
    ("Step 3 (Metrics):", "Fit splines to compute max growth rate (Vmax) and lag time."),
    ("Step 4 (Plot):", "Render 300-DPI facet grids grouped by strain and drug dose.")
], title_color=COLOR_ACCENT)

apply_card_animations(s20, [
    CardInfo(c20_left, tb20_left, header_paras=1, is_bullet_card=True),
    CardInfo(c20_right, tb20_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 22: Command 2: /grill-me — Interactive Alignment Interview
# ==============================================================================
s21 = prs.slides.add_slide(blank_layout)
add_header(s21, "Part V: The Best Practices Toolkit", "Slash Command: /grill-me — Resolve Ambiguities",
           "Instruct the agent to ask you about anything it's unsure of before execution.")

c21_left = add_card(s21, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb21_left = s21.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb21_left.text_frame, "How /grill-me Works", [
    ("Agent Interviews You:", "AGY asks clarifying questions instead of guessing silently."),
    ("Great With /plan:", "Use alongside `/plan` or by itself before complex analyses."),
    ("Catches Biology Details:", "Surfaces unspoken rules about blanks, dilutions, and endpoints."),
    ("Prevents Wrong Math:", "Eliminates bad assumptions before a single line of code runs.")
], title_color=COLOR_ACCENT)

c21_right = add_card(s21, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb21_right = s21.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb21_right.text_frame, "Example /grill-me Dialogue", [
    ("Prompt:", "/grill-me 'Calculate MIC50 for our fluconazole disk assays'"),
    ("Agent Q1:", "'Is 50% inhibition relative to untreated growth or media blank?'"),
    ("Agent Q2:", "'Trailing growth appears at 48h—should we score at 24h or 48h?'"),
    ("Agent Q3:", "'How should we handle resistant strains that never drop below 50%?'"),
    ("Impact:", "2 minutes of Q&A prevents hours of re-running flawed analyses.")
], title_color=COLOR_TEXT_PRIMARY)

apply_card_animations(s21, [
    CardInfo(c21_left, tb21_left, header_paras=1, is_bullet_card=True),
    CardInfo(c21_right, tb21_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 23: Command 3: /learn — Automating Knowledge Capture
# ==============================================================================
s22 = prs.slides.add_slide(blank_layout)
add_header(s22, "Part V: The Best Practices Toolkit", "Slash Command: /learn — Capture Lessons",
           "Tell the agent to learn from a completed task and update AGENTS.md or skills.")

c22_left = add_card(s22, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb22_left = s22.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb22_left.text_frame, "How /learn Builds Memory", [
    ("Run After Finishing:", "Instructs AGY to extract lessons from what it just accomplished."),
    ("Updates AGENTS.md:", "Writes new project rules into `AGENTS.md` or creates a new skill."),
    ("Fixes Stuck Loops:", "Especially useful if the agent got stuck in a loop you helped solve."),
    ("Zero Instructions Needed:", "Can be called as a bare `/learn` command all by itself!")
], title_color=COLOR_SUCCESS)

c22_right = add_card(s22, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_SUCCESS_BG, border_color=COLOR_SUCCESS)
tb22_right = s22.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb22_right.text_frame, "Case Study: Learning Slide Animations", [
    ("The Challenge:", "Teaching AGY how to animate cards and bullets in this presentation."),
    ("Running /learn:", "Once we fixed the XML logic, we ran `/learn` to save the pattern."),
    ("Saved Rule 1:", "Use the `pptx_animations` helper module—never write raw XML."),
    ("Saved Rule 2:", "1–2 cards: reveal bullets on click; 3+ cards: reveal whole cards."),
    ("Result:", "Every future chat followed our animation rules automatically!")
], title_color=COLOR_SUCCESS)

apply_card_animations(s22, [
    CardInfo(c22_left, tb22_left, header_paras=1, is_bullet_card=True),
    CardInfo(c22_right, tb22_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 24: Custom Skills: Modular Functions for Coding Agents
# ==============================================================================
s23 = prs.slides.add_slide(blank_layout)
add_header(s23, "Part V: The Best Practices Toolkit", "Skills: Reusable Functions for AI Agents",
           "Self-contained task instructions, guardrails, and tools that the AI can invoke.")

c23_left = add_card(s23, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb23_left = s23.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb23_left.text_frame, "What Are Agent Skills?", [
    ("Functions for AI:", "Describe a specific task, how to do it, guardrails, and tools."),
    ("Built-In & Imported:", "Use AGY's many built-in skills or import skills made by others."),
    ("Write Your Own:", "Package any recurring lab workflow into a custom skill."),
    ("Share Across the Lab:", "Commit skills to Git so every lab member can reuse them.")
], title_color=COLOR_TEXT_PRIMARY)

c23_right = add_card(s23, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb23_right = s23.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb23_right.text_frame, "My Palette Skill & Lab Ideas", [
    ("Master Palette Skill:", "Takes general rules or a color site URL and builds the master palette."),
    ("Auto-Config Sync:", "Populates the full figure configuration list automatically."),
    ("Judy's Writing Skill:", "Turn writing course lessons into a skill that checks manuscripts."),
    ("Assay QC Skills:", "Create shared skills for standardizing DDA and growth curve QC.")
], title_color=COLOR_ACCENT)

apply_card_animations(s23, [
    CardInfo(c23_left, tb23_left, header_paras=1, is_bullet_card=True),
    CardInfo(c23_right, tb23_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 25: Safety, Sandboxing & Data Hygiene
# ==============================================================================
s24 = prs.slides.add_slide(blank_layout)
add_header(s24, "Part V: The Best Practices Toolkit", "Safety, Security & Data Backups",
           "How to protect your raw data and review agent code before execution.")

c24_left = add_card(s24, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb24_left = s24.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb24_left.text_frame, "Safety Boundaries & Risks", [
    ("Outside Folder Prompts:", "AGY asks before touching outside files (can even get annoying!)."),
    ("Can Modify Local Files:", "Be careful—running tools CAN modify or delete files by mistake."),
    ("Hidden Access in Code:", "Generated scripts can still access outside paths when executed."),
    ("Review Before Running:", "Always check what generated code does before running it.")
], title_color=COLOR_DANGER)

c24_right = add_card(s24, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb24_right = s24.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb24_right.text_frame, "Practical Backup Strategy", [
    ("Copy Data Into Project:", "Put a working copy of raw data inside your local project folder."),
    ("Keep Archival Backups:", "Store untouched master copies on Dropbox, Google Drive, or NAS."),
    ("Read-Only data/ Rule:", "Tell AGY in `AGENTS.md` never to modify files inside `data/`."),
    ("Git + Cloud Archives:", "Use Git and Dropbox together so you can always recover files.")
], title_color=COLOR_ACCENT)

apply_card_animations(s24, [
    CardInfo(c24_left, tb24_left, header_paras=1, is_bullet_card=True),
    CardInfo(c24_right, tb24_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 26: Version Control: Git & GitHub
# ==============================================================================
s25 = prs.slides.add_slide(blank_layout)
add_header(s25, "Part V: The Best Practices Toolkit", "Git, GitHub & Parallel Worktrees",
           "Track changes, roll back mistakes, collaborate safely, and run tasks in parallel.")

c25_left = add_card(s25, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb25_left = s25.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb25_left.text_frame, "Why Use Git & GitHub?", [
    ("Track & Roll Back:", "Keep a full history of changes and undo broken edits in seconds."),
    ("Safe Collaboration:", "Private GitHub repos let Anastasia and me share code safely."),
    ("Never Store Secrets:", "Never commit passwords or API keys into your Git repository!"),
    ("Git Worktrees:", "Advanced feature to run multiple tasks in parallel without conflict.")
], title_color=COLOR_TEXT_PRIMARY)

# Right Card: Visual Git Tree Schematic for Newcomers (High-res rendered figure)
c25_right = add_card(s25, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb25_right = s25.shapes.add_textbox(Inches(7.1), Inches(2.05), Inches(5.2), Inches(0.60))
tf25_r = tb25_right.text_frame
tf25_r.word_wrap = True
p = tf25_r.paragraphs[0]
p.text = "Schematic: How Version Control Works"
p.font.name = FONT_FAMILY
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_SUCCESS

p2 = tf25_r.add_paragraph()
p2.text = "Visual model of commits, branches, merges, and instant rollbacks"
p2.font.name = FONT_FAMILY
p2.font.size = Pt(13)
p2.font.color.rgb = COLOR_TEXT_SECONDARY

# Insert git_tree_schematic.png (1560x1200, aspect ratio 1.30)
pic25_w = Inches(5.30)
pic25_h = Inches(5.30 * (1200 / 1560))
pic25_left = Inches(6.9 + (5.6 - 5.30) / 2)
pic25_git = s25.shapes.add_picture("git_tree_schematic.png", pic25_left, Inches(2.68), width=pic25_w, height=pic25_h)

apply_card_animations(s25, [
    CardInfo(c25_left, tb25_left, header_paras=1, is_bullet_card=True),
    CardInfo(c25_right, tb25_right, header_paras=2, is_bullet_card=False, extra_shapes=[pic25_git])
])


# ==============================================================================
# SLIDE 27: Berman Lab Member Missions & Shared Infrastructure
# ==============================================================================
s26 = prs.slides.add_slide(blank_layout)
add_header(s26, "Part VI: Next Steps", "Personal Missions & Legacy Lab Tools",
           "Immediate ideas for lab members and maintaining our classic lab software.")

c26_left = add_card(s26, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb26_left = s26.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb26_left.text_frame, "Lab Member Missions", [
    ("Ze'ev (DDA Panel Creator):", "Feed its GitHub issues to AGY—finish them all within a week!"),
    ("Shani (Consolidate Work):", "Move from Claude UI to AGY and unify your analysis project."),
    ("Sofya (Transposon QC):", "Build final quality assessments and exploratory plots for libraries."),
    ("Alona (Claude in AGY):", "Plug your Claude sub into AGY, or test the Claude Code CLI."),
    ("Judy (Data & Writing Skills):", "Analyze data independently, and turn writing rules into skills!")
], title_color=COLOR_TEXT_PRIMARY)

c26_right = add_card(s26, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb26_right = s26.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb26_right.text_frame, "Maintaining Classic Lab Tools", [
    ("Revive Older Tools:", "Maintain OD Normalizer, Scanalyzer, diskImageR, and even Ymap."),
    ("No Programmer Required:", "Fix and update legacy tools without needing the original author."),
    ("Doable With Setup:", "Requires initial project setup per tool, then steerable in plain English."),
    ("Shared Lab Assets:", "Keep our core yeast analysis tools alive and modern for everyone.")
], title_color=COLOR_ACCENT)

apply_card_animations(s26, [
    CardInfo(c26_left, tb26_left, header_paras=1, is_bullet_card=True),
    CardInfo(c26_right, tb26_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 28: Platform Economics & Navigating Subscriptions
# ==============================================================================
s27 = prs.slides.add_slide(blank_layout)
add_header(s27, "Part VI: Next Steps", "Subscriptions, Quotas & Model Options",
           "Navigating free tiers, TAU academic subscriptions, and API keys.")

c27_left = add_card(s27, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb27_left = s27.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb27_left.text_frame, "Gemini Tiers & TAU Pricing", [
    ("AGY App Is Free:", "Free tier included, though quotas can run out quickly with heavy use."),
    ("TAU Gemini Pro (550 NIS/yr):", "Check if TAU's license covers developer API quotas or just web chat."),
    ("Private Sub (~75 NIS/mo):", "Standard developer subscription or API billing is ~900 NIS/year."),
    ("Other TAU Access:", "TAU has OpenAI & Microsoft access (Copilot is weaker for code).")
], title_color=COLOR_TEXT_PRIMARY)

c27_right = add_card(s27, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9))
tb27_right = s27.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb27_right.text_frame, "Claude vs. Gemini for Biology", [
    ("Plug In Claude Models:", "Connect Claude 3.5/3.7 Sonnet or Opus to AGY via API keys."),
    ("Gemini Pro Is Enough:", "While Claude is praised for code, Gemini Pro met all my bio needs."),
    ("Very Cost-Effective:", "Gemini Pro handled both of my biology pipelines at low cost."),
    ("Try It Today:", "Start with the free tier on a small lab dataset this week!")
], title_color=COLOR_ACCENT)

apply_card_animations(s27, [
    CardInfo(c27_left, tb27_left, header_paras=1, is_bullet_card=True),
    CardInfo(c27_right, tb27_right, header_paras=1, is_bullet_card=True)
])


# ==============================================================================
# SLIDE 29: Bonus Cheatsheet: The 8 Core Agent Habits for Reliable Science
# ==============================================================================
s28 = prs.slides.add_slide(blank_layout)
add_header(s28, "Bonus Cheatsheet", "8 Habits for Reliable Agentic Science",
           "Keep this handy: essential rules of thumb for working with Antigravity.")

c28_left = add_card(s28, Inches(0.8), Inches(1.95), Inches(5.6), Inches(4.9))
tb28_left = s28.shapes.add_textbox(Inches(1.05), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb28_left.text_frame, "Setup & Context (Habits 1–4)", [
    ("1. Copy Raw Data In:", "Keep archival backups on Dropbox/NAS; keep local `data/` read-only."),
    ("2. Seed via AGENTS.md:", "Document goals, files, allowed tools, and style rules once."),
    ("3. Plan & Grill First:", "Run `/plan` and `/grill-me` before coding to align on details."),
    ("4. Many Short Chats:", "Use one short chat per feature or figure so context stays sharp.")
], title_color=COLOR_TEXT_PRIMARY)

c28_right = add_card(s28, Inches(6.9), Inches(1.95), Inches(5.6), Inches(4.9), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb28_right = s28.shapes.add_textbox(Inches(7.15), Inches(2.15), Inches(5.1), Inches(4.5))
add_bullet_list(tb28_right.text_frame, "Execution & Safety (Habits 5–8)", [
    ("5. Compact Wisely:", "Use `/compact` after a task finishes cleanly before context fills up."),
    ("6. Capture Fixes (/learn):", "Run `/learn` after solving bugs or loops to update `AGENTS.md`."),
    ("7. Edit Code, Not Images:", "Never edit figures directly—have AGY update the plotting script."),
    ("8. Use Git & Audit Math:", "Commit milestones to Git, and always verify core math on controls.")
], title_color=COLOR_ACCENT)

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
