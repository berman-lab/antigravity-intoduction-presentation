"""
Generate Antigravity Lab Presentation (.pptx)
Widescreen 16:9, clean white background, high contrast, 1-2 column layout.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

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

# Main container
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
# SLIDE 2: The Wet-Lab Computational Bottleneck
# ==============================================================================
s2 = prs.slides.add_slide(blank_slide_layout)
add_header(s2, "Part I: The Paradigm Shift", "The Wet-Lab Computational Bottleneck",
           "Why aren't we all coding? The barrier isn't the biological question—it's execution friction.")

# Left Column: The Scientific Reality
add_card(s2, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s2.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Wet-Lab Reality"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

bullets = [
    ("Rich Experimental Assays:", "Plate reader growth curves, disk diffusion, spot dilutions, MitoTracker microscopy, proteomics."),
    ("Clear Hypotheses:", "You know what biological question you're asking and what the output figure should show."),
    ("The Excel Trap:", "Manual copy-pasting into spreadsheets, inconsistent color formatting, hard-to-update plots when replicates change.")
]
for title, desc in bullets:
    p = tf.add_paragraph()
    p.text = f"• {title} "
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(10)
    # Add body text
    run = p.add_run()
    run.text = desc
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

# Right Column: The Two Old Barriers
add_card(s2, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s2.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Traditional Coding Barriers"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

barriers = [
    ("Barrier 1: Syntax & Environment Setup", "Learning Python/R, installing packages, managing environments, and memorizing syntax takes months of dedicated effort."),
    ("Barrier 2: Implementation Time Sink", "Even if you know some basics, writing boilerplate, parsing non-standard CSVs, and tweaking Matplotlib takes days away from the bench."),
    ("The Real Cost:", "Researchers run simpler, smaller experiments simply to avoid the data analysis bottleneck.")
]
for title, desc in barriers:
    p = tf.add_paragraph()
    p.text = f"• {title}\n  "
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = COLOR_DANGER if "Barrier" in title else COLOR_TEXT_PRIMARY
    p.space_after = Pt(10)
    run = p.add_run()
    run.text = desc
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY


# ==============================================================================
# SLIDE 3: Chatbots vs. Coding Agents
# ==============================================================================
s3 = prs.slides.add_slide(blank_slide_layout)
add_header(s3, "Part I: The Paradigm Shift", "Web Chatbots vs. Autonomous Coding Agents",
           "A fundamental difference in architecture, capability, and workflow.")

# Left Card: Web Chatbots
add_card(s3, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s3.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Web Chatbots (ChatGPT / Web UI)"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

cb_points = [
    ("Disconnected from files:", "Cannot see your workspace, raw data files, or existing scripts."),
    ("Copy-Paste Friction:", "You copy a prompt -> paste code to a file -> run in terminal -> copy error back to browser."),
    ("No Execution Power:", "The chatbot generates code in a vacuum without testing whether it actually runs or throws errors."),
    ("Mental Model:", "Like texting a friend for advice who can't see your computer screen.")
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

# Right Card: Antigravity Agents
add_card(s3, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb = s3.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Antigravity Coding Agent"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

ag_points = [
    ("Native Workspace Access:", "Directly reads your CSVs, image directories, config files, and scripts."),
    ("Autonomous Execution:", "Runs scripts directly in your local conda environment and inspects output."),
    ("Self-Correction Loop:", "If code throws a traceback or bad shape, the agent reads the error, modifies code, and re-runs."),
    ("Mental Model:", "Like sitting beside a junior bioinformatician who types and executes on your machine under your direction.")
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


# ==============================================================================
# SLIDE 4: The Mental Model
# ==============================================================================
s4 = prs.slides.add_slide(blank_slide_layout)
add_header(s4, "Part I: The Paradigm Shift", "Mental Model: Your On-Call Junior Bioinformatician",
           "A collaborative division of labor: You provide the biology; the agent provides the code.")

# 3 Horizontal Cards
card_w = Inches(3.64)
card_h = Inches(4.8)

# Card 1: Your Role
add_card(s4, Inches(0.8), Inches(2.0), card_w, card_h)
tb = s4.shapes.add_textbox(Inches(1.0), Inches(2.2), card_w - Inches(0.4), card_h - Inches(0.4))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "1. You (Senior Scientist)"
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(12)
items1 = [
    ("Biological Hypothesis", "What question are we testing with this mutant?"),
    ("Experimental Context", "Which strains, drugs, concentrations, replicates?"),
    ("Scientific Logic", "Formulas (doubling times, MIC50, WT normalization)."),
    ("Quality Control", "Reviewing figures, catching biological anomalies.")
]
for t, d in items1:
    p = tf.add_paragraph()
    p.text = f"• {t}:\n  {d}"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_SECONDARY
    p.space_after = Pt(8)

# Card 2: The Agent's Role
add_card(s4, Inches(4.84), Inches(2.0), card_w, card_h, bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb = s4.shapes.add_textbox(Inches(5.04), Inches(2.2), card_w - Inches(0.4), card_h - Inches(0.4))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "2. Agent (Junior Programmer)"
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(12)
items2 = [
    ("Syntax & Libraries", "Matplotlib, Seaborn, Pandas, Scipy, Statannotations."),
    ("Data Wrangling", "Parsing messy plate reader CSVs, transposing tables."),
    ("Execution & Debugging", "Running code, reading tracebacks, fixing bugs."),
    ("Styling & Formatting", "300-DPI PNGs, vector SVGs, font scaling, palette sync.")
]
for t, d in items2:
    p = tf.add_paragraph()
    p.text = f"• {t}:\n  {d}"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_SECONDARY
    p.space_after = Pt(8)

# Card 3: The Result
add_card(s4, Inches(8.88), Inches(2.0), card_w, card_h)
tb = s4.shapes.add_textbox(Inches(9.08), Inches(2.2), card_w - Inches(0.4), card_h - Inches(0.4))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "3. The Scientific Result"
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_SUCCESS
p.space_after = Pt(12)
items3 = [
    ("10x Faster Iteration", "Change a palette or statistical test in 15 seconds."),
    ("Reproducible Code", "Modular scripts committed to git, not mysterious Excel files."),
    ("Democratized Bioinformatics", "Every bench scientist can run sophisticated computational analyses."),
    ("Zero 'Vibe Coding'", "Code is audited, versioned, testable, and deterministic.")
]
for t, d in items3:
    p = tf.add_paragraph()
    p.text = f"• {t}:\n  {d}"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_SECONDARY
    p.space_after = Pt(8)


# ==============================================================================
# SLIDE 5: The Context Window & Token Hygiene
# ==============================================================================
s5 = prs.slides.add_slide(blank_slide_layout)
add_header(s5, "Part II: Steering the Agent", "Context Windows & Token Hygiene",
           "An agent's working memory is finite. Good token hygiene makes or breaks your results.")

add_card(s5, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s5.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The 'Marathon Chat' Trap"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

trap_points = [
    ("What is the Context Window?", "The total amount of conversation history, files, and tool outputs the agent can 'see' at any one moment."),
    ("Context Saturation:", "When a single chat runs for 50 turns across 5 different topics, older instructions get diluted or dropped."),
    ("Symptom: 'Lazy' or Confused Agent:", "The agent starts hallucinating old filenames, forgetting style rules, or taking longer to respond."),
    ("Cost & Latency:", "Re-sending 150,000 tokens on every prompt wastes time and compute.")
]
for t, d in trap_points:
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

add_card(s5, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s5.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Best Practices for Token Hygiene"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_SUCCESS
p.space_after = Pt(14)

good_points = [
    ("One Task / One Figure Per Session:", "Treat each conversation like an isolated lab notebook entry (e.g., 'Analyze Figure 2 Growth Curves')."),
    ("Start Fresh Freely:", "Code and output files live in your workspace. Starting a new chat doesn't delete your code!"),
    ("Use the /compact Command:", "If a conversation is going well but getting long, type /compact to summarize context and free working memory."),
    ("Rely on Files, Not Memory:", "Put permanent instructions in AGENTS.md so fresh chats instantly know everything about your project.")
]
for t, d in good_points:
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


# ==============================================================================
# SLIDE 6: Choosing the Right Model
# ==============================================================================
s6 = prs.slides.add_slide(blank_slide_layout)
add_header(s6, "Part II: Steering the Agent", "Choosing the Right Model for the Job",
           "Match computational firepower to the complexity of the task.")

add_card(s6, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s6.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Lightweight / Fast Models (e.g., Flash)"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

flash_items = [
    ("Characteristics:", "Blazing fast response times (1-3 seconds), minimal token consumption, excellent for standard syntax."),
    ("Ideal Lab Tasks:", [
        "Tweaking plot aesthetics (changing colors, font sizes, labels)",
        "Parsing simple CSV formats or renaming columns",
        "Writing quick boilerplate or unit tests",
        "Explaining an existing error traceback"
    ]),
    ("When NOT to use:", "Deep architectural refactoring, complex multi-file debugging, or nuanced mathematical modeling.")
]
p = tf.add_paragraph()
p.text = "• Characteristics: "
p.font.bold = True
p.font.size = Pt(13)
p.add_run().text = flash_items[0][1]
p.space_after = Pt(10)

p = tf.add_paragraph()
p.text = "• Best Use Cases in the Lab:"
p.font.bold = True
p.font.size = Pt(13)
for sub in flash_items[1][1]:
    p2 = tf.add_paragraph()
    p2.text = f"    - {sub}"
    p2.font.size = Pt(12)
    p2.font.color.rgb = COLOR_TEXT_SECONDARY

add_card(s6, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s6.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Deep Reasoning Models (e.g., Pro / Thinking)"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

pro_items = [
    ("Characteristics:", "Generates an internal reasoning chain before outputting code; evaluates edge cases and alternative hypotheses."),
    ("Ideal Lab Tasks:", [
        "Designing an end-to-end figure pipeline from scratch",
        "Implementing complex curve-fitting (non-linear regression, Gompertz models)",
        "Cross-correlating disparate assays (e.g., proteomics + growth + microscopy)",
        "Diagnosing subtle bugs across multiple linked modules"
    ]),
    ("Best Practice:", "Use Pro/Thinking for the initial design & architecture; switch to Flash for fast iterative cosmetic adjustments.")
]
p = tf.add_paragraph()
p.text = "• Characteristics: "
p.font.bold = True
p.font.size = Pt(13)
p.add_run().text = pro_items[0][1]
p.space_after = Pt(10)

p = tf.add_paragraph()
p.text = "• Best Use Cases in the Lab:"
p.font.bold = True
p.font.size = Pt(13)
for sub in pro_items[1][1]:
    p2 = tf.add_paragraph()
    p2.text = f"    - {sub}"
    p2.font.size = Pt(12)
    p2.font.color.rgb = COLOR_TEXT_SECONDARY


# ==============================================================================
# SLIDE 7: Slash Commands
# ==============================================================================
s7 = prs.slides.add_slide(blank_slide_layout)
add_header(s7, "Part II: Steering the Agent", "Power Steering: Key Slash Commands",
           "Built-in shortcuts to automate complex agent workflows and enforce rigor.")

commands = [
    ("/plan", "Architect Before Coding",
     "Forces the agent to produce a structured, step-by-step roadmap before touching code.\n• Prevents hasty or broken implementations\n• Gives you a review checkpoint: approve the plan before execution begins.",
     COLOR_TEXT_PRIMARY),
    ("/grill-me", "Interactive Alignment Interview",
     "Instructs the agent to interrogate YOU with targeted questions to clarify ambiguities.\n• Uncovers missing parameters (e.g., 'What units are in column 3? Which strain is WT?')\n• Ensures the agent never makes incorrect biological assumptions.",
     COLOR_ACCENT),
    ("/learn", "Institutional Habit Learning",
     "Saves a successful workflow or correction into permanent agent memory.\n• 'Remember to always format yeast gene symbols with italics and CGD rules'\n• Turns hard-won solutions into automatic future habits.",
     COLOR_SUCCESS),
    ("/compact & /boost", "Memory & Deep Reasoning",
     "• /compact: Condenses lengthy chat context to refresh token budget.\n• /boost: Activates deep, multi-perspective strategic thinking for high-complexity problems.",
     COLOR_TEXT_SECONDARY)
]

for i, (cmd, sub, desc, col) in enumerate(commands):
    x = Inches(0.8 + (i % 2) * 6.1)
    y = Inches(2.0 + (i // 2) * 2.5)
    add_card(s7, x, y, Inches(5.6), Inches(2.2))
    tb = s7.shapes.add_textbox(x + Inches(0.25), y + Inches(0.2), Inches(5.1), Inches(1.8))
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


# ==============================================================================
# SLIDE 8: AGENTS.md
# ==============================================================================
s8 = prs.slides.add_slide(blank_slide_layout)
add_header(s8, "Part II: Steering the Agent", "Institutional Memory: AGENTS.md",
           "Teach your agent once; benefit across every future conversation.")

add_card(s8, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s8.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "What is AGENTS.md?"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(12)

md_points = [
    ("A Single Source of Truth:", "A plain Markdown file placed in your project root. Antigravity automatically reads it at the start of every chat."),
    ("Permanent Lab Rules:", "Define your preferred color palettes, statistical thresholds, naming conventions, and directory structures."),
    ("Zero Repetitive Prompting:", "No need to re-type: 'Remember our WT strain is CBS138 and all figures should be 300 DPI' every morning."),
    ("Lab-Wide Standardization:", "Commit AGENTS.md to git so every lab member's agent follows the exact same lab standards.")
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
add_card(s8, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_CODE_BG)
tb = s8.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
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
    "- Raw data in `data/` (NEVER modify raw files)\n"
    "- Reusable utilities in `utils/`\n"
    "- Master figure outputs in `output/`"
)
p = tf.add_paragraph()
p.text = sample_code
p.font.name = "Courier New"
p.font.size = Pt(10)
p.font.color.rgb = COLOR_TEXT_PRIMARY


# ==============================================================================
# SLIDE 9: Code-First Figures
# ==============================================================================
s9 = prs.slides.add_slide(blank_slide_layout)
add_header(s9, "Part II: Steering the Agent", "The Golden Rule: Code-First Figures",
           "Never manually edit scientific figures in Photoshop or Illustrator.")

add_card(s9, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s9.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Trap of Manual Figure Editing"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

manual_traps = [
    ("Un-reproducible Tweaks:", "Resizing an axis or changing a bar color in Illustrator leaves no record of how the figure was created."),
    ("The 'New Replicate' Nightmare:", "When you add 3 new biological replicates, you have to redo every manual adjustment from scratch."),
    ("Subtle Data Distortions:", "Accidentally nudging an error bar or alignment during manual assembly compromises scientific integrity."),
    ("Non-Transferable:", "Nobody else in the lab can reproduce or update your figure once you graduate or leave.")
]
for t, d in manual_traps:
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

add_card(s9, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb = s9.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Antigravity Approach: 100% Code"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

code_wins = [
    ("Every Pixel is Defined in Code:", "Colors, typography, significance stars (***), and composite panel layouts are fully scripted."),
    ("10-Second Figure Re-generation:", "New data arrives? Run: `python main.py --figure 2` and your complete 300-DPI figure updates instantly."),
    ("Centralized Palette Syncing:", "Change a color once in a config table, and all 15 figures across the paper update synchronously."),
    ("Reviewer-Proof Provenance:", "Reviewers or editors ask for raw code? It's all cleanly committed in your Git repository.")
]
for t, d in code_wins:
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


# ==============================================================================
# SLIDE 10: LIVE DEMO 1 (Break)
# ==============================================================================
s10 = prs.slides.add_slide(blank_slide_layout)
# Distinctive Live Demo Background Banner
add_card(s10, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9), bg_color=COLOR_ALERT_BG, border_color=COLOR_ALERT_BORDER)
tb = s10.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(10.9), Inches(5.1))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "LIVE DEMONSTRATION 1"
p.font.bold = True
p.font.size = Pt(14)
p.font.color.rgb = COLOR_ALERT_BORDER
p.space_after = Pt(8)

p = tf.add_paragraph()
p.text = "The Petite Mutant Figure Reproduction Pipeline"
p.font.bold = True
p.font.size = Pt(30)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(16)

p = tf.add_paragraph()
p.text = "Walkthrough of a real Berman Lab project: ~/dev/petite_paper"
p.font.size = Pt(16)
p.font.color.rgb = COLOR_TEXT_SECONDARY
p.space_after = Pt(24)

demo1_items = [
    ("1. Modular Architecture:", "Showing clean separation: raw data -> reusable utils -> modular figure scripts (`figures_growth.py`, `figures_microscopy.py`)."),
    ("2. Multi-Modal Wet-Lab Data:", "Growth curves + MIC50 (Fig 2), MitoTracker microscopy montages (Fig 3), Proteomics PCA & GO enrichment (Fig 4)."),
    ("3. Central Style & Palette Control:", "Demonstrating how Okabe-Ito / Dark2 palettes sync across all figures automatically."),
    ("4. One-Click Reproduction:", "Running `python main.py --figure 2` to generate publication-ready PNG & vector SVG.")
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


# ==============================================================================
# SLIDE 11: Git & GitHub
# ==============================================================================
s11 = prs.slides.add_slide(blank_slide_layout)
add_header(s11, "Part IV: Rigor, Safety & Scaling", "Git & GitHub: Your Scientific Safety Net",
           "Why version control is the essential partner for AI-assisted coding.")

add_card(s11, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s11.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Fear: 'What if it breaks my code?'"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

fear_points = [
    ("Common Hesitation:", "Non-programmers worry that an agent will silently modify or ruin a script that was working yesterday."),
    ("Without Git:", "Files named `analysis_v1.py`, `analysis_final.py`, `analysis_final_fixed2.py`—chaotic and fragile."),
    ("Loss of History:", "Impossible to know which prompt introduced a change or why an output looks different.")
]
for t, d in fear_points:
    p = tf.add_paragraph()
    p.text = f"• {t}\n  "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(10)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

add_card(s11, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s11.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "With Git: Total Fearlessness"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_SUCCESS
p.space_after = Pt(14)

git_points = [
    ("The Universal Undo Button:", "Every change made by Antigravity is tracked. If an experiment doesn't work, revert with one click."),
    ("Visual Diffs (Line-by-Line):", "Review exact green/red code diffs before accepting changes."),
    ("Auditability & Provenance:", "Complete scientific record: who wrote what, when, and what commit generated the published figure."),
    ("Collaboration:", "Share pipelines cleanly across lab members and compute clusters.")
]
for t, d in git_points:
    p = tf.add_paragraph()
    p.text = f"• {t}\n  "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(10)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY


# ==============================================================================
# SLIDE 12: Data Hygiene & Documentation
# ==============================================================================
s12 = prs.slides.add_slide(blank_slide_layout)
add_header(s12, "Part IV: Rigor, Safety & Scaling", "Data Hygiene & Documentation: Feeding the Agent",
           "An agent is only as smart as the context you provide. Good documentation pays 10x dividends.")

# 3 Columns for 3 Principles
c_w = Inches(3.64)
c_h = Inches(4.8)

# Col 1: Read-Only Raw Data
add_card(s12, Inches(0.8), Inches(2.0), c_w, c_h)
tb = s12.shapes.add_textbox(Inches(1.0), Inches(2.2), c_w - Inches(0.4), c_h - Inches(0.4))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "1. Immutable Raw Data"
p.font.bold = True
p.font.size = Pt(16)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(12)
c1_items = [
    ("Strict Separation", "Keep raw plate reader exports, microscopes, and sequencing files in a dedicated `data/` folder."),
    ("Never Overwrite Raw", "Instruct the agent that raw data is read-only. All parsed/cleaned files go to `processed/` or `output/`.")
]
for t, d in c1_items:
    p = tf.add_paragraph()
    p.text = f"• {t}:\n  {d}"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_SECONDARY
    p.space_after = Pt(10)

# Col 2: data/README.md
add_card(s12, Inches(4.84), Inches(2.0), c_w, c_h, bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb = s12.shapes.add_textbox(Inches(5.04), Inches(2.2), c_w - Inches(0.4), c_h - Inches(0.4))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "2. The data/README.md"
p.font.bold = True
p.font.size = Pt(16)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(12)
c2_items = [
    ("Column Definitions", "Explicitly state what 'OD600', 'slope', 'norm_val' represent."),
    ("Strain Metadata", "Map sample IDs (e.g., 'YJB123') to biological genotypes ('cgrho0', 'rho+')."),
    ("Units & Conditions", "Document media, drug concentrations (µg/mL), temperatures (30°C vs 37°C).")
]
for t, d in c2_items:
    p = tf.add_paragraph()
    p.text = f"• {t}:\n  {d}"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_SECONDARY
    p.space_after = Pt(10)

# Col 3: Centralized Paths
add_card(s12, Inches(8.88), Inches(2.0), c_w, c_h)
tb = s12.shapes.add_textbox(Inches(9.08), Inches(2.2), c_w - Inches(0.4), c_h - Inches(0.4))
tf = tb.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "3. Centralized Paths"
p.font.bold = True
p.font.size = Pt(16)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(12)
c3_items = [
    ("No Hardcoded Paths", "Never hardcode `/Users/alice/Desktop/test.csv` in plotting code."),
    ("Central Config", "Use a single `config/data_paths.py` or `.env` so any lab member can run the pipeline on their own laptop."),
    ("Instant Scalability", "Switching from test data to full paper data takes 1 line in config.")
]
for t, d in c3_items:
    p = tf.add_paragraph()
    p.text = f"• {t}:\n  {d}"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_SECONDARY
    p.space_after = Pt(10)


# ==============================================================================
# SLIDE 13: Parallelizing with Git Worktrees
# ==============================================================================
s13 = prs.slides.add_slide(blank_slide_layout)
add_header(s13, "Part IV: Rigor, Safety & Scaling", "Advanced Scale: Parallel Agents with Git Worktrees",
           "Why wait for one agent to finish? Run 3 independent agents on 3 figures simultaneously.")

add_card(s13, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s13.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Sequential Bottleneck"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

seq_points = [
    ("The Single-Branch Problem:", "If Agent A is refactoring Figure 2 while Agent B is writing Figure 4 on the same branch, files collide and git gets messy."),
    ("Waiting Around:", "Bench scientists shouldn't have to wait 10 minutes for an image analysis task to finish before plotting a plate reader curve.")
]
for t, d in seq_points:
    p = tf.add_paragraph()
    p.text = f"• {t}\n  "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_TEXT_PRIMARY
    p.space_after = Pt(10)
    run = p.add_run()
    run.text = d
    run.font.bold = False
    run.font.color.rgb = COLOR_TEXT_SECONDARY

add_card(s13, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb = s13.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Worktree Solution (As in petite_paper)"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

wt_points = [
    ("Isolated Parallel Workspaces:", "Git worktrees allow multiple directories linked to the same repository, each checked out to a different branch."),
    ("Zero Disk Bloat:", "Shares git history and large conda environments without duplicating gigabytes of files."),
    ("Multi-Agent Execution:", "Session 1: Agent works on Growth Curves (`fig2-branch`).\nSession 2: Agent works on Microscopy (`fig3-branch`)."),
    ("One-Command Merge:", "When done: `./utils/worktree.py merge fig2` cleanly merges back to main.")
]
for t, d in wt_points:
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


# ==============================================================================
# SLIDE 14: Prompting Like a Lab PI
# ==============================================================================
s14 = prs.slides.add_slide(blank_slide_layout)
add_header(s14, "Part IV: Rigor, Safety & Scaling", "How to Direct the Agent: Prompting Like a PI",
           "Treat the agent like a talented junior researcher: specific context yields publication-ready results.")

add_card(s14, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s14.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Vague 'Chatbot' Prompt"
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(10)

p = tf.add_paragraph()
p.text = '❌ "Plot my yeast growth curves from the CSV file."'
p.font.size = Pt(13)
p.font.italic = True
p.font.color.rgb = COLOR_DANGER
p.space_after = Pt(14)

vague_issues = [
    ("Agent must guess columns:", "Is OD600 raw or blank-subtracted? What is column 'Well_B2'?"),
    ("Generic, ugly styling:", "Produces default blue/orange Matplotlib lines, tiny fonts, unreadable legend."),
    ("Missing biology:", "Doesn't know which strain is WT or what drug concentration was tested."),
    ("Result:", "You spend 30 minutes going back and forth making minor corrections.")
]
for t, d in vague_issues:
    p = tf.add_paragraph()
    p.text = f"• {t} {d}"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_TEXT_SECONDARY
    p.space_after = Pt(8)

add_card(s14, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb = s14.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "The Directed 'Lab PI' Prompt"
p.font.bold = True
p.font.size = Pt(17)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(10)

p = tf.add_paragraph()
p.text = '✅ "Read data/202611_tecan_fluconazole.csv. Wells A1-A3 are WT (CBS138); B1-B3 are cgrho0. Subtract media blank (H12). Plot OD600 vs time (0-48h) with mean +/- SD shading. Use our Okabe-Ito palette. Save as Figure_2B.png (300 DPI)."'
p.font.size = Pt(11)
p.font.italic = True
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

good_elements = [
    ("Goal:", "Explicit scientific target (doubling time / growth curves)."),
    ("Data Inputs:", "Exact path and column/well layout."),
    ("Biological Logic:", "Blank subtraction, replicates, statistical aggregation."),
    ("Output Target:", "300-DPI publication file adhering to project styling.")
]
for t, d in good_elements:
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


# ==============================================================================
# SLIDE 15: LIVE DEMO 2 (Break)
# ==============================================================================
s15 = prs.slides.add_slide(blank_slide_layout)
add_card(s15, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9), bg_color=COLOR_ALERT_BG, border_color=COLOR_ALERT_BORDER)
tb = s15.shapes.add_textbox(Inches(1.2), Inches(1.2), Inches(10.9), Inches(5.1))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "LIVE DEMONSTRATION 2"
p.font.bold = True
p.font.size = Pt(14)
p.font.color.rgb = COLOR_ALERT_BORDER
p.space_after = Pt(8)

p = tf.add_paragraph()
p.text = "From Scratch: Raw Data to Publication Plot in 5 Minutes"
p.font.bold = True
p.font.size = Pt(30)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(16)

p = tf.add_paragraph()
p.text = "Demystifying the process for the skeptic: Watching an agent work live"
p.font.size = Pt(16)
p.font.color.rgb = COLOR_TEXT_SECONDARY
p.space_after = Pt(24)

demo2_items = [
    ("1. The Input:", "Drop a raw, unformatted yeast plate reader CSV or spot dilution table into a clean folder."),
    ("2. The Directive:", "Direct Antigravity to inspect columns, parse conditions, and design a plot."),
    ("3. Watch the Self-Correction:", "Observe how the agent runs python, encounters a minor data formatting hitch, inspects the error trace, and fixes it autonomously."),
    ("4. The Final Output:", "View the resulting high-res plot generated entirely by audited code.")
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


# ==============================================================================
# SLIDE 16: What This Unlocks for Your Science
# ==============================================================================
s16 = prs.slides.add_slide(blank_slide_layout)
add_header(s16, "Part V: The Path Forward", "What This Unlocks for Your Science",
           "Shifting the bottleneck from 'How do I code this?' to 'What is the most interesting question?'")

# 3 Pillars
p_w = Inches(3.64)
p_h = Inches(4.8)

pillars = [
    ("1. Speed of Iteration",
     "Explore alternative hypotheses instantly.\n\n• Test 5 different growth models in 20 minutes instead of 2 weeks.\n• Re-run complete pipelines across new drug concentrations with 1 terminal command.\n• Free up mental bandwidth for wet-lab creativity."),
    ("2. Experimental Scale",
     "Design bolder, higher-throughput screens.\n\n• Don't limit your assay to 2 strains because data analysis will be painful.\n• Screen 50 strains or combine multimodal assays (imaging + proteomics + genomics).\n• The analysis pipeline scales effortlessly."),
    ("3. Rigor & Publication Quality",
     "Publication-ready figures from Day 1.\n\n• Every figure in your paper backed by reproducible code.\n• Colors, typography, and statistics remain perfectly consistent across all figures.\n• Seamless paper revisions and reviewer responses.")
]

for i, (title, body) in enumerate(pillars):
    x = Inches(0.8 + i * 4.04)
    add_card(s16, x, Inches(2.0), p_w, p_h)
    tb = s16.shapes.add_textbox(x + Inches(0.2), Inches(2.2), p_w - Inches(0.4), p_h - Inches(0.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.bold = True
    p.font.size = Pt(17)
    p.font.color.rgb = COLOR_ACCENT if i == 0 else COLOR_TEXT_PRIMARY
    p.space_after = Pt(14)

    p2 = tf.add_paragraph()
    p2.text = body
    p2.font.size = Pt(12)
    p2.font.color.rgb = COLOR_TEXT_SECONDARY


# ==============================================================================
# SLIDE 17: Getting Started (Actionable Next Steps)
# ==============================================================================
s17 = prs.slides.add_slide(blank_slide_layout)
add_header(s17, "Part V: The Path Forward", "Getting Started: Low-Hanging Fruit for the Lab",
           "Three simple projects you can test this week to build confidence.")

add_card(s17, Inches(0.8), Inches(2.0), Inches(5.6), Inches(4.8))
tb = s17.shapes.add_textbox(Inches(1.1), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "3 Starter Projects to Try This Week"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_TEXT_PRIMARY
p.space_after = Pt(14)

starter_projects = [
    ("Project 1: The 'Excel Figure Makeover'", "Take an old Excel chart from a past lab meeting. Feed the raw data to Antigravity and ask it to generate a publication-ready Seaborn/Matplotlib plot."),
    ("Project 2: Machine Export Parser", "Automate parsing an annoying machine export (plate reader, qPCR, NanoDrop). Turn 30 minutes of manual formatting into a 2-second script."),
    ("Project 3: A Quick Exploration", "Ask the agent to pull a Candida gene sequence from NCBI or UniProt and check its domain architecture or homolog in S. cerevisiae.")
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

add_card(s17, Inches(6.9), Inches(2.0), Inches(5.6), Inches(4.8), bg_color=COLOR_ACCENT_BG, border_color=COLOR_ACCENT)
tb = s17.shapes.add_textbox(Inches(7.2), Inches(2.3), Inches(5.0), Inches(4.2))
tf = tb.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Summary Checklist: The Agent Habit"
p.font.bold = True
p.font.size = Pt(18)
p.font.color.rgb = COLOR_ACCENT
p.space_after = Pt(14)

checklist = [
    ("1. Isolate Raw Data:", "Put raw files in `data/` and keep them read-only."),
    ("2. Write an AGENTS.md:", "Give your agent your lab's rules, strains, and palettes."),
    ("3. Start with /plan or /grill-me:", "Align on the scientific plan before writing any code."),
    ("4. Keep Chats Focused:", "One task/figure per session to keep context clean."),
    ("5. Commit to Git:", "Treat code as an auditable scientific asset."),
    ("Questions & Lab Discussion:", "Let's explore your current bottlenecks!")
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


# ==============================================================================
# Save Presentation
# ==============================================================================
output_path = "antigravity_lab_presentation.pptx"
prs.save(output_path)
print(f"Presentation successfully saved to {output_path} (17 slides, 16:9 widescreen).")
