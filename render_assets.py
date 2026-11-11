"""
Render visual assets for Antigravity Lab Presentation:
1. paper-repro-flow.png — High-resolution faithful render of paper-repro-flow.svg
2. Figure_1_pangenome_analysis-AGY_cropped.png — Cleanly cropped AGY Figure 1 for side-by-side comparison
3. git_tree_schematic.png — Publication-grade Git Version Control Tree diagram for Slide 25 (now 26)
4. figure_1_paper_cropped.png — Original paper Figure 1 with panels c and f removed for Slide 8
"""

import math
from lxml import etree
from PIL import Image, ImageDraw, ImageFont, ImageOps

FONT_REG = "/System/Library/Fonts/Supplemental/Arial.ttf"
FONT_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
FONT_HELV = "/System/Library/Fonts/HelveticaNeue.ttc"


def get_font(size, bold=False):
    path = FONT_BOLD if bold else FONT_REG
    return ImageFont.truetype(path, int(round(size)))


def render_paper_repro_svg(svg_path="paper-repro-flow.svg", out_path="paper-repro-flow.png", scale=3.0):
    """Renders paper-repro-flow.svg at high resolution with exact coordinates and styling."""
    tree = etree.parse(svg_path)
    root = tree.getroot()

    width = float(root.attrib.get("width", 982.08))
    height = float(root.attrib.get("height", 1292.6))

    W = int(round(width * scale))
    H = int(round(height * scale))

    # Colors from paper-repro-flow.svg CSS variables
    BG = (24, 24, 27)             # #18181B
    FG = (244, 244, 245)          # #F4F4F5
    LINE = (113, 113, 122)        # #71717A
    MUTED = (161, 161, 170)       # #A1A1AA
    SURFACE = (39, 39, 42)        # #27272A
    BORDER = (63, 63, 70)         # #3F3F46
    GROUP_HDR = (34, 34, 38)      # color-mix(in srgb, #F4F4F5 5%, #18181B)
    ARROW = (210, 210, 214)       # color-mix(in srgb, #F4F4F5 85%, #18181B)

    img = Image.new("RGBA", (W, H), BG + (255,))
    draw = ImageDraw.Draw(img)

    ns = {"svg": "http://www.w3.org/2000/svg"}

    # 1. Draw subgraphs
    for g in root.findall(".//svg:g[@class='subgraph']", ns):
        rects = g.findall("svg:rect", ns)
        for idx, r in enumerate(rects):
            rx = float(r.attrib["x"]) * scale
            ry = float(r.attrib["y"]) * scale
            rw = float(r.attrib["width"]) * scale
            rh = float(r.attrib["height"]) * scale
            fill_col = BG if idx == 0 else GROUP_HDR
            draw.rectangle([rx, ry, rx + rw, ry + rh], fill=fill_col, outline=BORDER, width=max(1, int(round(1.2 * scale))))

        txt = g.find("svg:text", ns)
        if txt is not None and txt.text:
            tx = float(txt.attrib["x"]) * scale
            ty = float(txt.attrib["y"]) * scale
            fsize = float(txt.attrib.get("font-size", 12)) * scale
            font = get_font(fsize, bold=True)
            draw.text((tx, ty), txt.text, fill=MUTED, font=font, anchor="lm")

    # 2. Draw edges (polylines + arrowheads)
    for poly in root.findall(".//svg:polyline[@class='edge']", ns):
        pts_raw = poly.attrib["points"].strip().split()
        pts = []
        for p in pts_raw:
            px, py = p.split(",")
            pts.append((float(px) * scale, float(py) * scale))

        if len(pts) >= 2:
            draw.line(pts, fill=LINE, width=max(1, int(round(1.5 * scale))), joint="curve")
            # Draw arrowhead at the end
            if poly.attrib.get("data-arrow-end") == "true":
                x1, y1 = pts[-2]
                x2, y2 = pts[-1]
                angle = math.atan2(y2 - y1, x2 - x1)
                arr_len = 8.5 * scale
                arr_w = 3.2 * scale
                bx = x2 - arr_len * math.cos(angle)
                by = y2 - arr_len * math.sin(angle)
                left_pt = (bx - arr_w * math.sin(angle), by + arr_w * math.cos(angle))
                right_pt = (bx + arr_w * math.sin(angle), by - arr_w * math.cos(angle))
                draw.polygon([(x2, y2), left_pt, right_pt], fill=ARROW)

    # 3. Draw nodes
    for g in root.findall(".//svg:g[@class='node']", ns):
        r = g.find("svg:rect", ns)
        if r is not None:
            rx = float(r.attrib["x"]) * scale
            ry = float(r.attrib["y"]) * scale
            rw = float(r.attrib["width"]) * scale
            rh = float(r.attrib["height"]) * scale
            draw.rectangle([rx, ry, rx + rw, ry + rh], fill=SURFACE, outline=BORDER, width=max(1, int(round(1.2 * scale))))

        txt = g.find("svg:text", ns)
        if txt is not None and txt.text:
            tx = float(txt.attrib["x"]) * scale
            ty = float(txt.attrib["y"]) * scale
            fsize = float(txt.attrib.get("font-size", 13)) * scale
            font = get_font(fsize, bold=False)
            draw.text((tx, ty), txt.text, fill=FG, font=font, anchor="mm")

    img.convert("RGB").save(out_path, "PNG", optimize=True)
    print(f"Saved {out_path} ({W}x{H})")


def crop_agy_figure(in_path="Figure_1_pangenome_analysis-AGY.png", out_path="Figure_1_pangenome_analysis-AGY_cropped.png"):
    """Crops excessive outer whitespace from Figure_1_pangenome_analysis-AGY.png for side-by-side comparison."""
    img = Image.open(in_path).convert("RGB")
    inv = ImageOps.invert(img)
    bbox = inv.getbbox()
    if bbox:
        pad = 60
        left = max(0, bbox[0] - pad)
        top = max(0, bbox[1] - pad)
        right = min(img.width, bbox[2] + pad)
        bottom = min(img.height, bbox[3] + pad)
        cropped = img.crop((left, top, right, bottom))
        cropped.save(out_path, "PNG", optimize=True)
        print(f"Saved cropped AGY figure to {out_path} ({cropped.size})")


def crop_paper_figure(in_path="figure_1_paper.png", out_path="figure_1_paper_cropped.png"):
    """Removes panels c and f from figure_1_paper.png and crops trailing whitespace."""
    img = Image.open(in_path).convert("RGB")
    draw = ImageDraw.Draw(img)
    # Blank out panel c (top-right: x >= 885, y < 346) and panel f (bottom-right: x >= 640, y >= 346)
    draw.rectangle([885, 0, img.width, 345], fill=(255, 255, 255))
    draw.rectangle([640, 346, img.width, img.height], fill=(255, 255, 255))
    # Crop right edge just past panel b (ends at x=867) while preserving full height and left margin
    cropped = img.crop((0, 0, 882, img.height))
    cropped.save(out_path, "PNG", optimize=True)
    print(f"Saved cropped paper figure to {out_path} ({cropped.size})")


def cubic_bezier(p0, p1, p2, p3, steps=80):
    pts = []
    for i in range(steps + 1):
        t = i / steps
        mt = 1 - t
        x = (mt**3) * p0[0] + 3 * (mt**2) * t * p1[0] + 3 * mt * (t**2) * p2[0] + (t**3) * p3[0]
        y = (mt**3) * p0[1] + 3 * (mt**2) * t * p1[1] + 3 * mt * (t**2) * p2[1] + (t**3) * p3[1]
        pts.append((x, y))
    return pts


def draw_arrow_head(draw, p_from, p_to, color, size=18):
    angle = math.atan2(p_to[1] - p_from[1], p_to[0] - p_from[0])
    w = size * 0.48
    bx = p_to[0] - size * math.cos(angle)
    by = p_to[1] - size * math.sin(angle)
    left_pt = (bx - w * math.sin(angle), by + w * math.cos(angle))
    right_pt = (bx + w * math.sin(angle), by - w * math.cos(angle))
    draw.polygon([p_to, left_pt, right_pt], fill=color)


def render_git_tree_schematic(out_path="git_tree_schematic.png"):
    """
    Generates a clean, high-resolution (2x supersampled) schematic diagram
    illustrating Git version control (commits, branches, merges, and revert)
    for Slide 25/26 right card.
    Target card inner area aspect ratio: ~5.2 in x 4.0 in (1560 x 1200 px at 300 DPI, drawn at 3120 x 2400).
    """
    S = 2  # Supersampling factor
    W, H = 1560 * S, 1200 * S
    img = Image.new("RGBA", (W, H), (248, 250, 252, 255))  # Matches COLOR_CARD_BG (#F8FAFC)
    draw = ImageDraw.Draw(img)

    # Palette matching presentation
    C_WHITE = (255, 255, 255)
    C_DARK = (17, 24, 39)          # #111827
    C_SEC = (75, 85, 99)           # #4B5563
    C_MUTED = (107, 114, 128)      # #6B7280
    C_BORDER = (203, 213, 225)     # #CBD5E1
    C_CARD_BORDER = (226, 232, 240)# #E2E8F0

    C_BLUE = (37, 99, 235)         # #2563EB (main branch)
    C_BLUE_LIGHT = (219, 234, 254) # #DBEAFE
    C_GREEN = (16, 185, 129)       # #10B981 (feature branch)
    C_GREEN_LIGHT = (209, 250, 229)# #D1FAE5
    C_PURPLE = (147, 51, 234)      # #9333EA (merged commit)
    C_PURPLE_LIGHT = (243, 232, 255)# #F3E8FF
    C_AMBER = (245, 158, 11)       # #F59E0B (revert / undo highlight)
    C_AMBER_LIGHT = (254, 243, 199)# #FEF3C7

    # Top Diagram Container Card (white inner panel)
    draw.rounded_rectangle(
        [20 * S, 15 * S, (1560 - 20) * S, 725 * S],
        radius=22 * S,
        fill=C_WHITE,
        outline=C_CARD_BORDER,
        width=2 * S
    )

    # Subtle time arrow along top of diagram panel
    time_y = 62 * S
    draw.line([(310 * S, time_y), (1440 * S, time_y)], fill=(203, 213, 225), width=3 * S)
    draw_arrow_head(draw, (1420 * S, time_y), (1450 * S, time_y), (148, 163, 184), size=14 * S)
    draw.text((875 * S, (time_y - 18 * S)), "PROJECT TIMELINE  (Each node is a saved snapshot / Commit)",
              fill=C_MUTED, font=get_font(22 * S, bold=True), anchor="mm")

    # Y positions for lanes
    y_main = 235 * S
    y_feat = 505 * S

    # Subtle lane guide lines (dashed effect)
    for x_dash in range(300, 1450, 24):
        draw.line([(x_dash * S, y_main), ((x_dash + 12) * S, y_main)], fill=(241, 245, 249), width=2 * S)
        draw.line([(x_dash * S, y_feat), ((x_dash + 12) * S, y_feat)], fill=(241, 245, 249), width=2 * S)

    # Branch Label Badges on the left
    # 1. main badge
    draw.rounded_rectangle([55 * S, (y_main - 30 * S), 255 * S, (y_main + 30 * S)],
                           radius=14 * S, fill=C_BLUE, outline=None)
    draw.text((155 * S, y_main), "main branch", fill=C_WHITE, font=get_font(26 * S, bold=True), anchor="mm")
    draw.text((155 * S, y_main + 52 * S), "Production Pipeline", fill=C_SEC, font=get_font(21 * S, bold=False), anchor="mm")

    # 2. feature/palette badge
    draw.rounded_rectangle([55 * S, (y_feat - 30 * S), 275 * S, (y_feat + 30 * S)],
                           radius=14 * S, fill=C_GREEN, outline=None)
    draw.text((165 * S, y_feat), "feature/palette", fill=C_WHITE, font=get_font(25 * S, bold=True), anchor="mm")
    draw.text((165 * S, y_feat + 52 * S), "Safe Sandbox Branch", fill=C_SEC, font=get_font(21 * S, bold=False), anchor="mm")

    # X coordinates for commits
    x_c1 = 380 * S
    x_c2 = 620 * S
    x_c3 = 960 * S
    x_c6 = 1350 * S

    x_c4 = 820 * S
    x_c5 = 1130 * S

    # --- Draw Connecting Curves & Lines ---
    # 1. Branch out curve: C2 -> C4
    curve_out = cubic_bezier(
        (x_c2, y_main),
        (x_c2 + 110 * S, y_main),
        (x_c4 - 110 * S, y_feat),
        (x_c4, y_feat),
        steps=100
    )
    draw.line(curve_out, fill=C_GREEN, width=9 * S, joint="curve")

    # Pill label on Branch Out curve
    bx_lbl, by_lbl = (x_c2 + x_c4) // 2 - 15 * S, (y_main + y_feat) // 2
    draw.rounded_rectangle([bx_lbl - 68 * S, by_lbl - 20 * S, bx_lbl + 68 * S, by_lbl + 20 * S],
                           radius=10 * S, fill=C_GREEN_LIGHT, outline=C_GREEN, width=2 * S)
    draw.text((bx_lbl, by_lbl), "git branch", fill=(6, 95, 70), font=get_font(20 * S, bold=True), anchor="mm")

    # 2. Feature branch line: C4 -> C5
    draw.line([(x_c4, y_feat), (x_c5, y_feat)], fill=C_GREEN, width=9 * S)

    # 3. Merge curve: C5 -> C6
    curve_merge = cubic_bezier(
        (x_c5, y_feat),
        (x_c5 + 120 * S, y_feat),
        (x_c6 - 120 * S, y_main),
        (x_c6, y_main),
        steps=100
    )
    draw.line(curve_merge, fill=C_PURPLE, width=9 * S, joint="curve")

    # Pill label on Merge curve
    mx_lbl, my_lbl = (x_c5 + x_c6) // 2 + 15 * S, (y_main + y_feat) // 2
    draw.rounded_rectangle([mx_lbl - 68 * S, my_lbl - 20 * S, mx_lbl + 68 * S, my_lbl + 20 * S],
                           radius=10 * S, fill=C_PURPLE_LIGHT, outline=C_PURPLE, width=2 * S)
    draw.text((mx_lbl, my_lbl), "git merge", fill=(107, 33, 168), font=get_font(20 * S, bold=True), anchor="mm")

    # 4. Main branch line: C1 -> C2 -> C3 -> C6
    draw.line([(x_c1, y_main), (x_c2, y_main)], fill=C_BLUE, width=10 * S)
    draw.line([(x_c2, y_main), (x_c3, y_main)], fill=C_BLUE, width=10 * S)
    draw.line([(x_c3, y_main), (x_c6, y_main)], fill=C_BLUE, width=10 * S)

    # 5. Revert / Rollback curved dashed arc above C5 -> C4 to illustrate instant undo
    revert_pts = cubic_bezier(
        (x_c5 - 15 * S, y_feat - 42 * S),
        (x_c5 - 70 * S, y_feat - 115 * S),
        (x_c4 + 70 * S, y_feat - 115 * S),
        (x_c4 + 22 * S, y_feat - 42 * S),
        steps=60
    )
    for idx_r in range(0, len(revert_pts) - 2, 4):
        draw.line([revert_pts[idx_r], revert_pts[idx_r + 2]], fill=C_AMBER, width=5 * S)
    draw_arrow_head(draw, revert_pts[-6], revert_pts[-1], C_AMBER, size=16 * S)
    rx_mid = (x_c4 + x_c5) // 2
    ry_mid = y_feat - 108 * S
    draw.rounded_rectangle([rx_mid - 132 * S, ry_mid - 18 * S, rx_mid + 132 * S, ry_mid + 18 * S],
                           radius=9 * S, fill=C_AMBER_LIGHT, outline=C_AMBER, width=2 * S)
    draw.text((rx_mid, ry_mid), "Instant Rollback (git revert)", fill=(180, 83, 9), font=get_font(17 * S, bold=True), anchor="mm")

    # --- Draw Commit Nodes & Labels ---
    commits = [
        (x_c1, y_main, "C1", "C1: Setup", "Raw data & scripts", C_BLUE, C_BLUE_LIGHT, True),
        (x_c2, y_main, "C2", "C2: Baseline", "Working Fig 1–2", C_BLUE, C_BLUE_LIGHT, True),
        (x_c3, y_main, "C3", "C3: Stats Fix", "ANOVA annotations", C_BLUE, C_BLUE_LIGHT, True),
        (x_c6, y_main, "C6", "C6: Merged", "Publication Ready!", C_PURPLE, C_PURPLE_LIGHT, True),
        (x_c4, y_feat, "C4", "C4: New Colors", "Google Sheets sync", C_GREEN, C_GREEN_LIGHT, False),
        (x_c5, y_feat, "C5", "C5: Polish", "300-DPI layout", C_GREEN, C_GREEN_LIGHT, False),
    ]

    r_outer = 44 * S
    r_inner = 34 * S

    for cx, cy, cid, title, desc, col, ring_col, above in commits:
        # Outer halo
        draw.ellipse([cx - r_outer, cy - r_outer, cx + r_outer, cy + r_outer], fill=ring_col, outline=col, width=3 * S)
        # Inner solid circle
        draw.ellipse([cx - r_inner, cy - r_inner, cx + r_inner, cy + r_inner], fill=col)
        # Commit ID text
        draw.text((cx, cy), cid, fill=C_WHITE, font=get_font(25 * S, bold=True), anchor="mm")

        # Callout text above or below node
        if above:
            ty1 = cy - 92 * S
            ty2 = cy - 63 * S
        else:
            ty1 = cy + 66 * S
            ty2 = cy + 95 * S

        draw.text((cx, ty1), title, fill=C_DARK, font=get_font(24 * S, bold=True), anchor="mm")
        draw.text((cx, ty2), desc, fill=C_SEC, font=get_font(20 * S, bold=False), anchor="mm")

    # --- Bottom 2x2 Legend Grid ---
    legend_cards = [
        (20, 750, 765, 955, "commit", "Commit (Snapshot)", "An immutable checkpoint of all project files at a point in time. Jump back anytime.", C_BLUE, C_BLUE_LIGHT),
        (795, 750, 1540, 955, "branch", "Branch (Safe Sandbox)", "An isolated workspace to test new code or palettes without risking working scripts.", C_GREEN, C_GREEN_LIGHT),
        (20, 980, 765, 1185, "merge", "Merge (Integrate)", "Cleanly incorporates tested, verified improvements back into the main pipeline.", C_PURPLE, C_PURPLE_LIGHT),
        (795, 980, 1540, 1185, "revert", "Revert (Universal Undo)", "If an AI agent experiment breaks a plot, roll back to the previous commit in 1 second.", C_AMBER, C_AMBER_LIGHT),
    ]

    for x1, y1, x2, y2, icon_type, l_title, l_desc, accent_col, bg_tint in legend_cards:
        # Card background
        draw.rounded_rectangle([x1 * S, y1 * S, x2 * S, y2 * S],
                               radius=18 * S, fill=C_WHITE, outline=C_CARD_BORDER, width=2 * S)
        # Left color accent bar
        draw.rounded_rectangle([x1 * S, y1 * S, (x1 + 12) * S, y2 * S],
                               radius=6 * S, fill=accent_col)

        # Draw crisp vector icon badge
        ic_x, ic_y = (x1 + 46) * S, (y1 + 42) * S
        draw.ellipse([ic_x - 18 * S, ic_y - 18 * S, ic_x + 18 * S, ic_y + 18 * S], fill=bg_tint, outline=accent_col, width=2 * S)
        if icon_type == "commit":
            draw.ellipse([ic_x - 8 * S, ic_y - 8 * S, ic_x + 8 * S, ic_y + 8 * S], fill=accent_col)
        elif icon_type == "branch":
            draw.line([(ic_x - 8 * S, ic_y + 7 * S), (ic_x + 7 * S, ic_y - 7 * S)], fill=accent_col, width=4 * S)
            draw_arrow_head(draw, (ic_x, ic_y), (ic_x + 9 * S, ic_y - 9 * S), accent_col, size=10 * S)
        elif icon_type == "merge":
            draw.line([(ic_x - 8 * S, ic_y - 7 * S), (ic_x + 7 * S, ic_y + 7 * S)], fill=accent_col, width=4 * S)
            draw_arrow_head(draw, (ic_x, ic_y), (ic_x + 9 * S, ic_y + 9 * S), accent_col, size=10 * S)
        elif icon_type == "revert":
            col_dk = (180, 83, 9)
            draw.arc([ic_x - 10 * S, ic_y - 10 * S, ic_x + 10 * S, ic_y + 10 * S], start=200, end=40, fill=col_dk, width=3 * S)
            draw_arrow_head(draw, (ic_x - 4 * S, ic_y - 10 * S), (ic_x - 11 * S, ic_y - 4 * S), col_dk, size=9 * S)

        # Title
        draw.text(((x1 + 74) * S, (y1 + 42) * S), l_title,
                  fill=accent_col if accent_col != C_AMBER else (180, 83, 9),
                  font=get_font(26 * S, bold=True), anchor="lm")
        # Wrapped description (2 lines)
        words = l_desc.split()
        lines = []
        curr = ""
        font_desc = get_font(22 * S, bold=False)
        max_w = (x2 - x1 - 50) * S
        for w in words:
            test = (curr + " " + w).strip()
            if draw.textlength(test, font=font_desc) <= max_w:
                curr = test
            else:
                lines.append(curr)
                curr = w
        if curr:
            lines.append(curr)

        for idx_l, line_str in enumerate(lines):
            draw.text(((x1 + 32) * S, (y1 + 92 + idx_l * 34) * S), line_str,
                      fill=C_SEC, font=font_desc, anchor="lm")

    # Downsample with Lanczos for crisp anti-aliasing
    final_img = img.resize((1560, 1200), Image.Resampling.LANCZOS).convert("RGB")
    final_img.save(out_path, "PNG", optimize=True)
    print(f"Saved {out_path} ({final_img.size})")


if __name__ == "__main__":
    render_paper_repro_svg()
    crop_agy_figure()
    crop_paper_figure()
    render_git_tree_schematic()
