"""
BREAKING CHANGES — deck builder.

Builds new/BreakingChanges.pptx from the slide definitions in this file.
- Every word on a slide is real PowerPoint text (RTL Persian, LTR code).
- Speaker notes come from new/script.md: the k-th «اسلاید»/«کلیک» stage direction
  belongs to the k-th slide of the pptx.
- Background images are optional. Drop them into new/slides/ with the file names
  listed in new/image_prompts.md and run the build again; a slide whose image is
  missing gets a designed dark background instead.

Usage:
    python new/build_deck.py                 # writes new/BreakingChanges.pptx
    python new/build_deck.py --check         # only checks script.md against the slide list
    python new/build_deck.py --html DIR      # also writes an HTML preview of every slide to DIR
"""

import os
import re
import sys
import html
import argparse

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SLIDES_DIR = os.path.join(BASE_DIR, "slides")
SCRIPT_MD = os.path.join(BASE_DIR, "script.md")
OUTPUT = os.path.join(BASE_DIR, "BreakingChanges.pptx")
DEMO_VIDEO = "isr_demo.mp4"

# Slide canvas in pixels at 96 dpi: 1280 x 720 px = 13.333 x 7.5 in.
W, H = 1280, 720
EMU_PER_PX = 9525

# ---------------------------------------------------------------- fonts & colours
# The four weights of IRANYekanX are four separate font families on Windows.
# To present on a laptop with another Persian font, change these five names.
F_BLACK = "IRANYekanX Black"
F_BOLD = "IRANYekanX Bold"
F_REG = "IRANYekanX"
F_LIGHT = "IRANYekanX Light"
F_MONO = "Consolas"

BG = "0B0C10"
PANEL = "14161D"
PANEL2 = "1C1F29"
BORDER = "2E3340"
TEXT = "F5F2EA"
MUTED = "A6A9B4"
DIM = "6E7280"
GOLD = "F2C14E"
GREEN = "5BD68A"
RED = "FF6B6B"
ORANGE = "FFA24A"
BLUE = "7CC7FF"
BLACK = "000000"
WHITE = "FFFFFF"

COLORS = {"gold": GOLD, "green": GREEN, "red": RED, "orange": ORANGE, "blue": BLUE,
          "muted": MUTED, "dim": DIM, "white": TEXT}
FONTS = {"mono": F_MONO, "black": F_BLACK, "bold": F_BOLD, "reg": F_REG, "light": F_LIGHT}

# Natural line height of each family (ascent + descent, in em). Used by the HTML preview only.
LINE_HEIGHT = {F_MONO: 1.171}
DEFAULT_LINE_HEIGHT = 1.3

MIN_FONT_PT = 20


def fa(n):
    return str(n).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))


def find_image(name):
    """Path of a background image in new/slides/, whatever its extension, or None."""
    stem = os.path.splitext(name)[0]
    for ext in (".jpg", ".jpeg", ".png"):
        path = os.path.join(SLIDES_DIR, stem + ext)
        if os.path.exists(path):
            return path
    return None


# ---------------------------------------------------------------- scene model
SPAN = re.compile(r"\{([a-z0-9,]+)\|([^{}]*)\}")
WANTED_IMAGES = set()   # every background image the slides ask for, found or not


def parse_runs(line, base):
    """'plain {gold|coloured} {mono,red|code}' -> [(text, style), ...]"""
    runs, pos = [], 0
    for m in SPAN.finditer(line):
        if m.start() > pos:
            runs.append((line[pos:m.start()], base))
        style = dict(base)
        for tok in m.group(1).split(","):
            if tok in COLORS:
                style["color"] = COLORS[tok]
            elif tok in FONTS:
                style["font"] = FONTS[tok]
            elif tok == "b":
                style["bold"] = True
        runs.append((m.group(2), style))
        pos = m.end()
    if pos < len(line):
        runs.append((line[pos:], base))
    return runs or [(" ", base)]


class Slide:
    def __init__(self, key, bg=BG, fade=True):
        self.key = key          # e.g. "14.2" = slide 14, second click
        self.bg = bg
        self.fade = fade
        self.items = []
        self.video = None
        self.notes = ""

    def image(self, name, x=0, y=0, w=W, h=H):
        WANTED_IMAGES.add(name)
        path = find_image(name)
        if path:
            self.items.append(("image", path, x, y, w, h))
        return bool(path)

    def rect(self, x, y, w, h, fill=None, alpha=100, line=None, lw=1.5, r=0, rot=0, line_alpha=100):
        self.items.append(("rect", x, y, w, h, fill, alpha, line, lw, r, rot, line_alpha))

    def oval(self, x, y, w, h, fill=None, alpha=100, line=None, lw=1.5):
        self.items.append(("oval", x, y, w, h, fill, alpha, line, lw))

    def grad(self, x, y, w, h, stops, vertical=False):
        """stops: [(position %, hex colour, alpha %)], left-to-right or top-to-bottom."""
        self.items.append(("grad", x, y, w, h, stops, vertical))

    def line(self, x1, y1, x2, y2, color, lw=2, alpha=100, dash=False):
        self.items.append(("line", x1, y1, x2, y2, color, lw, alpha, dash))

    def poly(self, pts, color, lw=3, alpha=100):
        self.items.append(("poly", pts, color, lw, alpha))

    def play(self, x, y, w, h, fill):
        self.items.append(("play", x, y, w, h, fill))

    def text(self, x, y, w, h, content, size=28, color=TEXT, font=F_REG, align="r", anchor="m",
             rtl=True, spacing=1.0, bold=False, rot=0):
        if size < MIN_FONT_PT:
            raise ValueError(f"slide {self.key}: text below {MIN_FONT_PT}pt: {content!r}")
        base = {"font": font, "size": size, "color": color, "bold": bold}
        paras = [parse_runs(ln, base) for ln in content.split("\n")]
        self.items.append(("text", x, y, w, h, paras, align, anchor, rtl, spacing, rot))


# ---------------------------------------------------------------- components
def mono_w(label, size):
    return len(label) * size * 4 / 3 * 0.55


def pill(s, x, y, label, color, filled=False, size=20, w=None, h=44, font=F_MONO, rtl=False):
    if w is None:
        w = int(mono_w(label, size) + 36)
    s.rect(x, y, w, h, fill=color if filled else PANEL, alpha=16 if filled else 92, line=color, lw=1.5, r=10)
    s.text(x, y, w, h, label, size=size, color=color, font=font, bold=(font == F_MONO), align="c", anchor="m", rtl=rtl)
    return w


def rail(s, current):
    """Five dots at the bottom: v1 .. v4 and the unreleased one. Oldest on the left, like a git log."""
    gap, dot = 44, 10
    x0 = W / 2 - 2 * gap
    y = 694
    s.line(x0, y, x0 + 4 * gap, y, WHITE, lw=1.5, alpha=14)
    for i in range(5):
        cx = x0 + i * gap
        if i == current:
            s.oval(cx - 8, y - 8, 16, 16, fill=GOLD)
        elif i < current:
            s.oval(cx - dot / 2, y - dot / 2, dot, dot, fill=GOLD, alpha=45)
        else:
            s.oval(cx - dot / 2, y - dot / 2, dot, dot, fill=WHITE, alpha=20)


def chrome(s, ver=None, tag=None, tag_color=RED, year=None, at=None):
    """Version pill and tag top-left, year top-right, progress rail at the bottom."""
    x = 64
    if ver:
        x += pill(s, x, 34, ver, GOLD) + 12
    if tag:
        pill(s, x, 34, tag, tag_color, filled=True)
    if year:
        s.text(W - 64 - 360, 28, 360, 56, year, size=30, color=GOLD, font=F_BLACK, align="r")
    if at is not None:
        rail(s, at)


SEAM = [(170, -10), (196, 70), (160, 150), (214, 236), (182, 300), (250, 372), (214, 452), (286, 520), (262, 600), (330, 730)]
SEAM_BRANCHES = [
    [(214, 236), (290, 270), (330, 250), (392, 300)],
    [(250, 372), (190, 420), (150, 480)],
    [(286, 520), (360, 540), (410, 600)],
]


def seam(s, dx=0, alpha=75):
    """A gold-filled crack, drawn as shapes: the motif of the deck when no photo is present."""
    for path in [SEAM] + SEAM_BRANCHES:
        pts = [(x + dx, y) for x, y in path]
        s.poly(pts, GOLD, lw=9, alpha=int(alpha * 0.16))
        s.poly(pts, GOLD, lw=2.5 if path is SEAM else 2, alpha=alpha)


def backdrop(s, name, mode="hero"):
    """Full-bleed photo plus a scrim that keeps the text side dark. Falls back to a drawn background."""
    if name and s.image(name):
        if mode == "hero":        # subject on the left, text on the right
            s.grad(0, 0, W, H, [(0, BG, 8), (30, BG, 50), (50, BG, 88), (100, BG, 96)])
        elif mode == "dim":       # photo is only atmosphere behind a full slide of content
            s.rect(0, 0, W, H, fill=BG, alpha=85)
        elif mode == "top":       # text in the upper part, subject at the bottom
            s.grad(0, 0, W, H, [(0, BG, 94), (55, BG, 82), (100, BG, 30)], vertical=True)
        elif mode == "bottom":    # text at the bottom
            s.grad(0, 0, W, H, [(0, BG, 30), (50, BG, 62), (100, BG, 96)], vertical=True)
        elif mode == "center":    # text centred over a full-frame pattern
            s.rect(0, 0, W, H, fill=BG, alpha=66)
        return True
    s.grad(0, 0, W, H, [(0, "0B0C10", 100), (100, "14161E", 100)], vertical=True)
    if mode in ("hero", "bottom"):
        seam(s)
    elif mode == "center":
        seam(s, dx=-60, alpha=60)
        s.grad(0, 0, W, H, [(0, BG, 0), (22, BG, 86), (80, BG, 86), (100, BG, 0)], vertical=True)
    return False


def row(s, x, y, w, h, kind, content, size=28, rtl=True, color=None, font=None):
    """One line of a diff: '+' added, '-' removed, '~' changed/deprecated."""
    col = {"+": GREEN, "-": RED, "~": ORANGE}[kind]
    mark = {"+": "+", "-": "−", "~": "~"}[kind]
    s.rect(x, y, w, h, fill=col, alpha=12)
    if rtl:
        s.rect(x + w - 6, y, 6, h, fill=col, alpha=95)
        s.text(x + w - 60, y, 44, h, mark, size=size, color=col, font=F_MONO, bold=True, align="c", rtl=False)
        s.text(x + 20, y, w - 92, h, content, size=size, color=color or TEXT, font=font or F_REG, align="r")
    else:
        s.rect(x, y, 6, h, fill=col, alpha=95)
        s.text(x + 16, y, 44, h, mark, size=size, color=col, font=F_MONO, bold=True, align="c", rtl=False)
        s.text(x + 68, y, w - 88, h, content, size=size, color=color or TEXT, font=font or F_MONO, align="l", rtl=False)


def bubble(s, x, y, w, h, who, quote, status="open", size=26, note=None):
    """A review comment: who said it, the sentence, and whether it is still open or how it was resolved."""
    s.rect(x, y, w, h, fill=PANEL, alpha=95, line=BORDER, lw=1.5, r=16)
    chip_w = 112
    row_h = h if note is None else h * 0.52
    s.rect(x + w - 18 - chip_w, y + row_h / 2 - 20, chip_w, 40, fill=PANEL2, line=BORDER, lw=1, r=20)
    s.text(x + w - 18 - chip_w, y + row_h / 2 - 20, chip_w, 40, who, size=20, color=GOLD, font=F_BOLD, align="c")
    left = x + 22
    if status == "open":
        left += pill(s, x + 18, y + row_h / 2 - 20, "open", GOLD, h=40) + 14
    elif status == "true":
        left += pill(s, x + 18, y + row_h / 2 - 20, "درست بود", GOLD, filled=True, w=140, h=40, font=F_BOLD, rtl=True) + 14
    elif status == "false":
        left += pill(s, x + 18, y + row_h / 2 - 20, "غلط بود", GREEN, filled=True, w=140, h=40, font=F_BOLD, rtl=True) + 14
    elif status == "resolved":
        left += pill(s, x + 18, y + row_h / 2 - 20, "resolved", GREEN, filled=True, h=40) + 14
    right = x + w - 18 - chip_w - 18
    s.text(left, y + 4, right - left, row_h - 4, "«" + quote + "»", size=size, color=TEXT, font=F_BOLD, align="r")
    if note:
        s.line(x + 24, y + row_h, x + w - 24, y + row_h, BORDER, lw=1)
        s.text(x + 28, y + row_h, w - 56, h - row_h, note, size=24, color=MUTED, font=F_REG, align="r")


def card(s, x, y, w, rows, row_h=112, label_w=220):
    """Incident card. rows: (label, label colour, value, value size, value colour)."""
    s.rect(x, y, w, row_h * len(rows), fill=PANEL, alpha=93, line=BORDER, lw=1.5, r=18)
    for i, (label, col, value, vsize, vcol) in enumerate(rows):
        ry = y + i * row_h
        if i:
            s.line(x + 24, ry, x + w - 24, ry, BORDER, lw=1.5)
        latin = label.isascii()
        s.text(x + w - 30 - label_w, ry, label_w, row_h, label, size=24, color=col,
               font=F_MONO if latin else F_BOLD, bold=latin, align="r", rtl=not latin)
        s.text(x + 34, ry, w - label_w - 90, row_h, value, size=vsize, color=vcol, font=F_BOLD, align="r")


def window(s, x, y, w, h, title, badge=None, badge_color=GOLD):
    """A code / terminal window frame."""
    s.rect(x, y, w, h, fill="0E1016", alpha=96, line=BORDER, lw=1.5, r=16)
    for i, col in enumerate(("FF5F57", "FEBC2E", "28C840")):
        s.oval(x + 24 + i * 24, y + 21, 14, 14, fill=col)
    s.text(x + 110, y, 420, 56, title, size=20, color=MUTED, font=F_MONO, align="l", rtl=False)
    s.line(x + 1, y + 56, x + w - 1, y + 56, BORDER, lw=1.5)
    if badge:
        bw = int(mono_w(badge, 20) + 36)
        pill(s, x + w - 16 - bw, y + 7, badge, badge_color, h=42)


def code(s, x, y, w, lines, size=30, pitch=None, tint=None):
    """Code lines, left-to-right. tint: {line index: colour} paints a diff-style band behind that line."""
    pitch = pitch or size * 4 / 3 * 1.3
    for i, ln in enumerate(lines):
        if tint and i in tint:
            s.rect(x - 16, y + i * pitch, w + 32, pitch, fill=tint[i], alpha=14)
        s.text(x, y + i * pitch, w, pitch, ln, size=size, color=TEXT, font=F_MONO, align="l", rtl=False)


LOOP = ["{gold|for} ip {gold|in} {blue|$SERVERS}; {gold|do}",
        "  ssh {blue|$ip} {green|\"}{blue|$CMD}{green|\"}",
        "{gold|done}"]
LOOP_SHUTDOWN = [LOOP[0], "  ssh {blue|$ip} {red|\"shutdown -h now\"}", LOOP[2]]


# ---------------------------------------------------------------- the slides
def s01(step):
    s = Slide("1" if step == 0 else "1.1", fade=(step == 0))
    s.grad(0, 0, W, H, [(0, "0B0C10", 100), (100, "14161E", 100)], vertical=True)
    window(s, 140, 56, 1000, 356, "deploy.sh", "Review requested")
    code(s, 196, 146, 900, LOOP, size=38, pitch=78)
    pill(s, 375, 444, "Approve", GREEN, filled=True, size=22, w=210, h=58)
    pill(s, 605, 444, "Request changes", RED, filled=True, size=22, w=300, h=58)
    if step == 0:
        s.text(140, 536, 1000, 110, "این سه خط را approve می‌کنید؟", size=36, font=F_BOLD, align="c")
    else:
        s.text(40, 536, 1200, 110, "همین سه خط، یک روز {red|۷۰ سرورِ} یک بانک را خاموش کرد.", size=34, font=F_BLACK, align="c")
    return s


def s02(step):
    s = Slide("2")
    backdrop(s, "01_cover.jpg", "hero")
    s.text(420, 138, 796, 40, "TECHTALK  ·  CHANGELOG.md", size=20, color=GOLD, font=F_MONO, bold=True, align="r", rtl=False)
    s.text(380, 182, 836, 282, "BREAKING\nCHANGES", size=88, font=F_BLACK, align="r", rtl=False, spacing=0.86)
    s.rect(W - 64 - 132, 478, 132, 5, fill=GOLD)
    s.text(380, 500, 836, 60, "CHANGELOG یک مهندس؛ ۱۶ سال روی پروداکشن", size=28, align="r")
    s.text(380, 566, 836, 52, "علی امانی", size=28, color=GOLD, font=F_BOLD, align="r")
    return s


def s03(step):
    s = Slide("3")
    backdrop(s, None, "plain")
    window(s, 64, 84, 580, 552, "CHANGELOG.md")
    code(s, 104, 166, 500, ["{dim|#} {white,b|CHANGELOG}"], size=32, pitch=70)
    code(s, 104, 258, 500, [
        "{dim|## [5.0.0]} {orange|- Unreleased}",
        "{gold|## [4.0.0]} {muted|- 1404}",
        "{gold|## [3.0.0]} {muted|- 1398}",
        "{gold|## [2.0.0]} {muted|- 1394}",
        "{gold|## [1.0.0]} {muted|- 1389}",
    ], size=28, pitch=70)
    s.text(700, 104, 516, 160, "این ارائه\nیک CHANGELOG است.", size=36, font=F_BLACK, align="r")
    s.text(700, 276, 516, 56, "نه مالِ یک نرم‌افزار؛ مالِ یک آدم.", size=28, color=MUTED, align="r")
    s.rect(W - 64 - 132, 354, 132, 5, fill=GOLD)
    s.text(700, 384, 516, 110, "چهار نسخه‌ی اصلی،\nو یکی که هنوز release نشده.", size=30, align="r")
    bw = pill(s, W - 64 - 150, 540, "blameless", GREEN, filled=True, w=150)
    s.text(700, 536, 516 - bw - 18, 52, "اسم‌ها فقط برای تشکر.", size=26, color=TEXT, align="r")
    return s


def s04(step):
    s = Slide("4" if step == 0 else "4.1", fade=(step == 0))
    backdrop(s, None, "plain")
    s.text(64, 44, 1152, 64, "شماره‌ی نسخه سه عدد دارد", size=34, font=F_BOLD, align="c")
    cards = [("2", "MAJOR", "تغییری که با قبل\nسازگار نیست"),
             ("1", "MINOR", "چیزی اضافه شد؛\nچیزی نشکست"),
             ("3", "PATCH", "یک باگ درست شد")]
    for i, (digit, name, desc) in enumerate(cards):
        x = 64 + i * 392
        hot = step == 1 and i == 0
        s.rect(x, 176, 368, 330, fill=RED if hot else PANEL, alpha=10 if hot else 94,
               line=RED if hot else BORDER, lw=3 if hot else 1.5, r=20)
        s.text(x, 182, 368, 156, digit, size=96, color=RED if hot else GOLD, font=F_MONO, bold=True, align="c", rtl=False)
        s.text(x, 330, 368, 50, name, size=30, color=RED if hot else TEXT, font=F_MONO, bold=True, align="c", rtl=False)
        s.text(x + 20, 386, 328, 104, desc, size=26, color=TEXT if hot else MUTED, align="c")
        if i < 2:
            s.text(x + 340, 182, 80, 156, ".", size=96, color=MUTED, font=F_MONO, bold=True, align="c", rtl=False)
    if step == 1:
        pill(s, 64 + (368 - 252) // 2, 120, "BREAKING CHANGE", RED, filled=True, w=252)
        s.text(64, 540, 1152, 100, "با patch و minor راحتیم. از {red|major} فرار می‌کنیم.", size=38, font=F_BLACK, align="c")
    return s


def s05(step):
    s = Slide("5")
    backdrop(s, "02_java.jpg", "hero")
    chrome(s, "v1.0.0", "stable", GREEN, "۱۳۸۹", at=0)
    s.text(520, 124, 696, 110, "دولوپرِ جاوا", size=64, font=F_BLACK, align="r")
    s.text(520, 252, 696, 40, "### Added", size=22, color=GREEN, font=F_MONO, bold=True, align="r", rtl=False)
    row(s, 520, 300, 696, 66, "+", "پنج سال، فقط کد", size=30)
    row(s, 520, 376, 696, 66, "+", "کلاس، معماری، تست‌های سبز", size=30)
    s.text(520, 478, 696, 60, "پنج سال، فقط patch و minor.", size=30, color=MUTED, align="r")
    return s


COMMENTS_94 = [("بقیه", "حیفِ پنج سال سابقه‌ی جاوا نیست؟"),
               ("خودم", "با یه دستورِ اشتباه تو ترمینالِ پروداکشن،\nکل سرویس می‌خوابه."),
               ("خودم", "هویتِ دولوپری‌ت نابود می‌شه.")]


def s06(step):
    s = Slide("6" if step == 0 else "6.1", fade=(step == 0))
    backdrop(s, "03_door.jpg", "dim")
    chrome(s, "v2.0.0", "BREAKING CHANGE", RED, "۱۳۹۴", at=1)
    row(s, 64, 112, 1152, 70, "-", "نقش: دولوپرِ جاوا", size=32, font=F_BOLD)
    row(s, 64, 192, 1152, 70, "+", "نقش: عملیات، با ذهنیتِ برنامه‌نویس", size=32, font=F_BOLD)
    if step == 1:
        s.text(64, 282, 1152, 40, "سه کامنت روی این تغییر", size=22, color=MUTED, align="r")
        for i, (who, quote) in enumerate(COMMENTS_94):
            bubble(s, 64, 330 + i * 112, 1152, 100, who, quote)
    return s


def s07(step):
    s = Slide("7")
    backdrop(s, "04_queue.jpg", "hero")
    chrome(s, "v2.0.0", "Added", GREEN, "۱۳۹۴", at=1)
    s.text(500, 112, 716, 84, "صف پشتِ درِ عملیات", size=48, font=F_BLACK, align="r")
    row(s, 500, 218, 716, 100, "-", "هر کوئری و هر لاگ:\nیک درخواست به عملیات", size=28)
    row(s, 500, 330, 716, 132, "+", "{black|Query Executor}\nکوئریِ امن، دانلودِ لاگ، Live Tail", size=28)
    s.text(500, 484, 716, 46, "ایده و طراحیِ اولیه: آقای شریفی", size=26, color=GOLD, font=F_BOLD, align="r")
    s.text(500, 544, 716, 70, "صف تمام شد.", size=40, font=F_BLACK, align="r")
    return s


def s08(step):
    s = Slide("8" if step == 0 else "8.1", fade=(step == 0))
    backdrop(s, "05_key.jpg", "dim")
    chrome(s, "v2.0.1", "HOTFIX", RED, at=1)
    window(s, 64, 104, 1152, 330, "SQL", "production" if step == 1 else None, RED)
    s.text(104, 172, 1072, 40, "قرار بود:", size=24, color=MUTED, align="r")
    code(s, 104, 214, 1072, ["{gold|ALTER USER} {blue|dev_user} {gold|IDENTIFIED BY} {green|\"********\"};"], size=28, pitch=56)
    if step == 0:
        s.text(64, 476, 1152, 90, "یک خط کوئری. همین.", size=40, color=MUTED, font=F_BOLD, align="c")
    else:
        s.text(104, 292, 1072, 40, "زده شد:", size=24, color=RED, font=F_BOLD, align="r")
        code(s, 104, 334, 1072, ["{gold|ALTER USER} {red,b|main_user} {gold|IDENTIFIED BY} {green|\"********\"};"],
             size=28, pitch=56, tint={0: RED})
        s.text(64, 462, 1152, 56, "دیتابیسِ اصلی قطع شد.", size=30, color=MUTED, align="c")
        s.text(64, 522, 1152, 110, "بانک: {red|یک ساعت} پایین", size=60, font=F_BLACK, align="c")
    return s


def s09(step):
    s = Slide("9" if step == 0 else "9.1", fade=(step == 0))
    backdrop(s, "05_key.jpg", "dim")
    chrome(s, "v2.0.1", "HOTFIX", RED, at=1)
    rows = [("اثر", RED, "بانک، یک ساعت پایین", 36, TEXT),
            ("Trigger", GOLD, "یک خط کوئری، با دست", 36, TEXT)]
    if step == 0:
        rows.append(("Root cause", BLUE, "؟", 60, BLUE))
    else:
        rows.append(("Root cause", BLUE, "آن خط را با دست روی پروداکشن می‌زدیم.", 36, BLUE))
    card(s, 64, 110, 1152, rows, row_h=124)
    if step == 1:
        s.text(64, 508, 1152, 130, "اگر محتاط‌ترین آدمِ تیم می‌تواند،\n{gold|مشکل آدم نیست.}", size=38, font=F_BLACK, align="c", spacing=0.95)
    return s


def s10(step):
    s = Slide("10")
    backdrop(s, "06_scale.jpg", "hero")
    chrome(s, "v2.0.2", "HOTFIX", RED, at=1)
    s.text(500, 104, 716, 40, "پروژه‌ی بانک سپه", size=24, color=MUTED, align="r")
    s.text(500, 148, 716, 130, "الگوریتمِ لودبالانسر را روی سیستمِ لایو عوض کردم.", size=38, font=F_BLACK, align="r", spacing=0.95)
    s.text(500, 286, 716, 96, "کل بانک قطع شد.", size=56, color=RED, font=F_BLACK, align="r")
    card(s, 500, 404, 716, [("Trigger", GOLD, "خودم", 34, TEXT),
                            ("Root cause", BLUE, "تغییرِ مستقیم روی سیستمِ زنده", 28, BLUE)],
         row_h=110, label_w=196)
    return s


def s11(step):
    s = Slide("11")
    backdrop(s, None, "plain")
    chrome(s, "v2.0.0", "resolved", GREEN, "۱۳۹۴", at=1)
    data = [("خودم", "کل سرویس می‌خوابه.", "true", "خراب کردیم؛ جلوی همه هم خراب کردیم. کار همان‌جا تمام نشد."),
            ("خودم", "هویتِ دولوپری‌ت نابود می‌شه.", "false", "همان برنامه‌نویسی بود که صف پشتِ درِ عملیات را جمع کرد."),
            ("بقیه", "حیفِ پنج سال سابقه‌ی جاوا نیست؟", "resolved", "آن پنج سال دور ریخته نشد؛ Query Executor یک وب‌اپلیکیشنِ جاوا بود.")]
    for i, (who, quote, status, note) in enumerate(data):
        bubble(s, 64, 108 + i * 186, 1152, 170, who, quote, status=status, size=30, note=note)
    return s


def s12(step):
    s = Slide("12" if step == 0 else "12.1", fade=(step == 0))
    backdrop(s, "07_manuals.jpg", "hero")
    chrome(s, "v2.1.0", "Added" if step == 0 else "Removed", GREEN if step == 0 else RED, "۱۳۹۵", at=1)
    s.text(540, 100, 676, 104, "WebLogic", size=64, font=F_BLACK, align="r", rtl=False)
    row(s, 540, 216, 676, 62, "+", "۴ سرور، هر کدام ۶۴ گیگ رم", size=28)
    row(s, 540, 286, 676, 62, "+", "یک بار: حدودِ ۸۰٪ بارِ کل CMS", size=28)
    row(s, 540, 356, 676, 62, "+", "ماژولِ احراز هویت، از نو", size=28)
    s.text(540, 436, 676, 110, "۲ تا ۳ هزار صفحه داکیومنت،\nبا دیکشنری", size=32, color=GOLD, font=F_BLACK, align="r", spacing=0.95)
    if step == 1:
        s.rect(570, 244, 560, 116, fill=BG, alpha=88, line=RED, lw=4, r=14, rot=-7)
        s.text(570, 244, 560, 116, "@Deprecated", size=54, color=RED, font=F_MONO, bold=True, align="c", rtl=False, rot=-7)
        s.text(540, 560, 676, 96, "۱۴۰۰، پروژه‌ی سپه: کنار گذاشته شد.\nهر ابزار برای یک دوره است.", size=28, align="r", spacing=0.95)
    return s


def s13(step):
    s = Slide("13" if step == 0 else "13.1", fade=(step == 0))
    backdrop(s, "08_aisle.jpg", "dim")
    chrome(s, "v2.2.0", year="۱۳۹۶", at=1)
    window(s, 140, 104, 1000, 344, "deploy.sh", "no review", RED)
    if step == 0:
        code(s, 196, 188, 900, LOOP, size=38, pitch=78)
        s.text(64, 478, 1152, 130, "NFS جواب نداد. Ansible نداشتیم.\nهمین سه خط بود.", size=34, font=F_BOLD, align="c", spacing=0.95)
    else:
        code(s, 196, 188, 900, LOOP_SHUTDOWN, size=38, pitch=78, tint={1: RED})
        s.text(64, 478, 1152, 130, "یک روز {red|shutdown} رفت توی حلقه.", size=46, font=F_BLACK, align="c")
    return s


def s14(step):
    s = Slide("14" if step == 0 else f"14.{step}", fade=(step == 0))
    backdrop(s, None, "plain")
    chrome(s, "v2.2.0", year="۱۳۹۶", at=1)
    off = (0, 32, 70)[step]
    cell, gap = 60, 12
    x0, y0 = 142, 116
    for i in range(70):
        cx = x0 + (i % 14) * (cell + gap)
        cy = y0 + (i // 14) * (cell + gap)
        if i < off:
            s.rect(cx, cy, cell, cell, fill="15171E", line="262A34", lw=1.5, r=10)
            s.oval(cx + 40, cy + 10, 10, 10, fill="5A2C2C")
        else:
            s.rect(cx, cy, cell, cell, fill="113322", line=GREEN, lw=1.5, r=10, line_alpha=70)
            s.oval(cx + 40, cy + 10, 10, 10, fill=GREEN)
        for k in range(3):
            s.rect(cx + 10, cy + 28 + k * 9, 30, 3, fill=WHITE, alpha=10 if i < off else 28)
    num, col = (("۷۰", "green"), ("۳۸", "gold"), ("۰", "red"))[step]
    s.text(64, 478, 1152, 96, "{%s|%s} از ۷۰ سرور روشن" % (col, num), size=52, font=F_BLACK, align="c")
    sub = ("حلقه از کسی تأیید نمی‌گیرد.", "سرور به سرور…", "بانک: {red|۳۰ دقیقه} قطع")[step]
    s.text(64, 574, 1152, 70, sub, size=38 if step == 2 else 30, color=TEXT if step == 2 else MUTED,
           font=F_BLACK if step == 2 else F_REG, align="c")
    return s


def s15(step):
    s = Slide("15" if step == 0 else "15.1", fade=(step == 0))
    backdrop(s, "08_aisle.jpg", "dim")
    chrome(s, "v2.2.0", "Fixed" if step else None, GREEN, "۱۳۹۶", at=1)
    rows = [("اثر", RED, "۷۰ سرور خاموش · بانک ۳۰ دقیقه قطع", 34, TEXT),
            ("Trigger", GOLD, "یک دستورِ اشتباه", 34, TEXT)]
    if step == 0:
        rows.append(("Root cause", BLUE, "؟", 60, BLUE))
        card(s, 64, 110, 1152, rows, row_h=124)
        s.text(64, 508, 1152, 110, "root cause را خودتان اولِ ارائه گفتید.", size=32, color=MUTED, align="c")
    else:
        rows.append(("Root cause", BLUE, "دستورِ بدونِ بازبینی و تکرارناپذیر، روی پروداکشن", 30, BLUE))
        rows.append(("Fix", GREEN, "Git  ·  Review  ·  Pipeline", 36, GREEN))
        card(s, 64, 104, 1152, rows, row_h=112)
        s.text(64, 566, 1152, 70, "می‌توانست دستِ هر کدام از ما باشد.", size=30, color=MUTED, align="c")
    return s


def s16(step):
    s = Slide("16" if step == 0 else f"16.{step}", fade=(step == 0))
    backdrop(s, "09_pipes.jpg", "dim")
    chrome(s, "v2.2.0", "Fixed", GREEN, "۱۳۹۶", at=1)
    row(s, 64, 104, 1152, 62, "-", "حدودِ ۷۰ اسکریپت روی خودِ سرورها، بیرون از ورژن‌کنترل")
    row(s, 64, 174, 1152, 62, "+", "همه در Git؛ همه پایپ‌لاینِ Jenkins، از نوعِ scripted")
    row(s, 64, 244, 1152, 62, "+", "لاگِ دقیق برای هر سرور")
    row(s, 64, 314, 1152, 62, "+", "۶ بانک، یکپارچه از یک مستر")
    if step >= 1:
        bubble(s, 64, 404, 1152, 108, "همه‌ی ما", "اگه با یه کلیک دیپلوی بشه،\nمدیرها دیگه با ما چی کار دارن؟",
               status="open" if step == 1 else "resolved", size=28)
    if step == 2:
        s.text(64, 528, 1152, 116, "با Groovy فریم‌ورکی نوشتیم که کار را راحت کند، نه این‌که جای کسی را بگیرد.\n{gold|کسی را دور نینداختند.}",
               size=26, align="c", spacing=0.98)
    return s


def s17(step):
    s = Slide("17")
    backdrop(s, "10_lens.jpg", "hero")
    chrome(s, "v2.3.0", "Added", GREEN, "۱۳۹۷", at=1)
    s.text(440, 96, 776, 110, "ELK", size=64, font=F_BLACK, align="r", rtl=False)
    row(s, 440, 220, 776, 108, "-", "لاگ روی تک‌تکِ سرورها؛\nپیدا کردنِ یک خطا با SSH، دستی", size=28)
    row(s, 440, 340, 776, 66, "+", "همه‌ی لاگ‌ها یک‌جا", size=30, font=F_BOLD)
    row(s, 440, 418, 776, 108, "+", "یک کانفیگِ Filebeat برای همه‌ی سرویس‌ها\n{gold,bold|۵۰۰ تا ۶۰۰ خط}", size=26)
    return s


def s18(step):
    s = Slide("18")
    backdrop(s, None, "plain")
    chrome(s, "v2.x", at=1)
    s.text(64, 96, 1152, 76, "Trigger ≠ Root cause", size=44, font=F_MONO, bold=True, align="c", rtl=False)
    for i, (ver, who) in enumerate((("v2.0.1", "من و همکارم"), ("v2.0.2", "خودم"), ("v2.2.0", "یکی از ما"))):
        x = W - 64 - 368 - i * 392
        s.rect(x, 196, 368, 156, fill=PANEL, alpha=94, line=BORDER, lw=1.5, r=18)
        s.text(x + 24, 206, 150, 40, ver, size=20, color=MUTED, font=F_MONO, align="l", rtl=False)
        s.text(x + 194, 206, 150, 40, "Trigger", size=20, color=GOLD, font=F_MONO, bold=True, align="r", rtl=False)
        s.text(x, 250, 368, 90, who, size=40, font=F_BLACK, align="c")
    s.rect(64, 376, 1152, 112, fill=BLUE, alpha=8, line=BLUE, lw=2, r=18)
    s.text(W - 64 - 30 - 200, 376, 200, 112, "Root cause", size=24, color=BLUE, font=F_MONO, bold=True, align="r", rtl=False)
    s.text(96, 376, 880, 112, "تغییرِ مستقیم روی پروداکشن", size=36, font=F_BOLD, align="r")
    s.text(64, 520, 1152, 110, "آدم‌ها را عوض نکردیم؛ {gold|روشِ کار} را عوض کردیم.", size=42, font=F_BLACK, align="c")
    return s


STAGES = ["نیاز", "تحقیق\nو POC", "مستندسازی", "آماده‌سازیِ\nپروداکشن", "آموزش\nو تحویل"]


def s19(step):
    s = Slide("19")
    backdrop(s, "11_workbench.jpg", "dim")
    chrome(s, "v3.0.0", "BREAKING CHANGE", RED, "۱۳۹۸", at=2)
    s.text(64, 104, 1152, 76, "R&D Ops: هسته‌ی ۵ نفره", size=44, font=F_BLACK, align="r")
    s.text(64, 180, 1152, 46, "۴ برنامه‌نویسِ باسابقه + ۱ مهندسِ دوآپس", size=26, color=MUTED, align="r")
    nw, ng = 212, 23
    for i, title in enumerate(STAGES):
        x = W - 64 - nw - i * (nw + ng)
        last = i == len(STAGES) - 1
        s.rect(x, 262, nw, 136, fill=GOLD if last else PANEL, alpha=12 if last else 94, line=GOLD, lw=2 if last else 1.5,
               r=16, line_alpha=100 if last else 55)
        s.oval(x + nw - 50, 274, 36, 36, fill=GOLD)
        s.text(x + nw - 50, 274, 36, 36, fa(i + 1), size=20, color=BG, font=F_BLACK, align="c")
        s.text(x + 12, 306, nw - 24, 86, title, size=24, font=F_BOLD, align="c", spacing=0.92)
        if not last:
            s.line(x - ng + 3, 330, x - 3, 330, GOLD, lw=2, alpha=70)
    names = ["Kafka", "ActiveMQ", "ELK", "Kubernetes"]
    widths = [int(mono_w(n, 22) + 40) for n in names]
    x = (W - sum(widths) - 16 * 3) / 2
    for n, cw in zip(names, widths):
        s.rect(x, 438, cw, 48, fill=PANEL2, alpha=96, line=BORDER, lw=1.5, r=24)
        s.text(x, 438, cw, 48, n, size=22, color=TEXT, font=F_MONO, align="c", rtl=False)
        x += cw + 16
    s.text(64, 500, 1152, 44, "همه از همین مسیرِ پنج‌مرحله‌ای رد شدند.", size=24, color=MUTED, align="c")
    s.text(64, 566, 1152, 60, "تحویل به: تیم‌های عملیات و تیمِ middleware", size=28, font=F_BOLD, align="c")
    return s


def s20(step):
    s = Slide("20" if step == 0 else "20.1", fade=(step == 0))
    backdrop(s, "12_handover.jpg", "top")
    chrome(s, "v3.0.0", year="۱۳۹۸", at=2)
    s.text(40, 130, 1200, 150, "بساز. تحویل بده. {gold|رها کن.}", size=80, font=F_BLACK, align="c")
    s.text(64, 296, 1152, 50, "تحویل یعنی: آموزش + سند + آماده‌سازیِ پروداکشن + پشتیبانیِ فنی", size=26, color=MUTED, align="c")
    pill(s, (W - 222) // 2, 362, "enabling team", GOLD, w=222)
    if step == 1:
        s.text(64, 446, 1152, 100, "کیفیت مالِ ماست؛ مالکیت نه.", size=48, font=F_BLACK, align="c")
    return s


def s21(step):
    s = Slide("21" if step == 0 else "21.1", fade=(step == 0))
    backdrop(s, "13_pillar.jpg", "hero")
    chrome(s, "v3.0.0", year="۱۳۹۸", at=2)
    s.text(380, 112, 836, 50, "یکی از بچه‌ها به سیستم‌هایش می‌گفت «بچه‌م».", size=26, color=MUTED, align="r")
    s.text(380, 180, 836, 210, "«به این سیستم دست نزنید؛\nقلقش دستِ منه.»", size=44, font=F_BLACK, align="r")
    if step == 1:
        s.rect(W - 64 - 180, 428, 180, 70, fill=RED, alpha=16, line=RED, lw=2.5, r=14)
        s.text(W - 64 - 180, 428, 180, 70, "SPOF", size=40, color=RED, font=F_MONO, bold=True, align="c", rtl=False)
        s.text(520, 514, 696, 120, "برای سرویس چند instance بالا می‌آوریم؛ برای خودمان نه.", size=30, align="r", spacing=0.95)
    return s


def s22(step):
    s = Slide("22" if step == 0 else "22.1", fade=(step == 0))
    backdrop(s, "14_dusk_office.jpg", "bottom")
    chrome(s, "v3.1.0", year="۱۳۹۹", at=2)
    if step == 0:
        s.text(480, 540, 736, 96, "یک عصر، سالِ ۹۹", size=52, font=F_BLACK, align="r")
    else:
        s.text(360, 408, 856, 50, "پایپ‌لاین‌ها و اسکریپت‌هایی که تا دیروز کارِ من بود", size=28, color=MUTED, align="r")
        s.text(260, 462, 956, 100, "سریع‌تر. تمیزتر. بدونِ نیاز به من.", size=46, font=F_BLACK, align="r")
        s.text(480, 578, 736, 50, "یک عصر، سالِ ۹۹", size=26, color=MUTED, align="r")
    return s


def s23(step):
    s = Slide("23", bg=BLACK)
    pill(s, (W - 96) // 2, 196, "open", GOLD, w=96)
    s.text(110, 262, 1060, 230, "«فردا صبح این شرکت\nچرا باید به تو حقوق بده؟»", size=56, font=F_BLACK, align="c", spacing=0.98)
    return s


def s24(step):
    s = Slide("24")
    backdrop(s, None, "plain")
    chrome(s, "v3.1.0", "Deprecated", ORANGE, "۱۳۹۹", at=2)
    window(s, 64, 150, 580, 250, "CODEOWNERS")
    code(s, 104, 236, 500, ["{red|- pipelines/   @ali}", "{green|+ pipelines/   @teammate}"],
         size=26, pitch=62, tint={0: RED, 1: GREEN})
    s.text(64, 420, 580, 110, "کار را سپردم و رفتم سراغِ\nمعماریِ کلی و ELK.", size=28, color=MUTED, align="c")
    s.text(676, 140, 540, 80, "کدِ من deprecated شد.", size=34, font=F_BOLD, align="r")
    s.text(676, 226, 540, 190, "من نه.", size=104, color=GOLD, font=F_BLACK, align="r")
    s.text(676, 436, 540, 60, "من آن پایپ‌لاین نیستم.", size=30, color=MUTED, align="r")
    return s


def s25(step):
    s = Slide("25")
    backdrop(s, "15_bridge.jpg", "hero")
    chrome(s, "v4.0.0", year="۱۴۰۴", at=3)
    s.text(520, 100, 696, 88, "سه سال آماده‌سازی", size=52, font=F_BLACK, align="r")
    items = [("+", "Dockerfile سرویس‌ها"), ("+", "Deploymentها"), ("+", "آموزشِ تیم‌ها"),
             ("+", "ایمیج‌ها: از ۳ گیگ به ۱ گیگ"), ("-", "مهاجرت: انجام نشده")]
    for i, (kind, label) in enumerate(items):
        row(s, 520, 202 + i * 66, 696, 58, kind, label, size=26, font=F_BOLD if kind == "-" else F_REG)
    s.text(520, 546, 696, 96, "می‌خواستیم اول همه‌چیز\n{gold|صددرصد} آماده باشد.", size=30, font=F_BOLD, align="r", spacing=0.98)
    return s


def s26(step):
    s = Slide("26" if step == 0 else "26.1", fade=(step == 0))
    backdrop(s, "16_blackout.jpg", "hero")
    chrome(s, "v4.0.0", "BREAKING CHANGE", RED, "خرداد ۱۴۰۴", at=3)
    s.text(360, 150, 856, 50, "اینسیدنت وقتِ قبلی نمی‌گیرد.", size=28, color=MUTED, align="r")
    s.text(360, 206, 856, 120, "کل زیرساخت از بین رفت.", size=56, font=F_BLACK, align="r")
    if step == 1:
        s.text(360, 372, 856, 56, "تعمیر یا مهاجرت؟ {gold,black|مهاجرت.}", size=34, align="r")
        s.text(360, 440, 856, 140, "۳ سال {mono,muted|←} ۳ شب", size=84, color=GOLD, font=F_BLACK, align="r")
    return s


NIGHTS = [("سرویس‌ها چطور با هم حرف بزنند؟", "سه کلاستر و یک استانداردِ ارتباطی برای همه"),
          ("Nexus کند شد.", "خودِ Nexus نبود؛\nپهنای باندِ سرورش پر شده بود."),
          ("اپ‌های stateful", "همان‌طور stateful، در کلاستر.\nدیتابیس‌ها سرِ جایشان ماندند.")]


def s27(step):
    s = Slide("27" if step == 0 else f"27.{step}", fade=(step == 0))
    backdrop(s, "17_night_ops.jpg", "dim")
    chrome(s, "v4.0.0", year="خرداد ۱۴۰۴", at=3)
    s.text(64, 96, 1152, 64, "مسئله‌های آن سه شب", size=40, font=F_BLACK, align="r")
    for i, (problem, fix) in enumerate(NIGHTS[:step + 1]):
        y = 176 + i * 156
        s.rect(64, y, 1152, 140, fill=PANEL, alpha=94, line=BORDER, lw=1.5, r=18)
        s.oval(W - 64 - 76, y + 46, 48, 48, fill=GOLD)
        s.text(W - 64 - 76, y + 46, 48, 48, fa(i + 1), size=26, color=BG, font=F_BLACK, align="c")
        s.text(640, y, 478, 140, problem, size=30, font=F_BLACK, align="r", spacing=0.95)
        s.text(584, y, 48, 140, "←", size=30, color=DIM, font=F_MONO, align="c", rtl=False)
        s.text(92, y, 484, 140, fix, size=26, color=GREEN, font=F_BOLD, align="r", spacing=0.95)
    return s


def s28(step):
    s = Slide("28" if step == 0 else "28.1", fade=(step == 0))
    backdrop(s, "18_padlock.jpg", "hero")
    chrome(s, "v4.0.0", year="خرداد ۱۴۰۴", at=3)
    s.text(500, 116, 716, 170, "اپِ اصلی\nبالا نمی‌آمد.", size=58, font=F_BLACK, align="r", spacing=0.92)
    s.text(500, 296, 716, 54, "داشتیم برمی‌گشتیم روی VMها.", size=30, color=MUTED, align="r")
    if step == 1:
        s.text(500, 386, 716, 40, "Root cause", size=24, color=BLUE, font=F_MONO, bold=True, align="r", rtl=False)
        s.text(500, 424, 716, 140, "یک دسترسی.", size=88, color=GOLD, font=F_BLACK, align="r")
        s.text(500, 572, 716, 50, "سه سال آماده‌سازی؛ و یک دسترسی.", size=26, color=MUTED, align="r")
    return s


def s29(step):
    s = Slide("29" if step == 0 else f"29.{step}", fade=(step == 0))
    if s.image("19_dawn.jpg"):
        s.rect(0, 0, W, H, fill=BG, alpha=(86, 86, 70)[step])
    else:
        s.grad(0, 0, W, H, [(0, "0B0C10", 100), (100, ("14161E", "14161E", "10221A")[step], 100)], vertical=True)
    chrome(s, "v4.0.0", year="خرداد ۱۴۰۴", at=3)
    clock = ("06:00", "06:15", "06:30")[step]
    s.text(64, 96, 640, 230, clock, size=150, color=GREEN if step == 2 else TEXT, font=F_MONO, bold=True, align="l", rtl=False)
    caption = ("همه‌چیز بالا. بار: صفر.", "به هیچ‌چیز دست نزدیم.", "شعبه‌ها باز شدند.")[step]
    s.text(64, 326, 600, 56, caption, size=32, color=MUTED if step < 2 else TEXT, font=F_BOLD, align="l")
    # schematic load chart (no real numbers)
    gx, gy, gw, gh = 724, 110, 492, 286
    s.rect(gx, gy, gw, gh, fill=PANEL, alpha=90, line=BORDER, lw=1.5, r=16)
    s.text(gx + 22, gy + 8, 200, 40, "load", size=20, color=MUTED, font=F_MONO, align="l", rtl=False)
    axis_y, base_y = gy + 226, gy + 212
    s.line(gx + 26, axis_y, gx + gw - 26, axis_y, BORDER, lw=2)
    ticks = [gx + 66, gx + 226, gx + 386]
    for tx, label in zip(ticks, ("06:00", "06:15", "06:30")):
        s.line(tx, axis_y, tx, axis_y + 8, DIM, lw=2)
        s.text(tx - 50, axis_y + 10, 100, 36, label, size=20, color=MUTED, font=F_MONO, align="c", rtl=False)
    end_x = ticks[step]
    s.poly([(gx + 26, base_y), (end_x, base_y)], GOLD, lw=4)
    if step == 2:
        rise = [(end_x, base_y), (end_x + 22, gy + 150), (end_x + 40, gy + 92), (end_x + 58, gy + 66), (gx + gw - 30, gy + 60)]
        s.poly(rise, GREEN, lw=5)
        s.oval(gx + gw - 38, gy + 52, 16, 16, fill=GREEN)
    else:
        s.oval(end_x - 8, base_y - 8, 16, 16, fill=GOLD)
    if step == 0:
        s.text(64, 470, 1152, 110, "منتظرِ باز شدنِ شعبه‌ها", size=40, color=MUTED, font=F_BOLD, align="c")
    elif step == 1:
        bubble(s, 64, 424, 1152, 104, "خودم", "اگه زیرِ بارِ صبح پادها بخوابن،\nجلوی کل کشور آبروی هممون می‌ره.", size=26)
        bubble(s, 64, 540, 1152, 104, "خودم", "اگه بالا نیاد، ۱۵ سال سابقه‌ت کشک بود.", size=28)
    else:
        s.text(64, 432, 1152, 190, "نمودار بالا رفت؛\n{green|و کلاستر نگهش داشت.}", size=50, font=F_BLACK, align="c")
    return s


def s30(step):
    s = Slide("30")
    if s.image("19_dawn.jpg"):
        s.rect(0, 0, W, H, fill=BG, alpha=78)
    else:
        backdrop(s, None, "plain")
    chrome(s, "v4.0.0", "Released", GREEN, "۱۴۰۴", at=3)
    row(s, 64, 108, 1152, 66, "+", "همه‌ی سرویس‌ها روی Kubernetes، جز ویندوزی‌ها", size=30, font=F_BOLD)
    row(s, 64, 184, 1152, 104, "~", "{mono,orange,b|Known issue}\nدامنه‌ی سرویس‌ها با دامنه‌ی داخلیِ Kubernetes یکی بود؛ عوضش کردیم.", size=24)
    s.text(40, 318, 1200, 120, "آمادگیِ صددرصد هیچ‌وقت نمی‌رسد.", size=54, color=GOLD, font=F_BLACK, align="c")
    s.text(64, 448, 1152, 60, "آن سه سال لازم بود؛ بدونِ آن، آن سه شب ممکن نبود.", size=30, align="c")
    s.text(64, 546, 1152, 70, "سه شب بی‌خوابی افتخار نیست.", size=36, color=MUTED, font=F_BOLD, align="c")
    return s


def s31(step):
    s = Slide("31")
    backdrop(s, "20_graph.jpg", "hero")
    chrome(s, "Unreleased", year="امروز", at=4)
    s.text(500, 80, 716, 120, "ISR", size=68, font=F_BLACK, align="r", rtl=False)
    s.text(500, 190, 716, 40, "Intelligent SRE", size=24, color=GOLD, font=F_MONO, bold=True, align="r", rtl=False)
    items = [("+", "n8n: ارکستریشنِ فرایندها"), ("+", "اجرای جاب‌های Jenkins"), ("+", "تگِ نسخه‌ها در Git"),
             ("+", "ایجنت‌های AI: تحلیلِ لاگ و عیب‌یابی"), ("+", "ژنراتورِ توپولوژی"),
             ("~", "تأییدِ انسان: فعلاً تقریباً همه‌جا")]
    for i, (kind, label) in enumerate(items):
        row(s, 500, 248 + i * 64, 716, 56, kind, label, size=26, font=F_BOLD if kind == "~" else F_REG)
    return s


def s32(step):
    s = Slide("32")
    backdrop(s, None, "plain")
    chrome(s, "Unreleased", "demo", GOLD, "امروز", at=4)
    if os.path.exists(os.path.join(SLIDES_DIR, DEMO_VIDEO)):
        s.video = os.path.join(SLIDES_DIR, DEMO_VIDEO)
    s.oval(W / 2 - 80, 150, 160, 160, fill=GOLD, alpha=10, line=GOLD, lw=3)
    s.play(W / 2 - 22, 196, 60, 68, GOLD)
    s.text(64, 344, 1152, 110, "دموی ISR", size=64, font=F_BLACK, align="c")
    s.text(64, 456, 1152, 56, "از هشدار تا رفع", size=30, color=MUTED, align="c")
    return s


def s33(step):
    s = Slide("33" if step == 0 else "33.1", bg=BLACK, fade=(step == 0))
    s.text(64, 110, 1152, 270, "۳ ثانیه", size=170, color=GOLD, font=F_BLACK, align="c")
    s.text(64, 384, 1152, 56, "اولین کدِ تمیزی که AI تحویلم داد", size=30, color=MUTED, align="c")
    if step == 1:
        bubble(s, 110, 500, 1060, 110, "خودم", "یعنی این همه سال سینتکس حفظ‌کردن کشک بود؟", size=30)
    return s


def s34(step):
    s = Slide("34")
    backdrop(s, None, "plain")
    chrome(s, "Unreleased", year="امروز", at=4)
    for x, label, quote, col in ((676, "دیروز", "به این سیستم دست نزنید؛ قلقش دستِ منه.", TEXT),
                                 (64, "امروز", "AI چرت‌وپرت می‌نویسه؛ هیچ ماشینی جای کد زدنِ من رو نمی‌گیره.", TEXT)):
        s.rect(x, 130, 540, 330, fill=PANEL, alpha=94, line=BORDER, lw=1.5, r=20)
        s.text(x + 30, 146, 480, 44, label, size=24, color=GOLD, font=F_BOLD, align="r")
        s.text(x + 30, 196, 480, 250, "«" + quote + "»", size=32, color=col, font=F_BLACK, align="r")
    s.text(604, 130, 72, 330, "==", size=40, color=GOLD, font=F_MONO, bold=True, align="c", rtl=False)
    s.text(64, 500, 1152, 100, "همان جمله، با یک لباسِ دیگر.", size=44, font=F_BLACK, align="c")
    return s


def s35(step):
    s = Slide("35")
    backdrop(s, "21_box.jpg", "hero")
    chrome(s, "v5.0.0", "Unreleased", ORANGE, "امروز", at=4)
    s.text(440, 122, 776, 52, "## [5.0.0] - Unreleased", size=30, color=GOLD, font=F_MONO, bold=True, align="r", rtl=False)
    s.text(440, 178, 776, 46, "### BREAKING CHANGE", size=26, color=RED, font=F_MONO, bold=True, align="r", rtl=False)
    s.text(400, 262, 816, 130, "هنوز نوشته نشده.", size=64, font=F_BLACK, align="r")
    s.text(440, 410, 776, 60, "این یکی را خودمان می‌نویسیم.", size=34, color=MUTED, align="r")
    return s


LEDGER = [("دولوپرِ جاوا", "۱۳۹۴: رفتم عملیات"),
          ("WebLogic", "۱۴۰۰، سپه: کنار رفت"),
          ("حلقه‌ی for و ۷۰ اسکریپت", "۱۳۹۶: Git و Jenkins"),
          ("پایپ‌لاین‌های دیروزِ من", "۱۳۹۹: همکارم بازنویسی کرد"),
          ("سرویس‌ها روی VM", "۱۴۰۴: Kubernetes")]


def s36(step):
    s = Slide("36" if step == 0 else "36.1", fade=(step == 0))
    backdrop(s, None, "plain")
    s.text(64, 40, 1152, 64, "یک بار، کل CHANGELOG از بالا", size=36, font=F_BLACK, align="r")
    s.text(684, 118, 532, 40, "+ Added", size=22, color=GREEN, font=F_MONO, bold=True, align="r", rtl=False)
    s.text(64, 118, 532, 40, "− Deprecated / Removed", size=22, color=RED, font=F_MONO, bold=True, align="r", rtl=False)
    for i, (added, removed) in enumerate(LEDGER):
        y = 166 + i * 72
        row(s, 684, y, 532, 62, "+", added, size=28, font=F_BOLD)
        s.text(596, y, 88, 62, "←", size=28, color=DIM, font=F_MONO, align="c", rtl=False)
        row(s, 64, y, 532, 62, "-", removed, size=28)
    if step == 1:
        s.rect(64, 548, 1152, 124, fill=GOLD, alpha=10, line=GOLD, lw=2, r=18)
        s.text(64, 554, 1152, 40, "چیزی که در هیچ نسخه‌ای Removed نشد", size=22, color=MUTED, align="c")
        s.text(64, 590, 1152, 76, "بساز · تحویل بده · رها کن · برو سراغِ بعدی", size=40, color=GOLD, font=F_BLACK, align="c")
    return s


def s37(step):
    s = Slide("37")
    backdrop(s, None, "plain")
    s.rect(64, 48, 1152, 624, fill=PANEL, alpha=94, line=BORDER, lw=1.5, r=20)
    s.text(104, 62, 500, 64, "## [Unreleased] - Saturday", size=24, color=GOLD, font=F_MONO, bold=True, align="l", rtl=False)
    s.text(620, 62, 556, 64, "CHANGELOG شما، از شنبه", size=32, font=F_BLACK, align="r")
    s.line(88, 134, W - 88, 134, BORDER, lw=1.5)
    blocks = [("### Added", GREEN, "+", "این هفته یک کارِ تکراری را بده به AI؛\nنتیجه را به تیم نشان بده."),
              ("### Deprecated", ORANGE, "~", "«فقط من بلدم»: مستندش کن و تحویل بده."),
              ("### Removed", RED, "-", "سالی یک بار، چیزی را که مرجعش هستی رها کن.")]
    for i, (label, col, kind, textline) in enumerate(blocks):
        y = 150 + i * 172
        s.text(104, y, 1072, 44, label, size=24, color=col, font=F_MONO, bold=True, align="r", rtl=False)
        row(s, 104, y + 48, 1072, 108, kind, textline, size=28, font=F_BOLD)
    return s


def s38(step):
    s = Slide("38" if step == 0 else "38.1", fade=(step == 0))
    backdrop(s, "22_gold_network.jpg", "center")
    s.text(64, 140, 1152, 96, "کد deprecated می‌شود.", size=54, font=F_BLACK, align="c")
    s.text(64, 240, 1152, 96, "ابزار deprecated می‌شود.", size=54, font=F_BLACK, align="c")
    if step == 1:
        s.text(40, 370, 1200, 110, "{gold|شما نه؛} اگر نسخه‌ی بعدی را بدهید.", size=52, font=F_BLACK, align="c")
        s.rect(W / 2 - 66, 506, 132, 5, fill=GOLD)
        s.text(64, 532, 1152, 80, "آخرین major شما کِی بود؟", size=40, font=F_BOLD, align="c")
    return s


def s39(step):
    s = Slide("39")
    backdrop(s, "01_cover.jpg", "hero")
    s.text(420, 196, 796, 160, "پرسش و پاسخ", size=92, font=F_BLACK, align="r")
    s.rect(W - 64 - 132, 372, 132, 5, fill=GOLD)
    s.text(420, 396, 796, 52, "## [Unreleased]", size=30, color=GOLD, font=F_MONO, bold=True, align="r", rtl=False)
    s.text(420, 456, 796, 52, "علی امانی", size=28, color=MUTED, font=F_BOLD, align="r")
    return s


# (builder, number of steps). Each step is one slide in the pptx and one press of the clicker.
DECK = [(s01, 2), (s02, 1), (s03, 1), (s04, 2), (s05, 1), (s06, 2), (s07, 1), (s08, 2), (s09, 2), (s10, 1),
        (s11, 1), (s12, 2), (s13, 2), (s14, 3), (s15, 2), (s16, 3), (s17, 1), (s18, 1), (s19, 1), (s20, 2),
        (s21, 2), (s22, 2), (s23, 1), (s24, 1), (s25, 1), (s26, 2), (s27, 3), (s28, 2), (s29, 3), (s30, 1),
        (s31, 1), (s32, 1), (s33, 2), (s34, 1), (s35, 1), (s36, 2), (s37, 1), (s38, 2), (s39, 1)]


def build_slides():
    return [fn(step) for fn, steps in DECK for step in range(steps)]


# ---------------------------------------------------------------- speaker notes from script.md
def load_cues(path):
    """Returns [(kind, cue, [lines])] — one entry per «اسلاید» or «کلیک» stage direction, in order."""
    cues, in_speech = [], False
    with open(path, encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if line.startswith("## "):
                in_speech = False
                continue
            if line.startswith("### "):
                in_speech = "🎙️" in line
                continue
            if not in_speech or not line or line == "---" or line.startswith("<"):
                continue
            m = re.fullmatch(r"\*\((.*)\)\*", line)
            if m and m.group(1).startswith(("اسلاید", "کلیک")):
                cues.append(("slide" if m.group(1).startswith("اسلاید") else "click", m.group(1), []))
            elif cues:
                cues[-1][2].append("[" + m.group(1) + "]" if m else line)
    return cues


def attach_notes(slides, cues):
    expected = [("click" if "." in s.key else "slide") for s in slides]
    got = [kind for kind, _, _ in cues]
    if expected != got:
        for i, (e, g) in enumerate(zip(expected, got)):
            if e != g:
                raise SystemExit(f"script.md and the deck disagree at press #{i + 1} (slide {slides[i].key}): "
                                 f"deck expects a {e}, script has «{cues[i][1]}».")
        raise SystemExit(f"script.md has {len(got)} slide/click directions, the deck has {len(expected)} slides.")
    for i, (s, (_, cue, lines)) in enumerate(zip(slides, cues)):
        nxt = cues[i + 1][1] if i + 1 < len(cues) else "پایان"
        s.notes = "▶ " + cue + "\n\n" + "\n\n".join(lines) + "\n\n⏭ بعدی: " + nxt


# ---------------------------------------------------------------- pptx writer
def write_pptx(slides, out_path):
    from pptx import Presentation
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
    from pptx.enum.dml import MSO_LINE
    from pptx.oxml import parse_xml

    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"

    def px(v):
        return Emu(int(round(v * EMU_PER_PX)))

    def fill_xml(color, alpha):
        if color is None:
            return f'<a:noFill xmlns:a="{A}"/>'
        return (f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="{color}">'
                f'<a:alpha val="{int(alpha * 1000)}"/></a:srgbClr></a:solidFill>')

    def line_xml(color, lw, alpha=100, dash=False):
        if color is None:
            return f'<a:ln xmlns:a="{A}"><a:noFill/></a:ln>'
        dash_xml = '<a:prstDash val="dash"/>' if dash else ""
        return (f'<a:ln xmlns:a="{A}" w="{int(lw * EMU_PER_PX)}"><a:solidFill><a:srgbClr val="{color}">'
                f'<a:alpha val="{int(alpha * 1000)}"/></a:srgbClr></a:solidFill>{dash_xml}</a:ln>')

    def restyle(shape, fill, line):
        """Replace fill, outline and effects of an autoshape, in schema order."""
        spPr = shape._element.spPr
        for child in list(spPr):
            if child.tag.split("}")[1] in ("solidFill", "noFill", "gradFill", "ln", "effectLst"):
                spPr.remove(child)
        spPr.append(parse_xml(fill))
        spPr.append(parse_xml(line))
        spPr.append(parse_xml(f'<a:effectLst xmlns:a="{A}"/>'))

    align_map = {"r": PP_ALIGN.RIGHT, "l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER}
    anchor_map = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}

    prs = Presentation()
    prs.slide_width, prs.slide_height = px(W), px(H)
    blank = prs.slide_layouts[6]

    for s in slides:
        slide = prs.slides.add_slide(blank)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(s.bg)
        shapes = slide.shapes

        for it in s.items:
            kind = it[0]
            if kind == "image":
                _, path, x, y, w, h = it
                pic = shapes.add_picture(path, px(x), px(y))
                iw, ih = pic.image.size
                want, have = w / h, iw / ih
                if have > want:
                    pic.crop_left = pic.crop_right = (1 - want / have) / 2
                else:
                    pic.crop_top = pic.crop_bottom = (1 - have / want) / 2
                pic.left, pic.top, pic.width, pic.height = px(x), px(y), px(w), px(h)
            elif kind == "rect":
                _, x, y, w, h, fill, alpha, line, lw, r, rot, line_alpha = it
                shp = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if r else MSO_SHAPE.RECTANGLE, px(x), px(y), px(w), px(h))
                if r:
                    shp.adjustments[0] = min(0.5, r / min(w, h))
                restyle(shp, fill_xml(fill, alpha), line_xml(line, lw, line_alpha))
                if rot:
                    shp.rotation = rot
            elif kind == "oval":
                _, x, y, w, h, fill, alpha, line, lw = it
                shp = shapes.add_shape(MSO_SHAPE.OVAL, px(x), px(y), px(w), px(h))
                restyle(shp, fill_xml(fill, alpha), line_xml(line, lw))
            elif kind == "play":
                _, x, y, w, h, fill = it
                # an isosceles triangle turned 90° clockwise points right; its box is given unrotated
                shp = shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, px(x + (w - h) / 2), px(y + (h - w) / 2), px(h), px(w))
                shp.rotation = 90
                restyle(shp, fill_xml(fill, 100), line_xml(None, 0))
            elif kind == "grad":
                _, x, y, w, h, stops, vertical = it
                shp = shapes.add_shape(MSO_SHAPE.RECTANGLE, px(x), px(y), px(w), px(h))
                gs = "".join(f'<a:gs pos="{int(p * 1000)}"><a:srgbClr val="{c}"><a:alpha val="{int(a * 1000)}"/></a:srgbClr></a:gs>'
                             for p, c, a in stops)
                grad = (f'<a:gradFill xmlns:a="{A}" rotWithShape="1"><a:gsLst>{gs}</a:gsLst>'
                        f'<a:lin ang="{5400000 if vertical else 0}" scaled="0"/></a:gradFill>')
                restyle(shp, grad, line_xml(None, 0))
            elif kind == "line":
                _, x1, y1, x2, y2, color, lw, alpha, dash = it
                cx = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, px(x1), px(y1), px(x2), px(y2))
                spPr = cx._element.spPr
                for child in list(spPr):
                    if child.tag.split("}")[1] == "ln":
                        spPr.remove(child)
                spPr.append(parse_xml(line_xml(color, lw, alpha, dash)))
            elif kind == "poly":
                _, pts, color, lw, alpha = it
                fb = shapes.build_freeform(px(pts[0][0]), px(pts[0][1]), scale=1.0)
                fb.add_line_segments([(px(x), px(y)) for x, y in pts[1:]], close=False)
                shp = fb.convert_to_shape()
                restyle(shp, fill_xml(None, 0), line_xml(color, lw, alpha))
            elif kind == "text":
                _, x, y, w, h, paras, align, anchor, rtl, spacing, rot = it
                tb = shapes.add_textbox(px(x), px(y), px(w), px(h))
                tf = tb.text_frame
                tf.word_wrap = True
                tf.auto_size = MSO_AUTO_SIZE.NONE
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                tf.vertical_anchor = anchor_map[anchor]
                for i, runs in enumerate(paras):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.alignment = align_map[align]
                    p.line_spacing = spacing
                    p._p.get_or_add_pPr().set("rtl", "1" if rtl else "0")
                    for txt, st in runs:
                        run = p.add_run()
                        run.text = txt
                        f = run.font
                        f.name = st["font"]
                        f.size = Pt(st["size"])
                        f.bold = bool(st.get("bold"))
                        f.color.rgb = RGBColor.from_string(st["color"])
                        rPr = run._r.get_or_add_rPr()
                        rPr.set("lang", "fa-IR" if rtl else "en-US")
                        latin = rPr.find(f"{{{A}}}latin")
                        ea = parse_xml(f'<a:ea xmlns:a="{A}" typeface="{st["font"]}"/>')
                        cs = parse_xml(f'<a:cs xmlns:a="{A}" typeface="{st["font"]}"/>')
                        latin.addnext(ea)
                        ea.addnext(cs)
                if rot:
                    tb.rotation = rot

        if s.video:
            shapes.add_movie(s.video, px(0), px(0), px(W), px(H), mime_type="video/mp4")

        if s.fade:
            # schema order inside <p:sld>: cSld, clrMapOvr, transition, timing (a video adds timing), extLst
            trans = parse_xml(f'<p:transition xmlns:p="{P}" spd="med"><p:fade/></p:transition>')
            later = [el for el in slide._element if el.tag in (f"{{{P}}}timing", f"{{{P}}}extLst")]
            if later:
                later[0].addprevious(trans)
            else:
                slide._element.append(trans)

        if s.notes:
            ntf = slide.notes_slide.notes_text_frame
            ntf.text = s.notes
            for p in ntf.paragraphs:
                p.alignment = PP_ALIGN.RIGHT
                p._p.get_or_add_pPr().set("rtl", "1")

    try:
        prs.save(out_path)
    except PermissionError:
        raise SystemExit(f"Cannot write '{out_path}': the file is open in PowerPoint. Close it and run the build again.")


# ---------------------------------------------------------------- html preview (for checking layout in a browser)
CSS_FONT = {F_BLACK: ("'IRANYekanX'", 900), F_BOLD: ("'IRANYekanX'", 700), F_REG: ("'IRANYekanX'", 400),
            F_LIGHT: ("'IRANYekanX'", 300), F_MONO: ("Consolas, monospace", 400)}


def rgba(color, alpha):
    r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
    return f"rgba({r},{g},{b},{alpha / 100:.3f})"


def slide_html(s):
    out = [f'<div class="slide" data-key="{s.key}" style="background:#{s.bg}">']
    svg = []
    for it in s.items:
        kind = it[0]
        if kind == "image":
            _, path, x, y, w, h = it
            out.append(f'<img src="file:///{path.replace(os.sep, "/")}" style="left:{x}px;top:{y}px;width:{w}px;height:{h}px;object-fit:cover">')
        elif kind == "rect":
            _, x, y, w, h, fill, alpha, line, lw, r, rot, line_alpha = it
            css = f"left:{x}px;top:{y}px;width:{w}px;height:{h}px;border-radius:{r}px;box-sizing:border-box;"
            if fill:
                css += f"background:{rgba(fill, alpha)};"
            if line:
                css += f"border:{lw}px solid {rgba(line, line_alpha)};"
            if rot:
                css += f"transform:rotate({rot}deg);"
            out.append(f'<div style="{css}"></div>')
        elif kind == "oval":
            _, x, y, w, h, fill, alpha, line, lw = it
            css = f"left:{x}px;top:{y}px;width:{w}px;height:{h}px;border-radius:50%;box-sizing:border-box;"
            if fill:
                css += f"background:{rgba(fill, alpha)};"
            if line:
                css += f"border:{lw}px solid #{line};"
            out.append(f'<div style="{css}"></div>')
        elif kind == "play":
            _, x, y, w, h, fill = it
            out.append(f'<div style="left:{x}px;top:{y}px;width:{w}px;height:{h}px;background:#{fill};clip-path:polygon(0 0,100% 50%,0 100%)"></div>')
        elif kind == "grad":
            _, x, y, w, h, stops, vertical = it
            g = ",".join(f"{rgba(c, a)} {p}%" for p, c, a in stops)
            out.append(f'<div style="left:{x}px;top:{y}px;width:{w}px;height:{h}px;background:linear-gradient({"to bottom" if vertical else "to right"},{g})"></div>')
        elif kind == "line":
            _, x1, y1, x2, y2, color, lw, alpha, dash = it
            svg.append((len(out), f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{rgba(color, alpha)}" stroke-width="{lw}"/>'))
            out.append("")
        elif kind == "poly":
            _, pts, color, lw, alpha = it
            p = " ".join(f"{x},{y}" for x, y in pts)
            svg.append((len(out), f'<polyline points="{p}" fill="none" stroke="{rgba(color, alpha)}" stroke-width="{lw}" stroke-linejoin="miter"/>'))
            out.append("")
        elif kind == "text":
            _, x, y, w, h, paras, align, anchor, rtl, spacing, rot = it
            just = {"t": "flex-start", "m": "center", "b": "flex-end"}[anchor]
            css = f"left:{x}px;top:{y}px;width:{w}px;height:{h}px;justify-content:{just};"
            if rot:
                css += f"transform:rotate({rot}deg);"
            inner = []
            for runs in paras:
                base = runs[0][1]
                lh = LINE_HEIGHT.get(base["font"], DEFAULT_LINE_HEIGHT) * spacing
                spans = []
                for txt, st in runs:
                    fam, weight = CSS_FONT[st["font"]]
                    if st.get("bold") and st["font"] == F_MONO:
                        weight = 700
                    spans.append(f'<span style="font-family:{fam};font-weight:{weight};font-size:{st["size"]}pt;color:#{st["color"]}">{html.escape(txt)}</span>')
                direction = "rtl" if rtl else "ltr"
                talign = {"r": "right", "l": "left", "c": "center"}[align]
                inner.append(f'<div style="direction:{direction};text-align:{talign};'
                             f'line-height:{lh:.3f};font-size:{base["size"]}pt">{"".join(spans)}</div>')
            out.append(f'<div class="t" style="{css}">{"".join(inner)}</div>')
    # lines and polylines keep their z-order by being wrapped one svg per item
    for idx, frag in svg:
        out[idx] = f'<svg width="{W}" height="{H}" style="left:0;top:0">{frag}</svg>'
    if s.video:
        out.append(f'<div style="left:0;top:0;width:{W}px;height:{H}px;background:rgba(0,0,0,.6)"></div>')
    out.append("</div>")
    return "\n".join(out)


PAGE = """<!doctype html><html><head><meta charset="utf-8"><style>
html,body{margin:0;background:#000}
.slide{position:relative;width:%dpx;height:%dpx;overflow:hidden}
.slide>*{position:absolute}
.slide>img{display:block}
.t{display:flex;flex-direction:column;white-space:pre-wrap}
.sheet{display:grid;grid-template-columns:repeat(2,640px);gap:0;width:1280px}
.sheet .cell{width:640px;height:360px;overflow:hidden;position:relative}
.sheet .slide{zoom:.5}
.sheet .tag{position:absolute;left:6px;bottom:4px;font:700 13px Consolas;color:#fff;background:rgba(0,0,0,.65);padding:1px 6px;z-index:9}
</style></head><body>%s
<script>
// overflow: content clearly taller/wider than its box. fragile: 4%% less width adds a line (PowerPoint may wrap there).
document.querySelectorAll('.t').forEach(function(t){
  var key=t.closest('.slide').dataset.key, label=key+': '+t.textContent.slice(0,40);
  function height(){var h=0; t.childNodes.forEach(function(c){h+=c.offsetHeight;}); return h;}
  var h0=height(), wide=false;
  t.childNodes.forEach(function(c){ if(c.scrollWidth>t.clientWidth+1) wide=true; });
  if(h0>t.clientHeight*1.15+6||wide){ t.setAttribute('data-over',label); t.style.outline='2px solid #f0f'; }
  var w=t.style.width; t.style.width=(t.clientWidth*0.96)+'px';
  if(height()>h0+2){ t.setAttribute('data-fragile',label); t.style.outline='2px dashed #0ff'; }
  t.style.width=w;
});
</script></body></html>"""


def write_html(slides, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    frags = [slide_html(s) for s in slides]
    for i, frag in enumerate(frags, 1):
        with open(os.path.join(out_dir, f"s{i:02d}.html"), "w", encoding="utf-8") as f:
            f.write(PAGE % (W, H, frag))
    for n, start in enumerate(range(0, len(frags), 4), 1):
        cells = "".join(f'<div class="cell"><span class="tag">#{start + j + 1} · {slides[start + j].key}</span>{frag}</div>'
                        for j, frag in enumerate(frags[start:start + 4]))
        with open(os.path.join(out_dir, f"sheet{n:02d}.html"), "w", encoding="utf-8") as f:
            f.write(PAGE % (W, H, f'<div class="sheet">{cells}</div>'))
    with open(os.path.join(out_dir, "all.html"), "w", encoding="utf-8") as f:
        f.write(PAGE % (W, H, "\n".join(frags)))


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Build the BREAKING CHANGES deck")
    ap.add_argument("--out", default=OUTPUT, help="output .pptx path")
    ap.add_argument("--check", action="store_true", help="only check script.md against the slide list")
    ap.add_argument("--html", metavar="DIR", help="also write an HTML preview of every slide to DIR")
    ap.add_argument("--images", metavar="DIR", help="read background images from DIR instead of new/slides/")
    args = ap.parse_args()

    global SLIDES_DIR
    if args.images:
        SLIDES_DIR = os.path.abspath(args.images)

    slides = build_slides()
    attach_notes(slides, load_cues(SCRIPT_MD))
    logical = sum(1 for s in slides if "." not in s.key)
    print(f"deck: {len(slides)} slides in the pptx = {logical} slides + {len(slides) - logical} clicks; script.md matches.")

    missing = sorted(n for n in WANTED_IMAGES if not find_image(n))
    print(f"background images: {len(WANTED_IMAGES) - len(missing)} of {len(WANTED_IMAGES)} found in {SLIDES_DIR}")
    if missing:
        print("  not there yet (those slides use the drawn background): " + ", ".join(missing))
    if os.path.exists(os.path.join(SLIDES_DIR, DEMO_VIDEO)):
        print(f"  demo video embedded: {DEMO_VIDEO}")
    else:
        print(f"  demo video not there yet ({DEMO_VIDEO}); slide 32 is a placeholder.")
    if args.check:
        return
    if args.html:
        write_html(slides, args.html)
        print(f"html preview: {args.html}")
    write_pptx(slides, args.out)
    print(f"saved: {args.out}")


if __name__ == "__main__":
    main()
