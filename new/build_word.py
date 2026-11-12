# -*- coding: utf-8 -*-
"""
Builds BreakingChanges_Script.docx, the Persian (RTL) rehearsal copy of the talk, from script.md.

    C:\\work\\eraee-ali\\eraealiamani\\.venv\\Scripts\\python.exe build_word.py

script.md is the source of truth and is only read, never written. The document is built in memory and
written to disk in one go; if BreakingChanges_Script.docx is locked (open in Word) the script says so and
exits with code 2 instead of writing a differently named file.

RTL notes (verified against Word 16 on this machine):
  * Every paragraph carries <w:bidi/>, every run <w:rtl/> plus ascii/hAnsi/eastAsia/cs fonts, every table <w:bidiVisual/>,
    the section <w:bidi/>.
  * <w:jc> is deliberately NOT set on body paragraphs. In a bidi paragraph Word reads jc="right" as "end edge"
    (= flush LEFT) and jc="left" as "start edge". No jc at all means start-aligned, i.e. flush right. Only
    jc="center" (symmetric) is used.
  * In a bidi paragraph <w:ind w:left> is the START edge (right side) and <w:ind w:right> the END edge, while
    <w:pBdr> left/right are physical sides. All indents/borders that matter here are symmetric.
  * IRANYekanX has no glyphs for the markers and a few punctuation marks, so those characters get their own runs
    in Segoe UI Symbol / Segoe UI instead of letting Word pick a fallback font (it picks an emoji-style box for the
    play triangle and a hairline em dash).
"""
from __future__ import annotations

import datetime as _dt
import hashlib
import io
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from xml.sax.saxutils import escape

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Twips
from lxml import etree

HERE = Path(__file__).resolve().parent
SRC = HERE / "script.md"
OUT = HERE / "BreakingChanges_Script.docx"

# --------------------------------------------------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------------------------------------------------
W = nsdecls("w")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

FONT = "IRANYekanX"
FONT_B = "IRANYekanX Bold"          # the bold weight is a separate family: use it instead of the bold flag
FONT_SYM = "Segoe UI Symbol"        # markers (not in IRANYekanX)
FONT_LAT = "Segoe UI"               # dashes / middle dot / arrows (not in IRANYekanX)

ZW = "\u200c"                       # ZWNJ
FA_COMMA = "\u060c"
FA_QMARK = "\u061f"
FA_DECIMAL = "\u066b"

MARK_SLIDE, MARK_CLICK, MARK_CROWD = "\u25b6", "\u25cf", "\u25c6"        # ▶ ● ◆
SYMBOL_CHARS = set("\u25b6\u25cf\u25c6\u25a0")
LATIN_FALLBACK_CHARS = set("\u2014\u2013\u00b7\u2190\u2192")             # — – · ← →
FALLBACK_CHARS = SYMBOL_CHARS | LATIN_FALLBACK_CHARS

# Page: A4, 2 cm margins -> 9638 twips of text width
PAGE_W, PAGE_H, MARGIN = 11906, 16838, 1134
CONTENT_W = PAGE_W - 2 * MARGIN

# Palette
INK = "1F2937"        # body text, dark grey
MUTED = "4B5563"
SLATE = "475569"
NAVY = "0F172A"       # header box / table header
ACCENT = "B42318"     # act titles (the "breaking change" red)
GOLD = "9A6A00"       # slide direction band (white text: contrast ~4.7)
GOLD_SOFT = "F5D487"
TEAL = "0F766E"       # audience direction
BAND_GREY = "E5E7EB"  # click direction band
GUIDE_FILL = "F1F3F5"
GUIDE_LINE = "D6DAE0"
ZEBRA = "F3F4F6"
TOTAL_FILL = "E2E8F0"
PLACEHOLDER_FILL = "FFE24A"

GUIDE_LABEL = "راهنمای اسلایدها"
SPOKEN_LABEL_WORD = "متن گفتاری"
TIMING_TITLE = f"زمان{ZW}بندی"
TIMING_HEAD = ("پرده", "موضوع", "اسلایدها", "زمان تقریبی")

# The timing table is given verbatim by the brief (columns right-to-left: first tuple item is the rightmost column).
TIMING_ROWS = [
    ("صفر", f"LGTM{FA_QMARK} و عنوان", "۱ تا ۲", "۳ دقیقه (با تعامل جمع)"),
    ("اول", f"CHANGELOG یک آدم{FA_COMMA} نسخه{ZW}ی ۱.۰.۰", "۳ تا ۵", f"۲{FA_DECIMAL}۵ دقیقه"),
    ("دوم", f"نسخه{ZW}ی ۲.۰.۰: عملیات{FA_COMMA} Query Executor{FA_COMMA} دو اینسیدنت", "۶ تا ۱۱", f"۵{FA_DECIMAL}۵ دقیقه"),
    ("سوم", f"وب{ZW}لاجیک{FA_COMMA} همان سه خط{FA_COMMA} جنکینز{FA_COMMA} ELK", "۱۲ تا ۱۸", f"۵{FA_DECIMAL}۵ دقیقه"),
    ("چهارم", f"نسخه{ZW}ی ۳.۰.۰: R&D و «رها کن»", "۱۹ تا ۲۴", f"۴{FA_DECIMAL}۵ دقیقه"),
    ("پنجم", f"نسخه{ZW}ی ۴.۰.۰: مدیریت{FA_COMMA} از scale-up به scale-out", "۲۵ تا ۲۸", "۴ دقیقه"),
    ("ششم", f"نسخه{ZW}ی ۵.۰.۰: سه سال{FA_COMMA} سه شب و ساعت ۶ صبح", "۲۹ تا ۳۴", f"۵{FA_DECIMAL}۵ دقیقه"),
    ("هفتم", f"Unreleased: ISR{FA_COMMA} دمو{FA_COMMA} AI و صندلی شاگردی", "۳۵ تا ۴۲", f"۷{FA_DECIMAL}۵ دقیقه (با دموی ۹۰ ثانیه{ZW}ای)"),
    ("هشتم", "CHANGELOG شما", "۴۳ تا ۴۷", "۴ دقیقه"),
]
TIMING_TOTAL = ("جمع", "", f"۴۷ اسلاید{FA_COMMA} ۷۸ کلیک", "حدود ۴۲ دقیقه + پرسش و پاسخ")
HEADER_SUBTITLE = "CHANGELOG یک مهندس \u2014 متن کامل ارائه"

EXPECTED = {"slide": 47, "click": 31, "crowd": 5}
EXPECTED_PLACEHOLDERS = 2      # the «[؟ علی: …]» blanks the speaker still has to fill in


# --------------------------------------------------------------------------------------------------------------------
# Markdown model + parser
# --------------------------------------------------------------------------------------------------------------------
@dataclass
class Block:
    kind: str   # 'p' | 'slide' | 'click' | 'crowd'
    text: str


@dataclass
class Act:
    title: str
    guide: list = field(default_factory=list)       # [(label, text)]
    spoken_heading: str = ""
    blocks: list = field(default_factory=list)


@dataclass
class Script:
    h1: str = ""
    meta: list = field(default_factory=list)        # [(label, value)]
    acts: list = field(default_factory=list)
    notes: list = field(default_factory=list)       # things the parser could not place / had to guess


DIR_RE = re.compile(r"^\*\((.*)\)\*$")
BULLET_RE = re.compile(r"^[*-]\s+\*\*(.+?)\*\*\s*(.*)$")
DIR_KINDS = {"اسلاید": "slide", "کلیک": "click", "جمع": "crowd"}


def clean_heading(s: str) -> str:
    """Drop leading emoji/punctuation (and VS16) and the trailing colon of a ### heading."""
    s = re.sub(r"^[\W_]+", "", s.strip())
    return s.rstrip(":：").strip()


def direction_kind(text: str):
    first = re.split(r"[\s:：]", text.strip(), maxsplit=1)[0].strip("\u200c\u200e\u200f")
    return DIR_KINDS.get(first)


def parse(md: str) -> Script:
    sc = Script()
    cur = None
    mode = None            # None (preamble) | 'guide' | 'spoken'
    buf: list = []

    def flush():
        nonlocal buf
        if not buf:
            return
        if cur is not None and mode == "spoken":
            if len(buf) > 1:
                sc.notes.append(f"{len(buf)} consecutive text lines were joined into one paragraph: {buf[0][:40]}...")
            cur.blocks.append(Block("p", " ".join(buf)))
        else:
            sc.notes.append(f"text outside a spoken section was skipped: {buf[0][:50]}")
        buf = []

    for raw in md.splitlines():
        line = raw.strip()
        if not line or line.startswith("<div") or line == "</div>" or line == "---":
            flush()
            continue
        if line.startswith("### "):
            flush()
            head = clean_heading(line[4:])
            if cur is None:
                sc.notes.append(f"### heading before the first act skipped: {head}")
            elif GUIDE_LABEL in head:
                mode = "guide"
            elif SPOKEN_LABEL_WORD in head:
                mode = "spoken"
                cur.spoken_heading = head
            else:
                mode = None
                sc.notes.append(f"unknown ### heading skipped: {head}")
            continue
        if line.startswith("## "):
            flush()
            cur = Act(title=line[3:].strip())
            sc.acts.append(cur)
            mode = None
            continue
        if line.startswith("# "):
            flush()
            sc.h1 = line[2:].strip()
            continue
        m = DIR_RE.match(line)
        if m:
            flush()
            text = m.group(1).strip()
            kind = direction_kind(text)
            if cur is None or mode != "spoken":
                sc.notes.append(f"stage direction outside a spoken section skipped: {text[:40]}")
                continue
            if kind is None:
                sc.notes.append(f"stage direction of unknown kind, rendered like a click: {text[:40]}")
                kind = "click"
            cur.blocks.append(Block(kind, text))
            continue
        m = BULLET_RE.match(line)
        if m:
            flush()
            label, text = m.group(1).strip(), m.group(2).strip()
            if text.startswith(":") and not label.endswith(":"):
                label, text = label + ":", text[1:].strip()
            if cur is None:
                sc.meta.append((label.rstrip(":").strip(), text))
            elif mode == "guide":
                cur.guide.append((label, text))
            else:
                sc.notes.append(f"bullet outside the guide section skipped: {label}")
            continue
        buf.append(line)
    flush()
    return sc


def count_kinds(sc: Script) -> dict:
    c = {"slide": 0, "click": 0, "crowd": 0}
    for a in sc.acts:
        for b in a.blocks:
            if b.kind in c:
                c[b.kind] += 1
    return c


_FA2EN = {0x06F0 + i: 0x30 + i for i in range(10)}
_EN2FA = {0x30 + i: 0x06F0 + i for i in range(10)}


def fa_to_int(s: str):
    m = re.search(r"[0-9\u06f0-\u06f9]+", s)
    return int(m.group(0).translate(_FA2EN)) if m else None


def to_fa(n: int) -> str:
    return str(n).translate(_EN2FA)


def lint(sc: Script) -> list:
    """Cross-checks that protect the speaker from typos: slide numbering, guide vs. directions, timing table ranges."""
    w = []
    k = count_kinds(sc)
    for kind, want in EXPECTED.items():
        if k[kind] != want:
            w.append(f"{kind} directions: found {k[kind]}, expected {want}")
    expect_next = 1
    ranges = []
    for a in sc.acts:
        nums = [fa_to_int(b.text) for b in a.blocks if b.kind == "slide"]
        g = [fa_to_int(lbl) for lbl, _ in a.guide]
        if nums != g:
            w.append(f"{a.title[:25]}: slide numbers in directions {nums} differ from the guide {g}")
        for n in nums:
            if n != expect_next:
                w.append(f"{a.title[:25]}: slide {n} found where {expect_next} was expected")
            expect_next = (n or 0) + 1
        ranges.append((nums[0], nums[-1]) if nums else None)
    if len(ranges) == len(TIMING_ROWS):
        for r, row in zip(ranges, TIMING_ROWS):
            want = f"{to_fa(r[0])} تا {to_fa(r[1])}" if r else ""
            if row[2] != want:
                w.append(f"timing table slides cell '{row[2]}' != '{want}' computed from the script ({row[0]})")
    else:
        w.append(f"{len(ranges)} acts in the script but {len(TIMING_ROWS)} rows in the timing table")
    # consecutive directions of the same kind would merge into one box in Word
    for a in sc.acts:
        for x, y in zip(a.blocks, a.blocks[1:]):
            if x.kind == y.kind and x.kind != "p":
                w.append(f"{a.title[:25]}: two '{x.kind}' directions in a row (Word would merge their boxes)")
    return w


# --------------------------------------------------------------------------------------------------------------------
# XML helpers
# --------------------------------------------------------------------------------------------------------------------
@dataclass(frozen=True)
class Fmt:
    size: float
    color: str = INK
    font: str = FONT
    bold: bool = False          # whole run in the bold family
    italic: bool = False
    spacing: int = 0            # character spacing, twentieths of a point


def _hp(size: float) -> int:
    return int(round(size * 2))


def rpr_xml(font, size, color, flag_bold=False, italic=False, highlight=None, shade=None, spacing=0, rtl=True) -> str:
    x = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>'
    if flag_bold:
        x += "<w:b/><w:bCs/>"
    if italic:
        x += "<w:i/><w:iCs/>"
    x += f'<w:color w:val="{color}"/>'
    if spacing:
        x += f'<w:spacing w:val="{spacing}"/>'
    x += f'<w:sz w:val="{_hp(size)}"/><w:szCs w:val="{_hp(size)}"/>'
    if highlight:
        x += f'<w:highlight w:val="{highlight}"/>'
    if shade:
        x += f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>'
    if rtl:
        x += "<w:rtl/>"
    return x


def run_xml(text: str, **kw) -> str:
    inner = ""
    for i, part in enumerate(text.split("\t")):
        if i:
            inner += "<w:tab/>"
        if part:
            inner += f'<w:t xml:space="preserve">{escape(part)}</w:t>'
    return f"<w:r><w:rPr>{rpr_xml(**kw)}</w:rPr>{inner}</w:r>"


INLINE_RE = re.compile(r"(\*\*.+?\*\*|`[^`]+`|\[؟ علی:[^\]]*\])")


def parse_inline(s: str):
    out = []
    for part in INLINE_RE.split(s):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            out.append((part[2:-2], "bold"))
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            out.append((part[1:-1], "code"))
        elif part.startswith("[؟ علی:"):
            out.append((part, "placeholder"))
        else:
            out.append((part, "plain"))
    return out


# A "Latin island" is a stretch of Latin words/digits plus the neutral characters BETWEEN them ("Unreleased: ISR",
# "MAJOR، MINOR، PATCH", "slides/isr_demo.mp4", "R&D"). It starts and ends with an ASCII letter or digit, so neutrals at
# its edges ("." after a word, brackets, the Persian comma before a Persian word) stay with the surrounding Persian run.
#
# Why islands are emitted WITHOUT <w:rtl/> (verified in Word 16): on a run flagged rtl, Word forces the neutral characters
# inside the run to RTL, which swaps ": " into " :" in "Unreleased: ISR", and it shapes ASCII digits with the Persian
# bidi language ("n8n" -> "n۸n", "mp4" -> "mp۴"). A plain run follows the normal Unicode bidi rules and keeps the digits.
#
# Inside an island any character may appear except Arabic-script letters/marks/digits, ZWNJ/ZWJ and line breaks: that is
# exactly what the Unicode bidi algorithm does with neutrals between two Latin letters (rule N1: they take the Latin
# direction), so "Trigger ≠ Root cause" keeps its reading order. The Arabic comma U+060C is a neutral separator
# (bidi class CS) and therefore allowed inside; the Arabic semicolon/question mark are strong RTL and are not.
# Guillemets are the exception: they delimit separate quoted items that continue in Persian, so they stay outside islands
# (otherwise the opening « of «CHANGELOG ... » would land on the wrong side of the quoted Persian text).
LATIN_ISLAND_RE = re.compile(
    r"[A-Za-z0-9](?:[^؀-؋؍-ۿݐ-ݿﭐ-﷿ﹰ-﻿‌‍«»\r\n]*"
    r"[A-Za-z0-9])?")
ISLANDS = True          # tests flip this to compare against the all-rtl behaviour


def split_islands(text: str):
    """Yield (chunk, is_island) covering the whole text."""
    if not ISLANDS:
        yield text, False
        return
    pos = 0
    for m in LATIN_ISLAND_RE.finditer(text):
        if m.start() > pos:
            yield text[pos:m.start()], False
        yield m.group(0), True
        pos = m.end()
    if pos < len(text):
        yield text[pos:], False


def build_runs(text: str, fmt: Fmt, inline: bool = True) -> str:
    """Text -> runs. Handles **bold**, `code` (plain text), the [؟ علی: ...] placeholder (highlighted), keeps Latin
    islands out of the rtl flag, and gives characters missing from IRANYekanX their own explicitly-fonted runs."""
    segs = parse_inline(text) if inline else [(text, "plain")]
    xml = ""
    for seg, kind in segs:
        use_bold = fmt.bold or kind in ("bold", "placeholder")
        font = FONT_B if use_bold else fmt.font
        hl = "yellow" if kind == "placeholder" else None
        shade = PLACEHOLDER_FILL if kind == "placeholder" else None
        color = "000000" if kind == "placeholder" else fmt.color
        for part, island in split_islands(seg):
            rtl = not island
            chunk = ""

            def emit_chunk():
                nonlocal xml, chunk
                if chunk:
                    xml += run_xml(chunk, font=font, size=fmt.size, color=color, italic=fmt.italic,
                                   highlight=hl, shade=shade, spacing=fmt.spacing, rtl=rtl)
                    chunk = ""

            for ch in part:
                if ch in FALLBACK_CHARS:
                    emit_chunk()
                    if ch in SYMBOL_CHARS:
                        xml += run_xml(ch, font=FONT_SYM, size=fmt.size, color=color, highlight=hl, shade=shade,
                                       rtl=rtl)
                    else:
                        xml += run_xml(ch, font=FONT_LAT, size=fmt.size, color=color, flag_bold=use_bold,
                                       italic=fmt.italic, highlight=hl, shade=shade, spacing=fmt.spacing, rtl=rtl)
                else:
                    chunk += ch
            emit_chunk()
    return xml


def pbdr_box(color: str, top_bottom_pt: int, left_right_pt: int, sz: int = 4) -> str:
    """Same-colour border around a shaded paragraph = padding inside the box.
    NB: w:space is in POINTS (0..31). Word reads larger numbers modulo 32, so never pass twips here."""
    assert 0 <= top_bottom_pt <= 31 and 0 <= left_right_pt <= 31
    side = lambda s, sp: f'<w:{s} w:val="single" w:sz="{sz}" w:space="{sp}" w:color="{color}"/>'
    top_bottom, left_right = top_bottom_pt, left_right_pt
    return f"<w:pBdr>{side('top', top_bottom)}{side('left', left_right)}{side('bottom', top_bottom)}{side('right', left_right)}</w:pBdr>"


def para(runs: str, mark: Fmt, *, before=0, after=0, line=None, line_rule="auto", keep_next=False, keep_lines=False,
         page_break_before=False, outline=None, ind_start=None, ind_end=None, hanging=None, jc=None,
         borders="", shade=None) -> str:
    """One RTL paragraph. Children are emitted in any order; normalize_order() sorts them into schema order later."""
    ppr = ""
    if keep_next:
        ppr += "<w:keepNext/>"
    if keep_lines:
        ppr += "<w:keepLines/>"
    if page_break_before:
        ppr += "<w:pageBreakBefore/>"
    ppr += borders
    if shade:
        ppr += f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>'
    ppr += "<w:bidi/>"
    ln = f' w:line="{line}" w:lineRule="{line_rule}"' if line else ""
    ppr += f'<w:spacing w:before="{before}" w:after="{after}"{ln}/>'
    if ind_start is not None or ind_end is not None or hanging:
        h = f' w:hanging="{hanging}"' if hanging else ""
        # in a bidi paragraph Word reads w:left as the START edge (right side) and w:right as the END edge
        ppr += f'<w:ind w:left="{ind_start or 0}" w:right="{ind_end or 0}"{h}/>'
    if jc:
        ppr += f'<w:jc w:val="{jc}"/>'
    if outline is not None:
        ppr += f'<w:outlineLvl w:val="{outline}"/>'
    ppr += f"<w:rPr>{rpr_xml(mark.font if not mark.bold else FONT_B, mark.size, mark.color, rtl=False)}</w:rPr>"
    return f"<w:p {W}><w:pPr>{ppr}</w:pPr>{runs}</w:p>"


def tc_xml(width: int, paras: str, *, fill=None, margins=None, valign=None, borders=None) -> str:
    pr = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if borders:
        pr += f"<w:tcBorders>{borders}</w:tcBorders>"
    if fill:
        pr += f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    if margins:
        t, r, b, l = margins
        pr += (f'<w:tcMar><w:top w:w="{t}" w:type="dxa"/><w:left w:w="{l}" w:type="dxa"/>'
               f'<w:bottom w:w="{b}" w:type="dxa"/><w:right w:w="{r}" w:type="dxa"/></w:tcMar>')
    if valign:
        pr += f'<w:vAlign w:val="{valign}"/>'
    return f"<w:tc><w:tcPr>{pr}</w:tcPr>{paras}</w:tc>"


def tr_xml(cells: str, *, cant_split=True, header=False, height=None) -> str:
    pr = ""
    if cant_split:
        pr += "<w:cantSplit/>"
    if height:
        pr += f'<w:trHeight w:val="{height}" w:hRule="atLeast"/>'
    if header:
        pr += "<w:tblHeader/>"
    return f"<w:tr><w:trPr>{pr}</w:trPr>{cells}</w:tr>"


def tbl_xml(widths, rows: str, *, borders: str = "", cell_mar=(0, 108, 0, 108)) -> str:
    t, r, b, l = cell_mar
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    pr = (f'<w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:bidiVisual/>'
          + (f"<w:tblBorders>{borders}</w:tblBorders>" if borders else "")
          + '<w:tblLayout w:type="fixed"/>'
          + f'<w:tblCellMar><w:top w:w="{t}" w:type="dxa"/><w:left w:w="{l}" w:type="dxa"/>'
            f'<w:bottom w:w="{b}" w:type="dxa"/><w:right w:w="{r}" w:type="dxa"/></w:tblCellMar>'
          + '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>')
    return f"<w:tbl {W}><w:tblPr>{pr}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>"


def line_border(side: str, color: str, sz: int = 4) -> str:
    return f'<w:{side} w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'


# Schema order of the children of the property containers we build by hand (ECMA-376 Part 1, wml.xsd)
SEQ = {
    "pPr": "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs "
           "suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd "
           "snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment "
           "textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange".split(),
    "rPr": "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid "
           "vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs "
           "em lang eastAsianLayout specVanish oMath".split(),
    "tcPr": "cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark".split(),
    "tblPr": "tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd "
             "tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription".split(),
    "sectPr": "headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType "
              "pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings".split(),
}


def normalize_order(root) -> None:
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        q = etree.QName(el)
        if q.namespace != W_NS or q.localname not in SEQ:
            continue
        idx = {n: i for i, n in enumerate(SEQ[q.localname])}
        kids = list(el)
        kids.sort(key=lambda c: idx.get(etree.QName(c).localname, len(idx)) if isinstance(c.tag, str) else len(idx))
        for c in kids:
            el.remove(c)
        for c in kids:
            el.append(c)


# --------------------------------------------------------------------------------------------------------------------
# Document pieces
# --------------------------------------------------------------------------------------------------------------------
F_BODY = Fmt(13, INK)
F_TITLE_BIG = Fmt(30, "FFFFFF", bold=True, spacing=30)
F_TITLE_SUB = Fmt(14, GOLD_SOFT, bold=True)
F_TITLE_BY = Fmt(11.5, "CBD5E1")
F_META = Fmt(11, INK)
F_META_LBL = Fmt(11, NAVY, bold=True)
F_SECTION = Fmt(14, NAVY, bold=True)
F_ACT = Fmt(18, ACCENT, bold=True)
F_GUIDE_LBL = Fmt(10.5, SLATE, bold=True)
F_GUIDE = Fmt(10.5, "374151")
F_GUIDE_BOLD = Fmt(10.5, NAVY, bold=True)
F_SPOKEN_LBL = Fmt(11.5, SLATE, bold=True)
F_SLIDE = Fmt(12, "FFFFFF", bold=True)
F_CLICK = Fmt(11, "374151")
F_CROWD = Fmt(12, TEAL, italic=True)
F_CELL = Fmt(10, INK)
F_CELL_BOLD = Fmt(10, NAVY, bold=True)
F_CELL_TOTAL = Fmt(10, NAVY)
F_CELL_HEAD = Fmt(10, "FFFFFF", bold=True)
F_FOOT = Fmt(9, "6B7280")


def box_table(paras: str, fill: str, *, margins, border_color=None, width=CONTENT_W) -> str:
    borders = ""
    if border_color:
        borders = "".join(line_border(s, border_color) for s in ("top", "left", "bottom", "right"))
    row = tr_xml(tc_xml(width, paras, fill=fill, margins=margins), cant_split=True)
    return tbl_xml([width], row, borders=borders)


def header_box(script: Script) -> str:
    by = ""
    for label, value in script.meta:
        if "ارائه" in label and "دهنده" in label:
            by = f"{label}: {value}"
    by = by or f"ارائه{ZW}دهنده: علی امانی"
    # The title is all caps without descenders, so an exact line height is safe and keeps the box compact
    paras = (
        para(build_runs("BREAKING CHANGES", F_TITLE_BIG, inline=False), F_TITLE_BIG, before=40, after=0,
             line=780, line_rule="exact", jc="center")
        + para(build_runs(HEADER_SUBTITLE, F_TITLE_SUB, inline=False), F_TITLE_SUB, before=20, after=20, jc="center")
        + para(build_runs(by, F_TITLE_BY, inline=False), F_TITLE_BY, before=40, after=40, jc="center")
    )
    return box_table(paras, NAVY, margins=(220, 360, 200, 360))


def bullet_para(label: str, text: str, fmt_text: Fmt, fmt_label: Fmt, *, hanging: int, before=0, after=60, line=None,
                bullet_color=SLATE, keep_next=False) -> str:
    runs = run_xml("\u2022\t", font=FONT, size=fmt_text.size, color=bullet_color)
    runs += build_runs(label, fmt_label, inline=False)
    if text:
        runs += build_runs(" " + text, fmt_text)
    return para(runs, fmt_text, before=before, after=after, line=line, ind_start=hanging, hanging=hanging,
                keep_next=keep_next)


def meta_list(script: Script) -> str:
    out = ""
    for i, (label, value) in enumerate(script.meta):
        out += bullet_para(label + ":", value, F_META, F_META_LBL, hanging=300, before=(280 if i == 0 else 0),
                           after=50, line=288)
    return out


def timing_table() -> str:
    widths = [780, 4000, 1850, 3008]          # right-to-left: act | topic | slides | time (measured: nothing wraps but row 2)
    head_cells = "".join(
        tc_xml(w, para(build_runs(t, F_CELL_HEAD, inline=False), F_CELL_HEAD, before=30, after=30), fill=NAVY,
               valign="center", margins=(30, 100, 30, 100))
        for w, t in zip(widths, TIMING_HEAD)
    )
    rows = tr_xml(head_cells, header=True, height=380)
    body_rows = list(TIMING_ROWS) + [TIMING_TOTAL]
    for i, row in enumerate(body_rows):
        total = i == len(body_rows) - 1
        fill = TOTAL_FILL if total else (ZEBRA if i % 2 else "FFFFFF")
        cells = ""
        for j, (w, txt) in enumerate(zip(widths, row)):
            # bold glyphs are wider: the total row keeps its text cells regular so nothing wraps, only the label is bold
            fmt = F_CELL_BOLD if j == 0 else (F_CELL_TOTAL if total else F_CELL)
            runs = build_runs(txt, fmt, inline=False) if txt else ""
            cells += tc_xml(w, para(runs, fmt, before=30, after=30), fill=fill, valign="center",
                            margins=(30, 100, 30, 100),
                            borders=(line_border("top", NAVY, 8) if total else None))
        rows += tr_xml(cells, height=360)
    borders = (line_border("bottom", "CBD5E1") + line_border("insideH", "E2E8F0"))
    return tbl_xml(widths, rows, borders=borders)


def act_title_para(title: str) -> str:
    runs = build_runs(title, F_ACT, inline=False)
    return para(runs, F_ACT, before=0, after=200, line=300, keep_next=True, page_break_before=True, outline=0,
                borders=f'<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="6" w:color="{ACCENT}"/></w:pBdr>')


def guide_box(act: Act) -> str:
    paras = para(build_runs(GUIDE_LABEL, F_GUIDE_LBL, inline=False), F_GUIDE_LBL, before=0, after=90, keep_next=True)
    for label, text in act.guide:
        paras += bullet_para(label, text, F_GUIDE, F_GUIDE_BOLD, hanging=280, before=0, after=60, line=300)
    return box_table(paras, GUIDE_FILL, margins=(140, 200, 120, 200), border_color=GUIDE_LINE)


def spoken_label(text: str) -> str:
    return para(build_runs(text, F_SPOKEN_LBL, inline=False), F_SPOKEN_LBL, before=300, after=140, keep_next=True,
                borders='<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="3" w:color="CBD5E1"/></w:pBdr>')


def block_xml(b: Block) -> str:
    if b.kind == "p":
        return para(build_runs(b.text, F_BODY), F_BODY, before=0, after=190, line=360)
    if b.kind == "slide":
        runs = run_xml(MARK_SLIDE, font=FONT_SYM, size=F_SLIDE.size, color=F_SLIDE.color) \
            + build_runs("  " + b.text, F_SLIDE, inline=False)
        return para(runs, F_SLIDE, before=300, after=150, line=288, keep_next=True, keep_lines=True,
                    borders=pbdr_box(GOLD, 4, 5), shade=GOLD, ind_start=110, ind_end=110)
    if b.kind == "click":
        runs = run_xml(MARK_CLICK, font=FONT_SYM, size=F_CLICK.size, color=F_CLICK.color) \
            + build_runs("  " + b.text, F_CLICK, inline=False)
        return para(runs, F_CLICK, before=170, after=110, line=276, keep_next=True, keep_lines=True,
                    borders=pbdr_box(BAND_GREY, 3, 5), shade=BAND_GREY, ind_start=110 + 340, ind_end=110)
    # crowd
    runs = run_xml(MARK_CROWD, font=FONT_SYM, size=F_CROWD.size, color=F_CROWD.color) \
        + build_runs("  " + b.text, F_CROWD, inline=False)
    return para(runs, F_CROWD, before=110, after=150, line=300, keep_next=True, keep_lines=True, ind_start=340)


def field_runs(instr: str, shown: str, fmt: Fmt) -> str:
    rpr = rpr_xml(fmt.font, fmt.size, fmt.color)
    return (f'<w:r><w:rPr>{rpr}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t>{shown}</w:t></w:r>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:fldChar w:fldCharType="end"/></w:r>')


def footer_para() -> str:
    runs = (build_runs("BREAKING CHANGES  \u00b7  ", F_FOOT, inline=False)
            + build_runs("صفحه ", F_FOOT, inline=False) + field_runs("PAGE", "1", F_FOOT)
            + build_runs(" از ", F_FOOT, inline=False) + field_runs("NUMPAGES", "1", F_FOOT))
    return para(runs, F_FOOT, jc="center")


# --------------------------------------------------------------------------------------------------------------------
# Package-level tuning (styles, settings, section, properties)
# --------------------------------------------------------------------------------------------------------------------
def tune_styles(doc) -> None:
    styles = doc.styles.element
    dd = styles.find(qn("w:docDefaults"))
    new_dd = parse_xml(
        f'<w:docDefaults {W}><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}"/>'
        f'<w:color w:val="{INK}"/><w:sz w:val="26"/><w:szCs w:val="26"/>'
        f'<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="fa-IR"/>'
        f'</w:rPr></w:rPrDefault>'
        f'<w:pPrDefault><w:pPr><w:bidi/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        f'</w:docDefaults>')
    dd.getparent().replace(dd, new_dd)
    for st in styles.findall(qn("w:style")):
        if st.get(qn("w:styleId")) == "Normal":
            st.append(parse_xml(f'<w:pPr {W}><w:bidi/><w:spacing w:before="0" w:after="190" w:line="360" w:lineRule="auto"/></w:pPr>'))
            st.append(parse_xml(
                f'<w:rPr {W}><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}" w:cs="{FONT}"/>'
                f'<w:color w:val="{INK}"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>'))


def tune_settings(doc) -> None:
    st = doc.settings.element
    for cs in st.iter(qn("w:compatSetting")):
        if cs.get(qn("w:name")) == "compatibilityMode":
            cs.set(qn("w:val"), "15")                    # modern layout: table edges sit on the text margins
    z = st.find(qn("w:zoom"))
    if z is not None:
        z.attrib.clear()
        z.set(qn("w:percent"), "100")
    ps = st.find(qn("w:proofState"))
    if ps is not None:                                   # no red squiggles all over a rehearsal script
        ps.addprevious(parse_xml(f"<w:hideSpellingErrors {W}/>"))
        ps.addprevious(parse_xml(f"<w:hideGrammaticalErrors {W}/>"))
    tfl = st.find(qn("w:themeFontLang"))
    if tfl is not None:
        tfl.set(qn("w:bidi"), "fa-IR")


def tune_section(doc) -> None:
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Twips(PAGE_W), Twips(PAGE_H)
    sec.top_margin = sec.bottom_margin = Twips(MARGIN)
    sec.left_margin = sec.right_margin = Twips(MARGIN)
    sec.header_distance = sec.footer_distance = Twips(510)
    sec.gutter = Twips(0)
    sect_pr = sec._sectPr
    if sect_pr.find(qn("w:bidi")) is None:
        sect_pr.append(parse_xml(f"<w:bidi {W}/>"))


def tune_properties(doc, script: Script) -> None:
    cp = doc.core_properties
    cp.title = script.h1 or "BREAKING CHANGES \u2014 متن کامل ارائه"
    cp.subject = "CHANGELOG یک مهندس"
    cp.author = "علی امانی"
    cp.last_modified_by = "علی امانی"
    cp.language = "fa-IR"
    cp.comments = "Generated by build_word.py from script.md"
    now = _dt.datetime.now()
    cp.created = now
    cp.modified = now


# --------------------------------------------------------------------------------------------------------------------
# Build
# --------------------------------------------------------------------------------------------------------------------
def build(script: Script) -> bytes:
    doc = Document()
    tune_styles(doc)
    tune_settings(doc)
    tune_section(doc)
    tune_properties(doc, script)
    body = doc.element.body
    sect_pr = body.find(qn("w:sectPr"))

    def add(xml: str) -> None:
        # a fragment may hold several sibling elements (e.g. a whole bullet list): wrap, then move them in order
        wrapper = parse_xml(f"<w:wrap {W}>{xml}</w:wrap>")
        for child in list(wrapper):
            sect_pr.addprevious(child)

    # ---- page 1: header box, metadata list, timing table
    add(header_box(script))
    add(meta_list(script))
    add(para(build_runs(TIMING_TITLE, F_SECTION, inline=False), F_SECTION, before=300, after=110, keep_next=True))
    add(timing_table())

    # ---- the acts, each on a new page
    for act in script.acts:
        add(act_title_para(act.title))
        add(guide_box(act))
        add(spoken_label(act.spoken_heading or SPOKEN_LABEL_WORD))
        for b in act.blocks:
            add(block_xml(b))

    # ---- footer
    sec = doc.sections[0]
    sec.footer.is_linked_to_previous = False
    ftr = sec.footer._element
    for ch in list(ftr):
        ftr.remove(ch)
    ftr.append(parse_xml(footer_para()))

    normalize_order(doc.element)
    normalize_order(ftr)
    normalize_order(doc.styles.element)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# --------------------------------------------------------------------------------------------------------------------
# Verification (re-opens the written file)
# --------------------------------------------------------------------------------------------------------------------
def verify(path: Path) -> bool:
    d = Document(str(path))
    body = d.element.body
    ok = True

    def check(cond: bool, msg: str, detail: str = "") -> None:
        nonlocal ok
        ok = ok and cond
        print(f"  [{'ok' if cond else 'FAIL'}] {msg}{(' -> ' + detail) if (detail and not cond) else ''}")

    def ptext(p) -> str:
        out = []
        for node in p.iter():
            if node.tag == qn("w:t"):
                out.append(node.text or "")
            elif node.tag == qn("w:tab"):
                out.append("\t")
        return "".join(out)

    paras = list(body.iter(qn("w:p")))
    texts = [ptext(p) for p in paras]
    print(f"verify: {len(paras)} paragraphs in body+tables, {len(d.tables)} tables")

    for bad in ("*(", ")*", "**", "`", "\ufe0f", "\U0001f4cc", "\U0001f399"):
        hits = [t[:50] for t in texts if bad in t]
        check(not hits, f"no {ascii(bad)} in any paragraph", "; ".join(hits[:3]))
    stray = [t[:50] for t in texts if "*" in t]
    check(not stray, "no stray asterisk at all", "; ".join(stray[:3]))
    arabic_variants = [t[:50] for t in texts if "\u064a" in t or "\u0643" in t or any(0x0660 <= ord(c) <= 0x0669 for c in t)]
    check(not arabic_variants, "no Arabic yeh/kaf/Arabic-Indic digits (Persian forms only)", "; ".join(arabic_variants[:3]))

    def starts(t, m):
        return t.lstrip().startswith(m)

    n_slide = sum(starts(t, MARK_SLIDE) for t in texts)
    n_click = sum(starts(t, MARK_CLICK) for t in texts)
    n_crowd = sum(starts(t, MARK_CROWD) for t in texts)
    check(n_slide == EXPECTED["slide"], f"slide directions = {n_slide} (want {EXPECTED['slide']})")
    check(n_click == EXPECTED["click"], f"click directions = {n_click} (want {EXPECTED['click']})")
    check(n_crowd == EXPECTED["crowd"], f"audience directions = {n_crowd} (want {EXPECTED['crowd']})")
    n_ph = sum(t.count("[؟ علی:") for t in texts)
    check(n_ph == EXPECTED_PLACEHOLDERS, f"placeholder '[؟ علی:' occurs {n_ph}x (want {EXPECTED_PLACEHOLDERS})")

    # every placeholder run must be highlighted
    ph_ok = []
    for r in body.iter(qn("w:r")):
        t = "".join((x.text or "") for x in r.findall(qn("w:t")))
        if "[؟ علی:" in t:
            rpr = r.find(qn("w:rPr"))
            ph_ok.append(rpr is not None and (rpr.find(qn("w:highlight")) is not None or rpr.find(qn("w:shd")) is not None))
    check(bool(ph_ok) and all(ph_ok), "placeholder runs are highlighted")

    # direction paragraphs carry paragraph shading (slide/click)
    shaded = {MARK_SLIDE: 0, MARK_CLICK: 0}
    for p, t in zip(paras, texts):
        for m in shaded:
            if starts(t, m):
                ppr = p.find(qn("w:pPr"))
                if ppr is not None and ppr.find(qn("w:shd")) is not None:
                    shaded[m] += 1
    check(shaded[MARK_SLIDE] == n_slide and shaded[MARK_CLICK] == n_click, "slide + click directions have paragraph shading")

    # RTL plumbing
    footer_runs = list(d.sections[0].footer._element.iter(qn("w:r")))
    body_runs = list(body.iter(qn("w:r")))
    runs = body_runs + footer_runs
    # same character classes as the island rule (the Arabic comma U+060C is a neutral and may live inside an island)
    arabic = re.compile("[؀-؋؍-ۿ‌ݐ-ݿﭐ-﷿ﹰ-﻿]")
    no_fonts = 0
    persian_without_rtl = []
    latin_with_rtl = []
    n_rtl = n_island = 0
    for r in runs:
        rpr = r.find(qn("w:rPr"))
        rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
        if rf is None or not all(rf.get(qn(a)) for a in ("w:ascii", "w:hAnsi", "w:cs")):
            no_fonts += 1
        has_rtl = rpr is not None and rpr.find(qn("w:rtl")) is not None
        text = "".join((t.text or "") for t in r.findall(qn("w:t")))
        n_rtl += has_rtl
        n_island += (not has_rtl) and bool(text)
        if arabic.search(text) and not has_rtl:
            persian_without_rtl.append(text[:30])
        if has_rtl and re.search("[A-Za-z0-9]", text) and r in body_runs:
            latin_with_rtl.append(text[:30])
    check(no_fonts == 0, f"all {len(runs)} runs set ascii/hAnsi/cs fonts", f"{no_fonts} runs without")
    check(not persian_without_rtl, f"every Persian-script run carries w:rtl ({n_rtl} rtl runs)", "; ".join(persian_without_rtl[:3]))
    check(not latin_with_rtl, f"no w:rtl run contains an ASCII letter/digit ({n_island} Latin-island runs without rtl)",
          "; ".join(latin_with_rtl[:3]))
    no_bidi = sum(1 for p in paras if p.find(qn("w:pPr")) is None or p.find(qn("w:pPr")).find(qn("w:bidi")) is None)
    check(no_bidi == 0, "all paragraphs carry w:bidi", f"{no_bidi} without")
    check(all(t.find(qn("w:tblPr")).find(qn("w:bidiVisual")) is not None for t in body.iter(qn("w:tbl"))), "all tables carry w:bidiVisual")
    check(d.sections[0]._sectPr.find(qn("w:bidi")) is not None, "section carries w:bidi")
    jc_bad = [j.get(qn("w:val")) for j in body.iter(qn("w:jc")) if j.get(qn("w:val")) not in ("center",)]
    check(not jc_bad, "no jc other than center (jc right/left is flipped in bidi paragraphs)", str(set(jc_bad)))
    fonts_used = sorted({rf.get(qn("w:ascii")) for rf in body.iter(qn("w:rFonts"))})
    print(f"  fonts used in the body: {fonts_used}")

    # schema order of hand-built containers
    wrong = 0
    for root in (body, d.sections[0].footer._element):
        for el in root.iter():
            if isinstance(el.tag, str) and etree.QName(el).namespace == W_NS and etree.QName(el).localname in SEQ:
                idx = {n: i for i, n in enumerate(SEQ[etree.QName(el).localname])}
                seq = [idx.get(etree.QName(c).localname, 999) for c in el if isinstance(c.tag, str)]
                if seq != sorted(seq):
                    wrong += 1
    check(wrong == 0, "pPr/rPr/tcPr/tblPr/sectPr children are in schema order", f"{wrong} containers out of order")
    return ok


# --------------------------------------------------------------------------------------------------------------------
def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    raw = SRC.read_bytes()
    print(f"source: {SRC.name}  {len(raw):,} bytes  md5={hashlib.md5(raw).hexdigest()}")
    script = parse(raw.decode("utf-8-sig"))
    k = count_kinds(script)
    print(f"parsed {len(script.acts)} acts, {len(script.meta)} metadata bullets; directions: "
          f"slide={k['slide']} click={k['click']} crowd={k['crowd']}; "
          f"guide bullets={sum(len(a.guide) for a in script.acts)}; "
          f"spoken paragraphs={sum(1 for a in script.acts for b in a.blocks if b.kind == 'p')}")
    for n in script.notes:
        print("  parser note:", n)
    for w in lint(script):
        print("  lint warning:", w)

    data = build(script)
    lock = OUT.with_name("~$" + OUT.name[2:])
    try:
        with open(OUT, "wb") as f:
            f.write(data)
    except PermissionError:
        print(f"ERROR: {OUT} is locked (probably open in Word). Close it and run again; nothing was written.")
        return 2
    print(f"wrote {OUT} ({len(data):,} bytes)")
    if lock.exists():
        print(f"  note: Word lock file {lock.name} exists next to the output (document may be open).")
    return 0 if verify(OUT) else 1


if __name__ == "__main__":
    sys.exit(main())
