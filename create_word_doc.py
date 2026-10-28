import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

def make_p_rtl(p):
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = p._p.get_or_add_pPr()
    # Check if bidi already exists
    if pPr.find(qn('w:bidi')) is None:
        pPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
    # Set text alignment to right in pPr as well
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="right"/>'))
    else:
        jc.set(qn('w:val'), 'right')

def make_run_rtl(run, font_name="Tahoma", font_size=11, bold=False, italic=False, color_rgb=(40, 44, 52)):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)
    
    rPr = run._r.get_or_add_rPr()
    if rPr.find(qn('w:rtl')) is None:
        rPr.append(parse_xml(f'<w:rtl {nsdecls("w")}/>'))
    
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>')
        rPr.append(rFonts)
    else:
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'), font_name)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def make_table_rtl(tbl):
    tblPr = tbl._tbl.tblPr
    if tblPr.find(qn('w:bidiVisual')) is None:
        tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))

doc = Document()

# Set Section to RTL
for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    sectPr = section._sectPr
    if sectPr.find(qn('w:bidi')) is None:
        sectPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))

# Normal Style default to RTL
style = doc.styles['Normal']
style.font.name = 'Tahoma'
style.font.size = Pt(11)
style_pPr = style._element.get_or_add_pPr()
if style_pPr.find(qn('w:bidi')) is None:
    style_pPr.append(parse_xml(f'<w:bidi {nsdecls("w")}/>'))
if style_pPr.find(qn('w:jc')) is None:
    style_pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="right"/>'))

# 1. Header Box
header_tbl = doc.add_table(rows=1, cols=1)
make_table_rtl(header_tbl)
header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
header_tbl.autofit = False
header_tbl.columns[0].width = Inches(6.9)

cell = header_tbl.cell(0, 0)
set_cell_background(cell, "0F172A")
set_cell_margins(cell, top=260, bottom=260, left=260, right=260)

hp1 = cell.paragraphs[0]
make_p_rtl(hp1)
hp1.alignment = WD_ALIGN_PARAGRAPH.CENTER
hp1.paragraph_format.space_before = Pt(6)
hp1.paragraph_format.space_after = Pt(4)
r = hp1.add_run("سناریو و متن کامل تک‌تاک")
make_run_rtl(r, font_size=13, bold=True, color_rgb=(56, 189, 248))

hp2 = cell.add_paragraph()
make_p_rtl(hp2)
hp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
hp2.paragraph_format.space_before = Pt(4)
hp2.paragraph_format.space_after = Pt(8)
r2 = hp2.add_run("«دایناسورهای دنیای تک و مهارتِ بقا»")
make_run_rtl(r2, font_size=19, bold=True, color_rgb=(255, 255, 255))

hp3 = cell.add_paragraph()
make_p_rtl(hp3)
hp3.alignment = WD_ALIGN_PARAGRAPH.CENTER
hp3.paragraph_format.space_before = Pt(2)
hp3.paragraph_format.space_after = Pt(6)
r3 = hp3.add_run("ارائه‌دهنده: علی  |  مدت زمان: ۳۵ تا ۴۰ دقیقه  |  نسخه نهایی بازنویسی‌شده")
make_run_rtl(r3, font_size=10, bold=False, color_rgb=(148, 163, 184))

# Space
p_space = doc.add_paragraph()
make_p_rtl(p_space)
p_space.paragraph_format.space_after = Pt(10)

# 2. Timeline Box
info_tbl = doc.add_table(rows=1, cols=1)
make_table_rtl(info_tbl)
info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
info_tbl.columns[0].width = Inches(6.9)
icell = info_tbl.cell(0, 0)
set_cell_background(icell, "F8FAFC")
set_cell_margins(icell, top=140, bottom=140, left=180, right=180)

ip = icell.paragraphs[0]
make_p_rtl(ip)
ir = ip.add_run("📌 ساختار روایی و توالی تاریخی (تایم‌لاین خطی):")
make_run_rtl(ir, font_size=11, bold=True, color_rgb=(15, 23, 42))

timeline_points = [
    "۱. افتتاحیه: استعاره دایناسورها و پستانداران چابک + معرفی رسمی «دوقلوهای ترس».",
    "۲. سال ۱۳۸۸ تا ۱۳۹۴: ۶ سال برنامه‌نویسی ناب جاوا ➔ تصمیم جسورانه ورود به عملیات در سال ۹۴ و خلق ابزار Query Executor.",
    "۳. سال ۱۳۹۵ تا ۱۳۹۷: میدان جنگ ابزارها (وب‌لاجیک در اسکیل Enterprise ➔ حلقه For مرگبار ➔ اتوماسیون با Jenkins و تجمیع لاگ در ELK).",
    "۴. سال ۱۳۹۸ تا ۱۴۰۰: تشکیل تیم تحقیق و توسعه (R&D Ops) ➔ چرخه ۵ مرحله‌ای از نیاز تا تحویل ➔ قانون طلایی: «حق نداری سرویسی که آوردی بالا رو نگه داری؛ تحویل بده و رها کن!»",
    "۵. نقطه اوج (سال ۱۴۰۰ - ساعت ۰۶:۰۰ صبح بانک): بحران ۳ شبانه‌روز بی‌خوابی ➔ امتحان نهایی ۱۶ ساله ➔ مهار دوقلوهای ترس و نجات پروداکشن.",
    "۶. زمان حال و آینده (۱۴۰۳ به بعد): معماری هوشمند ISR ➔ شهاب‌سنگ عظیم AI و دو رویکرد مهندسان ➔ پیام نهایی بقا."
]

for tp in timeline_points:
    tip = icell.add_paragraph()
    make_p_rtl(tip)
    tip.paragraph_format.space_before = Pt(2)
    tip.paragraph_format.space_after = Pt(2)
    tip.paragraph_format.line_spacing = 1.3
    tr = tip.add_run(tp)
    make_run_rtl(tr, font_size=10, bold=False, color_rgb=(51, 65, 85))

# Space
p_space2 = doc.add_paragraph()
make_p_rtl(p_space2)
p_space2.paragraph_format.space_after = Pt(12)

# Load Script Lines
with open(r"d:\work\erae-ali\techtalk_script.md", "r", encoding="utf-8") as f:
    lines = f.readlines()

for line in lines:
    raw = line.strip()
    if not raw or raw.startswith("<div") or raw.startswith("</div>"):
        continue
    if raw.startswith("# سناریو") or raw.startswith("- **ارائه‌دهنده"):
        continue
    if raw.startswith("## ساختار و توالی"):
        continue
    if raw.startswith("1. **افتتاحیه") or raw.startswith("2. **سال ۱۳۸۸") or raw.startswith("3. **سال ۱۳۹۵") or raw.startswith("4. **سال ۱۳۹۸") or raw.startswith("5. **نقطه اوج") or raw.startswith("6. **امروز و آینده"):
        continue
    if raw == "---":
        continue

    # Main Heading (##)
    if raw.startswith("## "):
        title_text = raw.replace("## ", "")
        p = doc.add_paragraph()
        make_p_rtl(p)
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(title_text)
        make_run_rtl(r, font_size=14, bold=True, color_rgb=(14, 116, 144))
        continue

    # Sub-heading (###)
    if raw.startswith("### "):
        sub_text = raw.replace("### ", "")
        p = doc.add_paragraph()
        make_p_rtl(p)
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(sub_text)
        make_run_rtl(r, font_size=12, bold=True, color_rgb=(30, 41, 59))
        continue

    # Stage guide & Key takeaways
    if raw.startswith("* **اسلاید") or raw.startswith("* **زبان بدن") or raw.startswith("* **درس‌آموخته"):
        p = doc.add_paragraph()
        make_p_rtl(p)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(3)
        parts = raw.split("**")
        if len(parts) >= 3:
            label = parts[1]
            content = "**".join(parts[2:]).lstrip(": ")
            r1 = p.add_run(f"• {label}: ")
            make_run_rtl(r1, font_size=10.5, bold=True, color_rgb=(180, 83, 9))
            r2 = p.add_run(content)
            make_run_rtl(r2, font_size=10.5, bold=False, color_rgb=(71, 85, 105))
        else:
            r = p.add_run(raw)
            make_run_rtl(r, font_size=10.5, bold=False, color_rgb=(71, 85, 105))
        continue

    # Dialogues / Speeches / Text
    p = doc.add_paragraph()
    make_p_rtl(p)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.35
    
    clean_text = raw.replace("**", "").replace("*", "")
    is_bold = "**" in raw
    is_italic = ("*" in raw and not "**" in raw) or raw.startswith("*")
    
    r = p.add_run(clean_text)
    if is_bold:
        make_run_rtl(r, font_size=11, bold=True, color_rgb=(15, 23, 42))
    elif is_italic:
        make_run_rtl(r, font_size=10.5, italic=True, color_rgb=(100, 116, 139))
    else:
        make_run_rtl(r, font_size=11, bold=False, color_rgb=(30, 41, 59))

output_doc = r"d:\work\erae-ali\TechTalk_Full_Script.docx"
doc.save(output_doc)
print(f"RTL Document saved successfully: {output_doc}")
