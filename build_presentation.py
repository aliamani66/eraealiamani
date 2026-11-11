"""
TechTalk Presentation Builder Engine
Generates the PowerPoint presentation based on slides_config.json.
Features:
- Navigation is by clicker only: no buttons and no clickable areas on the slides.
- A progress bar of station dots at the bottom of every station slide.
- Real PowerPoint text overlays (RTL Persian, LTR code) instead of text baked into images.
- Windows file lock handling.
"""

import os
import sys
import json
import shutil
import argparse
from datetime import datetime

# Ensure safe UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml


def load_config(config_path):
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config file not found: {config_path}")
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def validate_config(config, base_dir):
    print("🔍 Validating slides configuration...")
    slides = config.get("slides", [])
    slides_dir = os.path.join(base_dir, config.get("presentation", {}).get("slides_dir", "slides"))
    
    errors = []
    slide_ids = set()
    
    for idx, slide in enumerate(slides):
        s_id = slide.get("id")
        if not s_id:
            errors.append(f"Slide #{idx+1} is missing an 'id'.")
        elif s_id in slide_ids:
            errors.append(f"Duplicate slide id: '{s_id}'.")
        else:
            slide_ids.add(s_id)
            
        img_name = slide.get("file")
        if not img_name:
            if not slide.get("background"):
                errors.append(f"Slide '{s_id}' has no 'file' and no 'background' color.")
        else:
            path_in_slides = os.path.join(slides_dir, img_name)
            path_in_base = os.path.join(base_dir, img_name)
            if not (os.path.exists(path_in_slides) or os.path.exists(path_in_base)):
                errors.append(f"Slide '{s_id}': Image '{img_name}' not found in '{slides_dir}' or '{base_dir}'.")
                
        video_name = slide.get("video")
        if video_name and not (os.path.exists(os.path.join(slides_dir, video_name)) or os.path.exists(os.path.join(base_dir, video_name))):
            print(f"  ⚠️ Slide '{s_id}': video '{video_name}' not found; the slide keeps its text overlays (add the video later and rebuild).")

        for ov_idx, ov in enumerate(slide.get("text_overlays", [])):
            if not ov.get("text"):
                errors.append(f"Slide '{s_id}': text_overlays[{ov_idx}] has no 'text'.")
            size = ov.get("size_pt", MAIN_FONT_PT)
            if size < MIN_FONT_PT:
                errors.append(f"Slide '{s_id}': text_overlays[{ov_idx}] size {size}pt is below the {MIN_FONT_PT}pt minimum.")
            try:
                overlay_box(ov.get("position", "center"), 13.333, 7.5)
            except ValueError as e:
                errors.append(f"Slide '{s_id}': text_overlays[{ov_idx}]: {e}")

    roadmap_slide_id = config.get("roadmap_settings", {}).get("roadmap_slide_id")
    if roadmap_slide_id and roadmap_slide_id not in slide_ids:
        errors.append(f"Roadmap slide id '{roadmap_slide_id}' not found in slides list.")

    station_ids = set()
    for st_idx, st in enumerate(config.get("stations", [])):
        st_id = st.get("id")
        if not st_id or not st.get("year") or not st.get("title"):
            errors.append(f"stations[{st_idx}] needs 'id', 'year' and 'title'.")
        elif st_id in station_ids:
            errors.append(f"Duplicate station id: '{st_id}'.")
        else:
            station_ids.add(st_id)

    used_stations = set()
    for slide in slides:
        st_id = slide.get("station")
        if st_id is None:
            continue
        if st_id not in station_ids:
            errors.append(f"Slide '{slide.get('id')}' uses unknown station '{st_id}'.")
        used_stations.add(st_id)
    for st_id in sorted(station_ids - used_stations):
        print(f"  ⚠️ Station '{st_id}' has no slide yet; its dot will never light up.")

    if errors:
        print("❌ Configuration validation failed with errors:")
        for err in errors:
            print(f"   • {err}")
        return False

    print(f"✅ Configuration valid: {len(slides)} slides defined, {len(station_ids)} stations.")
    return True


MIN_FONT_PT = 20
MAIN_FONT_PT = 28
DEFAULT_FONT = "Tahoma"
CODE_FONT = "Consolas"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def style_run(run, font_name, size_pt, rgb, bold=False):
    """Sets font on the run itself (PowerPoint ignores paragraph-level defaults), Latin and complex-script."""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*rgb)
    rPr = run._r.get_or_add_rPr()
    for old in rPr.findall(f"{{{A_NS}}}cs"):
        rPr.remove(old)
    cs = parse_xml(f'<a:cs xmlns:a="{A_NS}" typeface="{font_name}"/>')
    latin = rPr.find(f"{{{A_NS}}}latin")
    latin.addnext(cs)


def add_paragraph(tf, first, text, font_name, size_pt, rgb, bold=False, align=PP_ALIGN.CENTER, rtl=True, space_after=0):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    run = p.add_run()
    run.text = text
    style_run(run, font_name, size_pt, rgb, bold)
    p.alignment = align
    p.space_after = Pt(space_after)
    p._element.get_or_add_pPr().set("rtl", "1" if rtl else "0")
    return p


def set_fill_and_line(shape, fill_hex=None, alpha_pct=100, line_xml=None):
    """Replaces fill and outline of an autoshape. fill_hex=None means no fill."""
    spPr = shape._element.spPr
    for child in list(spPr):
        if child.tag.endswith("solidFill") or child.tag.endswith("noFill") or child.tag.endswith("ln"):
            spPr.remove(child)
    if fill_hex:
        spPr.append(parse_xml(
            f'<a:solidFill xmlns:a="{A_NS}"><a:srgbClr val="{fill_hex}"><a:alpha val="{int(alpha_pct * 1000)}"/></a:srgbClr></a:solidFill>'
        ))
    else:
        spPr.append(parse_xml(f'<a:noFill xmlns:a="{A_NS}"/>'))
    spPr.append(parse_xml(line_xml or f'<a:ln xmlns:a="{A_NS}"><a:noFill/></a:ln>'))


def hex_to_rgb(value):
    value = value.lstrip("#")
    return (int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def overlay_box(position, slide_width_in, slide_height_in, width_in=None, height_in=None):
    """Returns (left, top, width, height, align) in inches for a named or explicit position."""
    margin = 0.4
    if isinstance(position, (list, tuple)):
        left, top, width, height = position
        return left, top, width, height, None
    w = width_in
    h = height_in
    if position == "top-right":
        w, h = w or 3.6, h or 1.1
        return slide_width_in - margin - w, margin, w, h, PP_ALIGN.RIGHT
    if position == "top-left":
        w, h = w or 3.6, h or 1.1
        return margin, margin, w, h, PP_ALIGN.LEFT
    if position == "bottom":
        w, h = w or 11.6, h or 1.4
        return (slide_width_in - w) / 2.0, slide_height_in - margin - h, w, h, PP_ALIGN.CENTER
    if position == "center":
        w, h = w or 11.0, h or 2.6
        return (slide_width_in - w) / 2.0, (slide_height_in - h) / 2.0, w, h, PP_ALIGN.CENTER
    raise ValueError(f"Unknown overlay position: {position!r}")


def add_text_overlay(slide, ov, font_name, slide_width_in, slide_height_in):
    """Real PowerPoint text with optional translucent dark bar. RTL by default, ltr=true for code."""
    ltr = ov.get("ltr", False)
    left, top, width, height, default_align = overlay_box(
        ov.get("position", "center"), slide_width_in, slide_height_in, ov.get("width_in"), ov.get("height_in")
    )
    align = {"left": PP_ALIGN.LEFT, "right": PP_ALIGN.RIGHT, "center": PP_ALIGN.CENTER}.get(ov.get("align"), default_align)
    if align is None:
        align = PP_ALIGN.LEFT if ltr else PP_ALIGN.RIGHT

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    bar = ov.get("bar")
    if bar:
        bar = {} if bar is True else bar
        set_fill_and_line(shape, bar.get("color", "0A0F1D"), bar.get("alpha_pct", 70))
    else:
        set_fill_and_line(shape)
    shape.shadow.inherit = False

    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.3)
    tf.margin_top = tf.margin_bottom = Inches(0.15)

    face = CODE_FONT if ltr else ov.get("font", font_name)
    size = ov.get("size_pt", MAIN_FONT_PT)
    rgb = hex_to_rgb(ov.get("color", "FFFFFF"))
    lines = ov["text"].split("\n")
    for i, line in enumerate(lines):
        add_paragraph(tf, i == 0, line, face, size, rgb, ov.get("bold", False), align, rtl=not ltr)
    return shape


def add_title_overlay(slide, overlay_cfg, font_name, slide_width_in, slide_height_in):
    """Creates a high-end keynote title card with RTL Persian typography."""
    card_w = overlay_cfg.get("width_in", 10.2)
    card_h = overlay_cfg.get("height_in", 3.4)
    card_top = overlay_cfg.get("top_in", 1.05)
    card_left = (slide_width_in - card_w) / 2.0

    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(card_left),
        Inches(card_top),
        Inches(card_w),
        Inches(card_h)
    )
    # Frosted dark obsidian backdrop (80% opacity) with glowing cyan accent border
    set_fill_and_line(
        card, "0A0F1D", 80,
        f'<a:ln xmlns:a="{A_NS}" w="19050"><a:solidFill><a:srgbClr val="06B6D4"><a:alpha val="65000"/></a:srgbClr></a:solidFill></a:ln>'
    )

    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.4)
    tf.margin_right = Inches(0.4)
    tf.margin_top = Inches(0.28)
    tf.margin_bottom = Inches(0.28)

    parts = [
        ("tag", "🎙️  رویداد ارائـه تخصصـی مهندسـی  |  TECHTALK", 20, (56, 189, 248), True, 8),
        ("title", "دایناسورهای دنیای تک و مهارتِ بقا", 36, (255, 255, 255), True, 10),
        ("subtitle", "", 20, (203, 213, 225), False, 14),
        ("presenter", "ارائه‌دهنده: علی امانی", 20, (251, 191, 36), True, 0),
    ]
    first = True
    for key, default, size, rgb, bold, gap in parts:
        text = overlay_cfg.get(key, default)
        if not text:
            continue
        add_paragraph(tf, first, text, font_name, size, rgb, bold, PP_ALIGN.CENTER, space_after=gap)
        first = False

    return card


def add_progress_bar(slide, station_count, current_idx, slide_width_in, slide_height_in):
    """Row of dots, one per station, first station on the right (RTL). No text.
    Current dot is bright, earlier ones half-bright, later ones faint."""
    dot = 0.18
    gap = 0.22
    pad_x, pad_y = 0.3, 0.07
    row_w = station_count * dot + (station_count - 1) * gap
    pill_w, pill_h = row_w + 2 * pad_x, dot + 2 * pad_y
    pill_left = (slide_width_in - pill_w) / 2.0
    pill_top = slide_height_in - 0.07 - pill_h

    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(pill_left), Inches(pill_top), Inches(pill_w), Inches(pill_h))
    pill.name = "progress-bar"
    pill.adjustments[0] = 0.5
    set_fill_and_line(pill, "0A0F1D", 60)
    pill.shadow.inherit = False

    for i in range(station_count):
        left = pill_left + pill_w - pad_x - dot - i * (dot + gap)
        shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(left), Inches(pill_top + pad_y), Inches(dot), Inches(dot))
        shape.name = f"progress-dot-{i + 1}"
        alpha = 100 if i == current_idx else (50 if i < current_idx else 20)
        set_fill_and_line(shape, "06B6D4" if i <= current_idx else "FFFFFF", alpha)
        shape.shadow.inherit = False


def add_roadmap(slide, stations, font_name, slide_width_in, slide_height_in):
    """Horizontal career line, first station on the right (RTL). Each station: circle on the line,
    year above (28pt) and short title below (20pt), all real PowerPoint text."""
    count = len(stations)
    margin_x = 0.6
    cell_w = (slide_width_in - 2 * margin_x) / count
    line_y = slide_height_in / 2.0 + 0.1
    circle = 0.5
    text_h = 0.75

    line_left = margin_x + cell_w / 2.0
    line_w = slide_width_in - 2 * margin_x - cell_w
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(line_left), Inches(line_y - 0.03), Inches(line_w), Inches(0.06))
    line.name = "roadmap-line"
    set_fill_and_line(line, "06B6D4", 80)
    line.shadow.inherit = False

    for i, st in enumerate(stations):
        cell_left = slide_width_in - margin_x - (i + 1) * cell_w
        center_x = cell_left + cell_w / 2.0
        last = i == count - 1

        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(center_x - circle / 2.0), Inches(line_y - circle / 2.0), Inches(circle), Inches(circle))
        dot.name = f"roadmap-station-{st['id']}"
        set_fill_and_line(
            dot, "FBBF24" if last else "06B6D4", 100,
            f'<a:ln xmlns:a="{A_NS}" w="38100"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>'
        )
        dot.shadow.inherit = False

        for name, text, top, size, rgb, bold in (
            ("year", st["year"], line_y - circle / 2.0 - 0.15 - text_h, 28, (251, 191, 36), True),
            ("title", st["title"], line_y + circle / 2.0 + 0.15, 20, (255, 255, 255), False),
        ):
            box = slide.shapes.add_textbox(Inches(cell_left), Inches(top), Inches(cell_w), Inches(text_h))
            box.name = f"roadmap-{name}-{st['id']}"
            tf = box.text_frame
            tf.word_wrap = False  # one-word labels: overflow evenly instead of wrapping
            tf.vertical_anchor = MSO_ANCHOR.BOTTOM if name == "year" else MSO_ANCHOR.TOP
            tf.margin_left = tf.margin_right = Inches(0.02)
            tf.margin_top = tf.margin_bottom = Inches(0.02)
            ltr = st["title"].isascii() and name == "title"
            add_paragraph(tf, True, text, font_name, size, rgb, bold, PP_ALIGN.CENTER, rtl=not ltr)


def build_presentation(config_path, output_override=None):
    base_dir = os.path.dirname(os.path.abspath(config_path))
    config = load_config(config_path)
    
    if not validate_config(config, base_dir):
        sys.exit(1)
        
    pres_meta = config.get("presentation", {})
    slides_dir = os.path.join(base_dir, pres_meta.get("slides_dir", "slides"))
    slide_width_in = pres_meta.get("width_inch", 13.333)
    slide_height_in = pres_meta.get("height_inch", 7.5)
    
    font_name = pres_meta.get("font", DEFAULT_FONT)

    prs = Presentation()
    prs.slide_width = Inches(slide_width_in)
    prs.slide_height = Inches(slide_height_in)
    blank_layout = prs.slide_layouts[6]
    
    slides_cfg = config.get("slides", [])
    station_order = [st["id"] for st in config.get("stations", [])]

    print("\n📦 Generating slides and transitions...")
    # Create all slides with full-bleed 16:9 images, text, progress bar, transitions, and notes
    for idx, scfg in enumerate(slides_cfg):
        s_id = scfg["id"]
        img_name = scfg.get("file")

        slide = prs.slides.add_slide(blank_layout)

        if img_name:
            # Background image 16:9 full-screen
            img_path = os.path.join(slides_dir, img_name)
            if not os.path.exists(img_path):
                img_path = os.path.join(base_dir, img_name)
            slide.shapes.add_picture(img_path, Inches(0), Inches(0), Inches(slide_width_in), Inches(slide_height_in))
        else:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(*hex_to_rgb(scfg["background"]))
        
        # Add title overlay if specified
        if "title_overlay" in scfg:
            add_title_overlay(slide, scfg["title_overlay"], font_name, slide_width_in, slide_height_in)
            print(f"  ✨ Added elegant title card overlay on Slide '{s_id}'.")

        # Optional full-slide video (slide key "video"). If the file is missing, the text overlays stay as a placeholder.
        video_embedded = False
        video_name = scfg.get("video")
        if video_name:
            video_path = os.path.join(slides_dir, video_name)
            if not os.path.exists(video_path):
                video_path = os.path.join(base_dir, video_name)
            if os.path.exists(video_path):
                slide.shapes.add_movie(video_path, Inches(0), Inches(0), Inches(slide_width_in), Inches(slide_height_in), mime_type="video/mp4")
                video_embedded = True

        for ov in ([] if video_embedded else scfg.get("text_overlays", [])):
            add_text_overlay(slide, ov, font_name, slide_width_in, slide_height_in)

        if scfg.get("type") == "roadmap":
            add_roadmap(slide, config.get("stations", []), font_name, slide_width_in, slide_height_in)

        if scfg.get("station"):
            add_progress_bar(slide, len(station_order), station_order.index(scfg["station"]), slide_width_in, slide_height_in)

        # Smooth fade transition, unless this slide is just the next step of the previous one
        if scfg.get("transition", "fade") != "none":
            trans_xml = '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:fade/></p:transition>'
            slide._element.append(parse_xml(trans_xml))
        
        # Add speaker notes
        notes_frame = slide.notes_slide.notes_text_frame
        notes_frame.text = f"{scfg.get('title', '')}\n\n{scfg.get('notes', '')}"
        
        print(f"  [{idx+1}/{len(slides_cfg)}] Slide '{s_id}' created with '{img_name or 'solid background'}'.")

    # Save the single presentation file
    print("\n💾 Saving presentation...")
    primary_out = output_override or pres_meta.get("output_primary", "TechTalk_Dinosaurs_Survival.pptx")
    primary_path = os.path.join(base_dir, primary_out)
    try:
        prs.save(primary_path)
    except PermissionError:
        print(f"❌ Cannot write '{primary_out}': the file is locked, most likely open in PowerPoint.")
        print("   Close it and run the build again. The file was not changed.")
        sys.exit(1)
    print(f"  ✅ Saved: {primary_path}")

    print("\n🎉 Presentation build completed successfully with pure minimal layout!")


def main():
    parser = argparse.ArgumentParser(description="Build clean interactive TechTalk PowerPoint decks")
    default_config = os.path.join(os.path.dirname(os.path.abspath(__file__)), "slides_config.json")
    parser.add_argument("--config", default=default_config, help="Path to slides_config.json")
    parser.add_argument("--output", default=None, help="Custom output PPTX filename")
    parser.add_argument("--validate", action="store_true", help="Only validate configuration without writing")
    args = parser.parse_args()
    
    if args.validate:
        base_dir = os.path.dirname(os.path.abspath(args.config))
        cfg = load_config(args.config)
        sys.exit(0 if validate_config(cfg, base_dir) else 1)
        
    build_presentation(args.config, args.output)


if __name__ == "__main__":
    main()
