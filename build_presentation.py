"""
TechTalk Presentation Builder Engine
Generates interactive, clickable PowerPoint presentations based on slides_config.json.
Supports adding new slides, modifying the career roadmap, adjusting hotspots,
and handling open file locks cleanly.
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
from pptx.enum.text import PP_ALIGN
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
            errors.append(f"Slide '{s_id}' has no 'file' specified.")
        else:
            # Check relative to base_dir or slides_dir
            path_in_slides = os.path.join(slides_dir, img_name)
            path_in_base = os.path.join(base_dir, img_name)
            if not (os.path.exists(path_in_slides) or os.path.exists(path_in_base)):
                errors.append(f"Slide '{s_id}': Image '{img_name}' not found in '{slides_dir}' or '{base_dir}'.")
                
    # Validate roadmap hotspots
    roadmap = config.get("roadmap_settings", {})
    roadmap_slide_id = roadmap.get("roadmap_slide_id")
    if roadmap_slide_id and roadmap_slide_id not in slide_ids:
        errors.append(f"Roadmap slide id '{roadmap_slide_id}' not found in slides list.")
        
    for hs in roadmap.get("hotspots", []):
        target = hs.get("target_slide_id")
        if target not in slide_ids:
            errors.append(f"Hotspot '{hs.get('name')}' targets nonexistent slide id: '{target}'.")

    quick_target = roadmap.get("quick_jump", {}).get("target_slide_id")
    if quick_target and quick_target not in slide_ids:
        errors.append(f"Roadmap quick jump targets nonexistent slide id: '{quick_target}'.")

    if errors:
        print("❌ Configuration validation failed with errors:")
        for err in errors:
            print(f"   • {err}")
        return False
    
    print(f"✅ Configuration valid: {len(slides)} slides defined, {len(roadmap.get('hotspots', []))} roadmap hotspots verified.")
    return True


def style_rtl_text(shape, text, font_size_pt=12, is_bold=True, text_color=RGBColor(255, 255, 255), align=PP_ALIGN.CENTER):
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = 'Vazirmatn'
    p.font.size = Pt(font_size_pt)
    p.font.bold = is_bold
    p.font.color.rgb = text_color
    p.alignment = align
    pPr = p._p.get_or_add_pPr()
    pPr.set('rtl', '1')


def add_button(slide, left, top, width, height, text, target_slide, font_size_pt=12, bg_rgb=(15, 23, 42), border_rgb=(6, 182, 212)):
    btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    btn.fill.solid()
    btn.fill.fore_color.rgb = RGBColor(*bg_rgb)
    btn.line.color.rgb = RGBColor(*border_rgb)
    btn.line.width = Pt(1.5)
    style_rtl_text(btn, text, font_size_pt=font_size_pt, is_bold=True, text_color=RGBColor(255, 255, 255), align=PP_ALIGN.CENTER)
    btn.click_action.target_slide = target_slide
    return btn


def add_transparent_hotspot(slide, left, top, width, height, target_slide):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    spPr = shape._element.spPr
    for child in list(spPr):
        if child.tag.endswith('solidFill') or child.tag.endswith('ln'):
            spPr.remove(child)
            
    # Add near-zero alpha fill (catches clicks reliably in presentation mode)
    fill_xml = parse_xml(
        '<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:srgbClr val="FFFFFF"><a:alpha val="100"/></a:srgbClr>'
        '</a:solidFill>'
    )
    spPr.append(fill_xml)
    ln_xml = parse_xml(
        '<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:noFill/>'
        '</a:ln>'
    )
    spPr.append(ln_xml)
    shape.click_action.target_slide = target_slide
    return shape


def build_presentation(config_path, output_override=None):
    base_dir = os.path.dirname(os.path.abspath(config_path))
    config = load_config(config_path)
    
    if not validate_config(config, base_dir):
        sys.exit(1)
        
    pres_meta = config.get("presentation", {})
    slides_dir = os.path.join(base_dir, pres_meta.get("slides_dir", "slides"))
    slide_width_in = pres_meta.get("width_inch", 13.333)
    slide_height_in = pres_meta.get("height_inch", 7.5)
    
    prs = Presentation()
    prs.slide_width = Inches(slide_width_in)
    prs.slide_height = Inches(slide_height_in)
    blank_layout = prs.slide_layouts[6]
    
    slides_cfg = config.get("slides", [])
    roadmap_cfg = config.get("roadmap_settings", {})
    
    # Mapping table: slide_id -> pptx_slide_object
    id_to_slide = {}
    id_to_idx = {}
    
    print("\n📦 Generating slides and transitions...")
    # Step 1: Create all slides with images, notes, transitions
    for idx, scfg in enumerate(slides_cfg):
        s_id = scfg["id"]
        img_name = scfg["file"]
        
        img_path = os.path.join(slides_dir, img_name)
        if not os.path.exists(img_path):
            img_path = os.path.join(base_dir, img_name)
            
        slide = prs.slides.add_slide(blank_layout)
        id_to_slide[s_id] = slide
        id_to_idx[s_id] = idx
        
        # Add background image 16:9 full-screen
        slide.shapes.add_picture(img_path, Inches(0), Inches(0), Inches(slide_width_in), Inches(slide_height_in))
        
        # Add smooth fade transition
        trans_xml = '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:fade/></p:transition>'
        slide._element.append(parse_xml(trans_xml))
        
        # Add speaker notes
        notes_frame = slide.notes_slide.notes_text_frame
        notes_frame.text = f"{scfg.get('title', '')}\n\n{scfg.get('notes', '')}"
        
        print(f"  [{idx+1}/{len(slides_cfg)}] Slide '{s_id}' created with '{img_name}'.")

    # Step 2: Configure Roadmap Slide
    roadmap_id = roadmap_cfg.get("roadmap_slide_id")
    roadmap_slide = id_to_slide.get(roadmap_id)
    
    if roadmap_slide:
        print("\n🗺️ Adding roadmap clickable hotspots and shortcuts...")
        for hs in roadmap_cfg.get("hotspots", []):
            target_id = hs.get("target_slide_id")
            target_slide = id_to_slide.get(target_id)
            if not target_slide:
                continue
                
            box_px = hs.get("box_px")
            ref_w, ref_h = hs.get("ref_image_size", [1376, 768])
            
            x_in = Inches((box_px[0] / ref_w) * slide_width_in)
            y_in = Inches((box_px[1] / ref_h) * slide_height_in)
            w_in = Inches((box_px[2] / ref_w) * slide_width_in)
            h_in = Inches((box_px[3] / ref_h) * slide_height_in)
            
            add_transparent_hotspot(roadmap_slide, x_in, y_in, w_in, h_in, target_slide)
            print(f"  • Hotspot '{hs.get('name')}' -> '{target_id}' (Slide #{id_to_idx[target_id]+1})")
            
        # Quick jump button
        qj = roadmap_cfg.get("quick_jump")
        if qj and qj.get("target_slide_id") in id_to_slide:
            q_target = id_to_slide[qj["target_slide_id"]]
            add_button(
                roadmap_slide,
                left=Inches(0.5),
                top=Inches(0.35),
                width=Inches(3.4),
                height=Inches(0.52),
                text=qj.get("text", "☄️ پرده پایانی ◀"),
                target_slide=q_target,
                font_size_pt=qj.get("font_size_pt", 11.5),
                bg_rgb=qj.get("bg_color", [26, 16, 37]),
                border_rgb=qj.get("border_color", [236, 72, 153])
            )
            
        # Interactive hint banner
        hint = roadmap_cfg.get("hint")
        if hint:
            hint_box = roadmap_slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(3.8),
                Inches(6.85),
                Inches(5.7),
                Inches(0.45)
            )
            hint_box.fill.solid()
            hint_box.fill.fore_color.rgb = RGBColor(*hint.get("bg_color", [10, 15, 28]))
            hint_box.line.color.rgb = RGBColor(*hint.get("border_color", [30, 58, 95]))
            hint_box.line.width = Pt(1)
            style_rtl_text(
                hint_box,
                hint.get("text", ""),
                font_size_pt=hint.get("font_size_pt", 11),
                is_bold=False,
                text_color=RGBColor(*hint.get("text_color", [148, 163, 184])),
                align=PP_ALIGN.CENTER
            )

    # Step 3: Add Navigation Buttons and Station Badges to slides
    print("\n🧭 Adding dynamic badges and navigation buttons...")
    for idx, scfg in enumerate(slides_cfg):
        s_id = scfg["id"]
        slide = id_to_slide[s_id]
        
        # 1. Back to Roadmap button
        if scfg.get("back_to_roadmap", False) and roadmap_slide:
            add_button(
                slide,
                left=Inches(0.5),
                top=Inches(0.35),
                width=Inches(3.1),
                height=Inches(0.52),
                text="🗺️ بازگشت به بزرگراه مسیر شغلی",
                target_slide=roadmap_slide,
                font_size_pt=12,
                bg_rgb=[15, 23, 42],
                border_rgb=[6, 182, 212]
            )
            
        # 2. Station Badge
        badge_text = scfg.get("badge")
        if badge_text:
            badge = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(4.5),
                Inches(0.35),
                Inches(8.3),
                Inches(0.52)
            )
            badge.fill.solid()
            badge.fill.fore_color.rgb = RGBColor(10, 15, 30)
            badge.line.color.rgb = RGBColor(56, 189, 248)
            badge.line.width = Pt(1.5)
            style_rtl_text(
                badge,
                badge_text,
                font_size_pt=12,
                is_bold=True,
                text_color=RGBColor(241, 245, 249),
                align=PP_ALIGN.CENTER
            )

        # 3. Next / Previous Buttons
        if scfg.get("nav_buttons", False):
            # Previous slide button
            if idx > 0 and slides_cfg[idx - 1].get("type") in ["station", "conclusion"]:
                prev_slide = id_to_slide[slides_cfg[idx - 1]["id"]]
                add_button(
                    slide,
                    left=Inches(10.6),
                    top=Inches(6.8),
                    width=Inches(2.2),
                    height=Inches(0.45),
                    text="▶ ایستگاه قبلی",
                    target_slide=prev_slide,
                    font_size_pt=11,
                    bg_rgb=[15, 23, 42],
                    border_rgb=[100, 116, 139]
                )
                
            # Next slide button
            if idx < len(slides_cfg) - 1:
                next_slide = id_to_slide[slides_cfg[idx + 1]["id"]]
                next_text = scfg.get("next_text") or "ایستگاه بعدی ◀"
                next_border = [236, 72, 153] if "☄️" in next_text or "🌍" in next_text else [6, 182, 212]
                add_button(
                    slide,
                    left=Inches(0.5),
                    top=Inches(6.8),
                    width=Inches(3.1 if len(next_text) > 16 else 2.2),
                    height=Inches(0.45),
                    text=next_text,
                    target_slide=next_slide,
                    font_size_pt=11,
                    bg_rgb=[15, 23, 42],
                    border_rgb=next_border
                )

    # Step 4: Save PowerPoint presentation safely
    print("\n💾 Saving presentations...")
    primary_out = output_override or pres_meta.get("output_primary", "TechTalk_Dinosaurs_Survival_Interactive.pptx")
    primary_path = os.path.join(base_dir, primary_out)
    fallback_out = pres_meta.get("output_fallback", "TechTalk_Dinosaurs_Survival.pptx")
    fallback_path = os.path.join(base_dir, fallback_out)
    
    saved_paths = []
    
    def try_save(target_path, desc):
        try:
            prs.save(target_path)
            print(f"  ✅ Saved {desc}: {target_path}")
            saved_paths.append(target_path)
            return True
        except PermissionError:
            print(f"  ⚠️ Note: '{os.path.basename(target_path)}' is currently open in PowerPoint (file locked).")
            return False
        except Exception as e:
            print(f"  ❌ Error saving '{target_path}': {e}")
            return False

    # Attempt primary
    saved_primary = try_save(primary_path, "Primary Deck")
    
    # Attempt fallback
    try_save(fallback_path, "Fallback Deck")
    
    # If primary was locked, save timestamped copy so work is never lost
    if not saved_primary:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt_path = os.path.join(base_dir, f"TechTalk_Dinosaurs_Survival_{ts}.pptx")
        try_save(alt_path, "Timestamped Alternative Deck")
        
    print("\n🎉 Presentation build completed successfully!")


def main():
    parser = argparse.ArgumentParser(description="Build interactive TechTalk PowerPoint decks")
    parser.add_argument("--config", default=r"d:\work\erae-ali\slides_config.json", help="Path to slides_config.json")
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
