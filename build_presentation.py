"""
TechTalk Presentation Builder Engine
Generates clean, interactive, clickable PowerPoint presentations based on slides_config.json.
Features:
- Pure minimal iconography for navigation buttons (no verbose text clutter).
- Zero top badges to preserve 100% of the image artwork headers.
- Dynamic ID-based hyperlinking for stations and roadmaps.
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
            path_in_slides = os.path.join(slides_dir, img_name)
            path_in_base = os.path.join(base_dir, img_name)
            if not (os.path.exists(path_in_slides) or os.path.exists(path_in_base)):
                errors.append(f"Slide '{s_id}': Image '{img_name}' not found in '{slides_dir}' or '{base_dir}'.")
                
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


def add_icon_button(slide, left, top, size, icon, target_slide, font_size_pt=14, bg_rgb=(15, 23, 42), border_rgb=(6, 182, 212)):
    """Creates a minimal, sleek icon-only navigation button with no text clutter."""
    btn = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, size, size)
    btn.fill.solid()
    btn.fill.fore_color.rgb = RGBColor(*bg_rgb)
    btn.line.color.rgb = RGBColor(*border_rgb)
    btn.line.width = Pt(1.2)
    
    tf = btn.text_frame
    tf.word_wrap = False
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)
    
    p = tf.paragraphs[0]
    p.text = icon
    p.font.name = 'Segoe UI Emoji'
    p.font.size = Pt(font_size_pt)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    
    btn.click_action.target_slide = target_slide
    return btn


def add_transparent_hotspot(slide, left, top, width, height, target_slide):
    """Creates an invisible hotspot rectangle that reliably captures mouse clicks."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    spPr = shape._element.spPr
    for child in list(spPr):
        if child.tag.endswith('solidFill') or child.tag.endswith('ln'):
            spPr.remove(child)
            
    # Alpha = 100 ensures mouse clicks are captured without any visible obstruction
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


def add_title_overlay(slide, overlay_cfg, slide_width_in, slide_height_in):
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
    spPr = card._element.spPr
    for child in list(spPr):
        if child.tag.endswith('solidFill') or child.tag.endswith('ln'):
            spPr.remove(child)
            
    # Frosted dark obsidian backdrop (80% opacity)
    fill_xml = parse_xml(
        '<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:srgbClr val="0A0F1D"><a:alpha val="80000"/></a:srgbClr>'
        '</a:solidFill>'
    )
    spPr.append(fill_xml)
    
    # Glowing cyan accent border
    ln_xml = parse_xml(
        '<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="19050">'
        '<a:solidFill><a:srgbClr val="06B6D4"><a:alpha val="65000"/></a:srgbClr></a:solidFill>'
        '</a:ln>'
    )
    spPr.append(ln_xml)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.4)
    tf.margin_right = Inches(0.4)
    tf.margin_top = Inches(0.28)
    tf.margin_bottom = Inches(0.28)
    
    # 1. Event Tag / Pill
    tag_text = overlay_cfg.get("tag", "🎙️  رویداد ارائـه تخصصـی مهندسـی  |  TECHTALK")
    if tag_text:
        p0 = tf.paragraphs[0]
        p0.text = tag_text
        p0.font.name = "Tahoma"
        p0.font.size = Pt(13)
        p0.font.bold = True
        p0.font.color.rgb = RGBColor(56, 189, 248)  # Cyan 400
        p0.alignment = PP_ALIGN.CENTER
        p0.space_after = Pt(8)
        pPr0 = p0._element.get_or_add_pPr()
        pPr0.set('rtl', '1')
        
    # 2. Main Title
    title_text = overlay_cfg.get("title", "دایناسورهای دنیای تک و مهارتِ بقا")
    if title_text:
        p1 = tf.add_paragraph() if tag_text else tf.paragraphs[0]
        p1.text = title_text
        p1.font.name = "Tahoma"
        p1.font.size = Pt(36)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(255, 255, 255)  # Crisp White
        p1.alignment = PP_ALIGN.CENTER
        p1.space_after = Pt(10)
        pPr1 = p1._element.get_or_add_pPr()
        pPr1.set('rtl', '1')
        
    # 3. Subtitle
    sub_text = overlay_cfg.get("subtitle", "")
    if sub_text:
        p2 = tf.add_paragraph()
        p2.text = sub_text
        p2.font.name = "Tahoma"
        p2.font.size = Pt(15)
        p2.font.bold = False
        p2.font.color.rgb = RGBColor(203, 213, 225)  # Slate 300
        p2.alignment = PP_ALIGN.CENTER
        p2.space_after = Pt(14)
        pPr2 = p2._element.get_or_add_pPr()
        pPr2.set('rtl', '1')
        
    # 4. Presenter Info
    pres_text = overlay_cfg.get("presenter", "ارائه‌دهنده: علی امانی   •   پاییز ۱۴۰۳")
    if pres_text:
        p3 = tf.add_paragraph()
        p3.text = pres_text
        p3.font.name = "Tahoma"
        p3.font.size = Pt(14)
        p3.font.bold = True
        p3.font.color.rgb = RGBColor(251, 191, 36)  # Amber 400
        p3.alignment = PP_ALIGN.CENTER
        pPr3 = p3._element.get_or_add_pPr()
        pPr3.set('rtl', '1')
        
    return card


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
    
    id_to_slide = {}
    id_to_idx = {}
    
    print("\n📦 Generating slides and transitions...")
    # Step 1: Create all slides with full-bleed 16:9 images, transitions, and notes
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
        
        # Add title overlay if specified
        if "title_overlay" in scfg:
            add_title_overlay(slide, scfg["title_overlay"], slide_width_in, slide_height_in)
            print(f"  ✨ Added elegant title card overlay on Slide '{s_id}'.")
            
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
        print("\n🗺️ Adding roadmap clickable hotspots and minimal shortcut...")
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
            
        # Minimal quick jump icon button (no verbose text)
        qj = roadmap_cfg.get("quick_jump")
        if qj and qj.get("target_slide_id") in id_to_slide:
            q_target = id_to_slide[qj["target_slide_id"]]
            add_icon_button(
                roadmap_slide,
                left=Inches(0.4),
                top=Inches(0.35),
                size=Inches(0.46),
                icon=qj.get("icon", "☄️"),
                target_slide=q_target,
                font_size_pt=14,
                bg_rgb=qj.get("bg_color", [26, 16, 37]),
                border_rgb=qj.get("border_color", [236, 72, 153])
            )

    # Step 3: Add Minimal Clean Navigation Buttons (NO TOP BADGES!)
    print("\n🧭 Adding clean icon-only navigation buttons...")
    btn_size = Inches(0.44)
    btn_y_bottom = Inches(6.7)
    
    for idx, scfg in enumerate(slides_cfg):
        s_id = scfg["id"]
        slide = id_to_slide[s_id]
        
        # 1. Back to Roadmap icon button (Top-Left, minimal icon [🗺️], no text)
        if scfg.get("back_to_roadmap", False) and roadmap_slide:
            add_icon_button(
                slide,
                left=Inches(0.4),
                top=Inches(0.35),
                size=btn_size,
                icon="🗺️",
                target_slide=roadmap_slide,
                font_size_pt=14,
                bg_rgb=[15, 23, 42],
                border_rgb=[6, 182, 212]
            )

        # 2. Next / Previous icon buttons (Bottom corners, minimal [◀] and [▶], no text)
        if scfg.get("nav_buttons", False):
            # Previous slide button (Bottom-Right, minimal icon [▶])
            if idx > 0:
                prev_slide = id_to_slide[slides_cfg[idx - 1]["id"]]
                add_icon_button(
                    slide,
                    left=Inches(12.45),
                    top=btn_y_bottom,
                    size=btn_size,
                    icon="▶",
                    target_slide=prev_slide,
                    font_size_pt=12,
                    bg_rgb=[15, 23, 42],
                    border_rgb=[100, 116, 139]
                )
                
            # Next slide button (Bottom-Left, minimal icon [◀])
            if idx < len(slides_cfg) - 1:
                next_slide = id_to_slide[slides_cfg[idx + 1]["id"]]
                is_finale = (idx >= len(slides_cfg) - 3)
                next_border = [236, 72, 153] if is_finale else [6, 182, 212]
                add_icon_button(
                    slide,
                    left=Inches(0.4),
                    top=btn_y_bottom,
                    size=btn_size,
                    icon="◀",
                    target_slide=next_slide,
                    font_size_pt=12,
                    bg_rgb=[15, 23, 42],
                    border_rgb=next_border
                )

    # Step 4: Save PowerPoint presentation safely
    print("\n💾 Saving presentations...")
    primary_out = output_override or pres_meta.get("output_primary", "TechTalk_Dinosaurs_Survival.pptx")
    primary_path = os.path.join(base_dir, primary_out)
    interactive_out = pres_meta.get("output_interactive", "TechTalk_Dinosaurs_Survival_Interactive.pptx")
    interactive_path = os.path.join(base_dir, interactive_out)
    clean_path = os.path.join(base_dir, "TechTalk_Dinosaurs_Survival_Clean.pptx")
    
    def try_save(target_path, desc):
        try:
            prs.save(target_path)
            print(f"  ✅ Saved {desc}: {target_path}")
            return True
        except PermissionError:
            print(f"  ⚠️ Note: '{os.path.basename(target_path)}' is currently open in PowerPoint (file locked).")
            return False
        except Exception as e:
            print(f"  ❌ Error saving '{target_path}': {e}")
            return False

    # Save to Clean guaranteed copy
    try_save(clean_path, "Clean Guaranteed Deck")
    
    # Save to primary
    try_save(primary_path, "Primary Deck")
    
    # Save to interactive (if not locked by user's PowerPoint)
    try_save(interactive_path, "Interactive Deck")
        
    print("\n🎉 Presentation build completed successfully with pure minimal layout!")


def main():
    parser = argparse.ArgumentParser(description="Build clean interactive TechTalk PowerPoint decks")
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
