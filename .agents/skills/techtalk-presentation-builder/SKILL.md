---
name: techtalk-presentation-builder
description: >-
  Interactive PowerPoint Presentation Builder for TechTalk "Dinosaurs Survival".
  Use when adding new slides, updating images, adding or modifying roadmap stations,
  adjusting speech notes, or regenerating the clickable interactive PowerPoint deck.
---

# TechTalk Interactive Presentation Builder Skill

This skill guides you through maintaining and extending the interactive PowerPoint presentation for the **«دایناسورهای دنیای تک و مهارتِ بقا»** TechTalk.

---

## 🏗️ Architecture Overview

The presentation is entirely declarative and managed via two primary files:
1. **`slides_config.json`**: The single source of truth manifest containing all slide metadata, roadmap hotspots, and navigation links.
2. **`build_presentation.py`**: The generation engine that validates assets, creates 16:9 widescreen slides, connects hyperlinks dynamically using string IDs, and applies transitions.

All slide background images are stored in `slides/` (or project root) at 16:9 widescreen ratio (e.g., 1376x768 px).

---

## ➕ How to Add a New Slide

To add a new slide anywhere in the deck:
1. Place your new 16:9 image inside `slides/` (e.g. `slides/my_new_station.jpg`).
2. Open `slides_config.json` and add an entry inside the `"slides"` list at your desired position:

```json
{
  "id": "slide_station_new",
  "file": "my_new_station.jpg",
  "type": "station",
  "title": "اسلاید جدید — عنوان ایستگاه",
  "badge": "📍 ایستگاه جدید (۱۴۰۱): عنوان دوره",
  "back_to_roadmap": true,
  "nav_buttons": true,
  "next_text": "ایستگاه بعدی ◀",
  "notes": "📌 راهنمای استیج:\n...\n🎙️ نکات کلیدی گفتار:\n..."
}
```

3. If this slide is a roadmap station that should be clickable from Slide 2 (the Highway Roadmap), register its hotspot in `"roadmap_settings" -> "hotspots"`:

```json
{
  "id": "station_new",
  "name": "ایستگاه جدید",
  "target_slide_id": "slide_station_new",
  "box_px": [700, 200, 200, 400],
  "ref_image_size": [1376, 768]
}
```

4. Run the builder:
```powershell
python build_presentation.py
```

All next/previous buttons, station badges, and roadmap links are **dynamically resolved by ID** — no hardcoded slide numbers will ever break!

---

## 🗺️ How to Update the Career Roadmap Image

If you regenerate or modify `slides/career_roadmap_highway.jpg`:
1. Save the new image in `slides/career_roadmap_highway.jpg`.
2. Inspect the new station bounding boxes (e.g., using `PIL` or coordinates).
3. Update `"box_px": [x, y, width, height]` for each station in `slides_config.json`.
4. Run:
```powershell
python build_presentation.py
```

---

## 🛠️ CLI Commands & Utilities

- **Build Presentation:**
  ```powershell
  python build_presentation.py
  ```
- **Validate Config & Assets Only (without building):**
  ```powershell
  python build_presentation.py --validate
  ```
- **Build with Custom Output File:**
  ```powershell
  python build_presentation.py --output "My_Custom_Deck.pptx"
  ```

---

## 🔒 File Lock Handling (Windows PowerPoint)

When a PowerPoint `.pptx` file is currently open in Microsoft PowerPoint, Windows locks the file exclusively. The builder handles this gracefully:
- It attempts to save to `TechTalk_Dinosaurs_Survival_Interactive.pptx`.
- It automatically updates `TechTalk_Dinosaurs_Survival.pptx`.
- If the primary file is locked by PowerPoint, it automatically saves a timestamped copy (e.g. `TechTalk_Dinosaurs_Survival_YYYYMMDD_HHMMSS.pptx`) so your progress is never blocked or lost!
