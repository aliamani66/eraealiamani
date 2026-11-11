---
name: techtalk-presentation-builder
description: >-
  Interactive PowerPoint Presentation Builder for TechTalk "Dinosaurs Survival".
  Use when adding new slides, updating images, adding or modifying roadmap stations,
  adjusting speech notes, or regenerating the PowerPoint deck (clicker navigation, station progress bar).
---

# TechTalk Interactive Presentation Builder Skill

This skill guides you through maintaining and extending the interactive PowerPoint presentation for the **«دایناسورهای دنیای تک و مهارتِ بقا»** TechTalk.

---

## 🏗️ Architecture Overview

The presentation is entirely declarative and managed via two primary files:
1. **`slides_config.json`**: The single source of truth manifest containing all slide metadata, the `stations` list, and per-slide text overlays.
2. **`build_presentation.py`**: The generation engine that validates assets, creates 16:9 widescreen slides, draws the station progress bar, and applies transitions. The deck has no buttons or clickable areas; it is driven by a clicker.

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
  "station": "cloud",
  "notes": "📌 راهنمای استیج:\n...\n🎙️ نکات کلیدی گفتار:\n..."
}
```

3. If this slide is a station, add the station to the top-level `"stations"` list (order = dots from right to left) and use its id in the slide's `"station"` key. The slide then gets a progress bar with one dot per station: current dot bright, earlier dots half-bright, later dots faint:

```json
{ "id": "cloud", "year": "۱۴۰۱", "title": "معماری کلود" }
```

An unknown station id fails validation; a station with no slide only prints a warning.

4. Run the builder:
```powershell
python build_presentation.py
```

The progress bar is **resolved by station ID** — no hardcoded slide numbers will ever break!

---

## 🗺️ Career Roadmap Image

`slides/career_roadmap_highway.jpg` is only a background now. There are no clickable areas on it, so replacing the image needs no coordinates. (A5 will rebuild the map with real PowerPoint text from `stations`.)

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

## 🔒 Output File and File Lock (Windows PowerPoint)

The builder writes exactly one file: `TechTalk_Dinosaurs_Survival.pptx` (or `--output` if given). Older copies (`_Backup`, `_v2`, `_Clean`, `_Interactive`) are not written anymore and are kept only as history.

When that file is open in PowerPoint, Windows locks it. The builder then prints a clear message, exits with a non-zero code, and leaves the file unchanged. Close the file in PowerPoint and run the build again. No timestamped copy is written.
