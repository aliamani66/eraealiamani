# -*- coding: utf-8 -*-
"""
Builds BreakingChanges_Audio.mp3 from BreakingChanges_Script.docx so it sounds like a live presenter
delivering the talk (complete with slide/click/audience pauses and Persian pronunciation normalization).

    C:\\work\\eraee-ali\\eraealiamani\\.venv\\Scripts\\python.exe new\\build_audio.py
"""
from __future__ import annotations

import argparse
import asyncio
import hashlib
import re
import sys
from dataclasses import dataclass
from pathlib import Path

import docx
import truststore

truststore.inject_into_ssl()
import edge_tts  # noqa: E402  (import after truststore injects Windows OS cert store)

HERE = Path(__file__).resolve().parent
DEFAULT_SRC = HERE / "BreakingChanges_Script.docx"
DEFAULT_OUT = HERE / "BreakingChanges_Audio.mp3"

VOICE_DEFAULT = "fa-IR-FaridNeural"
RATE_DEFAULT = "+2%"
PITCH_DEFAULT = "+0Hz"

# MPEG-2 Layer III, 48 kbps, 24 kHz, Mono -> 144 bytes = 24 ms per frame
SILENT_FRAME = b"\xff\xf3\x64\xc4" + (b"\x00" * 140)
FRAME_MS = 24.0

SKIP_PREFIXES = ("•", "زمان\u200cبندی", "متن گفتاری")

# Pause durations (milliseconds) aligned with live stage cues
PAUSE_ACT_MS = 1600
PAUSE_SLIDE_MS = 950
PAUSE_CLICK_MS = 600
PAUSE_CROWD_MS = 1900

DRAMATIC_PAUSES: tuple[tuple[str, int], ...] = (
    ("یکی\u200cیکی خاموش شدن", 1400),
    ("حقوق بده؟»", 1800),
    ("نمودار بار رو نگاه می\u200cکردیم", 1800),
    ("ویدئو رو ببینیم", 1400),
    ("چی می‌خوابه؟", 1500),
    ("شاگرد بودید کِی بود؟", 1500),
    ("جمله‌ی شما چیه؟", 1800),
    ("آخرین major شما", 1800),
)

# Longest-match-first English -> Persian phonetic pronunciation map
ENGLISH_TERMS: dict[str, str] = {
    "BREAKING CHANGES": "بِرِیکینگ چِینجِز",
    "BREAKING CHANGE": "بِرِیکینگ چِینج",
    "Intelligent SRE": "اینتِلیجِنت اِس آر ای",
    "Query Executor": "کوئری اِگزِکیوتر",
    "code generator": "کُد جِنِرِیتور",
    "enabling team": "اِنِیبلینگ تیم",
    "release notes": "ریلیز نوتس",
    "known issue": "نُوْن ایشو",
    "code review": "کُد ریویو",
    "Root cause": "روت کاز",
    "root cause": "روت کاز",
    "Live Tail": "لایو تِیل",
    "R&D Ops": "آر اَند دی آپس",
    "R&D": "آر اَند دی",
    "CHANGELOG": "چِینج‌لاگ",
    "Kubernetes": "کوبرنتیز",
    "Deprecated": "دیپریکِیتِد",
    "deprecated": "دیپریکِیتِد",
    "Unreleased": "آن‌ریلیزد",
    "middleware": "میدِل وِر",
    "blameless": "بِلِیم‌لِس",
    "scale-out": "اِسکِیل اَوت",
    "threshold": "تِرِش‌هولد",
    "scale-up": "اِسکِیل آپ",
    "ActiveMQ": "اَکتیو اِم کیو",
    "Filebeat": "فایل بیت",
    "shutdown": "شات داون",
    "scripted": "اسکریپتِد",
    "stateful": "استیت‌فول",
    "instance": "اینستَنس",
    "ChatGPT": "چَت جی پی تی",
    "Ansible": "اَنسیبِل",
    "approve": "اَپروو",
    "release": "ریلیز",
    "session": "سِشِن",
    "Removed": "ریمووْد",
    "Trigger": "تِریگِر",
    "trigger": "تِریگِر",
    "Cursor": "کِرسِر",
    "DevOps": "دِواپس",
    "Oracle": "اوراکِل",
    "commit": "کامیت",
    "Groovy": "گرووی",
    "legacy": "لِگَسی",
    "stable": "استِیبِل",
    "review": "ریویو",
    "Kafka": "کافکا",
    "Added": "اَدِد",
    "Fixed": "فیکسد",
    "state": "اِستِیت",
    "MAJOR": "مِیجِر",
    "major": "مِیجِر",
    "MINOR": "ماینِر",
    "minor": "ماینِر",
    "PATCH": "پَچ",
    "patch": "پَچ",
    "merge": "مِرج",
    "LGTM": "اِل جی تی اِم",
    "SPOF": "اِسپاف",
    "Bash": "بَش",
    "Git": "گیت",
    "SSH": "اِس اِس اِچ",
    "CMD": "سی اِم دی",
    "CMS": "سی اِم اِس",
    "NFS": "اِن اِف اِس",
    "ELK": "اِلک",
    "POC": "پی او سی",
    "ISR": "آی اِس آر",
    "NCG": "اِن سی جی",
    "diff": "دیف",
    "n8n": "اِن اِیت اِن",
    "dev": "دِیو",
    "fix": "فیکس",
    "for": "فُور",
    "IP": "آی پی",
    "VM": "وی اِم",
    "AI": "اِی آی",
}

SUFFIX_MAP = {"ه": "ِه", "ش": "ِش", "ی": "ی", "هاش": "‌هاش", "ها": "‌ها"}

FA_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
FA_NUM_WORDS = {0: "صفر", 1: "یک", 2: "دو", 3: "سه", 4: "چهار", 5: "پنج", 6: "شیش"}

PERSIAN_TWEAKS: tuple[tuple[str, str], ...] = (
    ("قصه‌ش", "قِصَّش"),
    ("بچه‌م", "بَچَّم"),
    ("بچه‌ش", "بَچَّش"),
    ("قلقش", "قِلِقِش"),
    ("فیکسد‌هاش", "فیکس‌هاش"),
    ("یه مستر ", "یه مَستِر "),
    ("ارکستریت", "اُرکِست‌رِیْت"),
    ("یکی\u200cیکی", "یکی یکی"),
    ("سه خط ۷۰ تا", "سه خط، ۷۰ تا"),
)

_TERM_KEYS = sorted(ENGLISH_TERMS, key=len, reverse=True)
_TERM_RE = re.compile(
    r"(?<![A-Za-z0-9])(" + "|".join(re.escape(k) for k in _TERM_KEYS) + r")(?![A-Za-z0-9])"
)
_ATTACHED_SUFFIX_RE = re.compile(
    r"(?<![A-Za-z0-9])("
    + "|".join(re.escape(k) for k in _TERM_KEYS)
    + r")(?:\s*ـ|\u200c)?(هاش|ها|ش|ه|ی)(?![\u0600-\u06FF])"
)
_TIME_RE = re.compile(r"(ساعت\s+)?([0-9\u06f0-\u06f9]{1,2}):([0-9\u06f0-\u06f9]{2})")
_VER3_RE = re.compile(r"([0-9\u06f0-\u06f9]+)\.([0-9\u06f0-\u06f9]+)\.([0-9\u06f0-\u06f9]+)")
_VER2_RE = re.compile(r"([0-9\u06f0-\u06f9]+)\.([0-9\u06f0-\u06f9]+)")


@dataclass
class Beat:
    pause_before_ms: int
    text: str
    pause_after_ms: int = 0
    label: str = ""


def _num_word(token: str) -> str:
    n = int(token.translate(FA_DIGITS))
    return FA_NUM_WORDS.get(n, str(n))


def _replace_time(m: re.Match[str]) -> str:
    prefix = m.group(1) or "ساعت "
    h = _num_word(m.group(2))
    mins = int(m.group(3).translate(FA_DIGITS))
    if mins == 0:
        return f"{prefix}{h}"
    if mins == 15:
        return f"{prefix}{h} و پونزده دقیقه"
    if mins == 30:
        return f"{prefix}{h} و سی دقیقه"
    return f"{prefix}{h} و {mins} دقیقه"


def normalize_for_tts(text: str) -> str:
    """Normalizes mixed Persian/English script text for natural Persian neural TTS."""
    s = re.sub(r"\s*\[؟[^\]]*\]", "", text).strip()
    s = _TIME_RE.sub(_replace_time, s)
    s = _VER3_RE.sub(lambda m: f"{_num_word(m.group(1))} نقطه {_num_word(m.group(2))} نقطه {_num_word(m.group(3))}", s)
    s = _VER2_RE.sub(lambda m: f"{_num_word(m.group(1))} و {_num_word(m.group(2))}", s)
    s = _ATTACHED_SUFFIX_RE.sub(
        lambda m: ENGLISH_TERMS[m.group(1)] + SUFFIX_MAP[m.group(2)], s
    )
    s = _TERM_RE.sub(lambda m: ENGLISH_TERMS[m.group(0)], s)
    for src, dst in PERSIAN_TWEAKS:
        s = s.replace(src, dst)
    return re.sub(r"\s{2,}", " ", s).strip()


def parse_docx_beats(docx_path: Path) -> list[Beat]:
    """Parses spoken beats and stage-cue pauses directly from BreakingChanges_Script.docx."""
    doc = docx.Document(str(docx_path))
    beats: list[Beat] = []
    pending_pause = 0
    current_lines: list[str] = []
    current_label = "مقدمه"
    seen_speech = False

    def flush(extra_after_ms: int = 0) -> None:
        nonlocal pending_pause, current_lines, seen_speech
        if not current_lines:
            return
        joined = normalize_for_tts("\n".join(current_lines))
        if joined:
            beats.append(Beat(pause_before_ms=pending_pause, text=joined, pause_after_ms=extra_after_ms, label=current_label))
            seen_speech = True
            pending_pause = 0
        current_lines = []

    for p in doc.paragraphs:
        raw = p.text.strip()
        if not raw or any(raw.startswith(pref) for pref in SKIP_PREFIXES):
            continue
        if raw.startswith("پرده "):
            flush()
            if seen_speech:
                pending_pause = max(pending_pause, PAUSE_ACT_MS)
            current_label = raw
            continue
        if raw.startswith("▶"):
            flush()
            if seen_speech:
                pending_pause = max(pending_pause, PAUSE_SLIDE_MS)
            current_label = raw.lstrip("▶ ").strip()
            continue
        if raw.startswith("●"):
            flush()
            if seen_speech:
                pending_pause = max(pending_pause, PAUSE_CLICK_MS)
            continue
        if raw.startswith("◆"):
            flush()
            if seen_speech:
                pending_pause = max(pending_pause, PAUSE_CROWD_MS)
            continue

        cleaned = re.sub(r"\s*\[؟[^\]]*\]", "", raw).strip()
        if not cleaned:
            continue
        current_lines.append(cleaned)
        for needle, pause_ms in DRAMATIC_PAUSES:
            if needle in cleaned:
                flush(extra_after_ms=pause_ms)
                break

    flush()
    return beats


def strip_mp3_wrappers(data: bytes) -> bytes:
    """Strips optional ID3v2 header and verifies raw MPEG-2 Layer III frames."""
    pos = 0
    if data[:3] == b"ID3" and len(data) >= 10:
        sz = ((data[6] & 0x7F) << 21) | ((data[7] & 0x7F) << 14) | ((data[8] & 0x7F) << 7) | (data[9] & 0x7F)
        pos = 10 + sz

    out = bytearray()
    bitrates = (0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160, 0)
    samplerates = (22050, 24000, 16000, 0)

    while pos + 4 <= len(data):
        if data[pos] != 0xFF or (data[pos + 1] & 0xE0) != 0xE0:
            break
        b1, b2 = data[pos + 1], data[pos + 2]
        ver = (b1 >> 3) & 3
        layer = (b1 >> 1) & 3
        br = bitrates[(b2 >> 4) & 0xF] * 1000
        sr = samplerates[(b2 >> 2) & 3]
        pad = (b2 >> 1) & 1
        if ver != 2 or layer != 1 or br == 0 or sr == 0:
            raise ValueError(f"Unexpected MPEG frame format: ver={ver} layer={layer} br={br} sr={sr}")
        flen = (72 * br // sr) + pad
        frame = data[pos : pos + flen]
        if len(frame) < flen:
            break
        if b"Xing" not in frame[:48] and b"Info" not in frame[:48]:
            out.extend(frame)
        pos += flen
    return bytes(out)


def silence_frames(ms: int) -> bytes:
    if ms <= 0:
        return b""
    count = round(ms / FRAME_MS)
    return SILENT_FRAME * count


async def synthesize_beat(
    beat: Beat, voice: str, rate: str, pitch: str, sem: asyncio.Semaphore, max_retries: int = 4
) -> bytes:
    async with sem:
        last_err: Exception | None = None
        for attempt in range(1, max_retries + 1):
            try:
                buf = bytearray()
                comm = edge_tts.Communicate(beat.text, voice, rate=rate, pitch=pitch)
                async for msg in comm.stream():
                    if msg["type"] == "audio":
                        buf.extend(msg["data"])
                frames = strip_mp3_wrappers(bytes(buf))
                if not frames:
                    raise RuntimeError(f"Empty audio returned for beat: {beat.label}")
                return frames
            except Exception as exc:  # noqa: BLE001
                last_err = exc
                await asyncio.sleep(0.8 * attempt)
        raise RuntimeError(f"Failed synthesizing '{beat.label}' after {max_retries} attempts: {last_err}")


async def build_all(beats: list[Beat], voice: str, rate: str, pitch: str, concurrency: int) -> bytes:
    sem = asyncio.Semaphore(concurrency)
    tasks = [synthesize_beat(b, voice, rate, pitch, sem) for b in beats]
    chunks = await asyncio.gather(*tasks)

    assembled = bytearray()
    pending_after = 0
    for beat, audio in zip(beats, chunks):
        gap = max(pending_after, beat.pause_before_ms)
        assembled.extend(silence_frames(gap))
        assembled.extend(audio)
        pending_after = beat.pause_after_ms
    assembled.extend(silence_frames(300))
    return bytes(assembled)


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(description="Convert BreakingChanges_Script.docx to a presenter MP3.")
    parser.add_argument("--src", type=Path, default=DEFAULT_SRC)
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT)
    parser.add_argument("--voice", default=VOICE_DEFAULT)
    parser.add_argument("--rate", default=RATE_DEFAULT)
    parser.add_argument("--pitch", default=PITCH_DEFAULT)
    parser.add_argument("--concurrency", type=int, default=4)
    parser.add_argument("--dry-run", action="store_true", help="Print normalized beats without synthesizing audio.")
    args = parser.parse_args(argv)

    if not args.src.exists():
        print(f"ERROR: Source file not found: {args.src}")
        return 1

    raw_bytes = args.src.read_bytes()
    print(f"source: {args.src.name}  {len(raw_bytes):,} bytes  md5={hashlib.md5(raw_bytes).hexdigest()}")

    beats = parse_docx_beats(args.src)
    leftover_latin = sorted({w for b in beats for w in re.findall(r"[A-Za-z0-9]+", b.text)})
    print(f"parsed {len(beats)} presentation beats; voice={args.voice} rate={args.rate} pitch={args.pitch}")
    if leftover_latin:
        print(f"  note: unmapped ASCII tokens (check pronunciation): {leftover_latin}")

    if args.dry_run:
        for i, b in enumerate(beats, 1):
            print(f"[{i:02d}] +{b.pause_before_ms}ms | {b.label} -> {b.text} | +{b.pause_after_ms}ms")
        return 0

    audio_bytes = asyncio.run(build_all(beats, args.voice, args.rate, args.pitch, args.concurrency))
    total_frames = len(audio_bytes) // len(SILENT_FRAME)
    dur_sec = total_frames * FRAME_MS / 1000.0
    mins, secs = divmod(int(round(dur_sec)), 60)

    try:
        args.out.write_bytes(audio_bytes)
    except PermissionError:
        print(f"ERROR: {args.out} is locked (probably open in a media player). Close it and run again.")
        return 2

    print(f"wrote {args.out} ({len(audio_bytes):,} bytes, {total_frames} frames, duration {mins}:{secs:02d})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
