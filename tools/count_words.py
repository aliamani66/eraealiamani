"""Count spoken words per act in the markdown script (stage directions excluded)."""
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main(path):
    text = open(path, encoding="utf-8").read()
    acts = re.split(r"\n## ", text)
    total = 0
    rows = []
    for act in acts[1:]:
        title = act.split("\n", 1)[0].strip()
        if "🎙️" not in act:
            continue
        spoken = act.split("🎙️", 1)[1]
        spoken = spoken.split("\n", 1)[1]  # drop the heading line
        spoken = re.sub(r"\*\([^)]*\)\*", " ", spoken)  # stage directions *( ... )*
        spoken = re.sub(r"</?div[^>]*>", " ", spoken)
        spoken = re.sub(r"[*_`#>«»\"“”—\-–]", " ", spoken)
        words = [w for w in spoken.replace("‌", "").split() if re.search(r"\w", w)]
        n = len(words)
        total += n
        rows.append((title, n))
    for title, n in rows:
        print(f"{n:5d} words | {n/150:4.1f} min @150 | {n/130:4.1f} min @130 | {n/110:4.1f} min @110 | {title}")
    print(f"{total:5d} words | {total/150:4.1f} min @150 | {total/130:4.1f} min @130 | {total/110:4.1f} min @110 | TOTAL")


if __name__ == "__main__":
    main(sys.argv[1])
