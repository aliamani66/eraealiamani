"""Dump structure of a .pptx using only the stdlib (zipfile + ElementTree)."""
import sys
import re
import zipfile
import hashlib
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
EMU = 914400.0


def rels_for(z, part):
    d, _, f = part.rpartition("/")
    rp = f"{d}/_rels/{f}.rels"
    out = {}
    if rp in z.namelist():
        root = ET.fromstring(z.read(rp))
        for r in root:
            out[r.get("Id")] = (r.get("Type").rsplit("/", 1)[-1], r.get("Target"))
    return out


def text_of(el):
    paras = []
    for p in el.iter(f"{{{NS['a']}}}p"):
        t = "".join(x.text or "" for x in p.iter(f"{{{NS['a']}}}t"))
        if t.strip():
            paras.append(t)
    return paras


def main(path, verbose):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    sz = pres.find("p:sldSz", NS)
    print(f"== {path}")
    print(f"   slide size: {int(sz.get('cx'))/EMU:.3f} x {int(sz.get('cy'))/EMU:.3f} in")
    prels = rels_for(z, "ppt/presentation.xml")
    order = []
    for sid in pres.find("p:sldIdLst", NS):
        rid = sid.get(f"{{{NS['r']}}}id")
        order.append("ppt/" + prels[rid][1])
    print(f"   slides: {len(order)}")
    media = {}
    for n in names:
        if n.startswith("ppt/media/"):
            data = z.read(n)
            media[n] = (len(data), hashlib.md5(data).hexdigest()[:10])
    for i, part in enumerate(order, 1):
        root = ET.fromstring(z.read(part))
        rels = rels_for(z, part)
        trans = root.find("p:transition", NS)
        tdesc = "none"
        if trans is not None:
            tdesc = ",".join(c.tag.split("}")[1] for c in trans) or "?"
        print(f"\n--- slide {i} ({part.split('/')[-1]}) transition={tdesc}")
        tree = root.find("p:cSld/p:spTree", NS)
        for sh in tree:
            tag = sh.tag.split("}")[1]
            if tag in ("nvGrpSpPr", "grpSpPr"):
                continue
            nv = sh.find("*/p:cNvPr", NS)
            name = nv.get("name") if nv is not None else "?"
            off = sh.find(".//a:off", NS)
            ext = sh.find(".//a:ext", NS)
            pos = ""
            if off is not None and ext is not None:
                pos = (f"x={int(off.get('x'))/EMU:.2f} y={int(off.get('y'))/EMU:.2f} "
                       f"w={int(ext.get('cx'))/EMU:.2f} h={int(ext.get('cy'))/EMU:.2f}")
            extra = []
            if tag == "pic":
                blip = sh.find(".//a:blip", NS)
                if blip is not None:
                    rid = blip.get(f"{{{NS['r']}}}embed")
                    tgt = rels.get(rid, ("?", "?"))[1]
                    key = "ppt/" + tgt.replace("../", "")
                    extra.append(f"img={tgt.split('/')[-1]} {media.get(key)}")
            for hl in sh.iter(f"{{{NS['a']}}}hlinkClick"):
                rid = hl.get(f"{{{NS['r']}}}id")
                act = hl.get("action", "")
                tgt = rels.get(rid, ("", ""))[1] if rid else ""
                extra.append(f"click->{tgt or act}")
            txt = text_of(sh)
            fonts = sorted({x.get('typeface') for x in sh.iter(f"{{{NS['a']}}}latin")} |
                           {x.get('typeface') for x in sh.iter(f"{{{NS['a']}}}cs")})
            sizes = sorted({x.get('sz') for x in sh.iter(f"{{{NS['a']}}}rPr") if x.get('sz')} |
                           {x.get('sz') for x in sh.iter(f"{{{NS['a']}}}defRPr") if x.get('sz')})
            line = f"  [{tag}] {name} {pos} {' '.join(extra)}"
            print(line)
            if txt:
                for t in txt:
                    print(f"       text: {t!r}  fonts={fonts} sz={sizes}")
        # notes
        for rid, (typ, tgt) in rels.items():
            if typ == "notesSlide":
                npart = "ppt/" + tgt.replace("../", "")
                nroot = ET.fromstring(z.read(npart))
                ntxt = text_of(nroot)
                total = sum(len(t) for t in ntxt)
                print(f"  notes: {total} chars, {len(ntxt)} paragraphs")
                if verbose:
                    for t in ntxt:
                        print(f"     | {t}")
    print("\n   media:")
    for n, (size, h) in sorted(media.items()):
        print(f"     {n} {size} {h}")


if __name__ == "__main__":
    main(sys.argv[1], len(sys.argv) > 2 and sys.argv[2] == "-v")
