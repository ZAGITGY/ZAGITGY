"""Check a built appendix against its handoff package.

Usage: python3 verify_appendix.py <handoff_dir> <output.docx>
Prints one line per check and exits non-zero if any check fails.
"""
import hashlib
import json
import os
import re
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {
    "w": W,
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
HEADING = "ภาคผนวก : รูปภาพกิจกรรมการปฏิบัติงาน นศท.จิตอาสา"

failures = []


def check(ok, msg):
    print(("PASS " if ok else "FAIL ") + msg)
    if not ok:
        failures.append(msg)


def text(el):
    return "".join(el.itertext()) if el is not None else ""


def ptext(p):
    return "".join(t.text or "" for t in p.iter("{%s}t" % W))


def body_children(z):
    root = etree.fromstring(z.read("word/document.xml"))
    return list(root.find("w:body", NS))


def appendix_start(els):
    return next(
        i for i, e in enumerate(els)
        if e.tag == "{%s}p" % W and ptext(e).strip() == HEADING
    )


def main():
    src, out = sys.argv[1], sys.argv[2]
    data = json.load(open(os.path.join(src, "01_ACTIVITY_DATA.json"), encoding="utf-8"))
    tpath = os.path.join(src, data["template"])
    check(hashlib.sha256(open(tpath, "rb").read()).hexdigest() == data["template_sha256"],
          "template matches template_sha256 (template untouched)")
    tz, oz = zipfile.ZipFile(tpath), zipfile.ZipFile(out)
    check(oz.testzip() is None, "output zip is intact")

    # 1. Pre-appendix body is byte-identical to the template.
    t_els, o_els = body_children(tz), body_children(oz)
    ts, os_ = appendix_start(t_els), appendix_start(o_els)
    same = ts == os_ and all(
        etree.tostring(a, method="c14n") == etree.tostring(b, method="c14n")
        for a, b in zip(t_els[:ts], o_els[:os_])
    )
    check(same, "pre-appendix body (%d elements) identical to template" % ts)
    check(etree.tostring(t_els[-1], method="c14n") == etree.tostring(o_els[-1], method="c14n"),
          "section properties (A4 landscape, margins) unchanged")
    for name in tz.namelist():
        if name.startswith("word/media/") or name in (
                "word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml",
                "docProps/app.xml", "docProps/core.xml"):
            continue
        check(name in oz.namelist() and tz.read(name) == oz.read(name), "part unchanged: " + name)

    # 2. Appendix content.
    app = o_els[os_:-1]
    paras = [e for e in app if e.tag == "{%s}p" % W]
    heads = [i for i, e in enumerate(app) if e.tag == "{%s}p" % W and ptext(e).strip() == HEADING]
    pages = -(-len(data["activities"]) // 3)
    check(len(heads) == pages, "appendix headings: %d (expected %d pages)" % (len(heads), pages))
    months = [ptext(app[i + 1]).strip() for i in heads]
    check(all(m == data["appendix_month"] for m in months),
          "month line under every heading = %r (%s)" % (data["appendix_month"], months))
    t_app = t_els[ts:-1]
    old_caps = {ptext(p).strip() for p in t_app if ptext(p).startswith("ภาพกิจกรรมที่")}
    left = [ptext(p).strip() for p in paras if ptext(p).strip() in old_caps]
    check(not left, "no template caption left in appendix %s" % left)
    if data["appendix_month"] != ptext(t_app[1]).strip():
        app_text = "".join(ptext(p) for p in paras)
        check(ptext(t_app[1]).strip() not in app_text, "no leftover template month line in appendix")
    captions = [ptext(p) for p in paras if ptext(p).startswith("ภาพกิจกรรมที่")]
    check(captions == [a["caption"] for a in data["activities"]],
          "captions exactly as given, in order (%d)" % len(captions))
    check(not re.search(r"\d+\s*(ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\.)",
                          "".join(captions)), "no dates in captions")
    check(sum("ตรวจถูกต้อง" in ptext(p) for p in paras) == 1, "one closing ตรวจถูกต้อง block")
    check(not app[-1].xpath(".//w:tbl", namespaces=NS) and "ตรวจถูกต้อง" in ptext(
        next(p for p in reversed(paras) if "ตรวจถูกต้อง" in ptext(p))), "signature block is last")

    # 3. Photos: each caption followed by a table of its 4 photos, unmodified, same
    #    frame size, filled by cropping (never stretched).
    frames = set()
    rels = etree.fromstring(oz.read("word/_rels/document.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels}
    seen = []
    tables = [e for e in app if e.tag == "{%s}tbl" % W]
    check(len(tables) == len(data["activities"]), "one photo table per activity")
    floating = sum(len(e.xpath(".//wp:anchor", namespaces=NS)) for e in app
                   if "ตรวจถูกต้อง" not in ptext(e))
    check(floating == 0, "no floating photos in appendix (only the template signature box floats)")
    photos = {p["file"]: p for p in data["photos"]}
    for act, tbl in zip(data["activities"], tables):
        prev = tbl.getprevious()
        check(ptext(prev) == act["caption"] and prev.find("w:pPr/w:keepNext", NS) is not None,
              "%s caption directly above its photos and kept with them" % act["code"])
        inl = tbl.findall(".//wp:inline", NS)
        names = [i.find("wp:docPr", NS).get("descr") for i in inl]
        check(names == act["main_files"], "%s photo order %s" % (act["code"], names))
        for i, name in zip(inl, names):
            blob = oz.read("word/" + target[i.find(".//a:blip", NS).get("{%s}embed" % NS["r"])])
            sha = hashlib.sha256(blob).hexdigest()
            check(sha == photos[name]["sha256"], "%s embedded byte-identical" % name)
            ext = i.find("wp:extent", NS)
            w, h = photos[name]["dimensions"]
            rect = i.find(".//a:srcRect", NS)
            l, t, r, b = (int(rect.get(k, 0)) / 100000 if rect is not None else 0 for k in "ltrb")
            shown = w * (1 - l - r) / (h * (1 - t - b))  # aspect of the uncropped part
            ratio = int(ext.get("cx")) / int(ext.get("cy"))
            check(abs(ratio - shown) / shown < 0.002,
                  "%s not stretched (frame %.3f, shown part %.3f, keeps %d%% x %d%%)"
                  % (name, ratio, shown, round(100 * (1 - l - r)), round(100 * (1 - t - b))))
            frames.add((ext.get("cx"), ext.get("cy")))
            seen.append(sha)
    check(len(frames) == 1, "all photos share one frame size %s" % sorted(frames))
    check(len(seen) == data["expected_main_photos"] == len(set(seen)),
          "%d photos, no duplicates" % len(seen))
    notes = {hashlib.sha256(open(os.path.join(src, a["note_file"]), "rb").read()).hexdigest()
             for a in data["activities"]}
    check(not notes & set(seen), "no SOURCE_NOTE image used as a photo")
    media = [n for n in oz.namelist() if n.startswith("word/media/")]
    check(len(media) == len(seen), "no leftover old media (%d files)" % len(media))
    ids = [d.get("id") for d in etree.fromstring(oz.read("word/document.xml")).iter("{%s}docPr" % NS["wp"])]
    check(len(ids) == len(set(ids)), "drawing ids unique")

    print("RESULT:", "ALL PASS" if not failures else "%d FAILED" % len(failures))
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
