"""Build the monthly RD volunteer appendix from REPORT_TEMPLATE.docx.

Usage: python3 build_appendix.py <handoff_dir> <output.docx> [image_height_mm]

Everything before the first appendix heading is left untouched. The old
appendix (headings, captions, photos) is replaced with the activities in
01_ACTIVITY_DATA.json, 3 activities per page, each caption above a single
centered row of 4 equal-size inline photos in a borderless table (photos of
another shape are cropped to fill, never stretched). The closing "ตรวจถูกต้อง"
signature block is carried over from the template unchanged.
"""
import copy
import json
import os
import sys
import zipfile

import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Mm

APPENDIX_HEADING = "ภาคผนวก : รูปภาพกิจกรรมการปฏิบัติงาน นศท.จิตอาสา"
PER_PAGE = 3
GAP = Mm(4)  # horizontal space around each photo inside its cell
FRAME_RATIO = 4 / 3  # every photo gets the same frame; most photos are 4:3
PORTRAIT_TOP_SHARE = 0.5  # share of a vertical trim taken from the top
EMU_PER_TWIP = 635


def page_starts(count):
    """Indexes of the activities that open an appendix page: PER_PAGE per page,
    except that the last page keeps room for the signature block below it."""
    sizes = [PER_PAGE] * (count // PER_PAGE) + ([count % PER_PAGE] if count % PER_PAGE else [])
    if sizes and sizes[-1] == PER_PAGE:
        sizes[-1:] = [PER_PAGE - 1, 1]
    return {sum(sizes[:i]) for i in range(len(sizes))}


def para_text(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def ensure_ppr_child(p, tag):
    ppr = p.find(qn("w:pPr"))
    if ppr.find(qn(tag)) is None:
        el = parse_xml("<%s %s/>" % (tag, nsdecls("w")))
        # keepNext / pageBreakBefore come right after pStyle in the schema order
        ppr.insert(1 if ppr.find(qn("w:pStyle")) is not None else 0, el)


def set_single_run_text(p, text):
    """Keep the paragraph's first run formatting, drop the other runs."""
    runs = p.findall(qn("w:r"))
    for r in runs[1:]:
        p.remove(r)
    for pe in p.findall(qn("w:proofErr")):
        p.remove(pe)
    r = runs[0]
    for child in list(r):
        if child.tag != qn("w:rPr"):
            r.remove(child)
    t = parse_xml('<w:t %s xml:space="preserve"/>' % nsdecls("w"))
    t.text = text
    r.append(t)


def fill_crop(img_w, img_h, box_ratio, top_share):
    """srcRect (l, t, r, b in 1/1000 %) that fills a box_ratio frame with the
    photo without distorting it; vertical trims put top_share of the cut on top."""
    ratio = img_w / img_h
    if ratio > box_ratio:
        side = round((1 - box_ratio / ratio) * 100000 / 2)
        return side, 0, side, 0
    cut = (1 - ratio / box_ratio) * 100000
    top = round(cut * top_share)
    return 0, top, 0, round(cut) - top


def photo_table(part, paths, height, top_shares):
    """One row of equal frames (FRAME_RATIO, `height` tall); photos with another
    shape are cropped in Word (reversible) to fill their frame."""
    from PIL import Image

    width = int(height * FRAME_RATIO)
    crops = []
    for path, top_share in zip(paths, top_shares):
        with Image.open(path) as im:
            crops.append(fill_crop(*im.size, FRAME_RATIO, top_share))
    sizes = [(width, int(height))] * len(paths)
    cols = [(w + GAP) // EMU_PER_TWIP for w, _ in sizes]
    grid = "".join('<w:gridCol w:w="%d"/>' % c for c in cols)
    tbl = parse_xml(
        "<w:tbl %s><w:tblPr>"
        '<w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/>'
        '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>'
        '<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
        '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
        "</w:tblPr><w:tblGrid>%s</w:tblGrid>"
        "<w:tr><w:trPr><w:cantSplit/><w:jc w:val=\"center\"/></w:trPr></w:tr></w:tbl>"
        % (nsdecls("w"), sum(cols), grid)
    )
    tr = tbl.find(qn("w:tr"))
    for path, (w, h), col, crop in zip(paths, sizes, cols, crops):
        inline = part.new_pic_inline(path, w, h)
        inline.docPr.set("descr", os.path.basename(path))
        if any(crop):
            blip_fill = inline.find(".//" + qn("pic:blipFill"))
            src_rect = parse_xml('<a:srcRect %s l="%d" t="%d" r="%d" b="%d"/>' % ((nsdecls("a"),) + crop))
            blip_fill.find(qn("a:blip")).addnext(src_rect)
        tc = parse_xml(
            "<w:tc %s><w:tcPr><w:tcW w:w=\"%d\" w:type=\"dxa\"/><w:vAlign w:val=\"center\"/></w:tcPr>"
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '<w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:drawing/></w:r></w:p></w:tc>'
            % (nsdecls("w"), col)
        )
        tc.find(".//" + qn("w:drawing")).append(inline)
        tr.append(tc)
    return tbl


def main():
    src, out = sys.argv[1], sys.argv[2]
    height = Mm(float(sys.argv[3]) if len(sys.argv) > 3 else 38)
    data = json.load(open(os.path.join(src, "01_ACTIVITY_DATA.json"), encoding="utf-8"))
    template = os.path.join(src, data["template"])
    assert os.path.abspath(out) != os.path.abspath(template)

    doc = docx.Document(template)
    body = doc.element.body
    els = list(body.iterchildren())
    start = next(
        i for i, e in enumerate(els)
        if e.tag == qn("w:p") and para_text(e).strip() == APPENDIX_HEADING
    )
    sect = els[-1]
    assert sect.tag == qn("w:sectPr")
    old = els[start:-1]

    heading_tpl = copy.deepcopy(els[start])
    month_tpl = copy.deepcopy(els[start + 1])
    caption_tpl = copy.deepcopy(els[start + 2])
    assert para_text(month_tpl).startswith("เดือน")
    assert para_text(caption_tpl).startswith("ภาพกิจกรรมที่ 1")
    sig = max(i for i, e in enumerate(old) if "ตรวจถูกต้อง" in para_text(e))
    closing = [copy.deepcopy(e) for e in old[sig:]]

    rids = set()
    for e in old:
        for blip in e.iter(qn("a:blip")):
            rids.add(blip.get(qn("r:embed")))
        body.remove(e)
    for rid in rids:
        doc.part.drop_rel(rid)

    new = []
    acts = data["activities"]
    crop_top = {p["file"]: p["crop_top"] for p in data.get("photos", []) if "crop_top" in p}
    starts = page_starts(len(acts))
    for n, act in enumerate(acts):
        if n in starts:
            h = copy.deepcopy(heading_tpl)
            for lr in h.iter(qn("w:lastRenderedPageBreak")):
                lr.getparent().remove(lr)
            if n:
                ensure_ppr_child(h, "w:pageBreakBefore")
            m = copy.deepcopy(month_tpl)
            set_single_run_text(m, data["appendix_month"])
            new += [h, m]
        c = copy.deepcopy(caption_tpl)
        set_single_run_text(c, act["caption"])
        ensure_ppr_child(c, "w:keepNext")
        if n not in starts:
            sp = c.find(qn("w:pPr")).find(qn("w:spacing"))
            sp.set(qn("w:before"), "160")
        new.append(c)
        paths = [os.path.join(src, f) for f in act["main_files"]]
        # a photo entry may set crop_top (0 = keep the top) when the default cut hides a face
        shares = [crop_top.get(f, PORTRAIT_TOP_SHARE) for f in act["main_files"]]
        new.append(photo_table(doc.part, paths, height, shares))
    new += closing

    for e in new:
        sect.addprevious(e)

    # new_pic_inline() picks its id before the picture is in the body, so every
    # photo got the same one; give each a unique id after the existing drawings.
    inlines = [d for e in new for d in e.iter(qn("wp:docPr")) if d.getparent().tag == qn("wp:inline")]
    others = [int(d.get("id")) for d in body.iter(qn("wp:docPr")) if d not in inlines]
    for i, d in enumerate(inlines, start=max(others, default=0) + 1):
        d.set("id", str(i))
    doc.save(out)
    restore_unchanged_parts(template, out)


def restore_unchanged_parts(template, out):
    """python-docx re-serializes every XML part; put back the template's exact
    bytes for parts whose content did not change."""
    from lxml import etree

    with zipfile.ZipFile(template) as tz:
        orig = {n: tz.read(n) for n in tz.namelist()}
    with zipfile.ZipFile(out) as oz:
        items = [(info, oz.read(info.filename)) for info in oz.infolist()]
    tmp = out + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as nz:
        for info, blob in items:
            name = info.filename
            if name in orig and name.endswith((".xml", ".rels")):
                c14n = lambda b: etree.tostring(etree.fromstring(b), method="c14n")
                if c14n(orig[name]) == c14n(blob):
                    blob = orig[name]
            nz.writestr(name, blob)
    os.replace(tmp, out)


if __name__ == "__main__":
    main()
