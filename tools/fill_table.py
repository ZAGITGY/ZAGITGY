"""Rewrite the report table (and the month in its titles) of a built appendix.

Usage: python3 fill_table.py <rows.json> <report.docx> [rows_on_first_page]

rows_on_first_page defaults to all rows for 3 or fewer (page 2 is then dropped),
else 6.

The template has two report tables, one per page, each with two header rows.
Data rows are rebuilt from the first data row of each table, so cell widths and
fonts stay as in the template. Only the two report titles, the two tables and
nothing else are changed; the appendix and signatures are left as they are.
"""
import copy
import json
import sys

import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

THAI_DIGITS = str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙")


def set_cell_text(tc, text, fallback_rpr):
    """Keep the cell's first paragraph and its first run formatting (or
    fallback_rpr when the template cell is empty)."""
    paras = tc.findall(qn("w:p"))
    for p in paras[1:]:
        tc.remove(p)
    p = paras[0]
    runs = p.findall(qn("w:r"))
    rpr = copy.deepcopy(runs[0].find(qn("w:rPr")) if runs else fallback_rpr)
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)
    r = parse_xml("<w:r %s/>" % nsdecls("w"))
    if rpr is not None:
        r.append(rpr)
    t = parse_xml('<w:t %s xml:space="preserve"/>' % nsdecls("w"))
    t.text = text
    r.append(t)
    p.append(r)


def set_title_month(p, month):
    """Titles end with runs like 'ก.ค.' ' ๖' '9'; rewrite from the month run on."""
    runs = p.findall(qn("w:r"))
    i = next(i for i, r in enumerate(runs) if "ก.ค." in "".join(t.text or "" for t in r.iter(qn("w:t"))))
    runs[i].find(qn("w:t")).text = month + " ๖๙"
    for r in runs[i + 1:]:
        p.remove(r)


def main():
    rows_path, report = sys.argv[1], sys.argv[2]
    data = json.load(open(rows_path, encoding="utf-8"))
    rows = data["rows"]
    first = int(sys.argv[3]) if len(sys.argv) > 3 else (len(rows) if len(rows) <= 3 else 6)

    doc = docx.Document(report)
    body = list(doc.element.body)
    titles = [e for e in body[:8] if e.tag == qn("w:p") and "ประจำเดือน" in "".join(t.text or "" for t in e.iter(qn("w:t")))]
    assert len(titles) == 2, titles
    for p in titles:
        set_title_month(p, data["month_short"])
    # the template relies on page 1 being full; with other rows the page-2 title
    # would slide up onto page 1 (dropped below when every row fits on page 1)
    ppr = titles[1].find(qn("w:pPr"))
    if ppr.find(qn("w:pageBreakBefore")) is None:
        ppr.insert(0, parse_xml("<w:pageBreakBefore %s/>" % nsdecls("w")))

    tables = doc.tables[:2]
    parts = [rows[:first], rows[first:]]
    n = 0
    for tbl, part in zip(tables, parts):
        trs = tbl._tbl.findall(qn("w:tr"))
        proto = copy.deepcopy(trs[2])
        for tr in trs[2:]:
            tbl._tbl.remove(tr)
        for row in part:
            n += 1
            tr = copy.deepcopy(proto)
            cells = tr.findall(qn("w:tc"))
            values = [str(n).translate(THAI_DIGITS) + ".", row["date"], row["activity"], row["result"],
                      row["category"], row["cadets"], row["public"], row.get("note", "")]
            base_rpr = cells[3].find(".//" + qn("w:rPr"))
            for tc, v in zip(cells, values):
                set_cell_text(tc, v, base_rpr)
            tbl._tbl.append(tr)
    body_el = doc.element.body
    if not parts[1]:
        # everything fits on page 1: drop the page-2 title, unit line and table
        title2 = titles[1]
        for e in [title2, title2.getnext(), tables[1]._tbl]:
            body_el.remove(e)
    # the template pads to the appendix with empty lines sized for its own
    # table; start the appendix on its own page instead
    children = list(body_el)
    head = next(i for i, e in enumerate(children)
                if "ภาคผนวก" in "".join(t.text or "" for t in e.iter(qn("w:t"))))
    sig = max(i for i, e in enumerate(children[:head])
              if "ตรวจถูกต้อง" in "".join(t.text or "" for t in e.iter(qn("w:t"))))
    for e in children[sig + 7:head]:  # keep 6 empty lines under the signature box
        if e.tag == qn("w:p") and not "".join(t.text or "" for t in e.iter(qn("w:t"))).strip():
            body_el.remove(e)
    hp = children[head].find(qn("w:pPr"))
    if hp.find(qn("w:pageBreakBefore")) is None:
        hp.insert(0, parse_xml("<w:pageBreakBefore %s/>" % nsdecls("w")))
    doc.save(report)
    print("filled %d rows (%d + %d)" % (n, len(parts[0]), len(parts[1])))


if __name__ == "__main__":
    main()
