"""Simplify the Setting sheet (and the Cek Kilang helpers) of the crude
availability template.

Run this on the output of clean_template.py.

Changes
-------
Setting
  * "Tanda 1" table: the wide K1..K6 Kelas/Volume pairs (R:AJ) are replaced by
    two simple tables:
      - Sheet kilang | Kilang | Status            (R:T, rows 9-24)
      - Sheet kilang | Kelas | Volume | Status    (V:Y, one need per row, rows 9-80)
    The Status columns say in plain words what is wrong (sheet not found,
    volume not a column header, missing input...) or show "✔ OK".
  * Kilang table: Tabel MM columns are picked by NAME from a dropdown instead
    of typing a column number. The text column F that only described those
    numbers is removed; the notes move into F.
  * Helper columns (AF:AJ) are hidden.
Refinery sheets
  * Their "tanda 1" COUNTIFS now read the new one-need-per-row list.
    Nothing else in them changes.
Cek Kilang - Detail
  * Reads the Tabel MM column names from Setting (instead of numbers).

Usage:
    python simplify_setting.py INPUT_CLEAN.xlsx OUTPUT.xlsx
"""

from __future__ import annotations

import argparse
import re
import sys
import warnings
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

warnings.filterwarnings("ignore")

RASIO = "'Impor - Rasio & Batasan'"
PP = "'Impor - Pour Point'"
SHEET_ROWS = range(9, 25)       # sheet kilang table (fixed rows, refinery sheets point at them)
NEED_ROWS = range(9, 81)        # one-need-per-row list
KILANG_ROWS = range(9, 17)

# style ids that already exist in this workbook's styles.xml (Setting sheet)
S_TITLE, S_NOTE, S_HEAD, S_LABEL, S_CALC, S_INPUT, S_GREY = 130, 117, 119, 121, 135, 138, 139


# ------------------------------------------------------------ xml helpers
def x_attr(s: str) -> str:
    return escape(s, {'"': "&quot;"})


def cell_xml(ref: str, value, style: int | None) -> str:
    s = f' s="{style}"' if style is not None else ""
    if value is None or value == "":
        return f'<c r="{ref}"{s}/>'
    if isinstance(value, str) and value.startswith("="):
        return f'<c r="{ref}"{s}><f>{escape(value[1:])}</f></c>'
    if isinstance(value, bool):
        return f'<c r="{ref}"{s} t="b"><v>{int(value)}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


# ------------------------------------------------------------ main
def simplify(src: Path, dst: Path, cib_max: float | None = None) -> dict:
    wb = openpyxl.load_workbook(src)  # read-only use: values + expanded formulas
    names = wb.sheetnames
    for need in ("Setting", "Cek Kilang - Detail", "Impor - Rasio & Batasan", "Impor - Pour Point"):
        if need not in names:
            sys.exit(f"Sheet '{need}' not found - run this on the cleaned template.")
    st = wb["Setting"]
    if st["T8"].value != "K1 Kelas":
        sys.exit("Setting sheet does not have the old K1..K6 layout (already simplified?).")

    rasio_hdr = [c.value for c in wb["Impor - Rasio & Batasan"][1]]
    pp_hdr = [c.value for c in wb["Impor - Pour Point"][1]]

    zin = zipfile.ZipFile(src)
    files = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    order = [i.filename for i in zin.infolist()]

    # map sheet name -> part
    wbx = files["xl/workbook.xml"].decode()
    rels = files["xl/_rels/workbook.xml.rels"].decode()
    rid2t = {m.group(1): m.group(2) for m in re.finditer(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)}
    rid2t.update({m.group(2): m.group(1) for m in re.finditer(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)})
    part = {}
    for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        part[m.group(1).replace("&amp;", "&")] = "xl/" + rid2t[m.group(2)].lstrip("/").removeprefix("xl/")

    raw = files[part["Setting"]].decode()
    raw_style = {r: int(s) for r, s in re.findall(r'<c r="([A-Z]+\d+)" s="(\d+)"', raw)}
    row_attr = {int(r): a for r, a in re.findall(r'<row r="(\d+)"([^>]*)>', raw)}

    # ---- read old content
    cells: dict[str, tuple] = {}
    for row in st.iter_rows():
        for c in row:
            if c.value is not None or c.coordinate in raw_style:
                cells[c.coordinate] = (c.value, raw_style.get(c.coordinate))

    old_sheet_rows = {r: (st[f"R{r}"].value, st[f"S{r}"].value) for r in SHEET_ROWS}
    needs = []
    for r in SHEET_ROWS:
        sheet = st[f"R{r}"].value
        for k in range(6):
            kc = get_column_letter(column_index_from_string("T") + 2 * k)
            vc = get_column_letter(column_index_from_string("U") + 2 * k)
            kel, vol = st[f"{kc}{r}"].value, st[f"{vc}{r}"].value
            if kel not in (None, "") or vol not in (None, ""):
                needs.append((sheet, kel, vol))
    if len(needs) > len(NEED_ROWS):
        sys.exit("Too many needs for the new list.")

    kil_cols = {}
    for r in KILANG_ROWS:
        def nm(col, hdr):
            n = st[f"{col}{r}"].value
            return hdr[int(n) - 1] if isinstance(n, (int, float)) and n > 0 else None
        kil_cols[r] = (nm("B", rasio_hdr), nm("C", rasio_hdr), nm("D", pp_hdr), nm("E", pp_hdr), st[f"G{r}"].value)

    # ---- wipe columns we redesign: F:G (kilang table desc/notes) and R:AJ
    def col_i(ref):
        return column_index_from_string(split_ref(ref)[0])
    for ref in list(cells):
        ci = col_i(ref)
        if ci in (6, 7) or ci >= column_index_from_string("R"):
            del cells[ref]

    def put(ref, val, style):
        cells[ref] = (val, style)

    if cib_max is not None:
        put("B5", int(cib_max) if float(cib_max).is_integer() else cib_max, cells.get("B5", (None, S_INPUT))[1])

    # ---- intro + kilang table
    put("A2", "Cara pakai: isi hanya sel KUNING. Sel lain dihitung otomatis, jangan diubah. "
              "Kolom 'Status' harus ✔ OK.", S_NOTE)
    put("B8", "Kolom batasan 1 di Tabel MM (pilih)", S_HEAD)
    put("C8", "Kolom batasan 2 (kosong = tidak ada)", S_HEAD)
    put("D8", "Kolom pour point terima (kosong = lewati cek PP)", S_HEAD)
    put("E8", "Kolom pour point CIB (kosong = tidak ada)", S_HEAD)
    put("F8", "Catatan", S_HEAD)
    for r in KILANG_ROWS:
        b, c, d, e, note = kil_cols[r]
        put(f"B{r}", b, S_INPUT); put(f"C{r}", c, S_INPUT)
        put(f"D{r}", d, S_INPUT); put(f"E{r}", e, S_INPUT)
        put(f"F{r}", note, S_GREY)

    # ---- tanda 1: sheet table
    put("R6", "Tanda 1 di sheet kilang", S_TITLE)
    put("R7", "1) Kiri: daftar sheet kilang + pilih kilangnya.  2) Kanan: satu kebutuhan per baris "
              "(sheet, kelas, volume). Sheet kilang memberi tanda 1 di kolom volume tsb untuk crude "
              "kelas itu yang lolos semua cek.  3) Status harus ✔ OK.", S_NOTE)
    put("R8", "Sheet kilang (nama tab persis)", S_HEAD)
    put("S8", "Kilang", S_HEAD)
    put("T8", "Status", S_HEAD)
    for r in SHEET_ROWS:
        sh, kil = old_sheet_rows[r]
        put(f"R{r}", sh, S_INPUT)
        put(f"S{r}", kil, S_INPUT)
        put(f"T{r}",
            f'=IF(R{r}="","",IF(NOT(ISREF(INDIRECT("\'"&R{r}&"\'!A1"))),"⚠ Sheet tidak ditemukan",'
            f'IF(S{r}="","⚠ Pilih kilang",IF(ISNA(MATCH(S{r},$A$9:$A$16,0)),"⚠ Kilang tidak ada di tabel kilang",'
            f'IF(COUNTIF($V$9:$V$80,R{r})=0,"⚠ Belum ada kebutuhan","✔ OK ("&COUNTIF($V$9:$V$80,R{r})&" kebutuhan)")))))',
            S_CALC)

    # ---- tanda 1: one need per row
    put("V8", "Sheet kilang", S_HEAD)
    put("W8", "Kelas", S_HEAD)
    put("X8", "Volume (MB)", S_HEAD)
    put("Y8", "Status", S_HEAD)
    for i, r in enumerate(NEED_ROWS):
        sh, kel, vol = needs[i] if i < len(needs) else (None, None, None)
        put(f"V{r}", sh, S_INPUT); put(f"W{r}", kel, S_INPUT); put(f"X{r}", vol, S_INPUT)
        put(f"Y{r}",
            f'=IF(AND(V{r}="",W{r}="",X{r}=""),"",IF(OR(V{r}="",W{r}="",X{r}=""),"⚠ Isi sheet, kelas dan volume",'
            f'IF(COUNTIF($R$9:$R$24,V{r})=0,"⚠ Sheet belum ada di tabel kiri",'
            f'IF(NOT(ISREF(INDIRECT("\'"&V{r}&"\'!A1"))),"⚠ Sheet tidak ditemukan",'
            f'IF(SUMPRODUCT(--(INDIRECT("\'"&V{r}&"\'!A1:CZ1")&""=X{r}&""))=0,'
            f'"⚠ Volume "&X{r}&" tidak ada sebagai judul kolom di sheet itu",'
            f'IF(COUNTIFS($V$9:$V$80,V{r},$W$9:$W$80,W{r},$X$9:$X$80,X{r})>1,"⚠ Dobel (baris sama sudah ada)","✔ OK"))))))',
            S_CALC)

    # ---- hidden helpers (same cells the refinery sheets already read: AF:AJ rows 9-24)
    for col, txt in zip(["AF", "AG", "AH", "AI", "AJ"],
                        ["No. kilang", "Kol Batasan 1", "Kol Batasan 2", "Kol PP terima", "Kol PP CIB"]):
        put(f"{col}8", txt + " (bantu)", S_HEAD)
    for r in SHEET_ROWS:
        put(f"AF{r}", f'=IF(S{r}="","",MATCH(S{r},$A$9:$A$16,0))', S_CALC)
        for col, src_col, hdr in (("AG", "B", RASIO), ("AH", "C", RASIO), ("AI", "D", PP), ("AJ", "E", PP)):
            v = f"INDEX(${src_col}$9:${src_col}$16,$AF{r})"
            put(f"{col}{r}", f'=IF(ISNUMBER($AF{r}),IF({v}="",0,IFERROR(MATCH({v},{hdr}!$1:$1,0),0)),0)', S_CALC)

    # ---- emit Setting sheetData
    by_row: dict[int, list] = {}
    for ref, (v, s) in cells.items():
        c, r = split_ref(ref)
        by_row.setdefault(r, []).append((column_index_from_string(c), ref, v, s))
    rows_xml = []
    for r in sorted(by_row):
        attrs = row_attr.get(r, "")
        attrs = re.sub(r'\sspans="[^"]*"', "", attrs)
        cs = "".join(cell_xml(ref, v, s) for _, ref, v, s in sorted(by_row[r]))
        rows_xml.append(f'<row r="{r}"{attrs}>{cs}</row>')
    sheet_data = "<sheetData>" + "".join(rows_xml) + "</sheetData>"

    widths = {1: 34, 2: 30, 3: 30, 4: 30, 5: 30, 6: 55, 7: 2, 8: 2, 9: 20, 10: 45, 11: 2, 12: 34, 13: 40,
              14: 2, 15: 32, 16: 36, 17: 2, 18: 26, 19: 30, 20: 30, 21: 2, 22: 26, 23: 8, 24: 12, 25: 44}
    cols = "".join(f'<col min="{c}" max="{c}" width="{w}" customWidth="1"/>' for c, w in widths.items())
    cols += '<col min="32" max="36" width="9" hidden="1" customWidth="1"/>'
    cols = f"<cols>{cols}</cols>"

    dv = (
        '<dataValidations count="4">'
        '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="S9:S24"><formula1>$A$9:$A$16</formula1></dataValidation>'
        '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="V9:V80"><formula1>$R$9:$R$24</formula1></dataValidation>'
        '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="W9:W80"><formula1>"H,L,M,SH"</formula1></dataValidation>'
        '<dataValidation type="decimal" operator="greaterThan" allowBlank="1" showErrorMessage="1" '
        'errorTitle="Volume" error="Isi angka volume (MB), sama dengan judul kolom di sheet kilang." sqref="X9:X80"><formula1>0</formula1></dataValidation>'
        '</dataValidations>'
    )
    # dropdowns that read another sheet go in the x14 extension, the way Excel itself stores them
    ext = (
        '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
        '<x14:dataValidations count="2" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
        f'<x14:dataValidation type="list" allowBlank="1" showErrorMessage="1"><x14:formula1><xm:f>{x_attr(RASIO)}!$P$1:$X$1</xm:f></x14:formula1><xm:sqref>B9:C16</xm:sqref></x14:dataValidation>'
        f'<x14:dataValidation type="list" allowBlank="1" showErrorMessage="1"><x14:formula1><xm:f>{x_attr(PP)}!$F$1:$N$1</xm:f></x14:formula1><xm:sqref>D9:E16</xm:sqref></x14:dataValidation>'
        '</x14:dataValidations></ext></extLst>'
    )

    # styles: add green / red dxfs for the Status columns
    sty = files["xl/styles.xml"].decode()
    m = re.search(r'<dxfs count="(\d+)">', sty)
    n_dxf = int(m.group(1))
    green = '<dxf><font><b/><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>'
    red = '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
    sty = sty.replace(m.group(0), f'<dxfs count="{n_dxf + 2}">', 1).replace("</dxfs>", green + red + "</dxfs>", 1)
    files["xl/styles.xml"] = sty.encode()
    cf = ""
    for k, (rng, tl) in enumerate((("T9:T24", "T9"), ("Y9:Y80", "Y9"))):
        cf += (f'<conditionalFormatting sqref="{rng}">'
               f'<cfRule type="expression" dxfId="{n_dxf}" priority="{2 * k + 1}"><formula>LEFT({tl},1)="✔"</formula></cfRule>'
               f'<cfRule type="expression" dxfId="{n_dxf + 1}" priority="{2 * k + 2}"><formula>LEFT({tl},1)="⚠"</formula></cfRule>'
               f'</conditionalFormatting>')

    head = raw[: raw.index("<dimension")]
    views = re.search(r"<sheetViews>.*?</sheetViews>", raw, re.S).group(0)
    fmt = re.search(r"<sheetFormatPr[^>]*/>", raw).group(0)
    margins = re.search(r"<pageMargins[^>]*/>", raw).group(0)
    new_setting = (head + '<dimension ref="A1:AJ80"/>' + views + fmt + cols + sheet_data + cf + dv + margins + ext + "</worksheet>")
    files[part["Setting"]] = new_setting.encode()

    # ---- refinery sheets: point the tanda-1 COUNTIFS at the new list
    replaced = {}
    for r in SHEET_ROWS:
        sh = old_sheet_rows[r][0]
        if not sh or sh not in part:
            continue
        p = part[sh]
        x = files[p].decode()
        a = f"COUNTIFS(Setting!$T${r}:$AD${r},"
        b = f",Setting!$U${r}:$AE${r},"
        na, nb = x.count(a), x.count(b)
        x = x.replace(a, f"COUNTIFS(Setting!$V$9:$V$80,Setting!$R${r},Setting!$W$9:$W$80,")
        x = x.replace(b, ",Setting!$X$9:$X$80,")
        if re.search(r"Setting!\$(?:[T-Z]|A[A-E])\$", x.replace("Setting!$V$9:$V$80", "").replace("Setting!$W$9:$W$80", "").replace("Setting!$X$9:$X$80", "")):
            sys.exit(f"Unexpected remaining reference to old tanda-1 cells in {sh}")
        files[p] = x.encode()
        replaced[sh] = (na, nb)

    # ---- Cek Kilang - Detail: column numbers now come from names
    dp = part["Cek Kilang - Detail"]
    x = files[dp].decode()
    for col, src_col, hdr in (("S", "B", RASIO), ("T", "C", RASIO), ("U", "D", PP), ("V", "E", PP)):
        old = f"<f>INDEX(Setting!${src_col}$9:${src_col}$16,$R$2)</f>"
        if old not in x:
            sys.exit(f"Detail {col}2 formula not as expected")
        v = f"INDEX(Setting!${src_col}$9:${src_col}$16,$R$2)"
        new = f"IF({v}=\"\",0,IFERROR(MATCH({v},{hdr}!$1:$1,0),0))"
        x = x.replace(old, f"<f>{escape(new)}</f>")
    files[dp] = x.encode()

    # ---- recalc everything when Excel opens it
    wbx = files["xl/workbook.xml"].decode()
    if "<calcPr" in wbx:
        wbx = re.sub(r"<calcPr([^>]*)/>", lambda m: "<calcPr" + re.sub(r'\sfullCalcOnLoad="\w+"', "", m.group(1)) + ' fullCalcOnLoad="1"/>', wbx, 1)
    else:
        wbx = wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    files["xl/workbook.xml"] = wbx.encode()

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for p in order:
            z.writestr(p, files[p])
    return {"needs": needs, "replaced": replaced, "kilang_cols": kil_cols}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--cib-max", type=float, help="set 'Maks volume via CIB (MB)' (Setting B5)")
    a = ap.parse_args(argv)
    if a.input.resolve() == a.output.resolve():
        sys.exit("Refusing to overwrite the input file.")
    r = simplify(a.input, a.output, a.cib_max)
    print(f"Wrote {a.output}")
    print(f"  Tanda-1 needs moved to one-per-row list: {len(r['needs'])}")
    for sh, (na, nb) in r["replaced"].items():
        print(f"  {sh}: {na} formulas repointed")


if __name__ == "__main__":
    main()
