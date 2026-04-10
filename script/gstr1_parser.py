#!/usr/bin/env python3
"""Consolidate all GSTR-1 JSON sections into one multi-sheet Excel workbook."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape
import zipfile


NUMERIC_KEYS = {
    "TaxableValue",
    "IGST",
    "CGST",
    "SGST",
    "CESS",
    "Rate",
    "Quantity",
    "ExemptAmount",
    "NilRatedAmount",
    "NonGSTAmount",
}


def n(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def col_ref(idx: int) -> str:
    s = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def xml_cell(ref: str, value: Any, force_number: bool = False) -> str:
    if value is None:
        value = ""
    if force_number:
        return f'<c r="{ref}"><v>{n(value)}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"><v>{value}</v></c>'
    text = escape(str(value))
    return f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'


def write_xlsx(path: Path, sheets: dict[str, list[dict[str, Any]]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        workbook_xml = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            '<sheets>',
        ]
        workbook_rels = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        ]
        content_types = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
            '<Default Extension="xml" ContentType="application/xml"/>',
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        ]

        root_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            '</Relationships>'
        )

        for i, (sheet_name, rows) in enumerate(sheets.items(), start=1):
            headers = list(rows[0].keys()) if rows else []
            sheet_lines = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
                '<sheetData>',
            ]

            if headers:
                cells = [xml_cell(f"{col_ref(c)}1", h) for c, h in enumerate(headers, start=1)]
                sheet_lines.append(f'<row r="1">{"".join(cells)}</row>')

            for r_i, row in enumerate(rows, start=2):
                cells = []
                for c_i, h in enumerate(headers, start=1):
                    ref = f"{col_ref(c_i)}{r_i}"
                    cells.append(xml_cell(ref, row.get(h, ""), force_number=h in NUMERIC_KEYS))
                sheet_lines.append(f'<row r="{r_i}">{"".join(cells)}</row>')

            sheet_lines.extend(["</sheetData>", "</worksheet>"])
            zf.writestr(f"xl/worksheets/sheet{i}.xml", "".join(sheet_lines))

            workbook_xml.append(f'<sheet name="{escape(sheet_name)}" sheetId="{i}" r:id="rId{i}"/>')
            workbook_rels.append(
                f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                f'Target="worksheets/sheet{i}.xml"/>'
            )
            content_types.append(
                f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            )

        workbook_xml.extend(["</sheets>", "</workbook>"])
        workbook_rels.append("</Relationships>")
        content_types.append("</Types>")

        zf.writestr("[Content_Types].xml", "".join(content_types))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", "".join(workbook_xml))
        zf.writestr("xl/_rels/workbook.xml.rels", "".join(workbook_rels))


def deduplicate_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[tuple[tuple[str, Any], ...]] = set()
    unique: list[dict[str, Any]] = []
    for row in rows:
        key = tuple(sorted(row.items()))
        if key not in seen:
            seen.add(key)
            unique.append(row)
    return unique


def parse_b2b(data: dict[str, Any], month: str, source_file: str) -> list[dict[str, Any]]:
    rows = []
    for party in data.get("b2b", []) or []:
        for inv in party.get("inv", []) or []:
            for item in inv.get("itms", []) or []:
                d = item.get("itm_det", {}) or {}
                rows.append({"Month": month, "SourceFile": source_file, "CTIN": party.get("ctin"), "InvoiceNumber": inv.get("inum"), "InvoiceDate": inv.get("idt"), "POS": inv.get("pos"), "Rate": n(d.get("rt")), "TaxableValue": n(d.get("txval")), "IGST": n(d.get("iamt")), "CGST": n(d.get("camt")), "SGST": n(d.get("samt")), "CESS": n(d.get("csamt"))})
    return rows


def parse_b2cl(data: dict[str, Any], month: str, source_file: str) -> list[dict[str, Any]]:
    rows = []
    for entry in data.get("b2cl", []) or []:
        for inv in entry.get("inv", []) or []:
            for item in inv.get("itms", []) or []:
                d = item.get("itm_det", {}) or {}
                rows.append({"Month": month, "SourceFile": source_file, "InvoiceNumber": inv.get("inum"), "InvoiceDate": inv.get("idt"), "POS": inv.get("pos", entry.get("pos")), "Rate": n(d.get("rt")), "TaxableValue": n(d.get("txval")), "IGST": n(d.get("iamt")), "CGST": n(d.get("camt")), "SGST": n(d.get("samt")), "CESS": n(d.get("csamt"))})
    return rows


def parse_b2cs(data: dict[str, Any], month: str, source_file: str) -> list[dict[str, Any]]:
    rows = []
    for item in data.get("b2cs", []) or []:
        rows.append({"Month": month, "SourceFile": source_file, "POS": item.get("pos"), "SupplyType": item.get("sply_ty") or item.get("typ"), "Rate": n(item.get("rt")), "TaxableValue": n(item.get("txval")), "IGST": n(item.get("iamt")), "CGST": n(item.get("camt")), "SGST": n(item.get("samt")), "CESS": n(item.get("csamt"))})
    return rows


def parse_exports(data: dict[str, Any], month: str, source_file: str) -> list[dict[str, Any]]:
    rows = []
    for entry in data.get("exp", []) or []:
        for inv in entry.get("inv", []) or []:
            for item in inv.get("itms", []) or []:
                d = item.get("itm_det", item) or {}
                rows.append({"Month": month, "SourceFile": source_file, "ExportType": entry.get("exp_typ"), "InvoiceNumber": inv.get("inum"), "InvoiceDate": inv.get("idt"), "POS": inv.get("pos") or entry.get("pos"), "Rate": n(d.get("rt")), "TaxableValue": n(d.get("txval")), "IGST": n(d.get("iamt")), "CGST": n(d.get("camt")), "SGST": n(d.get("samt")), "CESS": n(d.get("csamt"))})
    return rows


def parse_cdnr(data: dict[str, Any], month: str, source_file: str) -> list[dict[str, Any]]:
    rows = []
    for party in data.get("cdnr", []) or []:
        for note in party.get("nt", []) or []:
            for item in note.get("itms", []) or []:
                d = item.get("itm_det", {}) or {}
                rows.append({"Month": month, "SourceFile": source_file, "CTIN": party.get("ctin"), "DocumentNumber": note.get("nt_num"), "DocumentDate": note.get("nt_dt"), "OriginalInvoiceNumber": note.get("inum"), "OriginalInvoiceDate": note.get("idt"), "POS": note.get("pos"), "Rate": n(d.get("rt")), "TaxableValue": n(d.get("txval")), "IGST": n(d.get("iamt")), "CGST": n(d.get("camt")), "SGST": n(d.get("samt")), "CESS": n(d.get("csamt"))})
    return rows


def parse_cdnur(data: dict[str, Any], month: str, source_file: str) -> list[dict[str, Any]]:
    rows = []
    for note in data.get("cdnur", []) or []:
        for item in note.get("itms", []) or []:
            d = item.get("itm_det", {}) or {}
            rows.append({"Month": month, "SourceFile": source_file, "DocumentNumber": note.get("nt_num"), "DocumentDate": note.get("nt_dt"), "OriginalInvoiceNumber": note.get("inum"), "OriginalInvoiceDate": note.get("idt"), "POS": note.get("pos"), "Rate": n(d.get("rt")), "TaxableValue": n(d.get("txval")), "IGST": n(d.get("iamt")), "CGST": n(d.get("camt")), "SGST": n(d.get("samt")), "CESS": n(d.get("csamt"))})
    return rows


def parse_hsn(data: dict[str, Any], month: str, source_file: str) -> list[dict[str, Any]]:
    rows = []
    for item in ((data.get("hsn") or {}).get("data", []) or []):
        rows.append({"Month": month, "SourceFile": source_file, "HSNCode": item.get("hsn_sc"), "Description": item.get("desc"), "UQC": item.get("uqc"), "Quantity": n(item.get("qty")), "Rate": n(item.get("rt")), "TaxableValue": n(item.get("txval")), "IGST": n(item.get("iamt")), "CGST": n(item.get("camt")), "SGST": n(item.get("samt")), "CESS": n(item.get("csamt"))})
    return rows


def parse_nil(data: dict[str, Any], month: str, source_file: str) -> list[dict[str, Any]]:
    rows = []
    for item in ((data.get("nil") or {}).get("inv", []) or []):
        expt = n(item.get("expt_amt"))
        nilv = n(item.get("nil_amt"))
        ngst = n(item.get("ngsup_amt"))
        rows.append({"Month": month, "SourceFile": source_file, "SupplyType": item.get("sply_ty"), "TaxableValue": expt + nilv + ngst, "ExemptAmount": expt, "NilRatedAmount": nilv, "NonGSTAmount": ngst, "IGST": 0.0, "CGST": 0.0, "SGST": 0.0, "CESS": 0.0})
    return rows


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    data_dir = repo_root / "data"
    if not data_dir.exists():
        data_dir = repo_root / "Data"

    files = sorted(data_dir.glob("*.json"))
    sheet_rows = {"B2B": [], "B2CL": [], "B2CS": [], "EXPORTS": [], "B2B_CN": [], "B2C_CN": [], "HSN": [], "NIL": []}

    for file in files:
        payload = json.loads(file.read_text(encoding="utf-8"))
        month = payload.get("fp", "")
        source = file.name

        sheet_rows["B2B"].extend(parse_b2b(payload, month, source))
        sheet_rows["B2CL"].extend(parse_b2cl(payload, month, source))
        sheet_rows["B2CS"].extend(parse_b2cs(payload, month, source))
        sheet_rows["EXPORTS"].extend(parse_exports(payload, month, source))
        sheet_rows["B2B_CN"].extend(parse_cdnr(payload, month, source))
        sheet_rows["B2C_CN"].extend(parse_cdnur(payload, month, source))
        sheet_rows["HSN"].extend(parse_hsn(payload, month, source))
        sheet_rows["NIL"].extend(parse_nil(payload, month, source))

    sheet_rows = {k: deduplicate_rows(v) for k, v in sheet_rows.items()}

    output_file = Path("/output/GSTR1_Consolidated.xlsx")
    write_xlsx(output_file, sheet_rows)

    print(f"Files processed: {len(files)}")
    for name, rows in sheet_rows.items():
        print(f"{name}: {len(rows)} rows")
    print(f"Output written to: {output_file}")


if __name__ == "__main__":
    main()
