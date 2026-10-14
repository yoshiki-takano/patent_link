import re
from copy import copy
from dataclasses import dataclass
from io import BytesIO
from urllib.parse import quote

from openpyxl import load_workbook
from openpyxl.styles import Font

ESPACENET_BASE = "https://worldwide.espacenet.com/patent/search?q=pn%3D"
JPLATPAT_BASE = "https://www.j-platpat.inpit.go.jp/c1801/PU"
# Other countries listed in J-PlatPat's 文献蓄積情報; IDs use a 9-digit zero-padded number.
JPLATPAT_OTHER_COUNTRIES = {"GB", "DE", "DD", "FR", "CH", "CA"}

_PATTERN = re.compile(r"^([A-Z]{2})(\d+)([A-Z]\d?)$")


@dataclass(frozen=True)
class PatentNumber:
    country: str
    number: str
    kind: str


def parse_number(raw) -> PatentNumber | None:
    if raw is None:
        return None
    text = re.sub(r"[\s/,\-]", "", str(raw)).upper()
    m = _PATTERN.match(text)
    if not m:
        return None
    return PatentNumber(*m.groups())


def _is_granted(p: PatentNumber) -> bool:
    return p.kind.startswith(("B", "C"))


def _is_showa_examined(p: PatentNumber) -> bool:
    # 特公昭 is YY+6 digits; JP registration numbers in this data start with 0.
    return p.country == "JP" and p.kind.startswith("B") and len(p.number) == 8 and not p.number.startswith("0")


def _japanese_a_publication_parts(p: PatentNumber) -> tuple[str, str] | None:
    if p.country != "JP" or p.kind != "A" or len(p.number) not in (7, 8):
        return None
    return p.number[:-6].zfill(2), p.number[-6:]


def _normalize_korean_year_number(number: str) -> str:
    if len(number) == 9 and number[:2] in ("19", "20"):
        return number[2:4] + number[4:].zfill(7)
    return number


def _normalize_korean_grant_number(number: str) -> str:
    number = _normalize_korean_year_number(number)
    if len(number) == 6:
        return "100" + number
    if len(number) == 7:
        return "10" + number
    return number


def espacenet_number(p: PatentNumber) -> str:
    number = p.number
    if p.country == "DD":
        return f"{p.country}{number}"
    jp_a_parts = _japanese_a_publication_parts(p)
    if jp_a_parts:
        year, serial = jp_a_parts
        return f"JPH{year}{serial}{p.kind}"
    if p.country == "JP" and p.kind == "U" and len(number) == 8:
        number = number.lstrip("0")
    elif p.country == "JP" and _is_granted(p):
        number = number.lstrip("0")
    elif p.country == "KR" and _is_granted(p):
        number = _normalize_korean_grant_number(number)
    elif p.country == "KR" and p.kind == "Y1" and len(number) == 6:
        number = "200" + number
    elif p.country == "KR" and p.kind == "U":
        number = _normalize_korean_year_number(number)
    elif p.country == "KR" and p.kind.startswith("A") and len(number) == 10:
        # Source data drops a leading zero of the 7-digit serial: KR2026125202A -> KR20260125202A.
        number = number[:4] + number[4:].zfill(7)
    elif p.country == "US" and p.kind.startswith("A") and len(number) == 11:
        # Espacenet (DOCDB) uses a 6-digit serial: US2016222235A1, not US20160222235A1.
        number = number[:4] + number[4:].lstrip("0").zfill(6)
    return f"{p.country}{number}{p.kind}"


def espacenet_url(p: PatentNumber | None) -> str | None:
    if p is None:
        return None
    if _is_showa_examined(p):
        # DOCDB form is JPS4612531B1/B2; listing both kinds avoids matching the Y1 utility model.
        base = f"JPS{p.number[:2]}{int(p.number[2:])}"
        return ESPACENET_BASE + quote(f"{base}B1 OR pn={base}B2")
    return ESPACENET_BASE + quote(espacenet_number(p))


def _jplatpat_foreign_id(p: PatentNumber) -> str | None:
    kind, n = p.kind[0], p.number
    if p.country == "WO" and kind == "A" and len(n) == 10:
        return f"WO-A-{n[:4]}-{n[4:]}"
    if p.country == "US":
        if kind == "A" and len(n) == 7:
            return f"US-A-{n.zfill(9)}"
        if kind == "A" and len(n) == 11:
            return f"US-A-{n[:4]}-{n[4:]}"
        if kind == "B":
            return f"US-B-{n.zfill(9)}"
    if p.country == "EP" and kind in "AB":
        return f"EP-{kind}-{n.zfill(8)}"
    if p.country == "CN" and kind in "ABCUY":
        return f"CN-{kind}-{n}"
    if p.country == "KR":
        # J-PlatPat KR publication = 2-digit year + 7-digit serial.
        if kind == "A" and len(n) in (10, 11):
            return f"KR-A-{n[2:4]}{n[4:].zfill(7)}"
        if kind == "Y" and len(n) == 6:
            return f"KR-Y-200{n}"
        if kind == "U" and len(n) == 9:
            return f"KR-U-{_normalize_korean_year_number(n)}"
        if kind == "B":
            return f"KR-B-{_normalize_korean_grant_number(n)}"
    if p.country in JPLATPAT_OTHER_COUNTRIES:
        return f"{p.country}-{kind}-{n.zfill(9)}"
    return None


def jplatpat_url(p: PatentNumber | None) -> str | None:
    if p is None:
        return None
    if p.country == "JP":
        n = p.number
        if _is_showa_examined(p):
            return f"{JPLATPAT_BASE}/JP-S{n[:2]}-{n[2:]}/12/ja"
        if _is_granted(p):
            return f"{JPLATPAT_BASE}/JP-{p.number.lstrip('0').zfill(7)}/15/ja"
        if p.kind == "U" and len(n) == 8:
            return f"{JPLATPAT_BASE}/JP-{n.lstrip('0')}/25/ja"
        jp_a_parts = _japanese_a_publication_parts(p)
        if jp_a_parts:
            year, serial = jp_a_parts
            return f"{JPLATPAT_BASE}/JP-H{year}-{serial}/11/ja"
        if p.kind.startswith("A") and len(p.number) == 10:
            return f"{JPLATPAT_BASE}/JP-{p.number[:4]}-{p.number[4:]}/11/ja"
        return None
    doc_id = _jplatpat_foreign_id(p)
    return f"{JPLATPAT_BASE}/{doc_id}/50/ja" if doc_id else None


def build_links(values) -> list[tuple[str | None, str | None]]:
    links = []
    for v in values:
        p = parse_number(v)
        links.append((espacenet_url(p), jplatpat_url(p)))
    return links


def add_links_to_workbook(
    xlsx_bytes: bytes, sheet_name: str, src_col: int, header_row: int = 1
) -> bytes:
    """src_col and header_row are 1-based."""
    wb = load_workbook(BytesIO(xlsx_bytes))
    ws = wb[sheet_name]

    header_cell = ws.cell(row=header_row, column=src_col)
    esp_col = ws.max_column + 1
    jpp_col = esp_col + 1

    for col, title in ((esp_col, "Espacenet"), (jpp_col, "J-PlatPat")):
        cell = ws.cell(row=header_row, column=col, value=title)
        font = copy(header_cell.font)
        font.name = "Meiryo UI"
        font.sz = 10
        font.scheme = None
        cell.font = font
        cell.fill = copy(header_cell.fill)
        cell.border = copy(header_cell.border)
        alignment = copy(header_cell.alignment)
        alignment.vertical = "top"
        cell.alignment = alignment
        ws.column_dimensions[cell.column_letter].width = 20

    for row in range(header_row + 1, ws.max_row + 1):
        p = parse_number(ws.cell(row=row, column=src_col).value)
        for col, url, label in (
            (esp_col, espacenet_url(p), "Espacenet で開く"),
            (jpp_col, jplatpat_url(p), "J-PlatPat で開く"),
        ):
            # HYPERLINK formulas avoid Excel's 65,530 hyperlinks-per-sheet limit on large lists.
            value = f'=HYPERLINK("{url}","{label}")' if url else None
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = Font(
                name="Meiryo UI",
                size=10,
                color="0000FF" if url else None,
                underline="single" if url else None,
            )
            alignment = copy(cell.alignment)
            alignment.vertical = "top"
            cell.alignment = alignment

    wb.calculation.fullCalcOnLoad = True
    out = BytesIO()
    wb.save(out)
    return out.getvalue()
