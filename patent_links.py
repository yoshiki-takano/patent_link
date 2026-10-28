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
# Last 6-digit KR numbers Espacenet stores in the old "0"+6-digit form (verified on Espacenet).
KR_OLD_GRANT_MAX = 186214
KR_OLD_UTILITY_MAX = 140846
# Last S64 serial per (kind group, 5xxxxx series); source data may write S64 numbers with year "1".
JP_LAST_S64 = {
    ("A", False): 91300,
    ("A", True): 500558,
    ("B", False): 12440,
    ("Y", False): 7600,
    ("U", False): 57700,
    ("U", True): 500004,
}

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
    country, number, kind = m.groups()
    if country == "JP" and kind == "K1":
        # JP K1 follows the same conversion rules as A.
        kind = "A"
    return PatentNumber(country, number, kind)


def _is_granted(p: PatentNumber) -> bool:
    return p.kind.startswith(("B", "C"))


def _japanese_era_parts(p: PatentNumber) -> tuple[str, str, str] | None:
    """Return (era, 2-digit year, 6-digit serial) for Showa/Heisei A, B, Y and U numbers."""
    n = p.number
    if p.country != "JP" or n.startswith("0") or len(n) not in (7, 8):
        return None
    group = p.kind[0]
    if group == "B":
        if len(n) == 7 and p.kind not in ("B1", "B2"):
            return None
    elif p.kind not in ("A", "U", "Y", "Y1", "Y2"):
        return None
    if len(n) == 7:
        year, serial = n[0], n[1:]
        if year == "1" and int(serial) <= JP_LAST_S64.get((group, serial.startswith("5")), 0):
            return "S", "64", serial
        return "H", f"0{year}", serial
    year, serial = n[:2], n[2:]
    if group == "A" and year in ("10", "11"):
        return "H", year, serial
    return "S", year, serial


def _normalize_korean_year_number(number: str) -> str:
    if len(number) == 9 and number[:2] in ("19", "20"):
        return number[2:4] + number[4:].zfill(7)
    return number


def _normalize_korean_grant_number(number: str) -> str:
    number = _normalize_korean_year_number(number)
    if len(number) <= 5:
        # Only KR13B1 exists; Espacenet lists it as KR1000013B1.
        return "1" + number.zfill(6)
    if len(number) == 6:
        return ("0" if int(number) <= KR_OLD_GRANT_MAX else "100") + number
    if len(number) == 7:
        return "10" + number
    return number


def espacenet_number(p: PatentNumber) -> str:
    number = p.number
    if p.country == "DD":
        return f"{p.country}{number}"
    era_parts = _japanese_era_parts(p)
    if era_parts:
        era, year, serial = era_parts
        return f"JP{era}{year}{serial.lstrip('0') or '0'}{p.kind}"
    if p.country == "JP" and p.kind == "A1" and len(number) == 10 and number[:4] > "2002":
        # JP A1 = 再公表特許; Espacenet uses JPWO only from 2003, earlier ones stay JP.
        return f"JPWO{number}{p.kind}"
    if p.country == "JP" and p.kind == "U" and len(number) == 8:
        number = number.lstrip("0")
    elif p.country == "JP" and _is_granted(p):
        number = number.lstrip("0")
    elif p.country == "KR" and _is_granted(p):
        number = _normalize_korean_grant_number(number)
    elif p.country == "KR" and p.kind == "Y1" and len(number) == 6:
        number = ("0" if int(number) <= KR_OLD_UTILITY_MAX else "200") + number
    elif p.country == "KR" and p.kind == "U" and len(number) == 9 and number[:2] in ("19", "20"):
        number = number[:4] + number[4:].zfill(7)
    elif p.country == "KR" and p.kind.startswith("A") and len(number) == 10:
        # Source data drops a leading zero of the 7-digit serial: KR2026125202A -> KR20260125202A.
        number = number[:4] + number[4:].zfill(7)
    elif p.country == "EP" and p.kind in ("A1", "A2") and len(number) in (4, 5):
        number = number.zfill(7)
    elif p.country == "CN" and p.kind == "A" and len(number) == 10 and number[:2] in ("19", "20"):
        number = number[2:]
    elif p.country == "US" and p.kind.startswith("A") and len(number) == 11:
        # Espacenet (DOCDB) uses a 6-digit serial: US2016222235A1, not US20160222235A1.
        if not (p.kind == "A1" and number[:4] >= "2024"):
            number = number[:4] + number[4:].lstrip("0").zfill(6)
    return f"{p.country}{number}{p.kind}"


def espacenet_url(p: PatentNumber | None) -> str | None:
    if p is None:
        return None
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
    if p.country == "EP" and p.kind in ("A1", "A2") and 4 <= len(n) <= 7:
        return f"EP-A-{n.zfill(9)}"
    if p.country == "EP" and kind in "AB":
        return f"EP-{kind}-{n.zfill(8)}"
    if p.country == "CN" and kind == "A" and len(n) == 10 and n[:2] in ("19", "20"):
        return f"CN-A-0{n[2:]}"
    if p.country == "CN" and kind in "AY" and len(n) == 7:
        return f"CN-{kind}-{n.zfill(9)}"
    if p.country == "CN" and kind in "ABCUY":
        return f"CN-{kind}-{n}"
    if p.country == "KR":
        # J-PlatPat KR publication = 2-digit year + 7-digit serial.
        if kind == "A" and len(n) in (10, 11):
            return f"KR-A-{n[2:4]}{n[4:].zfill(7)}"
        if kind == "Y" and len(n) == 6:
            return f"KR-Y-{n.zfill(9)}"
        if kind == "U" and len(n) == 9:
            return f"KR-U-{_normalize_korean_year_number(n)}"
        if kind == "B":
            n = _normalize_korean_year_number(n)
            # J-PlatPat accepts zero-padded grant numbers without the "10" right code.
            return f"KR-B-{n.zfill(9) if len(n) <= 7 else n}"
    if p.country in JPLATPAT_OTHER_COUNTRIES:
        return f"{p.country}-{kind}-{n.zfill(9)}"
    return None


def jplatpat_url(p: PatentNumber | None) -> str | None:
    if p is None:
        return None
    if p.country == "JP":
        n = p.number
        era_parts = _japanese_era_parts(p)
        if era_parts:
            era, year, serial = era_parts
            if (era, year) == ("S", "64"):
                # J-PlatPat files all 1989 documents under H01.
                era, year = "H", "01"
            group = p.kind[0]
            if group == "U":
                code = "21" if serial.startswith("5") else "23"
            else:
                code = {"A": "11", "B": "12", "Y": "22"}[group]
            return f"{JPLATPAT_BASE}/JP-{era}{year}-{serial}/{code}/ja"
        if _is_granted(p):
            return f"{JPLATPAT_BASE}/JP-{p.number.lstrip('0').zfill(7)}/15/ja"
        if p.kind == "U" and len(n) == 8:
            return f"{JPLATPAT_BASE}/JP-{n.lstrip('0')}/25/ja"
        if p.kind == "A1" and len(n) == 10:
            return f"{JPLATPAT_BASE}/JP-{n[:4]}-{n[4:]}/19/ja"
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
