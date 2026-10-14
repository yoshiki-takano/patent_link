from hashlib import sha256
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from openpyxl import load_workbook

from patent_links import add_links_to_workbook, build_links, parse_number

st.set_page_config(page_title="公報番号リンク作成", page_icon=":material/link:", layout="wide")


@st.cache_data(ttl="1h", scope="session")
def read_sheets(xlsx_bytes: bytes) -> dict[str, list[list[str | None]]]:
    # Read raw rows (no header) so row numbers match Excel exactly.
    wb = load_workbook(BytesIO(xlsx_bytes), read_only=True, data_only=True)
    sheets = {
        ws.title: [[None if v is None else str(v) for v in row] for row in ws.iter_rows(values_only=True)]
        for ws in wb.worksheets
    }
    wb.close()
    return sheets


@st.cache_data(ttl="1h", scope="session")
def build_output(
    xlsx_bytes: bytes,
    sheet: str,
    src_col: int,
    header_row: int,
    link_rules_signature: str,
) -> bytes:
    return add_links_to_workbook(xlsx_bytes, sheet, src_col, header_row)


def detect_header_row(rows: list[list[str | None]], max_rows: int = 20) -> int:
    for i, row in enumerate(rows[:max_rows], start=1):
        if any(v is not None and v.strip() == "公報番号" for v in row):
            return i
    return 1


st.title(":material/link: 公報番号リンク作成")
st.caption("公報番号列から Espacenet / J-PlatPat へのリンク列を作成します。ただし、欠番、未蓄積、落丁の場合があります。")

st.markdown(
    """
**リンク作成ルール**

| リンク先 | 対応範囲・番号の扱い |
|---|---|
| Espacenet | 公報番号として解析できる国コード付き番号を検索します。日本の登録公報は先頭の0を除去、韓国の登録公報は西暦4桁の年を2桁にし通番を7桁に補正（例: KR199513437B1 → KR950013437B1）、米国の公開公報は年の後の先頭0を除去します。検索結果はEspacenetの収録状況に依存します。 |
| J-PlatPat（日本） | 公開・公表公報（A）は種別`11`、特許登録公報（B/B1/B2）は`15`、昭和の特許出願公告（例: JP46012531B → 特公昭46-012531）は`12`で照会します。 |
| J-PlatPat（外国） | WO（A）、US（A/B）、EP（A/B）、CN（A/B/C/U/Y）、KR（A/B/U。KR-Bは例: KR199513437B1 → KR-B-950013437、KR-Uは例: KR201405583U → KR-U-140005583）、GB・DE・DD・FR・CH・CAにリンクを作成します。外国公報の種別コードは`50`です。 |
| J-PlatPat（対象外） | 上記以外の国・種別（例: TW）はリンク欄を空欄にします。 |

リンクは各サービスの照会画面を開きます。対象文献がデータベースに収録されていない場合、検索結果が表示されないことがあります。
"""
)

uploaded = st.file_uploader("Excelファイル (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.stop()

data = uploaded.getvalue()
sheets = read_sheets(data)

c1, c2, c3 = st.columns(3)
sheet = c1.selectbox("シート", list(sheets))
rows = sheets[sheet]
if not rows:
    st.warning("選択したシートにデータがありません。")
    st.stop()

header_row = c2.number_input(
    "見出し行（Excelの行番号）",
    min_value=1,
    max_value=len(rows),
    value=detect_header_row(rows),
    key=f"header_row_{sheet}",
    help="「公報番号」を含む行を自動で初期選択します。",
)
width = max(len(r) for r in rows)
headers = rows[header_row - 1] + [None] * (width - len(rows[header_row - 1]))
labels = [h.strip() if h and h.strip() else f"（無題）{i + 1}列目" for i, h in enumerate(headers)]
col_idx = c3.selectbox(
    "公報番号列",
    range(width),
    index=labels.index("公報番号") if "公報番号" in labels else 0,
    format_func=lambda i: labels[i],
    key=f"column_{sheet}_{header_row}",
)
column = labels[col_idx]

values = pd.Series(
    [r[col_idx] if col_idx < len(r) else None for r in rows[header_row:]], dtype=object
)
df = pd.DataFrame({column: values})

links = build_links(df[column])
result = pd.DataFrame(
    {
        column: df[column],
        "Espacenet": [e for e, _ in links],
        "J-PlatPat": [j for _, j in links],
    }
)

total = int(df[column].notna().sum())
unparsed = int(sum(v is not None and not pd.isna(v) and parse_number(v) is None for v in df[column]))
no_jpp = int(result["J-PlatPat"].isna().sum() - df[column].isna().sum() - unparsed)

m1, m2, m3 = st.columns(3)
m1.metric("公報番号", total, border=True)
m2.metric("解析失敗", unparsed, border=True)
m3.metric("J-PlatPat 未対応", no_jpp, border=True)

st.dataframe(
    result,
    column_config={
        "Espacenet": st.column_config.LinkColumn("Espacenet", display_text="Espacenet で開く"),
        "J-PlatPat": st.column_config.LinkColumn("J-PlatPat", display_text="J-PlatPat で開く"),
    },
    hide_index=True,
)

try:
    link_rules_signature = sha256(Path(__file__).with_name("patent_links.py").read_bytes()).hexdigest()
    output = build_output(data, sheet, col_idx + 1, header_row, link_rules_signature)
except ValueError as e:
    # Streamlit Cloud redacts exception text; openpyxl messages only contain row/column numbers.
    st.error(
        f"リンク付きExcelの作成に失敗しました（シート: {sheet}, 見出し行: {header_row!r}, "
        f"列: {col_idx + 1!r}, 列数: {width}, 行数: {len(rows)}）: {e}"
    )
    st.stop()

st.download_button(
    "リンク付きExcelをダウンロード",
    data=output,
    file_name=f"{Path(uploaded.name).stem}_with_links.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    icon=":material/download:",
    type="primary",
)
