from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from patent_links import add_links_to_workbook, build_links, parse_number

st.set_page_config(page_title="公報番号リンク作成", page_icon=":material/link:", layout="wide")


@st.cache_data(ttl="1h", scope="session")
def read_sheets(xlsx_bytes: bytes) -> dict[str, pd.DataFrame]:
    return pd.read_excel(BytesIO(xlsx_bytes), sheet_name=None, dtype=str)


@st.cache_data(ttl="1h", scope="session")
def build_output(xlsx_bytes: bytes, sheet: str, column: str) -> bytes:
    return add_links_to_workbook(xlsx_bytes, sheet, column)


st.title(":material/link: 公報番号リンク作成")
st.caption("公報番号列から Espacenet / J-PlatPat へのリンク列を作成します。")

st.markdown(
    """
**リンク作成ルール**

| リンク先 | 対応範囲・番号の扱い |
|---|---|
| Espacenet | 公報番号として解析できる国コード付き番号を検索します。日本の登録公報は先頭の0を除去、韓国の登録公報は7桁番号に`10`を付加、米国の公開公報は年の後の先頭0を除去します。検索結果はEspacenetの収録状況に依存します。 |
| J-PlatPat（日本） | 公開・公表公報（A）は種別`11`、特許登録公報（B/B1/B2）は`15`で照会します。 |
| J-PlatPat（外国） | WO（A）、US（A/B）、EP（A/B）、CN（A/B/C/U/Y）、KR（A/B）にリンクを作成します。外国公報の種別コードは`50`です。 |
| J-PlatPat（対象外） | 上記以外の国・種別（例: TW）はリンク欄を空欄にします。 |

リンクは各サービスの照会画面を開きます。対象文献がデータベースに収録されていない場合、検索結果が表示されないことがあります。
"""
)

uploaded = st.file_uploader("Excelファイル (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.stop()

data = uploaded.getvalue()
sheets = read_sheets(data)

left, right = st.columns(2)
sheet = left.selectbox("シート", list(sheets))
df = sheets[sheet]
columns = list(df.columns)
column = right.selectbox(
    "公報番号列", columns, index=columns.index("公報番号") if "公報番号" in columns else 0
)

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

st.download_button(
    "リンク付きExcelをダウンロード",
    data=build_output(data, sheet, column),
    file_name=f"{Path(uploaded.name).stem}_with_links.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    icon=":material/download:",
    type="primary",
)
