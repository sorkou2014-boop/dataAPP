import io
import re
import time
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(page_title="行動檢修平台通用資料整理", layout="wide")


META_LABELS = {
    "工單編號": ["工單編號", "工單號碼"],
    "車號/最小成本": ["車號/最小成本單位", "車號", "最小成本單位"],
    "工項代碼": ["工項代碼"],
    "工項名稱": ["工項名稱"],
    "檢查開始日期": ["檢查開始日期"],
    "檢查結束日期": ["檢查結束日期"],
}

DETAIL_COLUMNS = [
    "進階分類",
    "檢查項目",
    "SCI",
    "檢查項目備註",
    "儀器編號",
    "軌道里程",
    "設備編號",
    "設備子編號",
    "感應點位置",
    "異常",
    "檢查結果",
    "單位",
    "異常原因",
    "處理對策",
    "處理說明",
    "備註",
    "執行者",
    "領班確認/SCI",
]


def read_excel(uploaded_file, header=None):
    uploaded_file.seek(0)
    try:
        return pd.read_excel(uploaded_file, header=header, engine="calamine")
    except Exception:
        uploaded_file.seek(0)
        try:
            return pd.read_excel(uploaded_file, header=header)
        except Exception:
            uploaded_file.seek(0)
            return read_xlsx_xml(uploaded_file, header=header)


def column_index(cell_ref):
    match = re.match(r"([A-Z]+)", cell_ref or "")
    if not match:
        return 0

    number = 0
    for char in match.group(1):
        number = number * 26 + ord(char) - ord("A") + 1
    return number - 1


def row_index(cell_ref):
    match = re.search(r"(\d+)", cell_ref or "")
    return int(match.group(1)) - 1 if match else 0


def read_shared_strings(zip_file, ns):
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []

    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    shared = []
    for item in root.findall("a:si", ns):
        texts = [node.text or "" for node in item.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]
        shared.append("".join(texts))
    return shared


def first_sheet_path(zip_file, ns):
    workbook = ET.fromstring(zip_file.read("xl/workbook.xml"))
    rels = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    first_sheet = workbook.find("a:sheets/a:sheet", ns)
    rel_id = first_sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
    target = rel_map[rel_id].lstrip("/")
    return target if target.startswith("xl/") else f"xl/{target}"


def cell_value(cell, shared, ns):
    cell_type = cell.attrib.get("t")

    if cell_type == "s":
        value = cell.find("a:v", ns)
        return shared[int(value.text)] if value is not None and value.text else ""

    if cell_type == "inlineStr":
        texts = [node.text or "" for node in cell.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]
        return "".join(texts)

    value = cell.find("a:v", ns)
    return value.text if value is not None and value.text is not None else ""


def read_xlsx_xml(uploaded_file, header=None):
    ns = {
        "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    data = uploaded_file.read()

    with zipfile.ZipFile(io.BytesIO(data)) as zip_file:
        shared = read_shared_strings(zip_file, ns)
        sheet_path = first_sheet_path(zip_file, ns)
        root = ET.fromstring(zip_file.read(sheet_path))

        cells = {}
        max_row = 0
        max_col = 0
        for cell in root.findall(".//a:c", ns):
            ref = cell.attrib.get("r", "")
            r_idx = row_index(ref)
            c_idx = column_index(ref)
            cells[(r_idx, c_idx)] = cell_value(cell, shared, ns)
            max_row = max(max_row, r_idx)
            max_col = max(max_col, c_idx)

    rows = []
    for r_idx in range(max_row + 1):
        rows.append([cells.get((r_idx, c_idx), "") for c_idx in range(max_col + 1)])

    if header is None:
        return pd.DataFrame(rows)

    headers = [normalize_text(value) or f"欄位{idx + 1}" for idx, value in enumerate(rows[header])]
    return pd.DataFrame(rows[header + 1 :], columns=headers)


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_date(value):
    if pd.isna(value) or normalize_text(value) == "":
        return ""
    try:
        return pd.to_datetime(value).strftime("%Y-%m-%d")
    except Exception:
        return normalize_text(value)


def extract_numeric(value):
    text = normalize_text(value).replace(",", "")
    if text == "":
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group(0)) if match else None


def find_next_value(df_raw, row_idx, col_idx):
    for offset in range(1, 4):
        next_col = col_idx + offset
        if next_col < len(df_raw.columns):
            value = df_raw.iloc[row_idx, next_col]
            if normalize_text(value) != "":
                return value
    return ""


def extract_metadata(df_raw):
    metadata = {key: "" for key in META_LABELS}

    for row_idx in range(len(df_raw)):
        for col_idx in range(len(df_raw.columns)):
            cell = normalize_text(df_raw.iloc[row_idx, col_idx])
            if not cell:
                continue

            for target_key, aliases in META_LABELS.items():
                if cell in aliases and metadata[target_key] == "":
                    metadata[target_key] = find_next_value(df_raw, row_idx, col_idx)

    metadata["工單編號"] = normalize_text(metadata["工單編號"])
    metadata["車號/最小成本"] = normalize_text(metadata["車號/最小成本"])
    metadata["工項代碼"] = normalize_text(metadata["工項代碼"])
    metadata["工項名稱"] = normalize_text(metadata["工項名稱"])
    metadata["檢查開始日期"] = normalize_date(metadata["檢查開始日期"])
    metadata["檢查結束日期"] = normalize_date(metadata["檢查結束日期"])
    return metadata


def locate_header_row(df_raw):
    required = {"進階分類", "檢查項目", "檢查結果"}
    for row_idx in range(min(len(df_raw), 30)):
        values = {normalize_text(v) for v in df_raw.iloc[row_idx].tolist()}
        if required.issubset(values):
            return row_idx
    return 2


def standardize_detail_columns(df_table):
    rename_map = {}
    for col in df_table.columns:
        col_text = normalize_text(col)

        if col_text in DETAIL_COLUMNS:
            rename_map[col] = col_text
            continue

        for expected in DETAIL_COLUMNS:
            if col_text.startswith(expected) and col_text != "檢查項目備註":
                rename_map[col] = expected
                break
    df_table = df_table.rename(columns=rename_map)

    df_table = df_table.loc[:, ~df_table.columns.duplicated()]

    for col in DETAIL_COLUMNS:
        if col not in df_table.columns:
            df_table[col] = ""

    return df_table[DETAIL_COLUMNS]


def remove_footer_rows(df_table):
    stop_keywords = ["工單結案人員", "檢查人員", "領班", "課長", "備註說明"]
    keep_rows = []

    for _, row in df_table.iterrows():
        joined = " ".join(normalize_text(v) for v in row.tolist())
        has_detail = normalize_text(row.get("進階分類")) or normalize_text(row.get("檢查項目")) or normalize_text(row.get("檢查結果"))

        if any(keyword in joined for keyword in stop_keywords) and not has_detail:
            break
        if has_detail:
            keep_rows.append(row)

    return pd.DataFrame(keep_rows, columns=df_table.columns)


def parse_workbook(uploaded_file):
    df_raw = read_excel(uploaded_file, header=None)
    metadata = extract_metadata(df_raw)
    header_row = locate_header_row(df_raw)
    df_table = read_excel(uploaded_file, header=header_row)
    df_table = standardize_detail_columns(df_table)
    df_table = remove_footer_rows(df_table)

    for key, value in metadata.items():
        df_table.insert(0, key, value)

    df_table.insert(0, "來源檔案", uploaded_file.name)
    df_table["檢查結果數值"] = df_table["檢查結果"].apply(extract_numeric)
    return df_table


def apply_keyword_filter(df, column, keyword):
    keyword = normalize_text(keyword)
    if not keyword:
        return df
    return df[df[column].astype(str).str.contains(keyword, case=False, na=False, regex=False)]


def to_excel_bytes(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    return output.getvalue()


def build_stats(df):
    numeric = pd.to_numeric(df["檢查結果數值"], errors="coerce").dropna()
    if numeric.empty:
        return pd.DataFrame(
            [{"項目": "檢查結果數值", "筆數": 0, "最小值": None, "最大值": None, "平均值": None}]
        )
    return pd.DataFrame(
        [
            {
                "項目": "檢查結果數值",
                "筆數": int(numeric.count()),
                "最小值": numeric.min(),
                "最大值": numeric.max(),
                "平均值": numeric.mean(),
            }
        ]
    )


def make_chart(df, chart_type, x_col, y_mode):
    if df.empty:
        return None

    chart_df = df.copy()
    if y_mode == "筆數":
        chart_df = chart_df.groupby(x_col, dropna=False).size().reset_index(name="筆數")
        y_col = "筆數"
    else:
        y_col = "檢查結果數值"
        chart_df[y_col] = pd.to_numeric(chart_df[y_col], errors="coerce")
        chart_df = chart_df.dropna(subset=[y_col])

    if chart_df.empty:
        return None

    title = f"{x_col} / {y_col}"
    if chart_type == "折線圖":
        return px.line(chart_df, x=x_col, y=y_col, markers=True, title=title)
    if chart_type == "圓餅圖":
        return px.pie(chart_df, names=x_col, values=y_col, title=title)
    return px.bar(chart_df, x=x_col, y=y_col, title=title)


st.title("行動檢修平台通用資料整理")
st.caption("多檔 Excel 匯入、進階分類/檢查項目篩選、結果清單、統計與圖表匯出")

uploaded_files = st.file_uploader(
    "選擇一份或多份 ISO 報表 Excel",
    type=["xlsx"],
    accept_multiple_files=True,
)

if "all_data" not in st.session_state:
    st.session_state.all_data = pd.DataFrame()

if uploaded_files and st.button("開始整理", use_container_width=True):
    start_time = time.time()
    frames = []
    progress = st.progress(0)

    with st.spinner("正在讀取與標準化資料..."):
        for idx, uploaded_file in enumerate(uploaded_files):
            if uploaded_file.name.startswith("~$"):
                continue
            try:
                frames.append(parse_workbook(uploaded_file))
            except Exception as exc:
                st.warning(f"{uploaded_file.name} 讀取失敗：{exc}")
            progress.progress((idx + 1) / len(uploaded_files))

    st.session_state.all_data = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    st.session_state.process_msg = f"已整理 {len(frames)} 份檔案，共 {len(st.session_state.all_data)} 筆明細，耗時 {time.time() - start_time:.1f} 秒"
    st.rerun()


all_data = st.session_state.all_data
if all_data.empty:
    st.info("請先上傳 Excel 並按下開始整理。")
    st.stop()

st.success(st.session_state.get("process_msg", "資料已整理完成。"))

with st.sidebar:
    st.header("篩選")

    category_options = sorted(v for v in all_data["進階分類"].dropna().astype(str).unique() if v.strip())
    selected_categories = st.multiselect("進階分類", category_options)
    category_keyword = st.text_input("進階分類關鍵字")

    item_options_source = all_data
    if selected_categories:
        item_options_source = item_options_source[item_options_source["進階分類"].isin(selected_categories)]
    item_options = sorted(v for v in item_options_source["檢查項目"].dropna().astype(str).unique() if v.strip())
    selected_items = st.multiselect("檢查項目", item_options)
    item_keyword = st.text_input("檢查項目關鍵字")

    result_keyword = st.text_input("檢查結果關鍵字")

filtered = all_data.copy()
if selected_categories:
    filtered = filtered[filtered["進階分類"].isin(selected_categories)]
if selected_items:
    filtered = filtered[filtered["檢查項目"].isin(selected_items)]

filtered = apply_keyword_filter(filtered, "進階分類", category_keyword)
filtered = apply_keyword_filter(filtered, "檢查項目", item_keyword)
filtered = apply_keyword_filter(filtered, "檢查結果", result_keyword)

summary_cols = st.columns(4)
summary_cols[0].metric("明細筆數", f"{len(filtered):,}")
summary_cols[1].metric("工單數", f"{filtered['工單編號'].nunique():,}")
summary_cols[2].metric("車號/最小成本", f"{filtered['車號/最小成本'].nunique():,}")
summary_cols[3].metric("可計算數值", f"{filtered['檢查結果數值'].notna().sum():,}")

tab_list, tab_stats, tab_chart = st.tabs(["結果清單", "統計", "圖表"])

with tab_list:
    display_cols = [
        "來源檔案",
        "工單編號",
        "車號/最小成本",
        "檢查結束日期",
        "進階分類",
        "檢查項目",
        "檢查結果",
        "單位",
        "異常",
        "異常原因",
        "處理對策",
        "執行者",
    ]
    st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)

with tab_stats:
    stats_df = build_stats(filtered)
    st.dataframe(stats_df, use_container_width=True, hide_index=True)

with tab_chart:
    chart_cols = [
        "工單編號",
        "車號/最小成本",
        "檢查結束日期",
        "進階分類",
        "檢查項目",
        "檢查結果",
        "單位",
    ]
    left, middle, right = st.columns(3)
    chart_type = left.selectbox("圖表類型", ["長條圖", "折線圖", "圓餅圖"])
    x_col = middle.selectbox("X 軸 / 分類", chart_cols, index=3)
    y_mode = right.selectbox("Y 軸 / 數值", ["筆數", "檢查結果數值"])

    fig = make_chart(filtered, chart_type, x_col, y_mode)
    if fig is None:
        st.info("目前篩選結果沒有可繪圖的資料。")
    else:
        st.plotly_chart(fig, use_container_width=True)

export_name = f"行動檢修平台整理_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
export_bytes = to_excel_bytes({"篩選結果": filtered, "統計": build_stats(filtered)})
st.download_button(
    "匯出目前篩選結果",
    data=export_bytes,
    file_name=export_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
