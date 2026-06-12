import io
import json
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


def split_category(value):
    text = normalize_text(value)
    if not text:
        return "", ""

    parts = [part.strip() for part in re.split(r"\s*/\s*|\n+", text) if part.strip()]
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " / ".join(parts[1:])


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
    df_table[["進階分類1", "進階分類2"]] = df_table["進階分類"].apply(lambda value: pd.Series(split_category(value)))
    df_table = df_table[df_table["檢查結果數值"].notna()].reset_index(drop=True)
    return df_table


def apply_keyword_filter(df, column, keyword):
    keyword = normalize_text(keyword)
    if not keyword:
        return df
    return df[df[column].astype(str).str.contains(keyword, case=False, na=False, regex=False)]


def compact_repeated_values(df, columns):
    compact_df = df.copy()
    for col in columns:
        if col in compact_df.columns:
            compact_df[col] = compact_df[col].mask(compact_df[col].eq(compact_df[col].shift()), "")
    return compact_df


def to_excel_bytes(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    return output.getvalue()


def build_stats(df, group_cols=None, lower_limit=None, upper_limit=None):
    numeric = pd.to_numeric(df["檢查結果數值"], errors="coerce").dropna()
    if numeric.empty:
        return pd.DataFrame(
            [{"項目": "檢查結果數值", "筆數": 0, "最小值": None, "最大值": None, "平均值": None, "低於下限": 0, "高於上限": 0}]
        )

    stats_source = df.copy()
    stats_source["檢查結果數值"] = pd.to_numeric(stats_source["檢查結果數值"], errors="coerce")
    stats_source = stats_source.dropna(subset=["檢查結果數值"])
    group_cols = group_cols or []

    if group_cols:
        stats_df = (
            stats_source.groupby(group_cols, dropna=False)["檢查結果數值"]
            .agg(筆數="count", 最小值="min", 最大值="max", 平均值="mean")
            .reset_index()
        )
    else:
        stats_df = pd.DataFrame(
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

    if lower_limit is not None or upper_limit is not None:
        def below_count(group):
            if lower_limit is None:
                return 0
            return int((group["檢查結果數值"] < lower_limit).sum())

        def above_count(group):
            if upper_limit is None:
                return 0
            return int((group["檢查結果數值"] > upper_limit).sum())

        if group_cols:
            limit_df = (
                stats_source.groupby(group_cols, dropna=False)
                .apply(lambda group: pd.Series({"低於下限": below_count(group), "高於上限": above_count(group)}), include_groups=False)
                .reset_index()
            )
            stats_df = stats_df.merge(limit_df, on=group_cols, how="left")
        else:
            stats_df["低於下限"] = below_count(stats_source)
            stats_df["高於上限"] = above_count(stats_source)
    else:
        stats_df["低於下限"] = 0
        stats_df["高於上限"] = 0

    return stats_df


def make_chart(df, chart_type, x_col, y_mode, bar_color=None, lower_limit=None, upper_limit=None):
    if df.empty:
        return None

    chart_df = df.copy()
    if y_mode == "統計結果數值":
        chart_df["檢查結果數值"] = pd.to_numeric(chart_df["檢查結果數值"], errors="coerce")
        chart_df = chart_df.dropna(subset=["檢查結果數值"])
        chart_df = chart_df.groupby(x_col, dropna=False)["檢查結果數值"].mean().reset_index(name="統計結果數值")
        y_col = "統計結果數值"
    else:
        y_col = "檢查結果數值"
        chart_df[y_col] = pd.to_numeric(chart_df[y_col], errors="coerce")
        chart_df = chart_df.dropna(subset=[y_col])

    if chart_df.empty:
        return None

    title = f"{x_col} / {y_col}"
    if chart_type == "折線圖":
        fig = px.line(chart_df, x=x_col, y=y_col, markers=True, title=title)
    elif chart_type == "圓餅圖":
        fig = px.pie(chart_df, names=x_col, values=y_col, title=title)
    else:
        fig = px.bar(chart_df, x=x_col, y=y_col, title=title)
        if bar_color:
            fig.update_traces(marker_color=bar_color)

    if chart_type != "圓餅圖":
        if lower_limit is not None:
            fig.add_hline(y=lower_limit, line_dash="dash", line_color="red", annotation_text="下限")
        if upper_limit is not None:
            fig.add_hline(y=upper_limit, line_dash="dash", line_color="green", annotation_text="上限")
    return fig


def setting_value(key, default):
    return st.session_state.get("saved_settings", {}).get(key, default)


def valid_default_list(key, options):
    saved = setting_value(key, [])
    if not isinstance(saved, list):
        return []
    option_set = set(options)
    return [value for value in saved if value in option_set]


def valid_index(key, options, default=0):
    saved = setting_value(key, default)
    if not isinstance(saved, int) or saved < 0 or saved >= len(options):
        return default
    return saved


def update_saved_settings(**kwargs):
    settings = st.session_state.get("saved_settings", {}).copy()
    settings.update(kwargs)
    st.session_state.saved_settings = settings


st.title("行動檢修平台通用資料整理")
st.caption("多檔 Excel 匯入、進階分類/檢查項目篩選、結果清單、統計與圖表匯出")

uploaded_files = st.file_uploader(
    "選擇一份或多份 ISO 報表 Excel",
    type=["xlsx"],
    accept_multiple_files=True,
)

if "all_data" not in st.session_state:
    st.session_state.all_data = pd.DataFrame()
if "saved_settings" not in st.session_state:
    st.session_state.saved_settings = {}

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

    settings_file = st.file_uploader("匯入設定檔", type=["json"], accept_multiple_files=False)
    if settings_file is not None and st.button("套用設定檔", use_container_width=True):
        try:
            st.session_state.saved_settings = json.loads(settings_file.getvalue().decode("utf-8"))
            st.rerun()
        except Exception as exc:
            st.warning(f"設定檔讀取失敗：{exc}")

    category1_options = sorted(v for v in all_data["進階分類1"].dropna().astype(str).unique() if v.strip())
    selected_category1 = st.multiselect("進階分類1", category1_options, default=valid_default_list("selected_category1", category1_options))
    category1_keyword = st.text_input("進階分類1關鍵字", value=setting_value("category1_keyword", ""))

    category2_source = all_data
    if selected_category1:
        category2_source = category2_source[category2_source["進階分類1"].isin(selected_category1)]
    category2_options = sorted(v for v in category2_source["進階分類2"].dropna().astype(str).unique() if v.strip())
    selected_category2 = st.multiselect("進階分類2", category2_options, default=valid_default_list("selected_category2", category2_options))
    category2_keyword = st.text_input("進階分類2關鍵字", value=setting_value("category2_keyword", ""))

    item_options_source = all_data
    if selected_category1:
        item_options_source = item_options_source[item_options_source["進階分類1"].isin(selected_category1)]
    if selected_category2:
        item_options_source = item_options_source[item_options_source["進階分類2"].isin(selected_category2)]
    item_options = sorted(v for v in item_options_source["檢查項目"].dropna().astype(str).unique() if v.strip())
    selected_items = st.multiselect("檢查項目", item_options, default=valid_default_list("selected_items", item_options))
    item_keyword = st.text_input("檢查項目關鍵字", value=setting_value("item_keyword", ""))

    result_keyword = st.text_input("檢查結果關鍵字", value=setting_value("result_keyword", ""))

    if st.button("保留目前設定", use_container_width=True):
        update_saved_settings(
            selected_category1=selected_category1,
            category1_keyword=category1_keyword,
            selected_category2=selected_category2,
            category2_keyword=category2_keyword,
            selected_items=selected_items,
            item_keyword=item_keyword,
            result_keyword=result_keyword,
        )
        st.success("已保留目前篩選設定。")

    st.download_button(
        "下載設定檔",
        data=json.dumps(st.session_state.saved_settings, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="行動檢修平台設定.json",
        mime="application/json",
        use_container_width=True,
    )

filtered = all_data.copy()
if selected_category1:
    filtered = filtered[filtered["進階分類1"].isin(selected_category1)]
if selected_category2:
    filtered = filtered[filtered["進階分類2"].isin(selected_category2)]
if selected_items:
    filtered = filtered[filtered["檢查項目"].isin(selected_items)]

filtered = apply_keyword_filter(filtered, "進階分類1", category1_keyword)
filtered = apply_keyword_filter(filtered, "進階分類2", category2_keyword)
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
        "進階分類1",
        "進階分類2",
        "檢查項目",
        "檢查結果數值",
        "單位",
        "執行者",
    ]
    display_df = compact_repeated_values(
        filtered[display_cols],
        ["來源檔案", "工單編號", "車號/最小成本", "檢查結束日期"],
    ).rename(columns={"檢查結果數值": "檢查結果"})
    st.dataframe(display_df, use_container_width=True, hide_index=True)

with tab_stats:
    vehicle_options = sorted(v for v in filtered["車號/最小成本"].dropna().astype(str).unique() if v.strip())
    stat_vehicle = st.multiselect("選擇要計算的車號/最小成本", vehicle_options, default=valid_default_list("stat_vehicle", vehicle_options))

    stat_item_options = sorted(v for v in filtered["檢查項目"].dropna().astype(str).unique() if v.strip())
    stat_items = st.multiselect("選擇要計算的檢查項目", stat_item_options, default=valid_default_list("stat_items", stat_item_options))

    limit_cols = st.columns(2)
    lower_limit_enabled = limit_cols[0].checkbox("啟用統計下限", value=setting_value("lower_limit_enabled", False))
    lower_limit = limit_cols[0].number_input("統計下限", value=float(setting_value("lower_limit", 0.0)), disabled=not lower_limit_enabled)
    upper_limit_enabled = limit_cols[1].checkbox("啟用統計上限", value=setting_value("upper_limit_enabled", False))
    upper_limit = limit_cols[1].number_input("統計上限", value=float(setting_value("upper_limit", 0.0)), disabled=not upper_limit_enabled)

    stats_source = filtered.copy()
    if stat_vehicle:
        stats_source = stats_source[stats_source["車號/最小成本"].isin(stat_vehicle)]
    if stat_items:
        stats_source = stats_source[stats_source["檢查項目"].isin(stat_items)]

    lower_value = lower_limit if lower_limit_enabled else None
    upper_value = upper_limit if upper_limit_enabled else None
    stats_df = build_stats(stats_source, ["車號/最小成本", "檢查項目"], lower_value, upper_value)
    st.dataframe(stats_df, use_container_width=True, hide_index=True)

    if st.button("保留統計設定", use_container_width=True):
        update_saved_settings(
            stat_vehicle=stat_vehicle,
            stat_items=stat_items,
            lower_limit_enabled=lower_limit_enabled,
            lower_limit=lower_limit,
            upper_limit_enabled=upper_limit_enabled,
            upper_limit=upper_limit,
        )
        st.success("已保留統計設定。")

with tab_chart:
    chart_cols = ["車號/最小成本", "檢查結束日期"]
    left, middle, right = st.columns(3)
    chart_types = ["長條圖", "折線圖", "圓餅圖"]
    y_modes = ["檢查結果數值", "統計結果數值"]
    chart_type = left.selectbox("圖表類型", chart_types, index=valid_index("chart_type_index", chart_types))
    x_col = middle.selectbox("X 軸 / 分類", chart_cols, index=valid_index("x_col_index", chart_cols))
    y_mode = right.selectbox("Y 軸 / 數值", y_modes, index=valid_index("y_mode_index", y_modes))
    bar_color = st.color_picker("長條圖柱體顏色", value=setting_value("bar_color", "#2563EB"))

    fig = make_chart(filtered, chart_type, x_col, y_mode, bar_color, lower_value, upper_value)
    if fig is None:
        st.info("目前篩選結果沒有可繪圖的資料。")
    else:
        st.plotly_chart(fig, use_container_width=True)
        export_cols = st.columns(2)
        chart_html = fig.to_html(include_plotlyjs="cdn").encode("utf-8")
        export_cols[0].download_button(
            "匯出圖表 HTML",
            data=chart_html,
            file_name=f"圖表_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
            mime="text/html",
            use_container_width=True,
        )
        try:
            chart_png = fig.to_image(format="png", scale=2)
            export_cols[1].download_button(
                "匯出圖表圖片",
                data=chart_png,
                file_name=f"圖表_{datetime.now().strftime('%Y%m%d_%H%M')}.png",
                mime="image/png",
                use_container_width=True,
            )
        except Exception:
            export_cols[1].info("若要匯出 PNG，請確認已安裝 kaleido。")

    if st.button("保留圖表設定", use_container_width=True):
        update_saved_settings(
            chart_type_index=chart_types.index(chart_type),
            x_col_index=chart_cols.index(x_col),
            y_mode_index=y_modes.index(y_mode),
            bar_color=bar_color,
        )
        st.success("已保留圖表設定。")

export_name = f"行動檢修平台整理_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
export_display = compact_repeated_values(
    filtered[display_cols],
    ["來源檔案", "工單編號", "車號/最小成本", "檢查結束日期"],
).rename(columns={"檢查結果數值": "檢查結果"})
export_bytes = to_excel_bytes({"篩選結果": export_display, "統計": stats_df})
st.download_button(
    "匯出目前篩選結果",
    data=export_bytes,
    file_name=export_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
