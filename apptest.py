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


st.set_page_config(page_title="自訂統計資料整理", layout="wide")

st.markdown(
    """
    <style>
    :root {
        --app-bg: #080D17;
        --panel-bg: #111827;
        --panel-soft: #172033;
        --text-main: #F8FAFC;
        --text-muted: #AAB6C7;
        --line: #263245;
        --primary: #38BDF8;
        --primary-soft: #0B2A3A;
        --accent: #34D399;
        --accent-soft: #0B2F28;
        --warning-soft: #34280A;
    }
    .main .block-container {
        max-width: 1480px;
        padding-top: 1.6rem;
        padding-bottom: 2rem;
    }
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(56, 189, 248, 0.16), transparent 34rem),
            radial-gradient(circle at top right, rgba(52, 211, 153, 0.12), transparent 30rem),
            linear-gradient(180deg, #0B1120 0%, var(--app-bg) 38rem, #070B13 100%);
        color: var(--text-main);
    }
    html, body, [class*="css"] {
        color: var(--text-main);
    }
    h1, h2, h3 {
        letter-spacing: 0;
        color: var(--text-main);
    }
    h1 {
        font-weight: 850;
        margin-bottom: 0.2rem;
    }
    p, label, span, div {
        color: inherit;
    }
    [data-testid="stMarkdownContainer"] p {
        color: var(--text-muted);
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: var(--panel-bg);
        border: 1px solid var(--line);
        border-radius: 12px;
        box-shadow: 0 14px 38px rgba(0, 0, 0, 0.26);
    }
    div[data-testid="stMetric"] {
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 14px 16px;
        background: var(--panel-bg);
        box-shadow: 0 10px 28px rgba(0, 0, 0, 0.24);
    }
    div[data-testid="stMetric"] label,
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: var(--text-main);
    }
    div[data-testid="stTabs"] button {
        font-weight: 700;
        color: var(--text-main);
    }
    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: var(--primary);
        border-bottom-color: var(--primary);
    }
    section[data-testid="stSidebar"] {
        border-right: 1px solid var(--line);
        background: #0B1120;
    }
    section[data-testid="stSidebar"] * {
        color: var(--text-main);
    }
    .stTextInput input,
    .stNumberInput input,
    .stSelectbox div[data-baseweb="select"],
    .stMultiSelect div[data-baseweb="select"] {
        background: #0F172A;
        color: var(--text-main);
        border-color: var(--line);
    }
    .stTextInput input::placeholder {
        color: #7B8798;
    }
    .stButton > button,
    .stDownloadButton > button {
        border-radius: 9px;
        font-weight: 750;
        border: 1px solid var(--primary);
        background: var(--primary);
        color: #FFFFFF;
    }
    .stButton > button:hover,
    .stDownloadButton > button:hover {
        border-color: #1D4ED8;
        background: #1D4ED8;
        color: #FFFFFF;
    }
    div[role="radiogroup"] label {
        background: #0F172A;
        border: 1px solid var(--line);
        border-radius: 999px;
        padding: 4px 10px;
        margin-right: 6px;
    }
    div[role="radiogroup"] label:has(input:checked) {
        background: var(--primary-soft);
        border-color: var(--primary);
        color: var(--primary);
    }
    .workflow {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 14px;
        margin: 16px 0 20px 0;
    }
    .workflow-step {
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 14px 16px;
        background: linear-gradient(180deg, #121C2E 0%, #0F172A 100%);
        min-height: 94px;
        box-shadow: 0 14px 34px rgba(0, 0, 0, 0.25);
    }
    .workflow-step strong {
        display: block;
        color: var(--text-main);
        margin-bottom: 4px;
    }
    .workflow-step span {
        color: var(--text-muted);
        font-size: 0.92rem;
    }
    .step-title {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 6px;
        font-size: 1.05rem;
        font-weight: 800;
        color: var(--text-main);
    }
    .step-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 26px;
        height: 26px;
        border-radius: 999px;
        background: linear-gradient(135deg, var(--primary), var(--accent));
        color: #FFFFFF;
        font-size: 0.9rem;
        font-weight: 800;
    }
    .step-help {
        color: var(--text-muted);
        font-size: 0.92rem;
        margin: 0 0 12px 0;
    }
    .soft-divider {
        height: 1px;
        background: var(--line);
        margin: 12px 0;
    }
    .stAlert {
        border-radius: 10px;
    }
    div[data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
        border: 1px solid var(--line);
        background: var(--panel-bg);
    }
    @media (max-width: 900px) {
        .workflow {
            grid-template-columns: 1fr;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


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


def format_number(value):
    if pd.isna(value):
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return value
    if number.is_integer():
        return str(int(number))
    return f"{number:.10f}".rstrip("0").rstrip(".")


def format_numeric_columns(df):
    display_df = df.copy()
    numeric_cols = display_df.select_dtypes(include="number").columns
    for col in numeric_cols:
        display_df[col] = display_df[col].apply(format_number)
    return display_df


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


def classify_traffic_light(value, traffic_ranges):
    if value is None or pd.isna(value) or not traffic_ranges:
        return ""

    for name in ["綠燈", "黃燈", "紅燈"]:
        lower, upper = traffic_ranges.get(name, (None, None))
        lower_ok = lower is None or value >= lower
        upper_ok = upper is None or value <= upper
        if lower_ok and upper_ok:
            return name
    return "未分類"


def calculate_compare_value(stats_df, compare_metric, operation_left=None, operation_right=None, operation_value=None):
    if compare_metric in ["原始值", "最小值", "最大值", "平均值"]:
        return stats_df[compare_metric]

    if compare_metric not in ["相加", "相減", "相乘", "相除"]:
        return None

    operation_left = operation_left if operation_left in stats_df.columns else "平均值"
    left_value = pd.to_numeric(stats_df[operation_left], errors="coerce")

    if operation_right == "自訂數值":
        right_value = operation_value
    elif operation_right in stats_df.columns:
        right_value = pd.to_numeric(stats_df[operation_right], errors="coerce")
    else:
        right_value = 0 if compare_metric in ["相加", "相減"] else 1

    if compare_metric == "相加":
        return left_value + right_value
    if compare_metric == "相減":
        return left_value - right_value
    if compare_metric == "相乘":
        return left_value * right_value
    if compare_metric == "相除":
        if isinstance(right_value, pd.Series):
            return left_value / right_value.replace(0, pd.NA)
        if right_value == 0:
            return pd.Series([pd.NA] * len(stats_df), index=stats_df.index)
        return left_value / right_value
    return None


def build_stats(
    df,
    group_cols=None,
    compare_metric="",
    lower_limit=None,
    upper_limit=None,
    traffic_ranges=None,
    operation_left=None,
    operation_right=None,
    operation_value=None,
):
    numeric = pd.to_numeric(df["檢查結果數值"], errors="coerce").dropna()
    if numeric.empty:
        return pd.DataFrame(
            [
                {
                    "項目": "檢查結果數值",
                    "筆數": 0,
                    "最小值": None,
                    "最大值": None,
                    "平均值": None,
                    "統計方式": compare_metric or "",
                    "統計結果數值": None,
                    "判定": "",
                }
            ]
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
    stats_df["原始值"] = stats_df["最小值"].where(stats_df["筆數"] == 1)

    if compare_metric not in ["原始值", "最小值", "最大值", "平均值", "相加", "相減", "相乘", "相除"]:
        compare_metric = ""

    stats_df["統計方式"] = compare_metric
    compare_value = calculate_compare_value(stats_df, compare_metric, operation_left, operation_right, operation_value)
    stats_df["統計結果數值"] = compare_value if compare_value is not None else None
    stats_df["判定"] = ""
    if compare_metric and traffic_ranges:
        stats_df["判定"] = stats_df["統計結果數值"].apply(lambda value: classify_traffic_light(value, traffic_ranges))
    elif compare_metric and lower_limit is not None:
        stats_df.loc[stats_df["統計結果數值"] < lower_limit, "判定"] = "低於下限"
    if compare_metric and upper_limit is not None:
        stats_df.loc[stats_df["統計結果數值"] > upper_limit, "判定"] = "高於上限"
    return stats_df


def make_chart(
    df,
    chart_type,
    x_col,
    y_col,
    bar_color=None,
    lower_limit=None,
    upper_limit=None,
    y_min=None,
    y_max=None,
    title=None,
    x_label=None,
    y_label=None,
    traffic_color_enabled=False,
    traffic_colors=None,
    pie_label_mode="標籤+百分比",
    series_col=None,
    series_colors=None,
    show_data_labels=False,
    y_tick_interval=None,
    limit_lines=None,
):
    if df.empty:
        return None

    chart_df = df.copy()
    if x_col not in chart_df.columns or y_col not in chart_df.columns:
        return None

    chart_df[y_col] = pd.to_numeric(chart_df[y_col], errors="coerce")
    chart_df = chart_df.dropna(subset=[y_col])

    if chart_df.empty:
        return None

    chart_df = chart_df.sort_values(by=[series_col, x_col] if series_col and series_col in chart_df.columns else x_col)
    title = title or f"{x_col} / {y_col}"
    if chart_type == "折線圖":
        fig = px.line(
            chart_df,
            x=x_col,
            y=y_col,
            color=series_col if series_col in chart_df.columns else None,
            color_discrete_map=series_colors or None,
            markers=True,
            title=title,
        )
    elif chart_type == "圓餅圖":
        fig = px.pie(chart_df, names=x_col, values=y_col, title=title)
        textinfo_map = {
            "標籤+百分比": "label+percent",
            "標籤+數值": "label+value",
            "百分比": "percent",
            "數值": "value",
            "全部": "label+percent+value",
        }
        fig.update_traces(textinfo=textinfo_map.get(pie_label_mode, "label+percent"))
    else:
        color_col = None
        color_map = None
        if traffic_color_enabled and "判定" in chart_df.columns:
            color_col = "判定"
            color_map = traffic_colors or None
        elif series_col and series_col in chart_df.columns:
            color_col = series_col
            color_map = series_colors or None
        fig = px.bar(
            chart_df,
            x=x_col,
            y=y_col,
            color=color_col,
            color_discrete_map=color_map,
            title=title,
            barmode="group",
        )
        if not color_col and bar_color:
            fig.update_traces(marker_color=bar_color)

    if chart_type != "圓餅圖":
        lines = limit_lines or []
        if not lines:
            if lower_limit is not None:
                lines.append({"value": lower_limit, "label": "下限", "color": "#EF4444"})
            if upper_limit is not None:
                lines.append({"value": upper_limit, "label": "上限", "color": "#10B981"})
        for line in lines:
            if line.get("value") is not None:
                fig.add_hline(
                    y=line["value"],
                    line_dash="dash",
                    line_color=line.get("color", "#64748B"),
                    annotation_text=line.get("label", ""),
                    annotation_position="top left",
                )
        if y_min is not None and y_max is not None and y_min >= y_max:
            y_min = None
            y_max = None
        if y_min is not None or y_max is not None:
            fig.update_yaxes(range=[y_min, y_max])
        if y_tick_interval:
            fig.update_yaxes(dtick=y_tick_interval)
        fig.update_layout(xaxis_type="category", xaxis_tickangle=-45)
        if x_label:
            fig.update_xaxes(title_text=x_label)
        if y_label:
            fig.update_yaxes(title_text=y_label)
        if show_data_labels:
            if chart_type == "折線圖":
                fig.update_traces(mode="lines+markers+text", texttemplate="%{y:g}", textposition="top center")
            else:
                fig.update_traces(texttemplate="%{y:g}", textposition="outside")
        fig.update_layout(showlegend=True, legend_title_text=series_col or ("判定" if traffic_color_enabled else "圖例"))
    elif show_data_labels and chart_type == "圓餅圖":
        fig.update_traces(textposition="inside")
    fig.update_layout(
        showlegend=True,
        paper_bgcolor="#111827",
        plot_bgcolor="#0F172A",
        font={"color": "#F8FAFC"},
        legend={"bgcolor": "rgba(17,24,39,0.85)", "bordercolor": "#263245", "borderwidth": 1},
        title_font={"color": "#F8FAFC"},
    )
    fig.update_xaxes(gridcolor="#263245", zerolinecolor="#334155", color="#DDE7F3")
    fig.update_yaxes(gridcolor="#263245", zerolinecolor="#334155", color="#DDE7F3")
    return fig


def build_limit_lines(range_mode, lower_value=None, upper_value=None, traffic_ranges=None):
    if range_mode == "上下限":
        return [
            {"value": lower_value, "label": "下限", "color": "#EF4444"},
            {"value": upper_value, "label": "上限", "color": "#10B981"},
        ]

    if range_mode == "紅黃綠燈" and traffic_ranges:
        lines = []
        color_map = {"紅燈": "#EF4444", "黃燈": "#F59E0B", "綠燈": "#10B981"}
        seen = set()
        for name in ["紅燈", "黃燈", "綠燈"]:
            lower, upper = traffic_ranges.get(name, (None, None))
            for label, value in [(f"{name}下限", lower), (f"{name}上限", upper)]:
                if value is None or value in seen:
                    continue
                seen.add(value)
                lines.append({"value": value, "label": label, "color": color_map[name]})
        return sorted(lines, key=lambda item: item["value"])

    return []


def setting_value(key, default):
    return st.session_state.get("saved_settings", {}).get(key, default)


def current_filter_settings():
    return {
        "selected_category1": selected_category1,
        "category1_keyword": category1_keyword,
        "selected_category2": selected_category2,
        "category2_keyword": category2_keyword,
        "selected_items": selected_items,
        "item_keyword": item_keyword,
    }


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


def step_header(number, title, help_text):
    st.markdown(
        f"""
        <div class="step-title"><span class="step-badge">{number}</span><span>{title}</span></div>
        <p class="step-help">{help_text}</p>
        """,
        unsafe_allow_html=True,
    )


def go_to_step(step):
    st.session_state.current_step = step
    st.rerun()


def step_nav(back_step=None, next_step=None, next_label="下一步"):
    cols = st.columns([1, 1, 4])
    if back_step is not None and cols[0].button("上一步", use_container_width=True):
        go_to_step(back_step)
    if next_step is not None and cols[1].button(next_label, use_container_width=True):
        go_to_step(next_step)


st.title("自訂統計資料整理")
st.caption("多檔 Excel 匯入、進階分類/檢查項目篩選、結果清單、統計與圖表匯出")
st.markdown(
    """
    <div class="workflow">
      <div class="workflow-step"><strong>1. 上傳資料</strong><span>一次選多份 ISO Excel，系統自動整理成數值明細。</span></div>
      <div class="workflow-step"><strong>2. 篩選項目</strong><span>用分類、檢查項目或關鍵字縮小資料範圍。</span></div>
      <div class="workflow-step"><strong>3. 設定統計</strong><span>選最小值、最大值、平均值、原始值或四則運算。</span></div>
      <div class="workflow-step"><strong>4. 產生圖表</strong><span>做長條圖、折線圖或圓餅圖，並匯出結果。</span></div>
    </div>
    """,
    unsafe_allow_html=True,
)

if "current_step" not in st.session_state:
    st.session_state.current_step = 1

uploaded_files = None
if st.session_state.current_step == 1:
    with st.container(border=True):
        step_header(1, "上傳資料", "先選擇一份或多份 ISO Excel，按下開始整理後，系統會只保留可計算的數值檢查結果。")
        uploaded_files = st.file_uploader(
            "選擇一份或多份 ISO 報表 Excel",
            type=["xlsx"],
            accept_multiple_files=True,
        )

if "all_data" not in st.session_state:
    st.session_state.all_data = pd.DataFrame()
if "saved_settings" not in st.session_state:
    st.session_state.saved_settings = {}
if "filter_profiles" not in st.session_state:
    st.session_state.filter_profiles = {}
if "current_step" not in st.session_state:
    st.session_state.current_step = 1

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
    st.session_state.current_step = 2
    st.rerun()


all_data = st.session_state.all_data
if all_data.empty:
    st.info("請先上傳 Excel 並按下開始整理。")
    st.stop()

st.success(st.session_state.get("process_msg", "資料已整理完成。"))

if st.session_state.current_step == 1:
    step_nav(next_step=2, next_label="下一步：篩選資料")
    st.stop()

selected_category1 = st.session_state.get("selected_category1", [])
category1_keyword = st.session_state.get("category1_keyword", "")
selected_category2 = st.session_state.get("selected_category2", [])
category2_keyword = st.session_state.get("category2_keyword", "")
selected_items = st.session_state.get("selected_items", [])
item_keyword = st.session_state.get("item_keyword", "")

if st.session_state.current_step == 2:
    with st.container(border=True):
        step_header(2, "篩選資料", "從分類、檢查項目或關鍵字縮小範圍；不輸入就代表全部。設定好後按下一步查看清單。")

        category1_options = sorted(v for v in all_data["進階分類1"].dropna().astype(str).unique() if v.strip())
        selected_category1 = st.multiselect("進階分類1", category1_options, default=[v for v in selected_category1 if v in category1_options])
        category1_keyword = st.text_input("進階分類1關鍵字", value=category1_keyword)

        category2_source = all_data
        if selected_category1:
            category2_source = category2_source[category2_source["進階分類1"].isin(selected_category1)]
        category2_options = sorted(v for v in category2_source["進階分類2"].dropna().astype(str).unique() if v.strip())
        selected_category2 = st.multiselect("進階分類2", category2_options, default=[v for v in selected_category2 if v in category2_options])
        category2_keyword = st.text_input("進階分類2關鍵字", value=category2_keyword)

        item_options_source = all_data
        if selected_category1:
            item_options_source = item_options_source[item_options_source["進階分類1"].isin(selected_category1)]
        if selected_category2:
            item_options_source = item_options_source[item_options_source["進階分類2"].isin(selected_category2)]
        item_options = sorted(v for v in item_options_source["檢查項目"].dropna().astype(str).unique() if v.strip())
        selected_items = st.multiselect("檢查項目", item_options, default=[v for v in selected_items if v in item_options])
        item_keyword = st.text_input("檢查項目關鍵字", value=item_keyword)

        st.session_state.selected_category1 = selected_category1
        st.session_state.category1_keyword = category1_keyword
        st.session_state.selected_category2 = selected_category2
        st.session_state.category2_keyword = category2_keyword
        st.session_state.selected_items = selected_items
        st.session_state.item_keyword = item_keyword

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
]

st.metric("搜尋結果工單數", f"{filtered['工單編號'].nunique():,}")

if st.session_state.current_step == 2:
    step_nav(back_step=1, next_step=3, next_label="下一步：確認清單")
    st.stop()

if st.session_state.current_step == 3:
    step_header(3, "確認結果清單", "這裡顯示目前篩選後的明細資料，可先確認資料是否符合預期。")
    display_df = compact_repeated_values(
        filtered[display_cols],
        ["來源檔案", "工單編號", "車號/最小成本", "檢查結束日期"],
    ).rename(columns={"檢查結果數值": "檢查結果"})
    display_df = format_numeric_columns(display_df)
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    step_nav(back_step=2, next_step=4, next_label="下一步：設定統計")
    st.stop()

if st.session_state.current_step == 4:
    step_header(4, "設定統計範圍", "先決定要統計哪些車號、哪些檢查項目，以及資料要依什麼欄位分組。未選車號或檢查項目時，系統會自動視為全選。")

    with st.container(border=True):
        st.markdown("##### 統計資料範圍")
        vehicle_options = sorted(v for v in filtered["車號/最小成本"].dropna().astype(str).unique() if v.strip())
        stat_vehicle = st.multiselect("選擇要計算的車號/最小成本", vehicle_options, default=valid_default_list("stat_vehicle", vehicle_options))

        stat_item_keyword = st.text_input("統計檢查項目關鍵字", value=setting_value("stat_item_keyword", "直徑"), help="例如輸入「直徑」會只計算檢查項目名稱包含直徑的資料。")
        stat_item_options = sorted(v for v in filtered["檢查項目"].dropna().astype(str).unique() if v.strip())
        if stat_item_keyword:
            stat_item_options = [item for item in stat_item_options if stat_item_keyword.lower() in item.lower()]
        stat_items = st.multiselect("選擇要計算的檢查項目", stat_item_options, default=valid_default_list("stat_items", stat_item_options))

        group_options = ["車號/最小成本", "檢查結束日期", "檢查項目", "進階分類1", "進階分類2"]
        stat_group_cols = st.multiselect(
            "統計分組欄位",
            group_options,
            default=valid_default_list("stat_group_cols", group_options) or ["車號/最小成本", "檢查項目"],
            help="分組欄位決定統計表每一列代表什麼。例如車號 + 檢查項目，就是每台車每個檢查項目各算一列。",
        )
        if len(stat_items) > 1 and "檢查項目" not in stat_group_cols:
            st.info("已選多個檢查項目；若要在統計或折線圖中分列顯示，建議將「檢查項目」加入統計分組欄位。")

    with st.container(border=True):
        st.markdown("##### 統計方式")
        compare_options = ["", "原始值", "最小值", "最大值", "平均值", "相加", "相減", "相乘", "相除"]
        compare_metric = st.radio("要拿來比較/畫圖的統計值", compare_options, horizontal=True, index=valid_index("compare_metric_index", compare_options))
        st.caption("常用：輪徑看最小值、溫度/電壓可看平均值；只有一筆資料時可選原始值。")
    operation_left = operation_right = None
    operation_value = None
    if compare_metric in ["相加", "相減", "相乘", "相除"]:
        operation_cols = st.columns([1, 1, 1])
        operation_base_cols = ["原始值", "最小值", "最大值", "平均值", "筆數"]
        operation_left = operation_cols[0].selectbox(
            "左側數值",
            operation_base_cols,
            index=valid_index("operation_left_index", operation_base_cols),
        )
        operation_right_options = operation_base_cols + ["自訂數值"]
        operation_right = operation_cols[1].selectbox(
            "右側數值",
            operation_right_options,
            index=valid_index("operation_right_index", operation_right_options),
        )
        operation_value = operation_cols[2].number_input(
            "自訂數值",
            value=float(setting_value("operation_value", 1.0)),
            disabled=operation_right != "自訂數值",
        )

    with st.container(border=True):
        st.markdown("##### 判定範圍")
        range_mode = st.radio(
            "統計判定方式",
            ["不判定", "上下限", "紅黃綠燈"],
            horizontal=True,
            index=valid_index("range_mode_index", ["不判定", "上下限", "紅黃綠燈"]),
        )

    lower_limit = upper_limit = None
    traffic_ranges = None
    if range_mode == "上下限":
        limit_cols = st.columns(2)
        lower_limit = limit_cols[0].number_input("統計下限", value=float(setting_value("lower_limit", 0.0)))
        upper_limit = limit_cols[1].number_input("統計上限", value=float(setting_value("upper_limit", 0.0)))
    elif range_mode == "紅黃綠燈":
        st.caption("範圍採含頭含尾判定。例：綠燈 788 到 850、黃燈 782 到 787、紅燈 775 到 781。")
        traffic_cols = st.columns(3)
        green_min = traffic_cols[0].number_input("綠燈最小值", value=float(setting_value("green_min", 788.0)))
        green_max = traffic_cols[0].number_input("綠燈最大值", value=float(setting_value("green_max", 850.0)))
        yellow_min = traffic_cols[1].number_input("黃燈最小值", value=float(setting_value("yellow_min", 782.0)))
        yellow_max = traffic_cols[1].number_input("黃燈最大值", value=float(setting_value("yellow_max", 787.0)))
        red_min = traffic_cols[2].number_input("紅燈最小值", value=float(setting_value("red_min", 775.0)))
        red_max = traffic_cols[2].number_input("紅燈最大值", value=float(setting_value("red_max", 781.0)))
        traffic_ranges = {
            "綠燈": (green_min, green_max),
            "黃燈": (yellow_min, yellow_max),
            "紅燈": (red_min, red_max),
        }

    stats_source = filtered.copy()
    if stat_vehicle:
        stats_source = stats_source[stats_source["車號/最小成本"].isin(stat_vehicle)]
    if stat_item_keyword:
        stats_source = apply_keyword_filter(stats_source, "檢查項目", stat_item_keyword)
    if stat_items:
        stats_source = stats_source[stats_source["檢查項目"].isin(stat_items)]

    lower_value = lower_limit if range_mode == "上下限" else None
    upper_value = upper_limit if range_mode == "上下限" else None
    stats_df = build_stats(
        stats_source,
        stat_group_cols,
        compare_metric,
        lower_value,
        upper_value,
        traffic_ranges,
        operation_left,
        operation_right,
        operation_value,
    )

    stat_summary_cols = st.columns(4)
    stat_summary_cols[0].metric("統計筆數", f"{len(stats_df):,}")
    if range_mode == "紅黃綠燈":
        stat_summary_cols[1].metric("綠燈", f"{(stats_df['判定'] == '綠燈').sum():,}" if "判定" in stats_df else "0")
        stat_summary_cols[2].metric("黃燈", f"{(stats_df['判定'] == '黃燈').sum():,}" if "判定" in stats_df else "0")
        stat_summary_cols[3].metric("紅燈", f"{(stats_df['判定'] == '紅燈').sum():,}" if "判定" in stats_df else "0")
    else:
        stat_summary_cols[1].metric("低於下限", f"{(stats_df['判定'] == '低於下限').sum():,}" if "判定" in stats_df else "0")
        stat_summary_cols[2].metric("高於上限", f"{(stats_df['判定'] == '高於上限').sum():,}" if "判定" in stats_df else "0")
        stat_summary_cols[3].metric("統計方式", compare_metric or "未指定")

    st.session_state.stats_df = stats_df
    st.session_state.stats_source = stats_source
    st.session_state.range_mode_value = range_mode
    st.session_state.lower_value = lower_value
    st.session_state.upper_value = upper_value
    st.session_state.traffic_ranges = traffic_ranges

    st.dataframe(format_numeric_columns(stats_df), use_container_width=True, hide_index=True)
    step_nav(back_step=3, next_step=5, next_label="下一步：產生圖表")
    st.stop()

if st.session_state.current_step == 5:
    stats_df = st.session_state.get("stats_df", pd.DataFrame())
    stats_source = st.session_state.get("stats_source", filtered.copy())
    range_mode = st.session_state.get("range_mode_value", "不判定")
    lower_value = st.session_state.get("lower_value", None)
    upper_value = st.session_state.get("upper_value", None)
    traffic_ranges = st.session_state.get("traffic_ranges", None)
    if stats_df.empty:
        st.warning("尚未建立統計結果，請先回到上一步完成統計設定。")
        step_nav(back_step=4, next_label=None)
        st.stop()
    step_header(5, "產生圖表", "選資料來源與欄位後產生圖表；長條圖和折線圖會保留上下限或紅黃綠燈門檻線，並顯示圖例。")
    chart_data_mode = st.radio("圖表資料來源", ["統計結果", "篩選明細"], horizontal=True, index=valid_index("chart_data_mode_index", ["統計結果", "篩選明細"]))
    chart_source = stats_df if chart_data_mode == "統計結果" else stats_source
    chart_cols = [col for col in ["車號/最小成本", "檢查結束日期", "檢查項目", "進階分類1", "進階分類2"] if col in chart_source.columns]
    chart_types = ["長條圖", "折線圖", "圓餅圖"]
    candidate_y_modes = ["統計結果數值", "檢查結果數值", "原始值", "最小值", "最大值", "平均值", "筆數"]
    y_modes = [
        col
        for col in candidate_y_modes
        if col in chart_source.columns and pd.to_numeric(chart_source[col], errors="coerce").notna().any()
    ]
    if not chart_cols or not y_modes or chart_source.empty:
        st.info("目前統計或篩選範圍沒有可繪圖的資料。")
    else:
        with st.container(border=True):
            st.markdown("##### 圖表資料與欄位")
            chart_type = st.radio("圖表類型", chart_types, horizontal=True, index=valid_index("chart_type_index", chart_types))
            left, middle, right = st.columns(3)
            x_col = left.selectbox("X 軸 / 分類", chart_cols, index=valid_index("x_col_index", chart_cols))
            y_mode = middle.selectbox("Y 軸 / 數值", y_modes, index=valid_index("y_mode_index", y_modes))
            series_options = [""] + [col for col in ["檢查項目", "進階分類1", "進階分類2", "車號/最小成本"] if col in chart_source.columns and col != x_col]
            series_col = right.selectbox("系列欄位（多條線/多組柱體）", series_options, index=valid_index("series_col_index", series_options))
            st.caption("折線圖若要一次顯示多條線，請在系列欄位選「檢查項目」或其他分類欄位。")

        with st.container(border=True):
            st.markdown("##### 圖表文字")
            label_cols = st.columns(3)
            chart_title = label_cols[0].text_input("圖表標題", value=setting_value("chart_title", ""))
            x_axis_label = label_cols[1].text_input("X 軸名稱", value=setting_value("x_axis_label", x_col))
            y_axis_label = label_cols[2].text_input("Y 軸名稱", value=setting_value("y_axis_label", y_mode))

            pie_label_modes = ["標籤+百分比", "標籤+數值", "百分比", "數值", "全部"]
            pie_label_mode = "標籤+百分比"
            if chart_type == "圓餅圖":
                pie_label_mode = st.radio("圓餅圖標籤呈現方式", pie_label_modes, horizontal=True, index=valid_index("pie_label_mode_index", pie_label_modes))

        y_range_enabled = False
        y_min = None
        y_max = None
        if chart_type != "圓餅圖":
            with st.container(border=True):
                st.markdown("##### Y 軸與資料標籤")
                axis_cols = st.columns([1, 1, 1, 1])
                y_range_enabled = axis_cols[0].checkbox("自訂 Y 軸範圍", value=setting_value("y_range_enabled", False))
                y_min = axis_cols[1].number_input("Y 軸最小值", value=float(setting_value("y_min", 0.0)), disabled=not y_range_enabled)
                y_max = axis_cols[2].number_input("Y 軸最大值", value=float(setting_value("y_max", 850.0)), disabled=not y_range_enabled)
                y_tick_interval = axis_cols[3].number_input(
                    "Y 軸間距",
                    min_value=0.0,
                    value=float(setting_value("y_tick_interval", 0.0)),
                    step=0.1,
                    help="例如 1~7 每格 0.5，就輸入 0.5。填 0 代表自動。",
                )
        else:
            y_tick_interval = None

        show_data_labels = st.checkbox("顯示資料標籤", value=setting_value("show_data_labels", False))

        traffic_color_enabled = False
        green_color = setting_value("green_color", "#10B981")
        yellow_color = setting_value("yellow_color", "#F59E0B")
        red_color = setting_value("red_color", "#EF4444")
        with st.container(border=True):
            st.markdown("##### 顏色與圖例")
            st.caption("有選系列欄位或啟用燈號上色時，圖表會顯示圖例。")
            if chart_type == "長條圖":
                color_cols = st.columns(5)
                traffic_color_enabled = color_cols[0].checkbox(
                    "依紅黃綠燈上色",
                    value=setting_value("traffic_color_enabled", False),
                    disabled=range_mode != "紅黃綠燈" or "判定" not in chart_source.columns,
                )
                bar_color = color_cols[1].color_picker("一般柱體顏色", value=setting_value("bar_color", "#2563EB"))
                green_color = color_cols[2].color_picker("綠燈顏色", value=green_color)
                yellow_color = color_cols[3].color_picker("黃燈顏色", value=yellow_color)
                red_color = color_cols[4].color_picker("紅燈顏色", value=red_color)
            else:
                bar_color = setting_value("bar_color", "#2563EB")
                st.caption("長條圖可設定柱體顏色；折線圖可在系列顏色設定各線條顏色。")
        traffic_colors = {
            "綠燈": green_color,
            "黃燈": yellow_color,
            "紅燈": red_color,
            "未分類": "#94A3B8",
        }
        series_colors = {}
        if series_col:
            with st.container(border=True):
                st.markdown("##### 系列顏色")
                series_values = [value for value in chart_source[series_col].dropna().astype(str).unique().tolist() if value]
                default_palette = ["#2563EB", "#EF4444", "#10B981", "#F59E0B", "#8B5CF6", "#06B6D4", "#64748B", "#EC4899"]
                color_columns = st.columns(min(4, max(1, len(series_values))))
                for idx, series_value in enumerate(series_values):
                    key = f"series_color_{series_col}_{series_value}"
                    series_colors[series_value] = color_columns[idx % len(color_columns)].color_picker(
                        series_value,
                        value=setting_value(key, default_palette[idx % len(default_palette)]),
                    )

        limit_lines = build_limit_lines(range_mode, lower_value, upper_value, traffic_ranges)

        fig = make_chart(
            chart_source,
            chart_type,
            x_col,
            y_mode,
            bar_color,
            lower_value,
            upper_value,
            y_min if y_range_enabled else None,
            y_max if y_range_enabled else None,
            chart_title,
            x_axis_label,
            y_axis_label,
            traffic_color_enabled,
            traffic_colors,
            pie_label_mode,
            series_col if series_col else None,
            series_colors,
            show_data_labels,
            y_tick_interval if y_tick_interval and y_tick_interval > 0 else None,
            limit_lines,
        )
        if fig is None:
            st.info("目前篩選結果沒有可繪圖的資料。")
        else:
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("##### 匯出圖表")
            export_cols = st.columns(3)
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
                chart_svg = fig.to_image(format="svg")
                export_cols[2].download_button(
                    "匯出圖表 SVG",
                    data=chart_svg,
                    file_name=f"圖表_{datetime.now().strftime('%Y%m%d_%H%M')}.svg",
                    mime="image/svg+xml",
                    use_container_width=True,
                )
            except Exception:
                pass

    step_nav(back_step=4)

export_name = f"行動檢修平台整理_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
export_display = compact_repeated_values(
    filtered[display_cols],
    ["來源檔案", "工單編號", "車號/最小成本", "檢查結束日期"],
).rename(columns={"檢查結果數值": "檢查結果"})
export_display = format_numeric_columns(export_display)
export_stats = format_numeric_columns(stats_df)
export_bytes = to_excel_bytes({"篩選結果": export_display, "統計": export_stats})
st.download_button(
    "匯出目前篩選結果",
    data=export_bytes,
    file_name=export_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
