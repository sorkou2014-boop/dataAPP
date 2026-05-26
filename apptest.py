import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# ==========================================
# 🌟 1. 系統全域設定
# ==========================================
st.set_page_config(page_title="綜合維修數據分析系統", layout="wide", page_icon="📈")
st.title("📈 綜合維修數據分析與提取平台")
st.divider()

# ==========================================
# 🛠️ 2. 共用函式庫
# ==========================================
@st.cache_data
def convert_df_to_excel(df):
    """將 DataFrame 轉換為 Excel 供下載"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def render_chart_builder(df, prefix_key=""):
    """共用的動態圖表產生器"""
    st.subheader("📊 動態圖表產生器")
    if df.empty:
        st.warning("目前沒有資料可供繪製圖表。")
        return
        
    c1, c2, c3 = st.columns(3)
    with c1:
        chart_type = st.selectbox("圖表類型", ["長條圖", "折線圖", "圓餅圖"], key=f"type_{prefix_key}")
    with c2:
        x_axis = st.selectbox("選擇 X 軸 (分類)", df.columns, key=f"x_{prefix_key}")
    with c3:
        # 篩選數值欄位供 Y 軸選擇，若無則提供計數選項
        num_cols = df.select_dtypes(include=['number']).columns.tolist()
        y_axis = st.selectbox("選擇 Y 軸 (數值)", ["資料筆數 (Count)"] + num_cols, key=f"y_{prefix_key}")

    if st.button("✨ 產生圖表", key=f"btn_{prefix_key}"):
        try:
            if y_axis == "資料筆數 (Count)":
                plot_data = df[x_axis].value_counts().reset_index()
                plot_data.columns = [x_axis, '筆數']
                y_col = '筆數'
            else:
                plot_data = df.groupby(x_axis, as_index=False)[y_axis].sum()
                y_col = y_axis

            if chart_type == "長條圖":
                fig = px.bar(plot_data, x=x_axis, y=y_col, text_auto=True, title=f"{x_axis} 統計圖")
            elif chart_type == "折線圖":
                fig = px.line(plot_data, x=x_axis, y=y_col, markers=True, title=f"{x_axis} 趨勢圖")
            else:
                fig = px.pie(plot_data, names=x_axis, values=y_col, title=f"{x_axis} 佔比圖")
            
            st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.error(f"圖表產生失敗，請確認欄位格式是否正確。錯誤：{e}")

# ==========================================
# 🎛️ 3. 側邊欄：選擇表單類型
# ==========================================
st.sidebar.header("⚙️ 選擇分析模組")
app_mode = st.sidebar.radio(
    "請選擇您要處理的表單類型：",
    ["1️⃣ 行動檢修平台 (需處理特規表單)", "2️⃣ EFMS工法統計 (結構化資料分析)"]
)

# ==========================================
# 🚀 模組 1：行動檢修平台 (Attachment 1)
# ==========================================
if app_mode == "1️⃣ 行動檢修平台 (需處理特規表單)":
    st.header("🛠️ 行動檢修平台 - 多檔整合與分析")
    uploaded_files = st.file_uploader("📂 請上傳多個檢修平台 Excel 檔", type=["xlsx", "xls"], accept_multiple_files=True)
    
    if uploaded_files:
        @st.cache_data
        def process_iso_forms(files):
            all_records = []
            for file in files:
                # 這裡沿用你之前的邏輯：上層抓 Metadata，下層抓表格
                try:
                    df_raw = pd.read_excel(file, header=None)
                    meta_data = {"工單號碼": "", "車號/最小成本": "", "檢查結束日期": ""}
                    
                    # 抓取表頭資訊
                    for r in range(min(15, len(df_raw))):
                        for c in range(len(df_raw.columns)):
                            cell_val = str(df_raw.iloc[r, c]).strip()
                            if "工單編號" in cell_val: meta_data["工單號碼"] = df_raw.iloc[r, c+1]
                            elif "車號/最小成本" in cell_val: meta_data["車號/最小成本"] = df_raw.iloc[r, c+1]
                            elif "檢查結束日期" in cell_val: meta_data["檢查結束日期"] = df_raw.iloc[r, c+1]
                    
                    # 抓取表格區段 (假設從含有'進階分類'的那行開始)
                    df_table = pd.read_excel(file, header=2) # 依據實際狀況調整 header 行數
                    if '進階分類' in df_table.columns and '檢查項目' in df_table.columns:
                        for _, row in df_table.iterrows():
                            # 尋找檢查結果相關的欄位
                            result_col = [col for col in df_table.columns if '檢查結果' in str(col)]
                            res_val = row[result_col[0]] if result_col else ""
                            
                            record = meta_data.copy()
                            record.update({
                                "進階分類": row.get('進階分類', ''),
                                "檢查項目": row.get('檢查項目', ''),
                                "檢查結果": res_val
                            })
                            all_records.append(record)
                except Exception as e:
                    st.warning(f"檔案 {file.name} 處理略過。原因: {e}")
            return pd.DataFrame(all_records)
            
        with st.spinner("高速萃取與整併中..."):
            df_action = process_iso_forms(uploaded_files)
            
        if not df_action.empty:
            st.success(f"✅ 成功整合 {len(df_action)} 筆檢修紀錄！")
            
            # --- 篩選區 ---
            st.subheader("🔍 進階整合與篩選")
            col1, col2 = st.columns(2)
            with col1:
                adv_classes = df_action['進階分類'].dropna().unique().tolist()
                sel_class = st.multiselect("過濾 A.[進階分類]", adv_classes)
            with col2:
                items = df_action['檢查項目'].dropna().unique().tolist()
                sel_item = st.multiselect("過濾 B.[檢查項目]", items)
                
            filtered_act = df_action.copy()
            if sel_class: filtered_act = filtered_act[filtered_act['進階分類'].isin(sel_class)]
            if sel_item: filtered_act = filtered_act[filtered_act['檢查項目'].isin(sel_item)]
            
            # --- 預覽與匯出 ---
            st.markdown(f"**📋 C.[檢查結果] 整合清單預覽 (共 {len(filtered_act)} 筆)**")
            st.dataframe(filtered_act, use_container_width=True)
            st.download_button("📥 匯出整合清單 (Excel)", convert_df_to_excel(filtered_act), "行動檢修整合清單.xlsx")
            
            # --- 附屬計算功能 ---
            st.subheader("🧮 附屬計算功能 (數值結果)")
            # 嘗試將結果轉為數值以進行計算
            filtered_act['數值結果'] = pd.to_numeric(filtered_act['檢查結果'], errors='coerce')
            calc_df = filtered_act.dropna(subset=['數值結果'])
            
            if not calc_df.empty:
                stats = calc_df.groupby(['進階分類', '檢查項目'])['數值結果'].agg(['min', 'max', 'mean']).reset_index()
                stats.rename(columns={'min': '最小值', 'max': '最大值', 'mean': '平均值'}, inplace=True)
                st.dataframe(stats, use_container_width=True)
                st.download_button("📥 匯出計算結果 (Excel)", convert_df_to_excel(stats), "數值計算結果.xlsx")
            else:
                st.info("目前篩選的資料中沒有可計算的數值。")
                
            st.divider()
            render_chart_builder(filtered_act, prefix_key="act")

# ==========================================
# 🚀 模組 2：EFMS 工法統計 (Attachment 2)
# ==========================================
elif app_mode == "2️⃣ EFMS工法統計 (結構化資料分析)":
    st.header("📊 EFMS 工法統計分析")
    uploaded_efms = st.file_uploader("📂 請上傳 EFMS 統計 Excel 檔", type=["xlsx", "xls"], accept_multiple_files=True)
    
    if uploaded_efms:
        @st.cache_data
        def process_efms(files):
            df_list = [pd.read_excel(f) for f in files]
            df = pd.concat(df_list, ignore_index=True)
            
            # 🌟 核心需求：將分類 1, 2, 3... 分列表示 (Unpivot 展開)
            # 尋找基礎欄位
            base_cols = ['子統別名稱', '站別', '報修單號', '人工時', '報修日期', '報修等級', '報修單狀態', '報修症狀', '報修症狀描述']
            exist_base = [c for c in base_cols if c in df.columns]
            
            # 動態找出有幾組故障分類 (例如1, 2, 3...)
            melted_rows = []
            for i in range(1, 6): # 假設最多到分類5
                grp_cols = [f'故障分類{i}', f'故障原因{i}', f'故障情形{i}', f'處理動作{i}']
                if all(c in df.columns for c in grp_cols):
                    # 擷取基礎欄位 + 該組故障欄位
                    temp_df = df[exist_base + grp_cols].copy()
                    # 重新命名去除數字，以便整合
                    temp_df.rename(columns={
                        f'故障分類{i}': '故障分類', 
                        f'故障原因{i}': '故障原因', 
                        f'故障情形{i}': '故障情形', 
                        f'處理動作{i}': '處理動作'
                    }, inplace=True)
                    # 剃除空白的分類列
                    temp_df = temp_df.dropna(subset=['故障分類'])
                    melted_rows.append(temp_df)
                    
            if melted_rows:
                final_efms = pd.concat(melted_rows, ignore_index=True)
            else:
                final_efms = df # 若無後綴數字，保持原樣
            return final_efms
            
        with st.spinner("資料展開與分析中..."):
            df_efms = process_efms(uploaded_efms)
            
        st.success("✅ EFMS 資料讀取完畢 (包含故障分列處理)！")
        
        # --- 主要架構篩選 ---
        st.subheader("🔍 主要架構與進階歸類")
        c1, c2, c3, c4 = st.columns(4)
        with c1: f_sub = st.multiselect("過濾 [子統別名稱]", df_efms.get('子統別名稱', pd.Series()).dropna().unique())
        with c2: f_sta = st.multiselect("過濾 [站別]", df_efms.get('站別', pd.Series()).dropna().unique())
        with c3: f_cat = st.multiselect("歸類 A.[故障分類]", df_efms.get('故障分類', pd.Series()).dropna().unique())
        with c4: f_rea = st.multiselect("歸類 B.[故障原因]", df_efms.get('故障原因', pd.Series()).dropna().unique())
        
        filtered_efms = df_efms.copy()
        if f_sub: filtered_efms = filtered_efms[filtered_efms['子統別名稱'].isin(f_sub)]
        if f_sta: filtered_efms = filtered_efms[filtered_efms['站別'].isin(f_sta)]
        if f_cat: filtered_efms = filtered_efms[filtered_efms['故障分類'].isin(f_cat)]
        if f_rea: filtered_efms = filtered_efms[filtered_efms['故障原因'].isin(f_rea)]
        
        st.markdown(f"**📋 報修清單預覽 (共 {len(filtered_efms)} 筆)**")
        
        # 定義顯示欄位順序
        display_cols = ['報修單號', '故障分類', '故障原因', '故障情形', '處理動作', '報修症狀', '報修症狀描述', '人工時']
        exist_display = [c for c in display_cols if c in filtered_efms.columns]
        st.dataframe(filtered_efms[exist_display], use_container_width=True)
        st.download_button("📥 匯出 EFMS 清單 (Excel)", convert_df_to_excel(filtered_efms[exist_display]), "EFMS處理清單.xlsx")
        
        st.divider()
        render_chart_builder(filtered_efms, prefix_key="efms")