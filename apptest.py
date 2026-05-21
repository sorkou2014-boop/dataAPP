import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# ==========================================
# 🌟 1. 設定網頁標題與大版面
# ==========================================
st.set_page_config(page_title="自訂資料提取與圖表系統", layout="wide", page_icon="📊")

st.title("📊 業務資料動態提取與分析平台")
st.markdown("上傳一份或多份 Excel 檔案，即可自由選擇所需欄位、過濾條件，並動態生成專屬圖表。")
st.divider()

# ==========================================
# 🛠️ 2. 資料讀取與快取引擎
# ==========================================
# 使用 st.cache_data 確保切換條件時，不用重複讀取檔案，提升流暢度
@st.cache_data
def load_and_merge_data(uploaded_files):
    df_list = []
    for file in uploaded_files:
        # 直接讀取 Excel，不預設任何特殊格式
        df = pd.read_excel(file, engine='calamine')
        # 可選：新增一個欄位紀錄資料來源檔案
        # df['來源檔案'] = file.name 
        df_list.append(df)
    
    # 將多份檔案上下疊加合併
    merged_df = pd.concat(df_list, ignore_index=True)
    return merged_df

# ==========================================
# 📂 3. 檔案上傳區
# ==========================================
uploaded_files = st.file_uploader("📂 拖曳或選擇【格式相同】的 Excel 檔案 (可多選)", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    with st.spinner("資料讀取與整併中..."):
        raw_df = load_and_merge_data(uploaded_files)
    
    st.success(f"✅ 成功整併 {len(uploaded_files)} 份檔案，共讀取 {len(raw_df)} 筆資料列！")
    
    # ==========================================
    # ⚙️ 4. 客製化資料提取控制台
    # ==========================================
    st.header("⚙️ 第一步：選擇提取欄位與過濾條件")
    
    col_fields, col_filters = st.columns(2)
    
    with col_fields:
        st.subheader("1. 選擇要保留的欄位")
        all_columns = raw_df.columns.tolist()
        # 預設全選，讓主管可以自己打叉取消
        selected_cols = st.multiselect("📌 勾選您這次需要的欄位：", all_columns, default=all_columns)
        
        if not selected_cols:
            st.warning("⚠️ 請至少選擇一個欄位！")
            st.stop()
            
        # 依據選擇保留欄位
        filtered_df = raw_df[selected_cols]

    with col_filters:
        st.subheader("2. 設定資料過濾條件")
        # 自動抓取文字類型的欄位作為篩選依據
        cat_cols = filtered_df.select_dtypes(include=['object', 'category']).columns.tolist()
        
        if cat_cols:
            filter_col = st.selectbox("🔍 選擇要用來篩選的【條件欄位】(選填)", ["無"] + cat_cols)
            if filter_col != "無":
                unique_vals = filtered_df[filter_col].dropna().unique().tolist()
                selected_vals = st.multiselect(f"✔️ 勾選「{filter_col}」要保留的項目：", unique_vals, default=unique_vals)
                
                # 執行過濾邏輯
                if selected_vals:
                    filtered_df = filtered_df[filtered_df[filter_col].isin(selected_vals)]
                else:
                    st.warning("⚠️ 篩選條件為空，將不顯示任何資料。")
                    filtered_df = pd.DataFrame(columns=selected_cols)
        else:
            st.info("💡 目前勾選的欄位中，沒有適合用來做分類篩選的文字欄位。")

    # ==========================================
    # 📋 5. 提取結果預覽與下載
    # ==========================================
    st.markdown(f"### 📋 提取結果預覽 (共 {len(filtered_df)} 筆符合條件)")
    st.dataframe(filtered_df, use_container_width=True)

    # 處理 Excel 下載
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='自訂提取資料')
    
    st.download_button(
        label="📥 下載此客製化報表 (Excel)",
        data=buf.getvalue(),
        file_name=f"自訂資料提取表_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

    st.divider()

    # ==========================================
    # 📊 6. 動態圖表生成器
    # ==========================================
    st.header("📊 第二步：動態圖表生成")
    
    if not filtered_df.empty:
        col_chart1, col_chart2, col_chart3 = st.columns(3)
        
        # 找出數值與文字欄位供 X/Y 軸使用
        num_cols = filtered_df.select_dtypes(include=['number']).columns.tolist()
        non_num_cols = [c for c in selected_cols if c not in num_cols]
        
        with col_chart1:
            chart_type = st.selectbox("📈 選擇圖表類型", ["長條圖 (Bar Chart)", "折線圖 (Line Chart)", "圓餅圖 (Pie Chart)", "散佈圖 (Scatter Plot)"])
        with col_chart2:
            x_axis = st.selectbox("👉 選擇 X 軸 (分類/時間)", non_num_cols if non_num_cols else selected_cols)
        with col_chart3:
            y_axis = st.selectbox("☝️ 選擇 Y 軸 (數值/加總目標)", num_cols if num_cols else selected_cols)

        if st.button("✨ 立即生成圖表", type="primary", use_container_width=True):
            try:
                # 準備繪圖資料：如果選定的 X 軸和 Y 軸合法，先進行群組加總處理 (散佈圖除外)
                if chart_type in ["長條圖 (Bar Chart)", "折線圖 (Line Chart)", "圓餅圖 (Pie Chart)"]:
                    # 動態群組並加總
                    plot_data = filtered_df.groupby(x_axis, as_index=False)[y_axis].sum()
                    
                    if chart_type == "長條圖 (Bar Chart)":
                        fig = px.bar(plot_data, x=x_axis, y=y_axis, title=f"各【{x_axis}】的【{y_axis}】總和", text_auto='.0f')
                    elif chart_type == "折線圖 (Line Chart)":
                        fig = px.line(plot_data, x=x_axis, y=y_axis, markers=True, title=f"【{x_axis}】與【{y_axis}】趨勢圖")
                    elif chart_type == "圓餅圖 (Pie Chart)":
                        fig = px.pie(plot_data, names=x_axis, values=y_axis, title=f"【{x_axis}】的【{y_axis}】佔比圖")
                
                else:
                    # 散佈圖不需群組加總，直接使用過濾後的明細資料
                    fig = px.scatter(filtered_df, x=x_axis, y=y_axis, title=f"【{x_axis}】與【{y_axis}】資料點分佈")

                # 優化圖表外觀
                fig.update_layout(xaxis_tickangle=-45, title_x=0.5)
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.error("⚠️ 生成圖表時發生錯誤！請確認：(1) Y 軸是否為可計算的『數值』欄位 (2) 該欄位是否有空值。")
                st.exception(e) # 開發期間顯示詳細錯誤，上線前可註解掉
    else:
        st.info("無資料可供繪圖，請放寬上方的篩選條件。")

else:
    st.info("👈 請先從上方上傳 Excel 檔案以啟用控制台。")