import streamlit as st
import pandas as pd
import plotly.express as px

# 設定版面
st.set_page_config(page_title="自訂報表系統", layout="wide")
st.title("📊 自訂資料提取與圖表生成")

# ==========================================
# 1. 檔案上傳區
# ==========================================
uploaded_file = st.file_uploader("📂 請上傳一份 Excel 檔案", type=["xlsx"])

if uploaded_file:
    # 嘗試讀取資料，並明確指定引擎為 openpyxl
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("✅ 檔案讀取成功！")
        st.divider()
        
        # ==========================================
        # 2. 客製化提取：欄位選擇
        # ==========================================
        st.subheader("⚙️ 步驟一：選擇您需要的欄位")
        all_columns = df.columns.tolist()
        selected_cols = st.multiselect("請勾選要提取的欄位：", all_columns, default=all_columns)
        
        if selected_cols:
            # 依照選擇過濾資料表
            filtered_df = df[selected_cols]
            st.dataframe(filtered_df, use_container_width=True)
            st.divider()
            
            # ==========================================
            # 3. 動態圖表生成
            # ==========================================
            st.subheader("📊 步驟二：動態生成圖表")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                chart_type = st.selectbox("選擇圖表類型", ["長條圖 (Bar)", "折線圖 (Line)"])
            with col2:
                x_axis = st.selectbox("選擇 X 軸 (分類/時間)", selected_cols)
            with col3:
                y_axis = st.selectbox("選擇 Y 軸 (數值)", selected_cols)
                
            if st.button("✨ 生成圖表", type="primary"):
                try:
                    # 群組加總資料
                    plot_data = filtered_df.groupby(x_axis, as_index=False)[y_axis].sum()
                    
                    # 判斷圖表類型並繪圖
                    if chart_type == "長條圖 (Bar)":
                        fig = px.bar(plot_data, x=x_axis, y=y_axis, text_auto='.0f')
                    else:
                        fig = px.line(plot_data, x=x_axis, y=y_axis, markers=True)
                        
                    st.plotly_chart(fig, use_container_width=True)
                except Exception as e:
                    st.error("⚠️ 圖表生成失敗！請確認您選擇的【Y 軸】欄位裡面是否都是數字。")
                    
    except Exception as e:
        st.error(f"❌ 讀取 Excel 失敗，請確認檔案格式是否正確。詳細錯誤：{e}")