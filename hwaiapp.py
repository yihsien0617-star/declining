import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import plotly.express as px

# 設定頁面與標題
st.set_page_config(page_title="招生生源地理動態分析系統", layout="wide")
st.title("📊 各系所申請入學 生源與地域分析儀表板")

# 1. 檔案上傳區
uploaded_file = st.sidebar.file_uploader("📂 請上傳申請生資料 (Excel/CSV)", type=["xlsx", "csv"])

if uploaded_file:
    # 讀取資料 (假設欄位包含：入學學年度, 系(組)、學程名稱, 畢業學校, 地址, 階段狀態, 學校緯度, 學校經度, 居住地緯度, 居住地經度)
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
        
    st.sidebar.success("資料載入成功！")
    
    # 2. 動態篩選器
    years = df['入學學年度'].dropna().unique().tolist()
    selected_year = st.sidebar.selectbox("📅 選擇入學學年度", sorted(years, reverse=True))
    
    departments = df['系(組)、學程名稱'].dropna().unique().tolist()
    selected_dept = st.sidebar.selectbox("🎓 選擇系(組)、學程", ["全校總覽"] + departments)
    
    # 過濾資料
    filtered_df = df[df['入學學年度'] == selected_year]
    if selected_dept != "全校總覽":
        filtered_df = filtered_df[filtered_df['系(組)、學程名稱'] == selected_dept]

    # 3. 建立主要頁籤
    tab1, tab2, tab3 = st.tabs(["🗺️ 地理分佈地圖", "📈 生源學校統計", "🔄 一二階轉換率分析"])

    with tab1:
        st.subheader(f"{selected_year}學年度 - {selected_dept} 地理分佈")
        map_type = st.radio("選擇地圖視角：", ["高中就讀學校分佈", "學生居住地分佈"], horizontal=True)
        
        # 預設地圖中心 (以台南市仁德區為例)
        m = folium.Map(location=[22.934, 120.246], zoom_start=8)
        
        # 繪製地圖標記 (需確保資料內有對應的經緯度欄位)
        if map_type == "高中就讀學校分佈" and '學校緯度' in filtered_df.columns:
            school_counts = filtered_df.groupby(['畢業學校', '學校緯度', '學校經度']).size().reset_index(name='人數')
            for idx, row in school_counts.iterrows():
                folium.CircleMarker(
                    location=[row['學校緯度'], row['學校經度']],
                    radius=row['人數'] * 2, # 依人數調整大小
                    popup=f"{row['畢業學校']}: {row['人數']}人",
                    color="blue", fill=True
                ).add_to(m)
        elif map_type == "學生居住地分佈" and '居住地緯度' in filtered_df.columns:
             # 針對地址或郵遞區號轉換的經緯度進行繪製
             pass # 實作居住地標記邏輯
             
        st_folium(m, width=1000, height=500)

    with tab2:
        st.subheader("🏆 主要生源學校排行 (TOP 10)")
        top_schools = filtered_df['畢業學校'].value_counts().reset_index()
        top_schools.columns = ['畢業學校', '申請人數']
        fig = px.bar(top_schools.head(10), x='畢業學校', y='申請人數', text='申請人數', color='申請人數', color_continuous_scale='Viridis')
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(top_schools) # 提供系主任匯出或詳細查看

    with tab3:
        st.subheader("📉 第一階段 vs 第二階段 漏斗分析")
        st.markdown("觀察重點：是否距離較遠的縣市/高中的學生，在第一階段通過後，放棄參加第二階段的面試或報到？")
        # 假設資料中有 '階段狀態' 欄位紀錄：'一階通過', '二階報名', '錄取', '報到'
        # 此處可加入長條圖或漏斗圖來比較不同地區的轉換率
        # (實作略...)
else:
    st.info("請於左側上傳申請生資料檔 (.xlsx 或 .csv) 以開始分析。")
