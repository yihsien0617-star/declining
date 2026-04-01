import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import plotly.express as px

# --- 1. 頁面基本設定 ---
st.set_page_config(page_title="招生生源一二階轉換分析系統", layout="wide")
st.title("📊 申請入學 一、二階生源轉換與地理分析")

# --- 2. 側邊欄：檔案上傳區 ---
st.sidebar.header("📂 資料上傳區")
st.sidebar.markdown("請上傳已包含**學校緯度**與**學校經度**的檔案")

file_stage1 = st.sidebar.file_uploader("上傳【第一階段】名單 (CSV/Excel)", type=["csv", "xlsx"])
file_stage2 = st.sidebar.file_uploader("上傳【第二階段】名單 (CSV/Excel)", type=["csv", "xlsx"])

# --- 3. 資料讀取與合併功能 ---
@st.cache_data
def load_data(f1, f2):
    if f1.name.endswith('.csv'):
        df1 = pd.read_csv(f1)
    else:
        df1 = pd.read_excel(f1)
    df1['階段'] = '第一階段'

    if f2.name.endswith('.csv'):
        df2 = pd.read_csv(f2)
    else:
        df2 = pd.read_excel(f2)
    df2['階段'] = '第二階段'

    combined_df = pd.concat([df1, df2], ignore_index=True)
    return combined_df

# --- 4. 主程式邏輯 ---
if file_stage1 and file_stage2:
    try:
        df = load_data(file_stage1, file_stage2)
        st.sidebar.success("✅ 兩份名單載入成功！")
        
        # --- 篩選器 ---
        departments = df['系(組)、學程名稱'].dropna().unique().tolist()
        selected_dept = st.sidebar.selectbox("🎓 選擇分析系(組)、學程", ["全校總覽"] + departments)
        
        if selected_dept != "全校總覽":
            df = df[df['系(組)、學程名稱'] == selected_dept]
            st.subheader(f"目前顯示：{selected_dept}")
        else:
            st.subheader("目前顯示：全校總覽")

        # --- 新增：關鍵指標 (KPI) 區塊 ---
        stage1_count = len(df[df['階段'] == '第一階段'])
        stage2_count = len(df[df['階段'] == '第二階段'])
        conversion_rate = (stage2_count / stage1_count * 100) if stage1_count > 0 else 0
        
        col1, col2, col3 = st.columns(3)
        col1.metric("第一階段總人數", f"{stage1_count} 人")
        col2.metric("第二階段總人數", f"{stage2_count} 人")
        col3.metric("整體二階轉換率", f"{conversion_rate:.1f} %")
        st.markdown("---")

        # --- 建立三個頁籤 ---
        tab1, tab2, tab3 = st.tabs(["🗺️ 畢業學校地圖分析", "📉 區域轉換率圖表", "📋 詳細統計資料表"])

        with tab1:
            st.markdown("🔴 **紅點**：僅參加第一階段 (未進入二階) ｜ 🔵 **藍點**：進入第二階段")
            # 預設地圖中心 (以 Chung Hwa University of Medical Technology 概略座標為中心)
            m = folium.Map(location=[22.934, 120.246], zoom_start=8, tiles='CartoDB positron')
            
            for idx, row in df[df['階段'] == '第一階段'].iterrows():
                if pd.notnull(row.get('學校緯度')) and pd.notnull(row.get('學校經度')):
                    folium.CircleMarker(
                        location=[row['學校緯度'], row['學校經度']],
                        radius=4, color="red", fill=True, fill_color="red", fill_opacity=0.4,
                        popup=f"{row['畢業學校']} (一階)", tooltip=row['畢業學校']
                    ).add_to(m)

            for idx, row in df[df['階段'] == '第二階段'].iterrows():
                if pd.notnull(row.get('學校緯度')) and pd.notnull(row.get('學校經度')):
                    folium.CircleMarker(
                        location=[row['學校緯度'], row['學校經度']],
                        radius=4, color="blue", fill=True, fill_color="blue", fill_opacity=0.7,
                        popup=f"{row['畢業學校']} (二階)", tooltip=row['畢業學校']
                    ).add_to(m)

            st_folium(m, width=1000, height=600)

        with tab2:
            st.markdown("### 主要生源學校 一階 vs 二階 比較 (TOP 15)")
            school_stats = df.groupby(['畢業學校', '階段']).size().reset_index(name='人數')
            
            # 找出第一階段人數最多的前 15 所高中
            top_schools = school_stats[school_stats['階段'] == '第一階段'].nlargest(15, '人數')['畢業學校']
            filtered_stats = school_stats[school_stats['畢業學校'].isin(top_schools)]
            
            fig = px.bar(
                filtered_stats, x='畢業學校', y='人數', color='階段', 
                barmode='group',
                color_discrete_map={'第一階段': '#EF553B', '第二階段': '#00CC96'},
                text_auto=True # 自動在柱子上顯示數字
            )
            st.plotly_chart(fig, use_container_width=True)

        # --- 新增：統計資料與明細頁籤 ---
        with tab3:
            st.markdown(f"### 🏆 {selected_dept} - 各高中生源統計與轉換率明細")
            st.info("💡 提示：您可以點擊表格右上角的圖示，將這份統計資料下載為 CSV 檔。")
            
            # 製作一二階人數的樞紐分析表
            pivot_df = df.groupby(['畢業學校', '階段']).size().unstack(fill_value=0).reset_index()
            
            # 確保欄位存在 (避免某些系只有一階沒有二階資料會報錯)
            if '第一階段' not in pivot_df.columns: pivot_df['第一階段'] = 0
            if '第二階段' not in pivot_df.columns: pivot_df['第二階段'] = 0
                
            # 計算該高中的二階轉換率
            pivot_df['二階轉換率(%)'] = (pivot_df['第二階段'] / pivot_df['第一階段'] * 100).round(1)
            pivot_df['二階轉換率(%)'] = pivot_df['二階轉換率(%)'].fillna(0) # 處理分母為0的狀況
            
            # 依照第一階段人數由高到低排序
            pivot_df = pivot_df.sort_values(by='第一階段', ascending=False).reset_index(drop=True)
            
            # 顯示表格 (設定 use_container_width 讓表格自動適應螢幕寬度)
            st.dataframe(pivot_df, use_container_width=True)

    except Exception as e:
        st.error(f"資料處理發生錯誤，請確認上傳的檔案格式與欄位是否正確。錯誤訊息：{e}")
else:
    st.info("👈 請在左側上傳【第一階段】與【第二階段】的檔案，開始進行分析。")
