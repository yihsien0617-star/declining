import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import plotly.express as px

st.set_page_config(page_title="招生三階段漏斗分析系統", layout="wide")
st.title("📊 申請入學 三階段生源漏斗與地理分析")

# --- 1. 側邊欄：檔案上傳區 ---
st.sidebar.header("📂 資料上傳區")
st.sidebar.markdown("💡 **提示**：請確保三份檔案中都包含 `姓名` 欄位。")

file_stage1 = st.sidebar.file_uploader("1️⃣ 上傳【第一階段】總表 (需含經緯度與姓名)", type=["csv", "xlsx"])
file_stage2 = st.sidebar.file_uploader("2️⃣ 上傳【第二階段】名單 (需含姓名)", type=["csv", "xlsx"])
file_stage3 = st.sidebar.file_uploader("3️⃣ 上傳【最終入學】名單 (需含姓名)", type=["csv", "xlsx"])

# --- 2. 資料讀取與比對邏輯 (改用姓名比對並清除空白) ---
@st.cache_data
def load_and_merge_data(f1, f2, f3):
    # 讀取一階母體資料
    df = pd.read_csv(f1) if f1.name.endswith('.csv') else pd.read_excel(f1)
    
    # 清理一階姓名欄位：轉為字串並去除前後空白，避免比對失敗
    df['姓名'] = df['姓名'].astype(str).str.strip()
    
    # 預設所有人都是「1_僅通過一階」
    df['最終狀態'] = '1_僅通過一階'

    # 如果有上傳二階名單，進行姓名比對
    if f2:
        df2 = pd.read_csv(f2) if f2.name.endswith('.csv') else pd.read_excel(f2)
        # 清除空白並轉為 List
        stage2_names = df2['姓名'].astype(str).str.strip().tolist()
        # 將有在二階名單中的人，狀態升級
        df.loc[df['姓名'].isin(stage2_names), '最終狀態'] = '2_進入二階(未入學)'

    # 如果有上傳最終入學名單，進行姓名比對
    if f3:
        df3 = pd.read_csv(f3) if f3.name.endswith('.csv') else pd.read_excel(f3)
        # 清除空白並轉為 List
        stage3_names = df3['姓名'].astype(str).str.strip().tolist()
        # 將有在入學名單中的人，狀態升級到最高
        df.loc[df['姓名'].isin(stage3_names), '最終狀態'] = '3_最終入學'

    return df

# --- 3. 主程式介面 ---
if file_stage1:
    try:
        df = load_and_merge_data(file_stage1, file_stage2, file_stage3)
        st.sidebar.success("✅ 資料載入與姓名比對完成！")
        
        # --- 篩選器 ---
        departments = df['系(組)、學程名稱'].dropna().unique().tolist()
        selected_dept = st.sidebar.selectbox("🎓 選擇分析系(組)、學程", ["全校總覽"] + departments)
        
        if selected_dept != "全校總覽":
            df = df[df['系(組)、學程名稱'] == selected_dept]
            st.subheader(f"目前顯示：{selected_dept}")
        else:
            st.subheader("目前顯示：全校總覽")

        # --- KPI 儀表板 ---
        s1_count = len(df)
        s2_count = len(df[df['最終狀態'].isin(['2_進入二階(未入學)', '3_最終入學'])])
        s3_count = len(df[df['最終狀態'] == '3_最終入學'])
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("第一階段總人數", f"{s1_count} 人")
        col2.metric("進入第二階段", f"{s2_count} 人", f"轉換率 {(s2_count/s1_count*100):.1f}%" if s1_count else "0%")
        col3.metric("最終註冊入學", f"{s3_count} 人", f"報到率 {(s3_count/s2_count*100):.1f}%" if s2_count else "0%")
        col4.metric("一階至入學(總留存)", f"{(s3_count/s1_count*100):.1f} %" if s1_count else "0 %")
        st.markdown("---")

        # --- 頁籤 ---
        tab1, tab2, tab3 = st.tabs(["🗺️ 最終狀態地理分布", "📉 生源漏斗分析圖", "📋 高中端轉換率明細"])

        with tab1:
            st.markdown("🔴 **紅點**：僅通過一階 ｜ 🔵 **藍點**：進入二階(未入學) ｜ 🟢 **綠點**：最終入學")
            m = folium.Map(location=[22.934, 120.246], zoom_start=8, tiles='CartoDB positron')
            
            color_map = {
                '1_僅通過一階': 'red',
                '2_進入二階(未入學)': 'blue',
                '3_最終入學': 'green'
            }

            for status in ['1_僅通過一階', '2_進入二階(未入學)', '3_最終入學']:
                subset = df[df['最終狀態'] == status]
                for _, row in subset.iterrows():
                    if pd.notnull(row.get('學校緯度')) and pd.notnull(row.get('學校經度')):
                        folium.CircleMarker(
                            location=[row['學校緯度'], row['學校經度']],
                            radius=4 if status != '3_最終入學' else 5,
                            color=color_map[status], 
                            fill=True, 
                            fill_color=color_map[status], 
                            fill_opacity=0.6,
                            popup=f"{row['畢業學校']} ({status.split('_')[1]})", 
                            tooltip=row['畢業學校']
                        ).add_to(m)

            st_folium(m, width=1000, height=600)

        with tab2:
            st.markdown("### 主要生源學校 三階段漏斗比較 (TOP 15)")
            school_stats = df.groupby(['畢業學校', '最終狀態']).size().reset_index(name='人數')
            
            top_schools = df['畢業學校'].value_counts().nlargest(15).index
            filtered_stats = school_stats[school_stats['畢業學校'].isin(top_schools)]
            
            fig = px.bar(
                filtered_stats, x='畢業學校', y='人數', color='最終狀態', 
                barmode='stack',
                color_discrete_map={
                    '1_僅通過一階': '#EF553B', 
                    '2_進入二階(未入學)': '#636EFA',
                    '3_最終入學': '#00CC96'
                },
                text_auto=True
            )
            st.plotly_chart(fig, use_container_width=True)

        with tab3:
            st.markdown(f"### 🏆 {selected_dept} - 各高中三階段轉換率明細")
            
            pivot_df = df.groupby(['畢業學校', '最終狀態']).size().unstack(fill_value=0).reset_index()
            
            for col in ['1_僅通過一階', '2_進入二階(未入學)', '3_最終入學']:
                if col not in pivot_df.columns:
                    pivot_df[col] = 0
            
            pivot_df['[A]一階總人數'] = pivot_df['1_僅通過一階'] + pivot_df['2_進入二階(未入學)'] + pivot_df['3_最終入學']
            pivot_df['[B]二階總人數'] = pivot_df['2_進入二階(未入學)'] + pivot_df['3_最終入學']
            pivot_df['[C]最終入學人數'] = pivot_df['3_最終入學']
            
            pivot_df['一階轉二階(%)'] = (pivot_df['[B]二階總人數'] / pivot_df['[A]一階總人數'] * 100).round(1).fillna(0)
            pivot_df['二階轉入學(%)'] = (pivot_df['[C]最終入學人數'] / pivot_df['[B]二階總人數'] * 100).round(1).fillna(0)
            pivot_df['總留存率(%)'] = (pivot_df['[C]最終入學人數'] / pivot_df['[A]一階總人數'] * 100).round(1).fillna(0)
            
            display_cols = ['畢業學校', '[A]一階總人數', '[B]二階總人數', '[C]最終入學人數', '一階轉二階(%)', '二階轉入學(%)', '總留存率(%)']
            final_df = pivot_df[display_cols].sort_values(by='[A]一階總人數', ascending=False).reset_index(drop=True)
            
            st.dataframe(final_df, use_container_width=True)

    except Exception as e:
        st.error(f"資料處理發生錯誤，請確認上傳的檔案中是否都有包含「姓名」這個欄位。錯誤訊息：{e}")
else:
    st.info("👈 請在左側上傳【第一階段總表】以開始分析。")
