import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Chung Hwa University of Medical Technology - 招生生源分析系統", layout="wide")
st.title("📊 Chung Hwa University of Medical Technology 申請入學 - 跨年度生源與地理分析")

# --- 1. 側邊欄：檔案上傳區 ---
st.sidebar.header("📂 資料上傳區")
st.sidebar.markdown("""
**💡 資料欄位需求：**
- **一階總表**：需有 `入學學年度`、`姓名`、`報考科系`、`畢業學校` 及經緯度。
- **二階名單**：需有 `姓名`、`報考科系`。
- **三階入學**：需有 `姓名`。
""")

file_stage1 = st.sidebar.file_uploader("1️⃣ 上傳【第一階段】總表", type=["csv", "xlsx"])
file_stage2 = st.sidebar.file_uploader("2️⃣ 上傳【第二階段】名單 (選填)", type=["csv", "xlsx"])
file_stage3 = st.sidebar.file_uploader("3️⃣ 上傳【最終入學】名單 (選填)", type=["csv", "xlsx"])

# --- 2. 資料讀取與智慧比對邏輯 ---
@st.cache_data
def load_and_merge_data(f1, f2, f3):
    df = pd.read_csv(f1) if f1.name.endswith('.csv') else pd.read_excel(f1)
    df['姓名'] = df['姓名'].astype(str).str.strip()
    
    if '系(組)、學程名稱' in df.columns:
        df.rename(columns={'系(組)、學程名稱': '報考科系'}, inplace=True)
        
    df['最終狀態'] = '1_僅通過一階'
    
    # 確保學年度為文字格式，方便作為類別標籤
    if '入學學年度' in df.columns:
        df['入學學年度'] = df['入學學年度'].astype(str)
    else:
        df['入學學年度'] = '未知年度'

    if f2:
        df2 = pd.read_csv(f2) if f2.name.endswith('.csv') else pd.read_excel(f2)
        df2['姓名'] = df2['姓名'].astype(str).str.strip()
        dept_col_2 = '報考科系' if '報考科系' in df2.columns else ('系(組)、學程名稱' if '系(組)、學程名稱' in df2.columns else None)
        if dept_col_2:
            df2['複合鍵'] = df2['姓名'] + "_" + df2[dept_col_2].astype(str).str.strip()
            df['複合鍵'] = df['姓名'] + "_" + df['報考科系'].astype(str).str.strip()
            df.loc[df['複合鍵'].isin(df2['複合鍵'].tolist()), '最終狀態'] = '2_進入二階(未入學)'
        else:
            df.loc[df['姓名'].isin(df2['姓名'].tolist()), '最終狀態'] = '2_進入二階(未入學)'

    if f3:
        df3 = pd.read_csv(f3) if f3.name.endswith('.csv') else pd.read_excel(f3)
        stage3_names = df3['姓名'].astype(str).str.strip().tolist()
        df.loc[df['姓名'].isin(stage3_names), '最終狀態'] = '3_最終入學'

    return df

# --- 3. 主程式介面 ---
if file_stage1:
    try:
        df_all = load_and_merge_data(file_stage1, file_stage2, file_stage3)
        st.sidebar.success("✅ 資料載入與比對完成！")
        
        # --- 動態篩選器 ---
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 分析維度設定")
        
        # 學年度多選器
        available_years = sorted(df_all['入學學年度'].unique().tolist(), reverse=True)
        selected_years = st.sidebar.multiselect("📅 選擇入學學年度 (支援多選比較)", available_years, default=available_years)
        
        # 科系單選器
        departments = df_all['報考科系'].dropna().unique().tolist()
        # 預設選項優化
        default_dept = "醫學檢驗生物技術系" if "醫學檢驗生物技術系" in departments else "全校總覽"
        selected_dept = st.sidebar.selectbox("🎓 選擇分析系(組)、學程", ["全校總覽"] + departments, index=(["全校總覽"] + departments).index(default_dept) if default_dept in (["全校總覽"] + departments) else 0)
        
        # 套用篩選
        df = df_all[df_all['入學學年度'].isin(selected_years)]
        if selected_dept != "全校總覽":
            df = df[df['報考科系'] == selected_dept]
            
        st.subheader(f"目前顯示：{selected_dept} (學年度: {', '.join(selected_years)})")

        # --- 頁籤分類 ---
        tab1, tab2, tab3, tab4 = st.tabs(["📈 跨年度趨勢分析", "🗺️ 地理分佈地圖", "📉 單年度漏斗分析", "📋 綜合統計明細"])

        # ==========================================
        # 頁籤 1: 跨年度趨勢分析 (強化的視覺化圖表)
        # ==========================================
        with tab1:
            if len(selected_years) < 2:
                st.warning("⚠️ 請在左側選擇至少「兩個」學年度，以檢視跨年度比較圖表。")
            else:
                st.markdown("### 📊 歷年招生三階段人數趨勢")
                # 計算歷年各階段人數
                trend_data = df.groupby(['入學學年度', '最終狀態']).size().reset_index(name='人數')
                
                # 重新塑形，讓一、二、三階的數據分開
                pivot_trend = trend_data.pivot(index='入學學年度', columns='最終狀態', values='人數').fillna(0).reset_index()
                
                # 確保欄位存在並計算還原的總人數
                for col in ['1_僅通過一階', '2_進入二階(未入學)', '3_最終入學']:
                    if col not in pivot_trend.columns: pivot_trend[col] = 0
                
                pivot_trend['[A]第一階段總數'] = pivot_trend['1_僅通過一階'] + pivot_trend['2_進入二階(未入學)'] + pivot_trend['3_最終入學']
                pivot_trend['[B]第二階段總數'] = pivot_trend['2_進入二階(未入學)'] + pivot_trend['3_最終入學']
                pivot_trend['[C]最終入學人數'] = pivot_trend['3_最終入學']
                
                # 繪製歷年折線圖
                fig_trend = go.Figure()
                fig_trend.add_trace(go.Scatter(x=pivot_trend['入學學年度'], y=pivot_trend['[A]第一階段總數'], mode='lines+markers', name='一階總人數', line=dict(color='#EF553B', width=3), marker=dict(size=10)))
                fig_trend.add_trace(go.Scatter(x=pivot_trend['入學學年度'], y=pivot_trend['[B]第二階段總數'], mode='lines+markers', name='二階總人數', line=dict(color='#636EFA', width=3), marker=dict(size=10)))
                fig_trend.add_trace(go.Scatter(x=pivot_trend['入學學年度'], y=pivot_trend['[C]最終入學人數'], mode='lines+markers', name='最終入學', line=dict(color='#00CC96', width=4), marker=dict(size=12)))
                
                fig_trend.update_layout(title="報名與入學人數歷年變化", xaxis_title="學年度", yaxis_title="人數", hovermode="x unified")
                st.plotly_chart(fig_trend, use_container_width=True)

                st.markdown("---")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("### 🏆 主要生源高中 歷年報考人數對比")
                    # 找出歷年來總人數最多的前 10 所高中
                    top_10_schools = df['畢業學校'].value_counts().nlargest(10).index
                    school_year_df = df[df['畢業學校'].isin(top_10_schools)].groupby(['入學學年度', '畢業學校']).size().reset_index(name='報考人數')
                    
                    fig_school_bar = px.bar(school_year_df, x='畢業學校', y='報考人數', color='入學學年度', barmode='group', title="Top 10 高中跨年度報名變化")
                    st.plotly_chart(fig_school_bar, use_container_width=True)
                
                with col2:
                    st.markdown("### 📉 主要生源高中 歷年最終入學人數")
                    enrolled_df = df[df['最終狀態'] == '3_最終入學']
                    if not enrolled_df.empty:
                        enrolled_school_year = enrolled_df[enrolled_df['畢業學校'].isin(top_10_schools)].groupby(['入學學年度', '畢業學校']).size().reset_index(name='入學人數')
                        fig_enrolled_bar = px.bar(enrolled_school_year, x='畢業學校', y='入學人數', color='入學學年度', barmode='group', title="Top 10 高中跨年度實收變化", color_discrete_sequence=px.colors.qualitative.Pastel)
                        st.plotly_chart(fig_enrolled_bar, use_container_width=True)
                    else:
                        st.info("尚無入學名單資料，無法顯示此圖表。")

        # ==========================================
        # 頁籤 2: 地理分佈地圖
        # ==========================================
        with tab2:
            st.markdown("🔴 **紅點**: 僅一階 ｜ 🔵 **藍點**: 進入二階 ｜ 🟢 **綠點**: 最終入學")
            m = folium.Map(location=[22.934, 120.246], zoom_start=8, tiles='CartoDB positron')
            
            color_map = {'1_僅通過一階': 'red', '2_進入二階(未入學)': 'blue', '3_最終入學': 'green'}

            for status in ['1_僅通過一階', '2_進入二階(未入學)', '3_最終入學']:
                subset = df[df['最終狀態'] == status]
                for _, row in subset.iterrows():
                    if pd.notnull(row.get('學校緯度')) and pd.notnull(row.get('學校經度')):
                        folium.CircleMarker(
                            location=[row['學校緯度'], row['學校經度']],
                            radius=4 if status != '3_最終入學' else 6,
                            color=color_map[status], fill=True, fill_color=color_map[status], fill_opacity=0.6,
                            popup=f"{row['入學學年度']} - {row['畢業學校']}<br>狀態: {status.split('_')[1]}<br>姓名: {row['姓名']}"
                        ).add_to(m)

            st_folium(m, width=1000, height=500)

        # ==========================================
        # 頁籤 3: 單年度/合併漏斗分析
        # ==========================================
        with tab3:
            st.markdown("### 生源轉換漏斗圖")
            st.info("💡 呈現目前所選學年度/科系的總體轉換率。如果漏斗在某個階段急遽縮小，代表該階段流失率過高。")
            
            s1 = len(df)
            s2 = len(df[df['最終狀態'].isin(['2_進入二階(未入學)', '3_最終入學'])])
            s3 = len(df[df['最終狀態'] == '3_最終入學'])
            
            fig_funnel = go.Figure(go.Funnel(
                y=['第一階段通過', '進入第二階段', '最終報到入學'],
                x=[s1, s2, s3],
                textinfo="value+percent initial",
                marker={"color": ["#EF553B", "#636EFA", "#00CC96"]}
            ))
            st.plotly_chart(fig_funnel, use_container_width=True)

        # ==========================================
        # 頁籤 4: 綜合統計明細
        # ==========================================
        with tab4:
            st.markdown(f"### 🏆 各高中生源統計與轉換率明細")
            
            pivot_df = df.groupby(['畢業學校', '最終狀態']).size().unstack(fill_value=0).reset_index()
            for col in ['1_僅通過一階', '2_進入二階(未入學)', '3_最終入學']:
                if col not in pivot_df.columns: pivot_df[col] = 0
            
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
        st.error(f"發生錯誤。請確認您的檔案欄位是否齊全（特別是「入學學年度」）。錯誤詳情：{e}")
else:
    st.info("👈 請在左側上傳資料以開始分析。")
