import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

# --- 1. 環境設定 ---
st.set_page_config(page_title="社交工程演練分析工具", layout="wide")
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False
sns.set_theme(style="whitegrid", font='Microsoft JhengHei')

st.title("🛡️ 社交工程演練數據分析系統")

# --- 2. 側邊欄參數設定 ---
st.sidebar.header("參數設定")
uploaded_file = st.sidebar.file_uploader("選擇演練 Excel 檔案", type=["xlsx", "xls"])
people_denom = st.sidebar.number_input("受測人員總數 (分母)", min_value=1, value=100)
mail_denom = st.sidebar.number_input("總寄出郵件總封數 (分母)", min_value=1, value=100)

if uploaded_file:
    try:
        # 讀取數據
        df = pd.read_excel(uploaded_file)
        
        # 欄位映射
        df = df.rename(columns={
            '目標部門': 'Dept_Core', 
            '目標郵箱': 'Email',
            '郵件主旨': 'Template',
            '事件類型': 'Response'
        })

        # 事件類型標準化
        def normalize_response(r):
            low_r = str(r).lower().strip()
            if 'submit' in low_r: return '輸入帳密'
            if 'click' in low_r: return '點閱連結'
            if 'open' in low_r: return '開啟信件'
            return '其他行為'

        df['行為紀錄'] = df['Response'].apply(normalize_response)
        
        # 數據補強邏輯 (Click/Submit 自動算入「開啟信件」)
        opened_logic = df[['Email', 'Dept_Core', 'Template']].drop_duplicates()
        opened_logic['行為紀錄'] = '開啟信件'
        df_full = pd.concat([df[['Email', 'Dept_Core', 'Template', '行為紀錄']], opened_logic], ignore_index=True).drop_duplicates()

        # 預計算摘要數據
        summary_list = []
        for act in ['開啟信件', '點閱連結', '輸入帳密']:
            f_count = len(df_full[df_full['行為紀錄'] == act])
            p_count = df_full[df_full['行為紀錄'] == act]['Email'].nunique()
            summary_list.append({
                '項目': act,
                '總封數': f_count,
                '封數比率': f"{(f_count/mail_denom*100):.2f}%",
                '影響人數': p_count,
                '人數比率': f"{(p_count/people_denom*100):.2f}%"
            })
        sum_df = pd.DataFrame(summary_list)

        # --- 3. UI 標籤頁 ---
        tabs = st.tabs(["📊 數據總覽", "📈 及時分析圖表", "📋 原始資料明細"])

        with tabs[0]:
            col1, col2, col3 = st.columns(3)
            col1.metric("開啟人數", f"{sum_df.loc[0, '影響人數']} 人", sum_df.loc[0, '人數比率'], delta_color="inverse")
            col2.metric("點擊連結", f"{sum_df.loc[1, '影響人數']} 人", sum_df.loc[1, '人數比率'], delta_color="inverse")
            col3.metric("輸入帳密", f"{sum_df.loc[2, '影響人數']} 人", sum_df.loc[2, '人數比率'], delta_color="inverse")
            
            st.subheader("部門參與人數統計")
            dept_counts = df_full[df_full['行為紀錄'] == '開啟信件'].drop_duplicates(subset=['Email', 'Dept_Core'])
            dept_final = dept_counts.groupby('Dept_Core').size().reset_index(name='人數').sort_values(by='人數', ascending=False)
            st.dataframe(dept_final, use_container_width=True)

        with tabs[1]:
            st.subheader("及時分析圖表 (含數字標註)")
            
            c1, c2 = st.columns(2)
            
            # --- 圖二：影響人數比率 (加上數字標註) ---
            with c1:
                fig2, ax2 = plt.subplots(figsize=(8, 6))
                sns.barplot(data=sum_df, x='項目', y='影響人數', palette='viridis', ax=ax2)
                # 數字標註邏輯
                for p, pct in zip(ax2.patches, sum_df['人數比率']):
                    ax2.annotate(f'{int(p.get_height())}人\n({pct})', 
                                 (p.get_x() + p.get_width() / 2., p.get_height()), 
                                 ha='center', va='bottom', fontsize=11, fontweight='bold', color='black')
                plt.title("各項行為影響人數與比率", fontsize=14)
                plt.ylim(0, max(sum_df['影響人數']) * 1.2) # 留空間給文字
                st.pyplot(fig2)

            # --- 圖三：郵件主旨統計 (加上數字標註) ---
            with c2:
                fig3, ax3 = plt.subplots(figsize=(8, 6))
                temp_counts = df.drop_duplicates(['Email', 'Template'])['Template'].value_counts().reset_index()
                temp_counts.columns = ['主旨', '人數']
                sns.barplot(data=temp_counts, x='主旨', y='人數', palette='flare', ax=ax3)
                # 數字標註邏輯
                for p in ax3.patches:
                    val = int(p.get_height())
                    ax3.annotate(f'{val}人\n({(val/people_denom*100):.1f}%)', 
                                 (p.get_x() + p.get_width() / 2., p.get_height()), 
                                 ha='center', va='bottom', fontsize=10, fontweight='bold')
                plt.xticks(rotation=20, ha='right')
                plt.title("各郵件主旨受測人數佔比", fontsize=14)
                plt.ylim(0, max(temp_counts['人數']) * 1.2)
                st.pyplot(fig3)

            # --- 圖五~七：次數分佈 (加上數字標註) ---
            st.divider()
            st.markdown("### 🎯 行為次數分佈統計")
            cols = st.columns(3)
            for i, name in enumerate(['開啟信件', '點閱連結', '輸入帳密']):
                with cols[i]:
                    sub = df_full[df_full['行為紀錄'] == name]
                    dist = sub.groupby('Email').size().value_counts().reindex(range(1, 6), fill_value=0).reset_index()
                    dist.columns = ['次數', '人數']
                    
                    fig, ax = plt.subplots(figsize=(6, 5))
                    sns.barplot(data=dist, x='次數', y='人數', color='#5B9BD5', ax=ax)
                    ax.set_xticklabels([f'{int(x)}次' for x in dist['次數']])
                    # 數字標註
                    for p in ax.patches:
                        ax.annotate(f'{int(p.get_height())}', 
                                     (p.get_x() + p.get_width() / 2., p.get_height()), 
                                     ha='center', va='bottom', fontweight='bold')
                    plt.title(f"{name} - 次數分佈", fontsize=12)
                    st.pyplot(fig)

        with tabs[2]:
            st.subheader("處理後完整明細")
            st.dataframe(df_full, use_container_width=True)

        # --- 4. 側邊欄下載按鈕 ---
        st.sidebar.divider()
        st.sidebar.subheader("檔案導出")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            sum_df.to_excel(writer, sheet_name='數據摘要', index=False)
            dept_final.to_excel(writer, sheet_name='部門統計', index=False)
            df_full.to_excel(writer, sheet_name='分析明細', index=False)
        
        st.sidebar.download_button(
            label="📥 下載 Excel 分析報告",
            data=output.getvalue(),
            file_name="社交工程分析結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"分析程式出錯: {e}")
else:
    st.info("💡 請從左側選單上傳 Excel 演練原始檔案以開始分析。")