import streamlit as st
import pandas as pd
import altair as alt
import base64
import json

# --- 1. 頁面基本設定 ---
st.set_page_config(page_title="社交工程演練完整報告工具", layout="wide")
st.title("📊 社交工程演練統計工具 (全圖表預覽版)")

# --- 2. 側邊欄：參數設定 ---
st.sidebar.header("⚙️ 參數設定")
uploaded_file = st.sidebar.file_uploader("1. 上傳演練紀錄 (.xlsx)", type=["xlsx"])
config_file = st.sidebar.file_uploader("2. 上傳參數設定 (.txt)", type=["txt"])

total_accounts = 99
total_emails_sent = 99
full_subject_list = []
tags_map = {"開啟信件": [], "點閱連結": [], "開啟附件": [], "輸入帳密": []}

if config_file is not None:
    try:
        content = config_file.read().decode("utf-8")
        lines = [line.strip() for line in content.split('\n') if line.strip()]
        mode = None
        for line in lines:
            if "總帳號數" in line: total_accounts = int(line.split(':')[1])
            elif "總發送數" in line: total_emails_sent = int(line.split(':')[1])
            elif "行為標籤對應" in line: mode = "TAG"; continue
            elif "郵件主旨" in line: mode = "SUBJECT"; continue
            if mode == "TAG" and ":" in line:
                key, values = line.split(':')
                tags_map[key.strip()] = [v.strip() for v in values.split(',')]
            elif mode == "SUBJECT": full_subject_list.append(line)
        st.sidebar.success("✅ TXT 參數讀取成功！")
    except Exception as e:
        st.sidebar.error(f"TXT 解析失敗: {e}")

# --- 3. 工具函式 ---
def mask_pii(df, name_col, email_col):
    masked_df = df.copy()
    def mask_name(val):
        val = str(val)
        if len(val) <= 1: return val
        if len(val) == 2: return val[0] + "*"
        return val[0] + "*" + val[-1]
    def mask_email(val):
        val = str(val)
        if "@" not in val: return "****"
        prefix, domain = val.split("@")
        if len(prefix) <= 2: return prefix + "****@" + domain
        return prefix[:2] + "****@" + domain
    if name_col in masked_df.columns: masked_df[name_col] = masked_df[name_col].apply(mask_name)
    if email_col in masked_df.columns: masked_df[email_col] = masked_df[email_col].apply(mask_email)
    return masked_df

def draw_horizontal_label_chart(data, x_col, y_col, color="#4E79A7", is_export=False):
    plot_df = data.reset_index()
    chart_width = 600 if is_export else 800 
    bars = alt.Chart(plot_df).mark_bar(size=45).encode(
        x=alt.X(f"{x_col}:N", sort=None, axis=alt.Axis(labelAngle=0, labelFontSize=12, title=None)),
        y=alt.Y(f"{y_col}:Q", axis=alt.Axis(title=y_col)),
        tooltip=[x_col, y_col]
    )
    text = bars.mark_text(align='center', baseline='bottom', dy=-5, fontSize=12, fontWeight='bold').encode(text=f"{y_col}:Q")
    chart = (bars + text).properties(height=350, width=chart_width)
    
    # 關鍵：若非匯出模式，則直接在 Streamlit 渲染圖表
    if not is_export: 
        st.altair_chart(chart, use_container_width=True)
    return chart

def parse_device(ua):
    # 先統一轉小寫，避免大小寫不一致導致判斷失敗
    ua = str(ua).lower()
    
    # --- 1. 最優先判定：精準識別 Outlook / MS-Office 環境 ---
    # Mozilla/4.0 (compatible; ms-office; MSOffice 16...) 屬於此類
    if 'ms-office' in ua or 'microsoft outlook' in ua or 'msoffice' in ua:
        return "電腦 (Desktop)"
    
    # --- 2. 行動裝置判定 ---
    if 'ipad' in ua: 
        return "平板 (Tablet)"
    if 'android' in ua and 'mobile' not in ua: 
        return "平板 (Tablet)"
    if 'iphone' in ua or 'android' in ua or 'mobile' in ua: 
        return "手機 (Mobile)"
    
    # --- 3. 一般電腦 OS 判定 ---
    if 'windows' in ua or 'macintosh' in ua or 'linux' in ua: 
        return "電腦 (Desktop)"
        
    return "其他 (Unknown)"

def generate_professional_advice(df, total_accounts, sum2, sum4, final_s, sum7):
    # 計算關鍵數據
    click_rate = (sum2.loc[sum2['項目'] == '點閱連結', '人'].values[0] / total_accounts) * 100
    credential_rate = (sum2.loc[sum2['項目'] == '輸入帳密', '人'].values[0] / total_accounts) * 100
    top_dept = sum4.iloc[0] if not sum4.empty else None
    top_subject = final_s.iloc[0] if not final_s.empty else None
    mobile_rate = (sum7.loc[sum7['裝置類型'] == '手機 (Mobile)', '帳號數量'].values[0] / sum7['帳號數量'].sum() * 100) if '手機 (Mobile)' in sum7['裝置類型'].values else 0

    advice = []
    
    # 總體風險評估
    if click_rate > 10:
        advice.append(f"🔴 **高風險警示**：本次演練點閱率達 {click_rate:.1f}%，高於業界平均 (7-10%)。顯示同仁對於誘騙連結的警覺性仍有提升空間。")
    else:
        advice.append(f"🟢 **風險受控**：點閱率 {click_rate:.1f}% 表現良好，優於業界標準。")

    # 針對統計五：主旨攻擊面分析
    if top_subject is not None:
        advice.append(f"📝 **主旨分析**：最成功的誘餌為「{top_subject['郵件主旨']}」。這類「{ '公務相關' if '通知' in top_subject['郵件主旨'] else '行政福利' }」主題最易使同仁放下戒心，建議未來教育訓練應加強此類案例宣導。")

    # 針對統計四：高風險單位
    if top_dept is not None:
        advice.append(f"🏢 **重點強化單位**：{top_dept['單位']} 的遭誘騙人數比例最高。建議針對該部門進行小規模的「強化補測」或實體宣導。")

    # 針對統計七：載具安全性
    if mobile_rate > 30:
        advice.append(f"📱 **行動辦公風險**：行動裝置點閱占比達 {mobile_rate:.1f}%。由於手機螢幕較小，較難辨識完整郵件地址與連結 URL，建議評估導入行動端郵件過濾機制。")

    # 具體行動建議
    advice.append("""
    🛠️ **後續行動建議 (Next Steps)**：
    1. **針對性教育訓練**：對曾點閱連結之同仁發送「資安隨機測驗」或微學習教材。
    2. **強化輸入警示**：對本次「輸入帳密」之同仁進行權限檢查，並確認是否已啟用多因素驗證 (MFA)。
    3. **主旨情境優化**：下次演練可嘗試結合時事（如報稅、資通訊軟體更新）以測試更高層級的心理攻防。
    """)
    
    return "\n\n".join(advice)	
	
# --- 4. HTML 匯出函式 ---
def generate_html_report(report_items):
    html_content = f"""
    <html><head><meta charset="utf-8">
    <script src="https://cdn.jsdelivr.net/npm/vega@5"></script>
    <script src="https://cdn.jsdelivr.net/npm/vega-lite@5"></script>
    <script src="https://cdn.jsdelivr.net/npm/vega-embed@6"></script>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
        body {{ padding: 40px; background-color: #f8f9fa; font-family: sans-serif; }}
        .section {{ background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); margin-bottom: 40px; }}
        .chart-wrapper {{ margin-bottom: 20px; overflow-x: auto; display: flex; justify-content: center; }}
        details {{ background: #f1f3f5; padding: 10px; border-radius: 8px; cursor: pointer; margin-top: 10px; }}
        table {{ width: 100% !important; table-layout: auto !important; border-collapse: collapse; }}
        th, td {{ text-align: left !important; padding: 8px !important; word-break: break-all !important; vertical-align: middle; }}
    </style></head><body><div class="container"><h1 class="text-center mb-5">社交工程演練統計報告</h1>
    """
    for i, item in enumerate(report_items):
        chart_id = f"vis{i}"
        c_json = item["chart"].to_json() if item["chart"] else None
        html_content += f"""
        <div class="section">
            <h3 class="mb-4">{item['title']}</h3>
            {"<div class='chart-wrapper'><div id='" + chart_id + "'></div></div>" if c_json else ""}
            <details><summary>查看詳細名單數據</summary><div class="table-responsive mt-2">{item['df'].to_html(classes='table table-sm table-striped', index=False)}</div></details>
        </div>"""
        if c_json: html_content += f"<script>vegaEmbed('#{chart_id}', {c_json}, {{actions: false}});</script>"
    html_content += "</div></body></html>"
    b64 = base64.b64encode(html_content.encode()).decode()
    return f'<a href="data:text/html;base64,{b64}" download="演練統計結案報告.html" class="btn btn-primary w-100">📥 下載完整報告 (含各項詳細名單)</a>'

# --- 5. 主程式 ---
if uploaded_file is not None and config_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        email_col, event_col, subject_col, dept_col, ua_col = "目標郵箱", "事件類型", "郵件主旨", "單位", "用戶代理"
        name_col = "目標姓名" if "目標姓名" in df.columns else "目標郵箱"

        def map_to_std(x):
            for std_name, raw_names in tags_map.items():
                if x in raw_names: return std_name
            return "其他"
        df['std_tag'] = df[event_col].apply(map_to_std)
        report_items = []

        # --- 統計一：遭誘騙受測名單 ---
        st.subheader("🎯 統計一：遭誘騙受測名單")
        u_users = df[[name_col, email_col, dept_col]].drop_duplicates().reset_index(drop=True)
        st.metric("實測遭誘騙總人數", f"{len(u_users)} 人")
        with st.expander("🔍 查看詳細名單"): st.dataframe(u_users, use_container_width=True)
        report_items.append({"title": "統計一：遭誘騙受測名單", "df": mask_pii(u_users, name_col, email_col), "chart": None})

        # --- 統計二：個人行為統計 ---
        st.divider(); st.subheader("📈 統計二：個人行為統計")
        df_u2 = df[[email_col, 'std_tag']].drop_duplicates()
        active_u = set(df_u2[df_u2['std_tag'].isin(["點閱連結", "開啟附件", "輸入帳密"])][email_col])
        openers = set(df_u2[df_u2['std_tag'] == "開啟信件"][email_col])
        sum2 = pd.DataFrame({
            "項目": ["開啟信件", "點閱連結", "開啟附件", "輸入帳密"],
            "人": [len(openers | active_u), df_u2[df_u2['std_tag'] == "點閱連結"][email_col].nunique(), df_u2[df_u2['std_tag'] == "開啟附件"][email_col].nunique(), df_u2[df_u2['std_tag'] == "輸入帳密"][email_col].nunique()]
        })
        sum2["比率"] = sum2["人"].apply(lambda x: f"{(x/total_accounts)*100:.2f}%")
        st.table(sum2.set_index("項目"))
        draw_horizontal_label_chart(sum2, "項目", "人") # 網頁顯示
        c2_exp = draw_horizontal_label_chart(sum2, "項目", "人", is_export=True) # 報告用
        report_items.append({"title": "統計二：個人行為分布圖與數據", "df": sum2, "chart": c2_exp})

        # --- 統計三：郵件主旨行為統計 ---
        st.divider(); st.subheader("✉️ 統計三：郵件主旨行為統計")
        df_u3 = df[[email_col, 'std_tag', subject_col]].drop_duplicates()
        active_u3 = df_u3[df_u3['std_tag'].isin(["點閱連結", "開啟附件", "輸入帳密"])][[email_col, subject_col]].drop_duplicates()
        opens_u3 = df_u3[df_u3['std_tag'] == "開啟信件"][[email_col, subject_col]].drop_duplicates()
        sum3 = pd.DataFrame({
            "項目": ["開啟總次數", "點閱連結總數", "點閱附件總數", "輸入帳密總數"],
            "次數": [len(pd.concat([opens_u3, active_u3]).drop_duplicates()), len(df_u3[df_u3['std_tag'] == "點閱連結"]), len(df_u3[df_u3['std_tag'] == "開啟附件"]), len(df_u3[df_u3['std_tag'] == "輸入帳密"])]
        })
        sum3["比率"] = sum3["次數"].apply(lambda x: f"{(x/total_emails_sent)*100:.2f}%")
        st.table(sum3.set_index("項目"))
        draw_horizontal_label_chart(sum3, "項目", "次數", color="#ED7D31") # 網頁顯示
        c3_exp = draw_horizontal_label_chart(sum3, "項目", "次數", color="#ED7D31", is_export=True) # 報告用
        report_items.append({"title": "統計三：郵件行為總次數統計", "df": sum3, "chart": c3_exp})

        # --- 統計四：各單位受測人數分布 ---
        st.divider(); st.subheader("🏢 統計四：各單位受測人數分布")
        sum4_df = df[df['std_tag'] != "其他"][[dept_col, email_col]].drop_duplicates()
        sum4_result = sum4_df.groupby(dept_col).size().reset_index(name='人數').sort_values(by='人數', ascending=False)
        st.table(sum4_result.set_index(dept_col))
        draw_horizontal_label_chart(sum4_result, dept_col, "人數", color="#70AD47") # 網頁顯示
        c4_exp = draw_horizontal_label_chart(sum4_result, dept_col, "人數", color="#70AD47", is_export=True)
        report_items.append({"title": "統計四：各單位分布名單", "df": sum4_result, "chart": c4_exp})

        # --- 統計五：郵件主旨影響力分析 ---
        st.divider(); st.subheader("📑 統計五：郵件主旨影響力分析")
        actual_s = df[[subject_col, email_col]].drop_duplicates().groupby(subject_col)[email_col].count().reset_index(name='觸及人數')
        all_s_df = pd.DataFrame(list(set(df[subject_col].unique().tolist() + full_subject_list)), columns=[subject_col])
        final_s = pd.merge(all_s_df, actual_s, on=subject_col, how='left').fillna(0)
        final_s['觸及人數'] = final_s['觸及人數'].astype(int)
        final_s = final_s.sort_values(by='觸及人數', ascending=False)
        st.table(final_s.set_index(subject_col))
        draw_horizontal_label_chart(final_s, subject_col, "觸及人數", color="#A5A5A5") # 網頁顯示
        c5_exp = draw_horizontal_label_chart(final_s, subject_col, "觸及人數", color="#A5A5A5", is_export=True)
        report_items.append({"title": "統計五：主旨影響力詳細名單", "df": final_s, "chart": c5_exp})

        # --- 統計六：個人重複行為分析 (僅修正表格顯示) ---
        st.divider(); st.subheader("📍 統計六：個人重複行為分析")
        for tag in ["開啟信件", "點閱連結", "開啟附件", "輸入帳密"]:
            if tag == "開啟信件":
                a = df[df['std_tag'].isin(["點閱連結", "開啟附件", "輸入帳密"])][[email_col, name_col, dept_col, subject_col]].drop_duplicates()
                o = df[df['std_tag'] == "開啟信件"][[email_col, name_col, dept_col, subject_col]].drop_duplicates()
                det = pd.concat([a, o]).drop_duplicates().groupby([name_col, email_col, dept_col]).size().reset_index(name='次數')
            else:
                det = df[df['std_tag'] == tag].groupby([name_col, email_col, dept_col])[subject_col].nunique().reset_index(name='次數')
            
            f_dist = det['次數'].value_counts().reindex([1,2,3,4,5], fill_value=0).reset_index()
            f_dist.columns = ['次數', '帳號數量']
            f_dist['標籤'] = f_dist['次數'].apply(lambda x: f"{tag[:2]}{x}封信")
            
            st.markdown(f"#### 🏷️ 【{tag}】分佈")
            
            # --- 補回統計表格 ---
            st.table(f_dist[['標籤', '帳號數量']].set_index('標籤'))
            
            # 保持原有的圖表預覽
            draw_horizontal_label_chart(f_dist, "標籤", "帳號數量", color="#4472C4")
            c6_exp = draw_horizontal_label_chart(f_dist, "標籤", "帳號數量", color="#4472C4", is_export=True)
            
            # 保持原有的詳細清單展開
            with st.expander(f"🔍 查看【{tag}】詳細名單"): 
                st.dataframe(det.sort_values(by='次數', ascending=False), use_container_width=True)
            
            report_items.append({"title": f"統計六：【{tag}】行為名單明細", "df": mask_pii(det, name_col, email_col), "chart": c6_exp})

        # --- 統計七：受測裝置載具分析 ---
        st.divider(); st.subheader("📱 統計七：受測裝置載具分析")
        if ua_col in df.columns:
            device_df = df.sort_values(by=email_col).drop_duplicates(subset=[email_col], keep='last').copy()
            device_df['裝置類型'] = device_df[ua_col].apply(parse_device)
            sum7 = device_df['裝置類型'].value_counts().reset_index()
            sum7.columns = ['裝置類型', '帳號數量']
            st.table(sum7.set_index('裝置類型'))
            draw_horizontal_label_chart(sum7, "裝置類型", "帳號數量", color="#7294D4") # 網頁顯示
            c7_exp = draw_horizontal_label_chart(sum7, "裝置類型", "帳號數量", color="#7294D4", is_export=True)
            
            list_cols = [name_col, email_col, '裝置類型', ua_col]
            device_list = device_df[list_cols].copy().sort_values(by='裝置類型')
            with st.expander("🔍 查看載具詳細名單"): st.dataframe(device_list, use_container_width=True)
            report_items.append({"title": "統計七：載具分析名單 (含原始 UA)", "df": mask_pii(device_list, name_col, email_col), "chart": c7_exp})
        else:
            st.warning(f"Excel 中找不到『{ua_col}』欄位。")

        if st.sidebar.button("🚀 生成最終修正報告"):
            st.sidebar.markdown(generate_html_report(report_items), unsafe_allow_html=True)

    except Exception as e: st.error(f"分析失敗: {e}")
else: st.info("💡 請上傳檔案以開始分析。")