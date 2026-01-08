import streamlit as st
import pandas as pd
import altair as alt
import base64
import json
import markdown
import google.generativeai as genai  # 新增 Gemini SDK

# =================================================================
# 1. 頁面基本設定
# =================================================================
st.set_page_config(page_title="社交工程演練完整報告工具", layout="wide")

# =================================================================
# 2. 側邊欄：參數設定
# =================================================================
st.sidebar.header("⚙️ 參數設定")
uploaded_file = st.sidebar.file_uploader("1. 上傳演練紀錄 (.xlsx)", type=["xlsx"])
config_file = st.sidebar.file_uploader("2. 上傳參數設定 (.txt)", type=["txt"])

company_name = ""
total_accounts = 99
total_emails_sent = 99
full_subject_list = []
tags_map = {"開啟信件": [], "點閱連結": [], "開啟附件": [], "輸入帳密": []}

# --- 解析 TXT 參數 ---
if config_file is not None:
    try:
        content = config_file.read().decode("utf-8")
        lines = [line.strip() for line in content.split('\n') if line.strip()]
        mode = None
        for line in lines:
            processed_line = line.replace('：', ':')
            if ":" in processed_line:
                parts = processed_line.split(':')
                key = parts[0].strip()
                val = parts[1].strip() if len(parts) > 1 else ""
                if "單位名稱" in key or "公司名稱" in key:
                    company_name = val
                elif "總帳號數" in key:
                    total_accounts = int(val) if val else 99
                elif "總發送數" in key:
                    total_emails_sent = int(val) if val else 99
                elif "行為標籤對應" in key:
                    mode = "TAG"
                    continue
                elif "郵件主旨" in key:
                    mode = "SUBJECT"
                    continue
                if mode == "TAG":
                    tags_map[key] = [v.strip() for v in val.split(',')]
            elif mode == "SUBJECT":
                full_subject_list.append(line)
        st.sidebar.success(f"✅ 參數讀取成功：{company_name}")
    except Exception as e:
        st.sidebar.error(f"TXT 解析失敗: {e}")

# 動態顯示網頁大標題
if company_name:
    st.markdown(f"""
        <h1 style='text-align: left; margin-bottom: 0;'>📊 {company_name}</h1>
        <h2 style='text-align: left; margin-top: 0;'>社交工程演練統計報告</h2>
    """, unsafe_allow_html=True)
else:
    st.title("📊 社交工程演練統計報告")

# =================================================================
# 側邊欄：新增 API Key 設定
# =================================================================
st.sidebar.divider()
st.sidebar.header("🤖 AI 顧問設定")
gemini_api_key = st.sidebar.text_input("輸入 Gemini API Key", type="password", help="請至 Google AI Studio 申請")
enable_ai = st.sidebar.checkbox("開啟 AI 即時分析報告")
    
# =================================================================
# 3. 工具函式
# =================================================================
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
    if not is_export: 
        st.altair_chart(chart, use_container_width=True)
    return chart

def parse_device(ua):
    ua = str(ua).lower()
    if 'ms-office' in ua or 'microsoft outlook' in ua or 'msoffice' in ua:
        return "電腦 (Desktop)"
    if 'ipad' in ua: 
        return "平板 (Tablet)"
    if 'android' in ua and 'mobile' not in ua: 
        return "平板 (Tablet)"
    if 'iphone' in ua or 'android' in ua or 'mobile' in ua: 
        return "手機 (Mobile)"
    if 'windows' in ua or 'macintosh' in ua or 'linux' in ua: 
        return "電腦 (Desktop)"
    return "其他 (Unknown)"

def generate_professional_advice(df, total_accounts, sum2, sum4, final_s, sum7):
    click_rate = (sum2.loc[sum2['項目'] == '點閱連結', '人'].values[0] / total_accounts) * 100
    credential_rate = (sum2.loc[sum2['項目'] == '輸入帳密', '人'].values[0] / total_accounts) * 100
    top_dept = sum4.iloc[0] if not sum4.empty else None
    top_subject = final_s.iloc[0] if not final_s.empty else None
    mobile_rate = (sum7.loc[sum7['裝置類型'] == '手機 (Mobile)', '帳號數量'].values[0] / sum7['帳號數量'].sum() * 100) if '手機 (Mobile)' in sum7['裝置類型'].values else 0
    advice = []
# 總體風險評估
    if click_rate > 10:
        advice.append(f"🔴 高風險警示：本次演練點閱率達 {click_rate:.1f}%，高於業界平均 (7-10%)。顯示同仁對於誘騙連結的警覺性仍有提升空間。")
    else:
        advice.append(f"🟢 風險受控：點閱率 {click_rate:.1f}% 表現良好，優於業界標準。")
    # --- 新增：輸入帳密率警告邏輯 ---
    if credential_rate > 0:
        advice.append(f"⚠️ 憑證外洩警告：本次有 {credential_rate:.1f}% 的受測者輸入帳號密碼。這屬於極高風險行為，代表若為真實攻擊，同仁的存取權限已遭竊取，建議立即進行權限稽核與 MFA 宣導。")
    else:
        advice.append(f"✅ 安全意識達標：本次無人輸入帳號密碼，顯示同仁在關鍵步驟（輸入憑證）具有高度警覺。")
    # 針對統計五：主旨攻擊面分析
    if top_subject is not None:
        advice.append(f"📝 主旨分析：最成功的誘餌為「{top_subject['郵件主旨']}」。這類「{ '公務相關' if '通知' in top_subject['郵件主旨'] else '行政福利' }」主題最易使同仁放下戒心，建議未來教育訓練應加強此類案例宣導。")

    # 針對統計四：高風險單位
    if top_dept is not None:
        advice.append(f"🏢 重點強化單位：{top_dept['單位']} 的遭誘騙人數比例最高。建議針對該部門進行小規模的「強化補測」或實體宣導。")

    # 針對統計七：載具安全性
    if mobile_rate > 20:
        advice.append(f"📱 行動辦公風險：行動裝置點閱占比達 {mobile_rate:.1f}%。由於手機螢幕較小，較難辨識完整郵件地址與連結 URL，建議評估導入行動端郵件過濾機制。")
    advice.append("""
🛠️ 後續行動建議 (Next Steps)：
1. 針對性教育訓練：對曾點閱連結之同仁發送「資安隨機測驗」或微學習教材。
2. 強化輸入警示：對本次「輸入帳密」之同仁進行權限檢查。
3. 主旨情境優化：下次演練可嘗試結合時事。
""")
    return "\n\n".join(advice)	
# =================================================================
# 新增：Gemini AI 分析函式
# =================================================================
def ask_gemini_advisor(api_key, context_data):
    try:
        genai.configure(api_key=api_key)
        
        # 修正 404 錯誤：使用最新穩定的模型別名
        # 注意：gemini-1.5-flash-latest 或 gemini-1.5-flash 是目前最穩定的選擇
        model = genai.GenerativeModel('gemini-2.5-flash') 
        
        prompt = f"""
        你是一位資深資安顧問，請分析以下社交工程演練數據：
        - 單位：{context_data['company']}
        - 點閱連結：{context_data['click_count']} 人 (總數 {context_data['total_accounts']})
        - 輸入帳密：{context_data['login_count']} 人
        - 成功誘餌：{context_data['top_subject']}
        - 高風險單位：{context_data['top_dept']}
        
        請提供 3 點具體改善建議與一段員工資安宣導語。
        """
        
        # 呼叫 API
        response = model.generate_content(prompt)
        
        if response and response.text:
            return response.text
        return "AI 回傳內容為空，請稍後重試。"

    except Exception as e:
        # 修正 429 錯誤：當請求過快時給予友善提示
        if "429" in str(e):
            return "⚠️ 請求太頻繁了！免費版 API 有頻率限制，請等 60 秒後再點擊一次。"
        # 修正 404 錯誤：若模型名稱不對，提示檢查
        if "404" in str(e):
            return "❌ 找不到 AI 模型 (404)。請確保已更新 google-generativeai 套件。"
        return f"❌ AI 分析出錯：{str(e)}"

	
# =================================================================
# 4. HTML 匯出函式
# =================================================================
def generate_html_report(report_items, title_name=""):
    report_title = f"{title_name}<br>社交工程演練統計報告" if title_name else "社交工程演練統計報告"
    html_content = f"""
    <html><head><meta charset="utf-8">
    <script src="https://cdn.jsdelivr.net/npm/vega@5"></script>
    <script src="https://cdn.jsdelivr.net/npm/vega-lite@5"></script>
    <script src="https://cdn.jsdelivr.net/npm/vega-embed@6"></script>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
        @media print {{ 
            body {{ background-color: white !important; padding: 0 !important; }} 
            .btn {{ display: none !important; }} 
        }}
        body {{ padding: 40px; background-color: #f8f9fa; font-family: "Microsoft JhengHei", sans-serif; }}
        .section {{ background: white; padding: 25px; border-radius: 12px; margin-bottom: 40px; border: 1px solid #ddd; page-break-inside: avoid; }}
        .text-box {{ 
            border-left: 5px solid #0d6efd; /* 改成藍色，區分 AI 與一般建議 */
            padding: 20px; 
            /* white-space: pre-wrap; */ 
            background: #f0f7ff; /* 淺藍色背景 */
            line-height: 1.6;
            font-size: 1.05em;
            border-radius: 0 8px 8px 0;
        }}
        .metric-box {{ background: #e9ecef; padding: 10px 20px; border-radius: 8px; margin-bottom: 20px; display: inline-block; }}
        .metric-label {{ font-size: 0.9em; color: #666; margin-right: 15px; }}
        .metric-number {{ font-size: 1.5em; font-weight: bold; color: #0d6efd; }}
        h1 {{ color: #333; }}
    </style></head><body><div class="container">
    <h1 class="text-center mb-5">{report_title}</h1>
    """
    for i, item in enumerate(report_items):
        chart_id = f"vis{i}"
        c_json = item["chart"].to_json() if item.get("chart") else None
        html_content += f'<div class="section"><h3>{item["title"]}</h3>'
        if item.get("metric_value"):
            html_content += f'<div class="metric-box"><span class="metric-label">數據統計</span><span class="metric-number">{item["metric_value"]}</span></div>'
        # if item.get("text"):
            # display_text = str(item["text"]).replace("\\n", "\n")
            # html_content += f'<div class="text-box">{item["text"]}</div>'
        if item.get("text"):
            # 【重點修改區】：將 Markdown 轉為 HTML
            raw_text = str(item["text"]).replace("\\n", "\n")
            formatted_text = markdown.markdown(raw_text, extensions=['tables']) 
            html_content += f'<div class="text-box">{formatted_text}</div>'
        if c_json:
            html_content += f"<div id='{chart_id}' class='mb-4'></div>"
            html_content += f"<script>vegaEmbed('#{chart_id}', {c_json}, {{actions: false}});</script>"
        if item.get("df") is not None:
            html_content += f'<div class="mt-3 table-responsive">{item["df"].to_html(classes="table table-sm table-bordered", index=False)}</div>'
        html_content += "</div>"
    html_content += "</div></body></html>"
    b64 = base64.b64encode(html_content.encode('utf-8-sig')).decode()
    return f'<a href="data:text/html;base64,{b64}" download="演練報告_{title_name}.html" class="btn btn-success w-100 p-3">📥 下載完整報告</a>'

# =================================================================
# 5. 主程式分析區塊
# =================================================================
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
        count_val = len(u_users)
        st.metric("實測遭誘騙總人數", f"{len(u_users)} 人")
        analysis_1 =f"""
分析說明：

1. 本次演練中，實際產生風險行為（開啟、點閱或輸入資訊）的總人數為 {count_val}人。  

2. 以總受測人數 {total_accounts} 人計算，實測遭誘騙率約為 {(count_val/total_accounts)*100:.1f}%。此數據反映了第一線員工在面對疑似釣魚郵件時的防範意識，建議針對名單內人員進行後續輔導。
"""
        st.write(analysis_1)
        with st.expander("🔍 查看詳細名單"): st.dataframe(u_users, use_container_width=True)
        report_items.append({"title": "統計一：遭誘騙受測名單分析", "df": mask_pii(u_users, name_col, email_col), "text": analysis_1, "metric_value": f"{count_val} 人", "chart": None})

        # --- 統計二：個人行為統計 ---
        st.divider(); st.subheader("📈 統計二：個人行為統計")
        df_u2 = df[[email_col, 'std_tag']].drop_duplicates()
        active_u = set(df_u2[df_u2['std_tag'].isin(["點閱連結", "開啟附件", "輸入帳密"])][email_col])
        openers = set(df_u2[df_u2['std_tag'] == "開啟信件"][email_col])
        count_open = len(openers | active_u)
        count_click = df_u2[df_u2['std_tag'] == "點閱連結"][email_col].nunique()
        count_attach = df_u2[df_u2['std_tag'] == "開啟附件"][email_col].nunique()
        count_login = df_u2[df_u2['std_tag'] == "輸入帳密"][email_col].nunique()
        sum2 = pd.DataFrame({"項目": ["開啟信件", "點閱連結", "開啟附件", "輸入帳密"], "人": [count_open, count_click, count_attach, count_login]})
        sum2["比率"] = sum2["人"].apply(lambda x: f"{(x/total_accounts)*100:.2f}%")
        
        click_rate = (count_click / total_accounts) * 100
        analysis_2 = f"""
分析說明：

1. 整體風險評估：本次演練之「點閱連結率」為 {click_rate:.2f}%。一般企業警戒線通常設為 10%，若高於此數值，建議加強宣導辨識偽造 URL 之技巧。  

2. 關鍵弱點發現：在點閱連結的人員中，有 {count_login} 位人員進一步執行了「輸入帳密」的行為。這顯示同仁對於『偽造登入頁面』的識別能力較為薄弱，建議列為優先資安輔導對象。
"""
        st.write(analysis_2)
        draw_horizontal_label_chart(sum2, "項目", "人")
        st.table(sum2.set_index("項目"))
        report_items.append({"title": "統計二：個人行為分布圖與數據分析", "df": sum2, "chart": draw_horizontal_label_chart(sum2, "項目", "人", is_export=True), "text": analysis_2})

        # --- 統計三：郵件主旨行為統計 ---
        st.divider(); st.subheader("✉️ 統計三：郵件主旨行為統計")
        df_u3 = df[[email_col, 'std_tag', subject_col]].drop_duplicates()
        active_u3 = df_u3[df_u3['std_tag'].isin(["點閱連結", "開啟附件", "輸入帳密"])][[email_col, subject_col]].drop_duplicates()
        opens_u3 = df_u3[df_u3['std_tag'] == "開啟信件"][[email_col, subject_col]].drop_duplicates()
        count_total_open = len(pd.concat([opens_u3, active_u3]).drop_duplicates())
        count_total_click = len(df_u3[df_u3['std_tag'] == "點閱連結"])
        count_total_login = len(df_u3[df_u3['std_tag'] == "輸入帳密"])
        sum3 = pd.DataFrame({"項目": ["開啟總次數", "點閱連結總數", "點閱附件總數", "輸入帳密總數"], "次數": [count_total_open, count_total_click, len(df_u3[df_u3['std_tag'] == "開啟附件"]), count_total_login]})
        sum3["比率"] = sum3["次數"].apply(lambda x: f"{(x/total_emails_sent)*100:.2f}%")
        
        total_click_rate = (count_total_click / total_emails_sent) * 100
        analysis_3 = f"""
分析說明：

1. 郵件觸及分析：本次演練共發送 {total_emails_sent} 封郵件，開啟次數為 {count_total_open} 次。這反映了同仁對於演練郵件主旨具有初步的點擊好奇度。

2. 誘騙成功率：總體點閱率為 {total_click_rate:.2f}%。在已開啟郵件的行為中，點閱比例的高低直接反映了誘餌設計與釣魚連結對同仁的心理引導強度。

3. 威脅程度評估：本次「輸入帳密」總數為 {count_total_login} 次。由於此行為直接涉及機敏憑證外洩，建議針對高誘惑性主旨進行案例分享。
"""
        st.markdown(analysis_3)
        draw_horizontal_label_chart(sum3, "項目", "次數", color="#ED7D31")
        st.table(sum3.set_index("項目"))
        report_items.append({"title": "統計三：郵件主旨行為統計分析", "df": sum3, "chart": draw_horizontal_label_chart(sum3, "項目", "次數", color="#ED7D31", is_export=True), "text": analysis_3})

        # --- 統計四：各單位受測人數分布 ---
        st.divider(); st.subheader("🏢 統計四：各單位受測人數分布")
        sum4_df = df[df['std_tag'] != "其他"][[dept_col, email_col]].drop_duplicates()
        sum4_result = sum4_df.groupby(dept_col).size().reset_index(name='人數').sort_values(by='人數', ascending=False)
        draw_horizontal_label_chart(sum4_result, dept_col, "人數", color="#70AD47")
        st.table(sum4_result.set_index(dept_col))
        report_items.append({"title": "統計四：各單位分布名單", "df": sum4_result, "chart": draw_horizontal_label_chart(sum4_result, dept_col, "人數", color="#70AD47", is_export=True)})

        # --- 統計五：郵件主旨影響力分析 ---
        st.divider(); st.subheader("📑 統計五：郵件主旨影響力分析")
        actual_s = df[[subject_col, email_col]].drop_duplicates().groupby(subject_col)[email_col].count().reset_index(name='觸及人數')
        all_s_df = pd.DataFrame(list(set(df[subject_col].unique().tolist() + full_subject_list)), columns=[subject_col])
        final_s = pd.merge(all_s_df, actual_s, on=subject_col, how='left').fillna(0)
        final_s['觸及人數'] = final_s['觸及人數'].astype(int)
        final_s = final_s.sort_values(by='觸及人數', ascending=False)
        draw_horizontal_label_chart(final_s, subject_col, "觸及人數", color="#A5A5A5")
        st.table(final_s.set_index(subject_col))
        report_items.append({"title": "統計五：主旨影響力詳細名單", "df": final_s, "chart": draw_horizontal_label_chart(final_s, subject_col, "觸及人數", color="#A5A5A5", is_export=True)})

        # --- 統計六：個人重複行為分析 ---
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
            draw_horizontal_label_chart(f_dist, "標籤", "帳號數量", color="#4472C4")
            st.table(f_dist[['標籤', '帳號數量']].set_index('標籤'))
            with st.expander(f"🔍 查看【{tag}】詳細名單"): 
                st.dataframe(det.sort_values(by='次數', ascending=False), use_container_width=True)
            report_items.append({"title": f"統計六：【{tag}】行為名單明細", "df": mask_pii(det, name_col, email_col), "chart": draw_horizontal_label_chart(f_dist, "標籤", "帳號數量", color="#4472C4", is_export=True)})

        # --- 統計七：受測裝置載具分析 ---
        st.divider(); st.subheader("📱 統計七：受測裝置載具分析")
        if ua_col in df.columns:
            device_df = df.sort_values(by=email_col).drop_duplicates(subset=[email_col], keep='last').copy()
            device_df['裝置類型'] = device_df[ua_col].apply(parse_device)
            sum7 = device_df['裝置類型'].value_counts().reset_index()
            sum7.columns = ['裝置類型', '帳號數量']
            draw_horizontal_label_chart(sum7, "裝置類型", "帳號數量", color="#7294D4")
            st.table(sum7.set_index('裝置類型'))
            device_list = device_df[[name_col, email_col, '裝置類型', ua_col]].copy().sort_values(by='裝置類型')
            with st.expander("🔍 查看載具詳細名單"): st.dataframe(device_list, use_container_width=True)
            report_items.append({"title": "統計七：載具分析名單", "df": mask_pii(device_list, name_col, email_col), "chart": draw_horizontal_label_chart(sum7, "裝置類型", "帳號數量", color="#7294D4", is_export=True)})
        else:
            st.warning(f"Excel 中找不到『{ua_col}』欄位。")

        # if st.sidebar.button("🚀 生成報告"):
            # st.sidebar.markdown(generate_html_report(report_items, company_name), unsafe_allow_html=True)
    except Exception as e: st.error(f"分析失敗: {e}")
        # =================================================================
        # 主程式：在專家建議區塊整合 AI
        # =================================================================
        # (假設這是在分析完所有數據後)
    if uploaded_file and config_file:

            st.divider()
            st.subheader("🧠 專家分析建議")
            
            # 1. 獲取專家建議內容
            advice_text = generate_professional_advice(df, total_accounts, sum2, sum4_result, final_s, sum7)
            
            # 準備數據上下文
            mobile_val = (sum7.loc[sum7['裝置類型'] == '手機 (Mobile)', '帳號數量'].values[0] / sum7['帳號數量'].sum() * 100) if '手機 (Mobile)' in sum7['裝置類型'].values else 0
            
            context_data = {
                "company": company_name or "受測單位",
                "total_accounts": total_accounts,
                "click_count": count_click,
                "login_count": count_login,
                "top_subject": final_s.iloc[0][subject_col] if not final_s.empty else "未知",
                "top_dept": sum4_result.iloc[0][dept_col] if not sum4_result.empty else "未知",
                "mobile_rate": mobile_val
            }

            # 2. 判斷顯示方式 (AI 或 腳本分析)
            if enable_ai and gemini_api_key:
                with st.spinner("Gemini 顧問正在深入分析中..."):
                    ai_report = ask_gemini_advisor(gemini_api_key, context_data)
                    st.markdown("### 🤖 AI 即時分析回饋")
                    st.info(ai_report)
                    
                    # 存入匯出清單 (確保 type 被標記為 text 供後續 HTML 渲染)
                    report_items.append({
                        "title": "🤖 Gemini AI 深度分析建議", 
                        "text": ai_report,
                        "chart": None,
                        "df": None
                    })
            else:
                if advice_text:
                    # clean_text = str(advice_text).replace("\\n", "\n")
                    clean_text = advice_text
                    st.info(clean_text)
                else:
                    clean_text = "暫無建議"
                    st.warning("數據不足，無法生成分析建議。")
                    
                # 存入匯出清單
                report_items.append({
                    "title": "🧠 演練專業分析建議與對策", 
                    "text": clean_text,
                    "type": "text"
                })
# 確保這行放在程式碼的最底部
if st.sidebar.button("🛠️ 產製完整報告"):
        st.sidebar.markdown(generate_html_report(report_items, company_name), unsafe_allow_html=True)
        st.success("✅ 報告已產製，請點擊側邊欄按鈕下載。")


    
