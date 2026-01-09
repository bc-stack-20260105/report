import streamlit as st
import pandas as pd
import altair as alt
import base64
import json
import markdown
#import google.generativeai as genai  # 新增 Gemini SDK
from google import genai

# =================================================================
# 1. 頁面基本設定
# =================================================================
st.set_page_config(page_title="社交工程演練完整報告工具", layout="wide")
st.markdown("""
    <style>
    /* 針對 st.dataframe 或 st.table 的數值欄位強制靠左 */
    /* 這裡使用 div 選取器是為了確保覆蓋掉內建的數值靠右樣式 */
    div[data-testid="stTable"] td, 
    div[data-testid="stTable"] th,
    div[data-testid="stDataFrame"] td,
    div[data-testid="stDataFrame"] [style*="text-align: right"] {
        text-align: left !important;
        justify-content: flex-start !important;
    }

    /* 讓整個主頁面的容器寬度極大化，達成真正的 100% 佔滿感 */
    .main .block-container {
        max-width: 95% !important;
        padding-top: 2rem !important;
    }
    </style>
    """, unsafe_allow_html=True)
# =================================================================
# 2. 側邊欄：參數設定與 AI 配置
# =================================================================

# --- 區塊一：檔案上傳 ---
st.sidebar.header("📁 資料導入")
uploaded_file = st.sidebar.file_uploader("1. 上傳演練紀錄 (.xlsx)", type=["xlsx"])
config_file = st.sidebar.file_uploader("2. 上傳參數設定 (.txt)", type=["txt"])

# 初始化變數
company_name = ""
total_accounts = 99
total_emails_sent = 99
full_subject_list = []
tags_map = {"開啟信件": [], "點閱連結": [], "開啟附件": [], "輸入帳密": []}

# --- 區塊二：解析 TXT 參數 (核心邏輯保持不變) ---
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

# --- 區塊三：AI 顧問設定 ---
st.sidebar.divider()
st.sidebar.header("🤖 AI 顧問設定")
gemini_api_key = st.sidebar.text_input(
    "輸入 Gemini API Key", 
    type="password", 
    help="請至 Google AI Studio 申請"
)
enable_ai = st.sidebar.checkbox("開啟 AI 即時分析報告")
# 初始化 Client 邏輯
client = None
if enable_ai:
    if gemini_api_key:
        try:
            # 建立新版 SDK 的 Client 物件
            client = genai.Client(api_key=gemini_api_key)
            st.sidebar.success("🤖 AI 模式已就緒")
        except Exception as e:
            st.sidebar.error(f"AI 初始化失敗: {e}")
    else:
        st.sidebar.warning("⚠️ 請輸入 API Key 以啟用 AI 功能")
# --- 主畫面標題顯示 ---
if company_name:
    st.markdown(f"""
        <h1 style='text-align: left; margin-bottom: 0;'>📊 {company_name}</h1>
        <h2 style='text-align: left; margin-top: 0;'>社交工程演練統計報告</h2>
    """, unsafe_allow_html=True)
else:
    st.title("📊 社交工程演練統計報告")
    
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

    # 針對統計七：裝置安全性
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
        # 1. 初始化新版 Client
        client = genai.Client(api_key=api_key)
        
        # 2. 設定模型名稱 
        model_id = "gemini-2.5-flash" 
        
        prompt = f"""
        你是一位資深資安顧問，請分析以下社交工程演練數據：
        - 點閱連結：{context_data['click_count']} 人 (總數 {context_data['total_accounts']})
        - 輸入帳密：{context_data['login_count']} 人
        - 成功誘餌：{context_data['top_subject']}
        - 高風險單位：{context_data['top_dept']}
        
        請提供 3 點具體改善建議與一段員工資安宣導語。
        """
        
        # 3. 呼叫 API (新版語法)
        response = client.models.generate_content(
            model=model_id,
            contents=prompt
        )
        
        # 4. 取得內容 (新版直接存取 .text)
        if response and response.text:
            return response.text
        return "AI 回傳內容為空，請稍後重試。"

    except Exception as e:
        error_msg = str(e)
        # 修正 429 錯誤：配額限制
        if "429" in error_msg:
            return "⚠️ 請求太頻繁了！免費版 API 有頻率限制，請等 60 秒後再點擊一次。"
        # 修正 404 錯誤：模型名稱不對
        if "404" in error_msg:
            return f"❌ 找不到模型 '{model_id}'。請確認模型名稱是否正確（例如 gemini-2.5-flash）。"
        # 修正 401 錯誤：API Key 無效
        if "401" in error_msg:
            return "❌ API Key 無效，請檢查您的側邊欄設定。"
            
        return f"❌ AI 分析出錯：{error_msg}"

	
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
        /* 確保表格整體靠左，且文字靠左 */
        /* 表格滿版設計 */
        table {{
            width: 100% !important;
            border-collapse: collapse !important;
            margin: 20px 0 !important;
            table-layout: auto !important; /* 核心：自動根據文字長度調整欄位寬 */
        }}
        
        th, td {{
            text-align: left !important;
            padding: 12px !important;
            border: 1px solid #dee2e6 !important; /* 加強邊框線條 */
        }}
        
        th {{
            background-color: #f8f9fa !important;
            color: #333 !important;
        }}

        .table-responsive {{
            width: 100% !important;
            overflow-x: auto;
        }}
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
        with st.expander("🔍 查看詳細名單"): st.dataframe(u_users, use_container_width=True, hide_index=True)
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
        # st.table(sum2.set_index("項目"))
        # 找到統計二顯示表格的地方，改寫如下：
        st.write("數據明細：")
        # 建立一個複製品專門用來顯示
        display_sum2 = sum2.copy()

        # 將數值欄位轉為字串，這樣 Streamlit 就會預設靠左
        display_sum2['人'] = display_sum2['人'].astype(str)
        # 使用 dataframe 並透過 column_config 或是 CSS 控制
        st.dataframe(display_sum2, use_container_width=True, hide_index=True)
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
        # st.table(sum3.set_index("項目"))
        st.write("數據明細：")
        display_sum3=sum3.copy()
        display_sum3['次數'] = display_sum3['次數'].astype(str)
        # 使用 dataframe 並透過 column_config 或是 CSS 控制
        st.dataframe(display_sum3, use_container_width=True, hide_index=True)
        report_items.append({"title": "統計三：郵件主旨行為統計分析", "df": sum3, "chart": draw_horizontal_label_chart(sum3, "項目", "次數", color="#ED7D31", is_export=True), "text": analysis_3})

        # --- 統計四：各單位受測人數分布 ---
        st.divider(); st.subheader("🏢 統計四：各單位受測人數分布")
        sum4_df = df[df['std_tag'] != "其他"][[dept_col, email_col]].drop_duplicates()
        sum4_result = sum4_df.groupby(dept_col).size().reset_index(name='人數').sort_values(by='人數', ascending=False)
        # 準備統計四的分析文字
        top_dept_name = sum4_result.iloc[0][dept_col]
        top_dept_count = sum4_result.iloc[0]['人數']
        dept_count = len(sum4_result)

        analysis_4 = f"""
分析說明：

1. 曝險熱區分析：本次演練其中「{top_dept_name}」受測人數最多（{top_dept_count} 人），為本次演練的主要觀測對象。

2. 單位異質性觀察：各單位受測人數分佈不一，基數較大的單位其整體防禦意識對公司資安曝險程度影響最鉅。

3. 精準宣導策略：建議針對受測人數比例最高的前三個單位優先進行複訓，以達到最高成本效益的風險降級。
"""

        st.markdown(analysis_4)
        draw_horizontal_label_chart(sum4_result, dept_col, "人數", color="#70AD47")
        st.table(sum4_result.set_index(dept_col))
        report_items.append({"title": "統計四：各單位分布名單", "df": sum4_result, "chart": draw_horizontal_label_chart(sum4_result, dept_col, "人數", color="#70AD47", is_export=True), "text": analysis_4})

        # --- 統計五：郵件主旨影響力分析 ---
        st.divider(); st.subheader("📑 統計五：郵件主旨影響力分析")
        actual_s = df[[subject_col, email_col]].drop_duplicates().groupby(subject_col)[email_col].count().reset_index(name='觸及人數')
        all_s_df = pd.DataFrame(list(set(df[subject_col].unique().tolist() + full_subject_list)), columns=[subject_col])
        final_s = pd.merge(all_s_df, actual_s, on=subject_col, how='left').fillna(0)
        final_s['觸及人數'] = final_s['觸及人數'].astype(int)
        final_s = final_s.sort_values(by='觸及人數', ascending=False)
        # 取得關鍵數據
        top_subject = final_s.iloc[0][subject_col]  # 影響力最高的主旨
        top_subject_count = final_s.iloc[0]['觸及人數']
        
        analysis_5 = f"""
分析說明：

1. 核心誘餌識別：本次演練中「{top_subject}」主旨引發最強烈的反應，觸及人數高達 {top_subject_count} 人。這類主旨是目前最主要的資安破口。

2. 防禦加強方向：建議將排名第一的主旨作為「反面教材」進行案例解析，提醒同仁在收到類似內容時，應先確認發件者帳號而非僅看主旨。
"""
        # 顯示分析文字
        st.markdown(analysis_5)
        draw_horizontal_label_chart(final_s, subject_col, "觸及人數", color="#A5A5A5")
        st.table(final_s.set_index(subject_col))
        report_items.append({"title": "統計五：主旨影響力分析說明", "df": final_s, "chart": draw_horizontal_label_chart(final_s, subject_col, "觸及人數", color="#A5A5A5", is_export=True), "text": analysis_5})

        # --- 統計六：個人重複行為分析 ---
        # =================================================================
        # 📍 統計六：個人重複行為分析 (完整版)
        # =================================================================
        st.divider()
        st.subheader("📍 統計六：個人重複行為分析")

        # 1. 【準備階段】定義優先順序與暫存容器
        # 優先順序：輸入帳密 > 點閱連結 > 開啟附件 > 開啟信件
        priority_order = ["輸入帳密", "點閱連結", "開啟附件", "開啟信件"]
        best_tag_to_analyze = ""
        max_repeat_count = 0
        max_repeat_val = 0
        all_det_data = {}  # 存放各標籤的計算結果

        # 2. 【運算階段】先跑迴圈計算數據，但不顯示 UI
        for tag in priority_order:
            if tag == "開啟信件":
                # 開啟信件的邏輯：需包含後續所有動作
                a = df[df['std_tag'].isin(["點閱連結", "開啟附件", "輸入帳密"])][[email_col, name_col, dept_col, subject_col]].drop_duplicates()
                o = df[df['std_tag'] == "開啟信件"][[email_col, name_col, dept_col, subject_col]].drop_duplicates()
                det = pd.concat([a, o]).drop_duplicates().groupby([name_col, email_col, dept_col]).size().reset_index(name='次數')
            else:
                # 其他標籤的邏輯
                det = df[df['std_tag'] == tag].groupby([name_col, email_col, dept_col])[subject_col].nunique().reset_index(name='次數')
            
            # 存入暫存器供後續畫圖
            all_det_data[tag] = det
            
            # 尋找最嚴重的行為：如果該行為有人重複，且目前還沒選定分析對象，就選它
            repeats = len(det[det['次數'] >= 2])
            if repeats > 0 and best_tag_to_analyze == "":
                best_tag_to_analyze = tag
                max_repeat_count = repeats
                max_repeat_val = det['次數'].max()

        # 3. 【顯示階段 A：分析說明】將結論放在最前面
        if best_tag_to_analyze != "":
            analysis_6 = f"""
分析說明 (系統偵測本次最嚴重行為：【{best_tag_to_analyze}】)：

1. 行員警戒度落差：數據顯示有 {max_repeat_count} 位同仁在【{best_tag_to_analyze}】行為中出現 2 次（含）以上的重複行為。這代表單次的錯誤經驗未能有效轉化為警戒心，需加強此類人員的深度宣導。

2. 慣性風險識別：本次演練中，個人最高重複行為次數達 {max_repeat_val} 次。此數據顯示特定同仁對於多種不同誘餌主旨皆缺乏辨識力，屬於資安防護的最弱環節。

3. 差異化管理建議：建議將這 {max_repeat_count} 位重複發生者列為重點關懷對象，提供比一般同仁更高強度的實作訓練（如：強制收看資安宣導影片或參與補考），以降低未來真實攻擊中的中招機率。
"""
        else:
            analysis_6 = "分析說明：本次演練中，所有受測同仁在各項行為中均無重複中招之情形，顯示整體資安警戒心維持良好。"

        # 在所有圖表前顯示唯一的分析文字
        st.markdown(analysis_6)
        # 將結果存入 report_items 供匯出
        report_items.append({
            "title": f"統計六：【{tag}】行為名單與重複分析", 
            "text": analysis_6,
            "chart": None,
            "df": None
        })

        # 4. 【顯示階段 B：圖表與詳細名單】
        # 這裡照原順序跑出四個標籤的內容
        for tag in ["開啟信件", "點閱連結", "開啟附件", "輸入帳密"]:
            if tag not in all_det_data:
                continue
                
            det = all_det_data[tag]
            
            # 計算次數分佈 (1~5封)
            f_dist = det['次數'].value_counts().reindex([1,2,3,4,5], fill_value=0).reset_index()
            f_dist.columns = ['次數', '帳號數量']
            f_dist['標籤'] = f_dist['次數'].apply(lambda x: f"{tag[:2]}{x}封信")
            
            # 顯示標題與圖表
            st.markdown(f"#### 🏷️ 【{tag}】分佈")
            draw_horizontal_label_chart(f_dist, "標籤", "帳號數量", color="#4472C4")
            
            # 顯示統計表
            st.table(f_dist[['標籤', '帳號數量']].set_index('標籤'))
            
            # 詳細名單展開
            with st.expander(f"🔍 查看【{tag}】詳細名單 (含重複次數)"): 
                st.dataframe(det.sort_values(by='次數', ascending=False), use_container_width=True, hide_index=True)
            
            # # 將結果存入 report_items 供匯出
            report_items.append({
                "title": f"統計六：【{tag}】行為名單與重複分析", 
                "df": mask_pii(det, name_col, email_col), 
                "chart": draw_horizontal_label_chart(f_dist, "標籤", "帳號數量", color="#4472C4", is_export=True),
                "text": "" # 僅在核心標籤附帶分析文字
            })

        # --- 統計七：受測裝置分析 ---
        st.divider(); st.subheader("📱 統計七：受測裝置分析")
        if ua_col in df.columns:
            device_df = df.sort_values(by=email_col).drop_duplicates(subset=[email_col], keep='last').copy()
            device_df['裝置類型'] = device_df[ua_col].apply(parse_device)
            sum7 = device_df['裝置類型'].value_counts().reset_index()
            sum7.columns = ['裝置類型', '帳號數量']
            
            # --- 數據計算 ---
            total_clicks = sum7['帳號數量'].sum()
            # 判斷是否包含 Mobile 相關關鍵字
            mobile_mask = sum7['裝置類型'].str.contains('Mobile|手機|iOS|Android', case=False, na=False)
            mobile_count = sum7[mobile_mask]['帳號數量'].sum()
            mobile_ratio = (mobile_count / total_clicks * 100) if total_clicks > 0 else 0
            top_device = sum7.iloc[0]['裝置類型'] if not sum7.empty else "未知"

            # --- 動態分析文字判斷 ---
            if mobile_count > 0:
                mobile_analysis = f"行動辦公風險：行動裝置佔比為 {mobile_ratio:.1f}%。由於行動裝置螢幕限制，使用者難以第一時間辨識惡意連結的完整網址，此類裝置比例越高，代表越容易受到社交工程攻擊。"
                action_suggestion = "建議針對行動裝置使用者加強「長按連結預覽網址」的宣導，並提醒同仁在非固定辦公環境下處理郵件時應更加謹慎。"
            else:
                mobile_analysis = "裝置環境穩定：本次演練數據顯示，同仁全數使用桌面端（Desktop）裝置進行操作，並未偵測到行動裝置存取紀錄。"
                action_suggestion = "這顯示公司對於辦公裝置有良好的管控，或同仁已養成僅在公司標準工作站處理公務郵件的習慣，有助於降低因行動裝置螢幕限制造成的誤點風險。"

            analysis_7 = f"""
分析說明：

1. 主要存取途徑：本次演練中，同仁主要透過「{top_device}」裝置開啟郵件。這反映了企業內部目前的資訊使用習慣，可作為後續資安防護策略的重點佈署參考。

2. {mobile_analysis}

3. 宣導建議：{action_suggestion}
"""
            st.markdown(analysis_7)

            # --- 後續畫圖與名單程式碼 (保持不變) ---
            draw_horizontal_label_chart(sum7, "裝置類型", "帳號數量", color="#7294D4")
            st.table(sum7.set_index('裝置類型'))
            
            device_list = device_df[[name_col, email_col, '裝置類型', ua_col]].copy().sort_values(by='裝置類型')
            with st.expander("🔍 查看裝置詳細名單"): 
                st.dataframe(device_list, use_container_width=True, hide_index=True)
            
            report_items.append({
                "title": "統計七：裝置分析名單", 
                "df": mask_pii(device_list, name_col, email_col), 
                "chart": draw_horizontal_label_chart(sum7, "裝置類型", "帳號數量", color="#7294D4", is_export=True),
                "text": analysis_7
            })
        else:
            st.warning(f"Excel 中找不到『{ua_col}』欄位。")

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


    
