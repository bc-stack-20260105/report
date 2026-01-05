import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
import tkinter as tk
from tkinter import filedialog
from matplotlib.backends.backend_pdf import PdfPages
import math

# --- 0. 環境設定與中文修正 ---
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False
sns.set_theme(style="whitegrid", font='Microsoft JhengHei')

def select_file():
    root = tk.Tk(); root.withdraw()
    return filedialog.askopenfilename(title="選擇演練 Excel 檔案", filetypes=[("Excel files", "*.xlsx *.xls")])

file_path = select_file()

if not file_path:
    print("未選取檔案，程式結束。")
else:
    try:
        df = pd.read_excel(file_path)
        
        # 欄位映射：抓取「目標部門」但內部邏輯視為 City
        df = df.rename(columns={
            '目標部門': 'Dept_Core', 
            '目標郵箱': 'Email',
            '郵件主旨': 'Template',
            '事件類型': 'Response'
        })
        
        people_denom = int(input("1. 請輸入受測人員總數 (分母): "))
        mail_denom = int(input("2. 請輸入總寄出郵件總封數 (分母): "))

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

        # --- 檔案儲存路徑 ---
        pdf_name = '社交工程演練_完整分析報告.pdf'
        excel_name = '社交工程演練_詳細資料.xlsx'
        
        # 使用 With 區塊確保 PDF 與 Excel 在程式結束前都會正確存檔
        with PdfPages(pdf_name) as pdf, pd.ExcelWriter(excel_name, engine='openpyxl') as writer:
            
            # --- 圖一：部門統計表格 (自動分頁) ---
            dept_counts = df_full[df_full['行為紀錄'] == '開啟信件'].drop_duplicates(subset=['Email', 'Dept_Core'])
            dept_final = dept_counts.groupby('Dept_Core').size().reset_index(name='人數').sort_values(by='人數', ascending=False)
            
            rows_per_page = 22
            total_pages = math.ceil(len(dept_final) / rows_per_page)

            for p in range(total_pages):
                fig, ax = plt.subplots(figsize=(10, 8))
                ax.axis('off')
                plt.text(0.5, 0.95, f'圖一：各部門參與人數統計 (第 {p+1}/{total_pages} 頁)', ha='center', fontsize=18, fontweight='bold')
                
                subset = dept_final.iloc[p*rows_per_page : (p+1)*rows_per_page]
                table_data = [['部門', '人數']] + subset.values.tolist()
                table = ax.table(cellText=table_data, loc='center', cellLoc='center', colWidths=[0.6, 0.2])
                table.auto_set_font_size(False); table.set_fontsize(11); table.scale(1.2, 2.0)
                for (row, col), cell in table.get_celld().items():
                    if row == 0:
                        cell.set_facecolor('#4A7EBB'); cell.set_text_props(color='white', fontweight='bold')
                pdf.savefig(); plt.close()
            
            dept_final.to_excel(writer, sheet_name='圖1_部門人數統計', index=False)

            # --- 數據準備 (圖二 ~ 圖八) ---
            summary_list = []
            for act in ['開啟信件', '點閱連結', '輸入帳密']:
                f_count = len(df_full[df_full['行為紀錄'] == act])
                p_count = df_full[df_full['行為紀錄'] == act]['Email'].nunique()
                summary_list.append([act, f_count, f"{(f_count/mail_denom*100):.2f}%", p_count, f"{(p_count/people_denom*100):.2f}%"])
            sum_df = pd.DataFrame(summary_list, columns=['項目', '總封數', '封數比率', '影響人數', '人數比率'])
            sum_df.to_excel(writer, sheet_name='數據摘要(圖2-4)', index=False)

            # --- 圖二、圖四：行為與封數佔比 ---
            for i, (col_y, col_pct, title) in enumerate([('影響人數','人數比率','圖二：各項行為影響人數比率'),('總封數','封數比率','圖四：演練總結封數比率')], start=2):
                if i == 3: continue 
                plt.figure(figsize=(10, 6.5))
                ax = sns.barplot(data=sum_df, x='項目', y=col_y, palette='viridis' if i==2 else 'Set2')
                for p, pct in zip(ax.patches, sum_df[col_pct]):
                    ax.annotate(f'{int(p.get_height())}\n({pct})', (p.get_x()+p.get_width()/2., p.get_height()), ha='center', va='bottom', fontweight='bold')
                plt.title(title, fontsize=15, pad=25); plt.margins(y=0.2); plt.tight_layout()
                pdf.savefig(); plt.close()

            # --- 圖三：郵件主旨統計 ---
            plt.figure(figsize=(10, 6.5))
            temp_counts = df.drop_duplicates(['Email', 'Template'])['Template'].value_counts().reset_index()
            temp_counts.columns = ['主旨', '人數']
            ax = sns.barplot(data=temp_counts, x='主旨', y='人數', palette='flare')
            for p in ax.patches:
                val = int(p.get_height())
                ax.annotate(f'{val}\n({(val/people_denom*100):.1f}%)', (p.get_x()+p.get_width()/2., p.get_height()), ha='center', va='bottom', fontweight='bold')
            plt.title('圖三：各郵件主旨受測佔比', fontsize=15, pad=25); plt.xticks(rotation=15, ha='right'); plt.tight_layout()
            pdf.savefig(); plt.close()
            temp_counts.to_excel(writer, sheet_name='圖3_主旨統計', index=False)

            # --- 圖五 ~ 圖八：行為次數分佈 ---
            for i, name in enumerate(['開啟信件', '點閱連結', '輸入帳密'], start=5):
                plt.figure(figsize=(10, 6.5))
                sub = df_full[df_full['行為紀錄'] == name]
                dist = sub.groupby('Email').size().value_counts().reindex(range(1, 6), fill_value=0).reset_index()
                dist.columns = ['次數', '人數']
                ax = sns.barplot(data=dist, x='次數', y='人數', color='#5B9BD5')
                ax.set_xticklabels([f'{name}{int(x)}次' for x in dist['次數']])
                for p in ax.patches:
                    ax.annotate(f'{int(p.get_height())}', (p.get_x()+p.get_width()/2., p.get_height()), ha='center', va='bottom', fontweight='bold')
                plt.title(f'圖{i}：{name}次數分佈統計', fontsize=15, pad=25); plt.tight_layout()
                pdf.savefig(); plt.close()
                # 同時輸出明細到 Excel
                sub.to_excel(writer, sheet_name=f'圖{i}_{name[:4]}明細', index=False)

        print(f"\n--- 檔案產出成功 ---")
        print(f"1. PDF 檔案: {os.path.abspath(pdf_name)}")
        print(f"2. Excel 檔案: {os.path.abspath(excel_name)}")

    except Exception as e:
        print(f"發生錯誤: {e}")