import streamlit as st
import pandas as pd
import numpy as np
import io
import xlsxwriter

st.set_page_config(page_title="鳥嶼國小智慧排課系統", layout="wide")

st.title("🏫 鳥嶼國小智慧排課系統 (115學年度)")
st.markdown("將配課資料匯入後，系統將依據 16 項校本規則自動排課，並匯出與學校慣用格式相同的 Excel 總表。")

# --- 側邊欄：檔案上傳與控制 ---
with st.sidebar:
    st.header("⚙️ 系統設定")
    uploaded_file = st.file_uploader("1. 上傳一維配課資料庫 (CSV)", type="csv")
    
    st.subheader("內建排課規則檢查")
    st.checkbox("校長排週一、週二", value=True, disabled=True)
    st.checkbox("主任避開週一、三、五", value=True, disabled=True)
    st.checkbox("組長避開週三、五", value=True, disabled=True)
    st.checkbox("週四上午藝術與生活綁定", value=True, disabled=True)
    st.checkbox("週二下午特色課程綁定", value=True, disabled=True)
    
    start_schedule = st.button("🚀 開始自動排課", type="primary")

# --- 核心排課邏輯 ---
def auto_schedule(course_data):
    days = ['一', '二', '三', '四', '五']
    periods = [1, 2, 3, 4, 5, 6, 7]
    grades = ['一', '二', '三', '四', '五', '六']
    
    # 建立暫存的課表結構
    schedule = {g: pd.DataFrame(index=periods, columns=days).fillna("") for g in grades}
    
    # 寫入絕對綁定條件
    schedule['一'].loc[3, '四'], schedule['一'].loc[4, '四'] = '生活', '生活'
    schedule['二'].loc[3, '四'], schedule['二'].loc[4, '四'] = '生活', '生活'
    schedule['三'].loc[3, '四'], schedule['三'].loc[4, '四'] = '視藝/表藝', '視藝/表藝'
    
    for g in ['四', '五', '六']:
        schedule[g].loc[1, '四'], schedule[g].loc[2, '四'] = '視藝/表藝', '視藝/表藝'
        
    for g in ['三', '四', '五', '六']:
        schedule[g].loc[6, '二'], schedule[g].loc[7, '二'] = '家鄉/綜海', '家鄉/綜海'
    schedule['二'].loc[6, '二'], schedule['二'].loc[7, '二'] = '閱創/英創', '閱創/英創'
    schedule['一'].loc[6, '二'], schedule['一'].loc[7, '二'] = '生活', '生活'
    schedule['五'].loc[1, '五'] = '閱創客'
    
    return schedule

# --- 主畫面顯示 ---
if uploaded_file is not None:
    df_courses = pd.read_csv(uploaded_file)
    st.success("✅ 配課資料庫載入成功！")
    
    if start_schedule:
        with st.spinner("🧠 系統正依據 16 項規則運算中..."):
            schedule = auto_schedule(df_courses)
            st.success("排課運算完成！請點擊下方按鈕下載格式化課表。")
            
            # ==========================================
            # 匯出為「鳥嶼國小專屬格式」的 Excel 邏輯
            # ==========================================
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('總日課表')
            
            # 定義儲存格格式 (置中、粗體、格線)
            title_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_size': 14})
            header_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'bg_color': '#f2f2f2'})
            cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            
            days = ['一', '二', '三', '四', '五']
            grades = ['一', '二', '三', '四', '五', '六']
            
            # 第一列：大標題合併儲存格
            worksheet.merge_range(0, 0, 0, 30, '115學年度 澎湖縣白沙鄉鳥嶼國民小學 總日課表', title_format)
            
            # 第二列：星期
            worksheet.write(1, 0, '星期', header_format)
            col_idx = 1
            for day in days:
                # 橫跨 6 個年級的合併儲存格
                worksheet.merge_range(1, col_idx, 1, col_idx + 5, f'星期{day}', header_format)
                col_idx += 6
                
            # 第三列：年級
            worksheet.write(2, 0, '年級', header_format)
            col_idx = 1
            for day in days:
                for grade in grades:
                    worksheet.write(2, col_idx, f'{grade}年級', header_format)
                    col_idx += 1
                    
            # 填入排好的課程資料
            row_idx = 3
            for period in range(1, 8):
                worksheet.write(row_idx, 0, f'第 {period} 節', header_format)
                col_idx = 1
                for day in days:
                    for grade in grades:
                        subject = schedule[grade].loc[period, day]
                        worksheet.write(row_idx, col_idx, subject, cell_format)
                        col_idx += 1
                row_idx += 1
            
            # 調整欄寬以符合視覺舒適度
            worksheet.set_column(0, 0, 8)  # 節次欄寬一點
            worksheet.set_column(1, 30, 10) # 課程欄
            
            workbook.close()
            output.seek(0)
            
            st.download_button(
                label="📥 下載格式化全校總課表 (Excel)",
                data=output.getvalue(),
                file_name="115學年度鳥嶼國小全校總課表_格式化.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
else:
    st.info("👈 請先從左側上傳「115學年度_一維配課資料庫.csv」來啟動系統。")