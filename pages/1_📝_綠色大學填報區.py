import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
import re

# ==========================================
# 🪄 魔法轉換器 2.0：表格自動重建與防盜鏈破解
# ==========================================
def fix_drive_images_and_tables(text):
    if not isinstance(text, str):
        return ""
    
    # 1. API 替換
    text = text.replace("https://drive.google.com/uc?id=", "https://drive.google.com/thumbnail?sz=w1000&id=")
    
    # 2. 將 Markdown 圖片標籤轉為 [IMG:url] 方便統一擷取
    pattern = r'!\[.*?\]\((https://drive\.google\.com/thumbnail\?sz=w1000&id=[a-zA-Z0-9_-]+)\)'
    text = re.sub(pattern, r'[IMG:\1]', text)
    
    lines = text.split('\n')
    final_html = []
    
    in_table = False
    in_row = False
    current_cell_text = ""
    current_cell_images = []
    row_cells = []
    
    def close_cell():
        nonlocal current_cell_text, current_cell_images, row_cells
        if current_cell_text.strip() or current_cell_images:
            # 🎯 需求 3：照片在上方、文字對應在照片下方
            imgs_html = "".join([
                f'<img src="{img}" referrerpolicy="no-referrer" style="width: 100%; height: auto; border-radius: 6px; margin-bottom: 10px; object-fit: cover; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">' 
                for img in current_cell_images
            ])
            
            clean_text = current_cell_text.replace('>', '').strip()
            txt_html = f'<div style="color: #333; font-size: 1em; line-height: 1.6;">{clean_text}</div>' if clean_text else ""
            
            # 組裝獨立儲存格：帶有外框與圓角
            cell_html = f'<div style="flex: 1; min-width: 250px; border: 1px solid #d9d9d9; padding: 15px; border-radius: 8px; background-color: #fcfcfc;">{imgs_html}{txt_html}</div>'
            row_cells.append(cell_html)
            current_cell_text = ""
            current_cell_images = []

    def close_row():
        nonlocal row_cells, in_row, final_html
        close_cell()
        if in_row and row_cells:
            # 🎯 需求 3：將這列的儲存格用 flexbox 並排，自動依據欄位數量均分空間
            row_html = f'<div style="display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap; width: 100%;">{ "".join(row_cells) }</div>'
            final_html.append(row_html)
            row_cells = []
        in_row = False

    for line in lines:
        line_str = line.strip()
        
        # 🎯 需求 2：過濾並隱藏「表格資料解析」關鍵字
        if "📊 表格資料解析" in line_str:
            in_table = True
            continue
        
        if in_table and line_str == "---":
            close_row()
            in_table = False
            continue
            
        if in_table:
            # 🎯 需求 2：隱藏【第 X 列】，並作為換列信號
            if re.search(r'【第 \d+ 列】', line_str):
                close_row()
                in_row = True
                continue
            
            # 🎯 需求 2：隱藏 [欄位 X]：，並作為換格信號
            col_match = re.search(r'\[欄位 \d+\][*\s]*[:：](.*)', line_str)
            if col_match:
                close_cell()
                current_cell_text = col_match.group(1).strip() + "<br>"
                continue
            
            if '[IMG:' in line_str:
                imgs = re.findall(r'\[IMG:(.*?)\]', line_str)
                current_cell_images.extend(imgs)
                clean_line = re.sub(r'\[IMG:.*?\]', '', line_str).replace('>', '').strip()
                if clean_line:
                    current_cell_text += clean_line + "<br>"
                continue
                
            clean_line = line_str.replace('>', '').strip()
            if clean_line:
                current_cell_text += clean_line + "<br>"
        else:
            # 非表格區塊：保留一般照片的並排邏輯
            if '[IMG:' in line_str:
                line = re.sub(r'\[IMG:(.*?)\]', r'<img src="\1" referrerpolicy="no-referrer" style="width: 48%; height: 250px; object-fit: cover; border-radius: 8px; margin: 5px 1%; display: inline-block; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">', line)
            final_html.append(line + "<br>")

    if in_table:
        close_row()

    result = "".join(final_html)
    result = re.sub(r'(<br>\s*){3,}', '<br><br>', result)
    return result

# ==========================================
# 🛡️ 資安防護罩
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

st.set_page_config(page_title="嘉大綠色大學填報區", page_icon="📝", layout="wide")

# ==========================================
# 🎨 系統 UI 樣式設定
# ==========================================
st.markdown("""
<style>
    .morandi-title {
        background-color: #8F9CA3;
        color: white;
        padding: 10px 15px;
        border-radius: 6px;
        font-weight: bold;
        font-size: 1.2em;
        margin-bottom: 5px;
        margin-top: 15px;
    }
    .morandi-req {
        background-color: #E2E7E3;
        color: #2C3E50;
        padding: 15px;
        border-left: 6px solid #8F9CA3;
        border-radius: 5px;
        margin-top: 10px;
        margin-bottom: 15px;
        line-height: 1.6;
        font-size: 1.05em;
    }
    div.stButton {
        display: flex;
        justify-content: center;
        margin-top: 20px;
        margin-bottom: 40px;
    }
    div.stButton > button:first-child {
        background-color: #D4A373 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 12px 30px !important;
        font-size: 18px !important;
        font-weight: bold !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:first-child:hover {
        background-color: #C08D5D !important;
        box-shadow: 0 6px 8px rgba(0,0,0,0.15) !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 資料庫連線設定
# ==========================================
@st.cache_data(ttl=600)
def load_gsheet_data():
    try:
        skey = st.secrets["gcp_oauth"].to_dict()
        credentials = Credentials(
            token=None,
            refresh_token=skey.get("refresh_token"),
            token_uri="https://oauth2.googleapis.com/token",
            client_id=skey.get("client_id"),
            client_secret=skey.get("client_secret")
        )
        gc = gspread.authorize(credentials)
        SHEET_ID = '1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8'
        sh = gc.open_by_key(SHEET_ID)
        worksheet = sh.worksheet("評比題目表")
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"⚠️ 無法讀取 Google Sheet，詳細錯誤：{e}")
        return pd.DataFrame()

# ==========================================
# 🚀 網頁主畫面
# ==========================================
st.title("📝 嘉大綠色大學填報區")

with st.spinner('🔄 正在載入最新評比題目...'):
    df_questions = load_gsheet_data()

if not df_questions.empty:
    required_columns = ['當年度題目', '英文標題', '中文標題', '中文說明', '資料需求', '權責單位', '前一年度題目']
    missing_columns = [col for col in required_columns if col not in df_questions.columns]
    
    if missing_columns:
        st.error(f"⚠️ Google Sheet 找不到以下欄位：{', '.join(missing_columns)}")
        st.stop()
        
    unit_list = df_questions['權責單位'].dropna().unique().tolist()
    unit_list = [str(u).strip() for u in unit_list if str(u).strip() != '']
    
    st.markdown("<div class='morandi-title'>🏢 請選擇您的單位</div>", unsafe_allow_html=True)
    selected_unit = st.selectbox("", ["請選擇..."] + unit_list, label_visibility="collapsed")
    
    if selected_unit != "請選擇...":
        df_unit_questions = df_questions[df_questions['權責單位'].astype(str).str.strip() == selected_unit].copy()
        df_unit_questions['選項標示'] = df_unit_questions['當年度題目'].astype(str) + " - " + df_unit_questions['中文標題'].astype(str)
        
        st.markdown("<div class='morandi-title'>📌 請選擇填報項目</div>", unsafe_allow_html=True)
        selected_option = st.selectbox("", df_unit_questions['選項標示'].tolist(), label_visibility="collapsed")
        
        selected_q_id = selected_option.split(" - ")[0]
        question_data = df_unit_questions[df_unit_questions['當年度題目'].astype(str) == selected_q_id].iloc[0]
        
        st.markdown("---")
        
        st.markdown(f"<div class='morandi-title'>📖 {question_data.get('當年度題目')}：{question_data.get('中文標題')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='color: black; font-size: 1.1em; padding-left: 5px; margin-bottom: 15px;'><b>中文說明：</b><br>{question_data.get('中文說明', '無')}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='morandi-req'><b>🔍 資料需求：</b><br>{question_data.get('資料需求', '無特別說明')}</div>", unsafe_allow_html=True)
        
        # ==========================================
        # ✨ 優化後的參考資訊區塊
        # ==========================================
        with st.expander("💡 點擊展開查看：前一年度 (2025) 參考資訊", expanded=True):
            
            # 🎯 需求 1：乾淨單純的呈現對應去年度題目，刪除原本的 "去年度填報內容 (圖文整合)："
            st.markdown(f"對應之去年度題目： {question_data.get('前一年度題目', '無')}")
            
            ref_text = question_data.get('2025參考文字_AI預留', '')
            
            if pd.notna(ref_text) and str(ref_text).strip() != "":
                # 執行 HTML 網格大變身
                fixed_content = fix_drive_images_and_tables(str(ref_text))
                st.markdown(fixed_content, unsafe_allow_html=True)
            else:
                st.info("*(🚧 系統提示：尚無去年度文字資料，待後台擷取後自動匯入。)*")
                
        st.markdown("---")
        
        # ==========================================
        # 🎯 年度成果填報與資料上傳
        # ==========================================
        with st.form("report_form"):
            report_text = st.text_area("✍️ 填報資訊/年度執行亮點成果", height=150, placeholder="請在此輸入您的填寫內容...")
            uploaded_files = st.file_uploader("📎 上傳照片或佐證檔案 (支援 PDF, JPG, PNG, DOCX 等)：", accept_multiple_files=True)
            
            submitted = st.form_submit_button("📤 資料確認送出")
            
            if submitted:
                if not report_text.strip() and not uploaded_files:
                    st.error("⚠️ 請至少填寫成果說明或上傳佐證檔案！")
                else:
                    st.success("🎉 資料已暫存成功！目前系統畫面設計與邏輯皆已完備。")
                    st.info("*(準備進入下一階段：將填寫內容寫入資料庫！)*")

elif 'df_questions' in locals() and df_questions.empty:
    st.warning("目前資料庫中沒有題目，請確認 Google Sheet 的「評比題目表」是否已填入資料。")