import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
import re

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
    div.stButton > button:first-child {
        border-radius: 6px !important;
        font-weight: bold !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 資料庫連線與快取控制 (🌟 新增快取清除機制)
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
# 🪄 超級渲染引擎：將純文字轉換為 Streamlit 原生排版
# ==========================================
def render_native_content(text):
    if not isinstance(text, str) or not text.strip():
        st.info("*(🚧 系統提示：尚無去年度文字資料，待後台擷取後自動匯入。)*")
        return

    lines = text.split('\n')
    in_table = False
    current_table = []
    current_row = []
    current_cell = {"text": "", "images": []}
    
    ui_blocks = [] # 儲存解析後的區塊
    
    for line in lines:
        line_str = line.strip()
        if line_str.startswith('>'):
            line_str = line_str[1:].strip()
            
        # 隱藏並標記表格開始
        if "📊" in line_str or "表格資料解析" in line_str:
            in_table = True
            continue
            
        # 表格結束
        if in_table and line_str == "---":
            if current_cell["text"].strip() or current_cell["images"]:
                current_row.append(current_cell)
            if current_row:
                current_table.append(current_row)
            if current_table:
                ui_blocks.append({"type": "table", "rows": current_table})
            
            in_table = False
            current_table = []
            current_row = []
            current_cell = {"text": "", "images": []}
            continue
            
        if in_table:
            # 隱藏【第 X 列】並作為換列觸發器
            if re.search(r'【第 \d+ 列】', line_str):
                if current_cell["text"].strip() or current_cell["images"]:
                    current_row.append(current_cell)
                if current_row:
                    current_table.append(current_row)
                    
                current_row = []
                current_cell = {"text": "", "images": []}
                continue
                
            # 隱藏 [欄位 X] 並擷取內容
            if "[欄位" in line_str or "[Column" in line_str:
                if current_cell["text"].strip() or current_cell["images"]:
                    current_row.append(current_cell)
                    
                parts = re.split(r'[：:]', line_str, maxsplit=1)
                text_content = parts[1].strip() if len(parts) > 1 else ""
                text_content = text_content.replace('**', '').replace('*', '').strip()
                
                current_cell = {"text": text_content + "\n", "images": []}
                continue
            
            # 擷取圖片網址並轉換為防盜鏈縮圖 API
            img_matches = re.findall(r'!\[.*?\]\((https://drive\.google\.com/[^\)]+)\)', line_str)
            if img_matches:
                for img_url in img_matches:
                    thumb_url = img_url.replace("uc?id=", "thumbnail?sz=w1000&id=")
                    current_cell["images"].append(thumb_url)
                
                # 移除 Markdown 圖片語法，只保留純文字
                clean_line = re.sub(r'!\[.*?\]\(.*?\)', '', line_str).replace('**', '').replace('*', '').strip()
                if clean_line:
                    current_cell["text"] += clean_line + "\n"
                continue
            
            clean_line = line_str.replace('**', '').replace('*', '').strip()
            if clean_line:
                current_cell["text"] += clean_line + "\n"
                
        else:
            # 處理非表格的一般文字與圖片
            img_matches = re.findall(r'!\[.*?\]\((https://drive\.google\.com/[^\)]+)\)', line_str)
            if img_matches:
                for img_url in img_matches:
                    thumb_url = img_url.replace("uc?id=", "thumbnail?sz=w1000&id=")
                    ui_blocks.append({"type": "image", "url": thumb_url})
                
                clean_line = re.sub(r'!\[.*?\]\(.*?\)', '', line_str).replace('**', '').replace('*', '').strip()
                if clean_line:
                    ui_blocks.append({"type": "text", "content": clean_line})
            else:
                clean_line = line_str.replace('**', '').replace('*', '').strip()
                if clean_line:
                    ui_blocks.append({"type": "text", "content": clean_line})

    # 補足最後一筆資料
    if in_table:
        if current_cell["text"].strip() or current_cell["images"]:
            current_row.append(current_cell)
        if current_row:
            current_table.append(current_row)
        if current_table:
            ui_blocks.append({"type": "table", "rows": current_table})

    # 🚀 真正繪製到 Streamlit 網頁上
    for block in ui_blocks:
        if block["type"] == "text":
            st.markdown(block["content"])
        elif block["type"] == "image":
            # 非表格內的獨立圖片，限制寬度避免過大
            col1, col2 = st.columns([1, 1])
            with col1:
                st.image(block["url"], use_container_width=True)
        elif block["type"] == "table":
            for row in block["rows"]:
                if not row: continue
                # 自動依據欄位數量切分等寬欄位
                cols = st.columns(len(row))
                for idx, cell in enumerate(row):
                    with cols[idx]:
                        # 🌟 使用 Streamlit 內建帶邊框的容器，完美模擬表格儲存格！
                        with st.container(border=True):
                            # 照片在上方
                            for img_url in cell["images"]:
                                st.image(img_url, use_container_width=True)
                            # 翻譯文字在下方
                            if cell["text"].strip():
                                st.markdown(cell["text"].strip())


# ==========================================
# 🚀 網頁主畫面
# ==========================================
col_title, col_btn = st.columns([3, 1])
with col_title:
    st.title("📝 嘉大綠色大學填報區")
with col_btn:
    st.write("") # 排版佔位
    if st.button("🔄 同步最新雲端資料", use_container_width=True):
        load_gsheet_data.clear()
        st.rerun()

with st.spinner('🔄 正在載入最新評比題目...'):
    df_questions = load_gsheet_data()

if not df_questions.empty:
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
        # ✨ 終極版：原生元件渲染區塊
        # ==========================================
        with st.expander("💡 點擊展開查看：前一年度 (2025) 參考資訊", expanded=True):
            st.markdown(f"**對應之去年度題目：** {question_data.get('前一年度題目', '無')}")
            st.markdown("---")
            
            ref_text = question_data.get('2025參考文字_AI預留', '')
            
            # 將後台整理好的文字，餵給原生渲染引擎
            render_native_content(ref_text)
                
        st.markdown("---")
        
        # ==========================================
        # 🎯 年度成果填報與資料上傳
        # ==========================================
        with st.form("report_form"):
            report_text = st.text_area("✍️ 填報資訊/年度執行亮點成果", height=150, placeholder="請在此輸入您的填寫內容...")
            uploaded_files = st.file_uploader("📎 上傳照片或佐證檔案 (支援 PDF, JPG, PNG, DOCX 等)：", accept_multiple_files=True)
            
            st.markdown("<div style='display: flex; justify-content: center;'>", unsafe_allow_html=True)
            submitted = st.form_submit_button("📤 資料確認送出", type="primary")
            st.markdown("</div>", unsafe_allow_html=True)
            
            if submitted:
                if not report_text.strip() and not uploaded_files:
                    st.error("⚠️ 請至少填寫成果說明或上傳佐證檔案！")
                else:
                    st.success("🎉 資料已暫存成功！目前系統畫面設計與邏輯皆已完備。")

elif 'df_questions' in locals() and df_questions.empty:
    st.warning("目前資料庫中沒有題目，請確認 Google Sheet 的「評比題目表」是否已填入資料。")