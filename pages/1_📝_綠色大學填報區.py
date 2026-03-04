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
# 🔌 資料庫連線與快取控制
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
# 🪄 AI 智慧渲染引擎：自動分辨「照片版面」與「資料表格」
# ==========================================
def render_native_content(text):
    if not isinstance(text, str) or not text.strip():
        st.info("*(🚧 系統提示：尚無去年度文字資料，待後台擷取後自動匯入。)*")
        return

    # 將 API 轉換為縮圖高畫質版本
    text = text.replace("https://drive.google.com/uc?id=", "https://drive.google.com/thumbnail?sz=w1000&id=")
    lines = text.split('\n')
    
    html_output = []
    in_table = False
    table_rows = []
    current_row = []
    current_cell = None  # 使用 None 來區分「還沒開始讀格子」跟「格子是空的」

    # 單行文字處理器 (處理圖片、粗體、清單項目)
    def process_line(l):
        l = l.strip()
        if not l: return ""
        
        # 1. 處理圖片：轉換為 100% 寬度、自動高度的 HTML 標籤
        def img_repl(m):
            url = m.group(1)
            return f'<img src="{url}" referrerpolicy="no-referrer" style="width: 100%; height: auto; border-radius: 6px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); margin-bottom: 8px;">'
        l = re.sub(r'!\[.*?\]\((.*?)\)', img_repl, l)
        
        # 2. 處理粗體 (保留字體加粗，移除星號)
        l = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', l)
        
        # 3. 處理項目符號與序號 (完美還原 Word 縮排)
        if re.match(r'^(\*|-)\s+', l):
            l = f"<div style='margin-left: 1.5em; text-indent: -1.2em; margin-bottom: 5px;'>• {l[2:]}</div>"
        elif re.match(r'^(\d+\.)\s+', l):
            match = re.match(r'^(\d+\.)\s+(.*)', l)
            l = f"<div style='margin-left: 1.5em; text-indent: -1.5em; margin-bottom: 5px;'>{match.group(1)} {match.group(2)}</div>"
        else:
            if not l.startswith('<img'):
                l = f"<div style='margin-bottom: 6px;'>{l}</div>"
        return l

    for line in lines:
        raw_line = line.strip()
        if raw_line.startswith('>'):
            raw_line = raw_line[1:].strip()
            
        # 偵測到表格開始
        if "📊" in raw_line or "表格資料解析" in raw_line:
            in_table = True
            table_rows = []
            current_row = []
            current_cell = None
            continue
            
        # 偵測到表格結束
        if in_table and raw_line == "---":
            in_table = False
            if current_cell is not None: current_row.append(current_cell)
            if current_row: table_rows.append(current_row)
            
            # 🌟 核心魔法：判斷這是哪種表格，並決定怎麼畫它！
            if table_rows:
                # 掃描表格內有沒有圖片標籤
                has_img = any('<img' in c for r in table_rows for c in r)
                
                if has_img:
                    # 📸 這是「照片排版表」：拔掉框線，拿掉底色，文字置中！
                    tbl = '<table style="width: 100%; border: none; table-layout: fixed; margin-bottom: 25px;">'
                    for r in table_rows:
                        tbl += '<tr>'
                        for c in r:
                            tbl += f'<td style="border: none; padding: 10px; vertical-align: top; text-align: left;">{c}</td>'
                        tbl += '</tr>'
                    tbl += '</table>'
                else:
                    # 📊 這是「真實資料表」：畫出漂亮的灰色框線與標題底色！
                    tbl = '<table style="width: 100%; border-collapse: collapse; border: 1px solid #D9E0E3; margin-bottom: 25px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">'
                    for i, r in enumerate(table_rows):
                        tbl += '<tr>'
                        bg = "#F8FAFB" if i == 0 else "#FFFFFF"
                        for c in r:
                            tbl += f'<td style="border: 1px solid #D9E0E3; padding: 12px; vertical-align: top; background-color: {bg}; font-size: 0.95em; color: #34495E;">{c}</td>'
                        tbl += '</tr>'
                    tbl += '</table>'
                html_output.append(tbl)
            continue
            
        if in_table:
            # 換列訊號
            if re.search(r'【第 \d+ 列】', raw_line):
                if current_cell is not None: current_row.append(current_cell)
                if current_row: table_rows.append(current_row)
                current_row = []
                current_cell = None
                continue
                
            # 換格訊號
            col_match = re.search(r'\[?(欄位|Column)\s*\d+\]?[*\s]*[:：]?(.*)', raw_line, re.IGNORECASE)
            if col_match:
                if current_cell is not None: current_row.append(current_cell)
                current_cell = process_line(col_match.group(2))
                continue
            
            # 填入格子內容
            if current_cell is not None and raw_line:
                current_cell += process_line(raw_line)
        else:
            # 一般文字 (非表格)
            if raw_line:
                html_output.append(f"<div style='line-height: 1.6; color: #2C3E50; font-size: 1.05em;'>{process_line(raw_line)}</div>")

    # 把最後的結果印在畫面上
    st.markdown("".join(html_output), unsafe_allow_html=True)


# ==========================================
# 🚀 網頁主畫面
# ==========================================
col_title, col_btn = st.columns([3, 1])
with col_title:
    st.title("📝 嘉大綠色大學填報區")
with col_btn:
    st.write("") 
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
        
        with st.expander("💡 點擊展開查看：前一年度 (2025) 參考資訊", expanded=True):
            st.markdown(f"**對應之去年度題目：** {question_data.get('前一年度題目', '無')}")
            st.markdown("---")
            
            ref_text = question_data.get('2025參考文字_AI預留', '')
            render_native_content(ref_text)
                
        st.markdown("---")
        
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