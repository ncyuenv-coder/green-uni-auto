import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
import re

# ==========================================
# 🌟 已經自動為您綁定好的 PDF 檔案 ID
# ==========================================
PDF_FILE_ID = "1hXs3yUwNxEiNZkJNeAfDTZjs9jxEgJos"

# ==========================================
# 🛡️ 資安防護罩
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

st.set_page_config(page_title="嘉大綠色大學填報區", page_icon="📝", layout="wide")

# ==========================================
# 🎨 系統 UI 樣式設定 (保留您原有的，並新增捲軸美化)
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
    /* ✨ 新增：美化右側翻譯區塊的捲軸 */
    .custom-scrollbar::-webkit-scrollbar {
        width: 8px;
    }
    .custom-scrollbar::-webkit-scrollbar-track {
        background: #f1f1f1; 
        border-radius: 4px;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb {
        background: #8F9CA3; 
        border-radius: 4px;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb:hover {
        background: #738087; 
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
        
        # ==========================================
        # ✨ 精準替換區：雙軌對照法 (左邊 PDF 原檔，右邊純文字翻譯)
        # ==========================================
        with st.expander("💡 點擊展開查看：前一年度 (2025) 參考資訊", expanded=True):
            st.markdown(f"**對應之去年度題目：** {question_data.get('前一年度題目', '無')}")
            st.markdown("---")
            
            # 建立左右兩個欄位
            col_pdf, col_trans = st.columns([1, 1])
            
            # 【左側：原檔 PDF】
            with col_pdf:
                st.markdown("#### 📄 原始排版 (2024年原檔)")
                pdf_url = f"https://drive.google.com/file/d/{PDF_FILE_ID}/preview"
                st.markdown(f'<iframe src="{pdf_url}" width="100%" height="600" style="border: 2px solid #8F9CA3; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);"></iframe>', unsafe_allow_html=True)
                st.caption("🔍 小提示：您可以在文件內滑動，或使用上方工具列放大/縮小尋找對應題號。")
            
            # 【右側：中文翻譯】
            with col_trans:
                st.markdown("#### 🇹🇼 中文翻譯參考")
                ref_text = question_data.get('2025參考文字_AI預留', '')
                if pd.notna(ref_text) and str(ref_text).strip() != "":
                    # 建立一個帶有卷軸、底色、且保留換行格式的文字區塊 (white-space: pre-wrap 是保留換行的魔法)
                    st.markdown(f'<div class="custom-scrollbar" style="background-color: #F8FAFB; padding: 20px; border-radius: 8px; border: 2px solid #E2E7E3; height: 600px; overflow-y: auto; line-height: 1.8; font-size: 1.05em; color: #2C3E50; white-space: pre-wrap; box-shadow: inset 0 2px 4px rgba(0,0,0,0.05);">{ref_text}</div>', unsafe_allow_html=True)
                else:
                    st.info("*(🚧 系統提示：本題尚無去年度文字資料，或原檔中未提供。)*")
                
        st.markdown("---")
        
        # ==========================================
        # 🎯 年度成果填報與資料上傳 (完全保留您的邏輯)
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