import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials

# ==========================================
# 🛡️ 資安防護罩：檢查是否已登入
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

st.set_page_config(page_title="嘉大綠色大學填報區", page_icon="📝", layout="wide")

# ==========================================
# 🔌 資料庫連線設定 (使用快取避免超過 API 限制)
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
        
        # 把所有欄位名稱的頭尾空白去掉
        df.columns = df.columns.str.strip()
        return df
        
    except Exception as e:
        st.error(f"⚠️ 無法讀取 Google Sheet，詳細錯誤：{e}")
        return pd.DataFrame()

# ==========================================
# 🚀 網頁主畫面：填報介面設計
# ==========================================
st.title("📝 嘉大綠色大學填報區")
st.info(f"👤 目前登入帳號：{st.session_state.get('username', '未知')}")

with st.spinner('🔄 正在從雲端載入最新評比題目，請稍候...'):
    df_questions = load_gsheet_data()

if not df_questions.empty:
    # 檢查必要欄位是否存在 (防呆機制)
    required_columns = ['2026年題目', '中文名稱', '資料需求', '權責單位', '2025年題目']
    missing_columns = [col for col in required_columns if col not in df_questions.columns]
    
    if missing_columns:
        st.error(f"⚠️ Google Sheet 找不到以下欄位：{', '.join(missing_columns)}")
        st.warning(f"🔍 目前實際讀到的欄位有：{df_questions.columns.tolist()}")
        st.stop()
        
    st.success("✅ 題目資料庫連線成功！")
    st.markdown("---")
    
    # ==========================================
    # 🎯 第一階段：篩選權責單位
    # ==========================================
    # 抓出所有不重複的權責單位，並過濾掉空白的
    unit_list = df_questions['權責單位'].dropna().unique().tolist()
    unit_list = [str(u).strip() for u in unit_list if str(u).strip() != '']
    
    selected_unit = st.selectbox(
        "🏢 步驟一：請選擇您的【權責單位】", 
        ["請選擇..."] + unit_list
    )
    
    if selected_unit != "請選擇...":
        # 依照選取的單位，過濾出專屬的題目
        df_unit_questions = df_questions[df_questions['權責單位'].astype(str).str.strip() == selected_unit].copy()
        
        # ==========================================
        # 🎯 第二階段：選擇題目
        # ==========================================
        # 組合「2026年題目 - 中文名稱」供選單顯示
        df_unit_questions['選項標示'] = df_unit_questions['2026年題目'].astype(str) + " - " + df_unit_questions['中文名稱'].astype(str)
        
        selected_option = st.selectbox(
            "📌 步驟二：請選擇要填報的【評比項目】", 
            df_unit_questions['選項標示'].tolist()
        )
        
        # 找出使用者選中的那一題的所有資料
        selected_q_id = selected_option.split(" - ")[0]
        question_data = df_unit_questions[df_unit_questions['2026年題目'].astype(str) == selected_q_id].iloc[0]
        
        st.markdown("---")
        
        # ==========================================
        # 🎯 第三階段：顯示題目詳細資訊與 2025 參考
        # ==========================================
        st.subheader(f"📖 {selected_option}")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**📄 資料需求 (2026)：**")
            st.info(question_data.get('資料需求', '無特別說明'))
            
        with col2:
            st.markdown("**💡 2025 年參考資訊：**")
            st.warning(question_data.get('2025年題目', '尚無去年度資料'))
            st.caption("*(備註：此處目前顯示 Google Sheet 內的文字，未來可擴充顯示 Word 萃取出的照片或完整翻譯)*")
            
        st.markdown("---")
        
        # ==========================================
        # 🎯 第四階段：本年度填報與上傳區
        # ==========================================
        st.subheader("📤 步驟三：本年度 (2026) 成果填報與上傳")
        
        with st.form("report_form"):
            # 填報文字框
            report_text = st.text_area(
                "✍️ 請填寫本年度執行成果與說明：", 
                height=150, 
                placeholder="請在此輸入您的填報內容..."
            )
            
            # 檔案上傳區 (可多選)
            uploaded_files = st.file_uploader(
                "📎 上傳照片或佐證檔案 (支援 PDF, JPG, PNG, DOCX 等，可多選)：", 
                accept_multiple_files=True
            )
            
            # 送出按鈕
            submitted = st.form_submit_button("✅ 送出填報資料")
            
            if submitted:
                if not report_text.strip() and not uploaded_files:
                    st.error("⚠️ 請至少填寫成果說明或上傳佐證檔案！")
                else:
                    st.success("🎉 資料已暫存成功！畫面上看起來已經完美了！")
                    st.info("*(開發中提示：下一個階段我們就會把這裡填寫的文字與檔案，真正寫入到您的「填報資料庫」與 Google Drive 中！)*")

elif 'df_questions' in locals() and df_questions.empty:
    st.warning("目前資料庫中沒有題目，請確認 Google Sheet 的「評比題目表」是否已填入資料。")