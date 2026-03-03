import streamlit as st
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials

# ==========================================
# 🛡️ 資安防護罩
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

st.set_page_config(page_title="翻譯校正後台", page_icon="🔧", layout="wide")

# ==========================================
# 🔌 資料庫連線與更新設定
# ==========================================
def get_gspread_client():
    skey = st.secrets["gcp_oauth"].to_dict()
    credentials = Credentials(
        token=None,
        refresh_token=skey.get("refresh_token"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=skey.get("client_id"),
        client_secret=skey.get("client_secret")
    )
    return gspread.authorize(credentials)

@st.cache_data(ttl=600)
def load_gsheet_data():
    gc = get_gspread_client()
    SHEET_ID = '1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8'
    worksheet = gc.open_by_key(SHEET_ID).worksheet("評比題目表")
    df = pd.DataFrame(worksheet.get_all_records())
    df.columns = df.columns.str.strip()
    return df

# ==========================================
# 🚀 網頁主畫面：寬敞舒適的校正介面
# ==========================================
st.title("🔧 2025 年度資料：AI 翻譯校正台")
st.markdown("這裡提供寬敞的編輯空間，讓您輕鬆校正 AI 翻譯的內容。修改後按下儲存，將同步更新至資料庫。")
st.markdown("---")

df_questions = load_gsheet_data()

if not df_questions.empty:
    # 建立選單標示
    df_questions['選項標示'] = df_questions['當年度題目'].astype(str) + " - " + df_questions['中文標題'].astype(str)
    
    # 選擇要校正的題目
    selected_option = st.selectbox("📌 請選擇要校正的題目：", df_questions['選項標示'].tolist())
    selected_q_id = selected_option.split(" - ")[0]
    
    # 找出該題目前的資料
    question_data = df_questions[df_questions['當年度題目'].astype(str) == selected_q_id].iloc[0]
    current_translation = question_data.get('2025參考文字_AI預留', '')
    
    # 使用 Form 確保使用者編輯完再送出
    with st.form("edit_form"):
        st.subheader(f"📖 編輯內容：{selected_q_id}")
        
        # 🌟 寬敞的文字編輯區 (高度設為 400，非常好閱讀)
        edited_text = st.text_area(
            "請在此進行校正 (支援多段落排版)：",
            value=current_translation,
            height=400
        )
        
        submitted = st.form_submit_button("💾 確認儲存修改", type="primary")
        
        if submitted:
            with st.spinner("🔄 正在將修改寫回 Google Sheet..."):
                try:
                    gc = get_gspread_client()
                    worksheet = gc.open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8').worksheet("評比題目表")
                    
                    # 找出這題在 Google Sheet 裡是第幾列 (Row)
                    # gspread 的 cell 尋找：找到題號所在的儲存格，就知道是哪一列
                    cell = worksheet.find(selected_q_id)
                    row_index = cell.row
                    
                    # 找出 "2025參考文字_AI預留" 是第幾欄 (Column)
                    header_row = worksheet.row_values(1)
                    col_index = header_row.index('2025參考文字_AI預留') + 1
                    
                    # 執行寫入覆蓋
                    worksheet.update_cell(row_index, col_index, edited_text)
                    
                    # 清除快取，讓畫面馬上顯示最新結果
                    st.cache_data.clear()
                    st.success("✅ 修改已成功儲存並同步至雲端資料庫！")
                except Exception as e:
                    st.error(f"⚠️ 儲存失敗：{e}")