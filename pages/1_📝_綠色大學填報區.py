import streamlit as st
import pandas as pd
import gspread
import io
import re
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ==========================================
# 🌟 全域參數設定
# ==========================================
PDF_FILE_ID = "1hXs3yUwNxEiNZkJNeAfDTZjs9jxEgJos"
DRIVE_UPLOAD_FOLDER_ID = "1C1EbqH2oL-i_uqB6aEAw6S6tEwTvpnNc"

# ==========================================
# 🛡️ 資安防護罩
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

st.set_page_config(page_title="嘉大綠色大學填報區", page_icon="📝", layout="wide")

# ==========================================
# 🎨 系統 UI 樣式設定 (全面升級莫蘭迪色系與字體)
# ==========================================
st.markdown("""
<style>
    /* 1. 選單小標題：莫蘭迪綠 (放大一號) */
    .morandi-select-title {
        background-color: #9DAB86;
        color: white;
        padding: 12px 15px;
        border-radius: 6px;
        font-weight: bold;
        font-size: 1.2em; /* 放大一號 */
        margin-bottom: 5px;
        margin-top: 15px;
    }
    
    /* 2. 題目標題：莫蘭迪紫灰 (放大兩號) */
    .morandi-question-title {
        background-color: #948B89;
        color: white;
        padding: 15px 18px;
        border-radius: 6px;
        font-weight: bold;
        font-size: 1.4em; /* 放大兩號 */
        margin-bottom: 15px;
        margin-top: 10px;
    }

    /* 7. 資料填報區：莫蘭迪橘 (置中大字) */
    .morandi-orange-title {
        background-color: #D4A373;
        color: white;
        padding: 15px 20px;
        border-radius: 6px;
        font-weight: bold;
        font-size: 1.6em; /* 大字顯示 */
        margin-bottom: 15px;
        margin-top: 30px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }

    /* 7(1) 提供資料及佐證說明：淺藍色底色 */
    .light-blue-box {
        background-color: #E6F0F9;
        color: #2C3E50;
        padding: 20px;
        border-left: 6px solid #8FAAB8;
        border-radius: 6px;
        margin-bottom: 15px;
        font-size: 1.05em;
        line-height: 1.8;
    }

    /* 7(2)(3)(4) 淺黃色底色區塊 */
    .light-yellow-box {
        background-color: #FDF6E3;
        color: #2C3E50;
        padding: 20px;
        border-left: 6px solid #E6C27A;
        border-radius: 6px;
        margin-bottom: 15px;
        line-height: 1.8;
    }

    /* 4. 點擊展開查看：放大一號字 */
    div[data-testid="stExpander"] details summary p {
        font-size: 1.2em !important;
        font-weight: bold;
        color: #2C3E50;
    }
    
    /* 8. 送出按鈕：放大兩號字 */
    div.stButton > button[kind="primary"] {
        border-radius: 8px !important;
        font-weight: bold !important;
        font-size: 1.4em !important; /* 放大兩號字 */
        padding: 12px 30px !important;
    }
    
    /* 下拉選單內部文字放大 */
    div[data-baseweb="select"] {
        font-size: 1.1em !important;
    }

    .custom-scrollbar::-webkit-scrollbar { width: 8px; }
    .custom-scrollbar::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 4px; }
    .custom-scrollbar::-webkit-scrollbar-thumb { background: #8F9CA3; border-radius: 4px; }
    .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #738087; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 初始化及資料庫連線
# ==========================================
def get_gcp_credentials():
    skey = st.secrets["gcp_oauth"].to_dict()
    return Credentials(
        token=None,
        refresh_token=skey.get("refresh_token"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=skey.get("client_id"),
        client_secret=skey.get("client_secret")
    )

@st.cache_data(ttl=600)
def load_gsheet_data():
    try:
        creds = get_gcp_credentials()
        gc = gspread.authorize(creds)
        sh = gc.open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8')
        worksheet = sh.worksheet("評比題目表")
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"⚠️ 無法讀取 Google Sheet，詳細錯誤：{e}")
        return pd.DataFrame()

# 上傳檔案至 Google Drive 的功能
def upload_file_to_drive(file_obj, filename, folder_id):
    try:
        creds = get_gcp_credentials()
        drive_service = build('drive', 'v3', credentials=creds)
        file_metadata = {'name': filename, 'parents': [folder_id]}
        media = MediaIoBaseUpload(io.BytesIO(file_obj.getvalue()), mimetype=file_obj.type, resumable=True)
        uploaded_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return uploaded_file.get('id')
    except Exception as e:
        st.error(f"上傳失敗 ({filename}): {e}")
        return None

# 初始化動態上傳區塊的數量 (預設 5 個)
if "upload_count" not in st.session_state:
    st.session_state.upload_count = 5

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
    
    # 1. 請選擇您的單位
    st.markdown("<div class='morandi-select-title'>🏢 請選擇您的單位</div>", unsafe_allow_html=True)
    selected_unit = st.selectbox("", ["請選擇..."] + unit_list, label_visibility="collapsed")
    
    if selected_unit != "請選擇...":
        df_unit_questions = df_questions[df_questions['權責單位'].astype(str).str.strip() == selected_unit].copy()
        df_unit_questions['選項標示'] = df_unit_questions['當年度題目'].astype(str) + " - " + df_unit_questions['中文標題'].astype(str)
        
        # 1. 請選擇填報項目
        st.markdown("<div class='morandi-select-title'>📌 請選擇填報項目</div>", unsafe_allow_html=True)
        selected_option = st.selectbox("", df_unit_questions['選項標示'].tolist(), label_visibility="collapsed")
        
        selected_q_id = selected_option.split(" - ")[0]
        question_data = df_unit_questions[df_unit_questions['當年度題目'].astype(str) == selected_q_id].iloc[0]
        
        st.markdown("---")
        
        # 2. 題目標題
        st.markdown(f"<div class='morandi-question-title'>📖 {question_data.get('當年度題目')}：{question_data.get('中文標題')}</div>", unsafe_allow_html=True)
        
        # 3. 中文說明 (去除前綴字，放大一號字)
        desc_text = str(question_data.get('中文說明', '無')).replace('中文說明：', '').replace('中文說明:', '').strip()
        st.markdown(f"<div style='color: #2C3E50; font-size: 1.2em; padding-left: 5px; margin-bottom: 25px;'>{desc_text}</div>", unsafe_allow_html=True)
        
        # ==========================================
        # 6. 去年度參考資訊 (判斷有無資料)
        # ==========================================
        prev_q = str(question_data.get('前一年度題目', '')).strip()
        has_prev_data = bool(prev_q and prev_q.lower() not in ['nan', '無', ''])
        
        if not has_prev_data:
            st.info("💡 本項目為今年度新增，尚無去年提報參考資料。")
        else:
            with st.expander("💡 點擊展開查看：前一年度參考資訊", expanded=True):
                st.markdown(f"**對應之去年度題目：** {prev_q}")
                st.markdown("---")
                
                col_pdf, col_trans = st.columns([1, 1])
                with col_pdf:
                    st.markdown("#### 📄 前一年度提報內容") # 5. 修改文字
                    pdf_id = str(question_data.get('單題PDF_ID', '')).strip()
                    if pdf_id and pdf_id.lower() != 'nan':
                        pdf_url = f"https://drive.google.com/file/d/{pdf_id}/preview"
                        st.markdown(f'<iframe src="{pdf_url}" width="100%" height="600" style="border: 2px solid #8F9CA3; border-radius: 8px;"></iframe>', unsafe_allow_html=True)
                    else:
                        st.info("*(🚧 系統提示：尚無本題專屬的 PDF 原檔可供預覽。)*")
                
                with col_trans:
                    st.markdown("#### 🇹🇼 中文翻譯參考") # 5. 修改文字
                    ref_text = question_data.get('2025參考文字_AI預留', '')
                    if pd.notna(ref_text) and str(ref_text).strip() != "":
                        st.markdown(f'<div class="custom-scrollbar" style="background-color: #F8FAFB; padding: 20px; border-radius: 8px; border: 2px solid #E2E7E3; height: 600px; overflow-y: auto; line-height: 1.8; font-size: 1.1em; color: #2C3E50; white-space: pre-wrap;">{ref_text}</div>', unsafe_allow_html=True)
                    else:
                        st.info("*(🚧 系統提示：本題尚無文字資料。)*")
        
        # ==========================================
        # 7. 🎯 資料填報區 (全新排版設計)
        # ==========================================
        st.markdown("<div class='morandi-orange-title'>資料填報區</div>", unsafe_allow_html=True)
        
        # 7(1) 提供資料及佐證示意照片說明
        st.markdown("""
        <div class='light-blue-box'>
            <strong style='font-size: 1.15em;'>提供資料及佐證示意照片說明：</strong><br>
            <ul>
                <li><b>成果說明文字：</b>請重點明確說明，包括質化、量化成果資訊。</li>
                <li><b>佐證示意照片：</b>請提供 .jpg 檔，至少提供 3-4 張 (可多提供)，如說明有指定數量，請對應提供。</li>
                <li>若需求資料僅有照片提供，則請您於「填報資訊 / 年度執行亮點成果」欄位填「<b>無須提供</b>」即可。</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        # 7(2) 🔍 請貴單位協助提供下列資料
        req_text = str(question_data.get('資料需求', '無特別說明')).replace('\n', '<br>')
        st.markdown(f"""
        <div class='light-yellow-box'>
            <strong style='font-size: 1.4em;'>🔍 請貴單位協助提供下列資料：</strong><br>
            <div style='font-size: 1.2em; margin-top: 10px; padding-left: 10px;'>{req_text}</div>
        </div>
        """, unsafe_allow_html=True)
        
        # 7(3) 填報資訊 / 年度執行亮點成果
        st.markdown("<div class='light-yellow-box'>", unsafe_allow_html=True)
        st.markdown("<strong style='font-size: 1.2em;'>✍️ 填報資訊 / 年度執行亮點成果</strong>", unsafe_allow_html=True)
        st.caption("💡 編輯小提示：您可直接使用數字 (1. 2.) 或連字號 (-) 來列出項目符號。")
        report_text = st.text_area("填寫區", height=200, placeholder="請在此輸入您的填寫內容...", label_visibility="collapsed")
        st.markdown("</div>", unsafe_allow_html=True)
        
        # 7(4) 上傳照片或佐證檔案
        st.markdown("<div class='light-yellow-box'>", unsafe_allow_html=True)
        st.markdown("<strong style='font-size: 1.2em;'>📎 上傳照片或佐證檔案 (支援 PDF, JPG, PNG, DOCX 等)</strong>", unsafe_allow_html=True)
        
        # 動態增減檔案區按鈕
        col_add, col_sub, _ = st.columns([2, 2, 6])
        if col_add.button("➕ 新增一筆檔案區"):
            st.session_state.upload_count += 1
            st.rerun()
        if col_sub.button("➖ 減少一筆檔案區"):
            if st.session_state.upload_count > 1:
                st.session_state.upload_count -= 1
                st.rerun()
        
        st.write("")
        upload_data = []
        for i in range(st.session_state.upload_count):
            c1, c2 = st.columns([1, 2])
            with c1:
                desc = st.text_input(f"📌 檔案 {i+1} 資料說明：", key=f"desc_{i}", placeholder="例如：智慧路燈系統畫面")
            with c2:
                file = st.file_uploader(f"📂 選擇檔案 {i+1}", key=f"file_{i}", label_visibility="collapsed")
            upload_data.append((desc, file))
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # ==========================================
        # 📤 資料送出與自動改名上傳邏輯
        # ==========================================
        st.markdown("<br><div style='display: flex; justify-content: center;'>", unsafe_allow_html=True)
        submitted = st.button("📤 資料確認送出", type="primary", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        if submitted:
            valid_files = [f for d, f in upload_data if f is not None]
            
            if not report_text.strip() and not valid_files:
                st.error("⚠️ 請至少填寫成果說明，或上傳一份佐證檔案！")
            else:
                with st.spinner("⏳ 正在將檔案安全上傳至 Google Drive..."):
                    upload_success_count = 0
                    for i, (desc, file) in enumerate(upload_data):
                        if file is not None:
                            # 自動組合新檔名：[當年度題目]-[檔案資料說明].[副檔名]
                            ext = file.name.split('.')[-1]
                            safe_desc = desc.strip() if desc.strip() else f"未命名附件{i+1}"
                            new_filename = f"{selected_q_id}-{safe_desc}.{ext}"
                            
                            # 執行上傳
                            file_id = upload_file_to_drive(file, new_filename, DRIVE_UPLOAD_FOLDER_ID)
                            if file_id:
                                upload_success_count += 1
                                
                if valid_files and upload_success_count == len(valid_files):
                    st.success(f"🎉 填報成功！共上傳 {upload_success_count} 份檔案至專屬資料夾。")
                elif valid_files:
                    st.warning(f"⚠️ 填報完成，但部分檔案上傳失敗 (成功 {upload_success_count}/{len(valid_files)} 份)。")
                else:
                    st.success("🎉 文字資料已成功暫存送出！")

elif 'df_questions' in locals() and df_questions.empty:
    st.warning("目前資料庫中沒有題目，請確認 Google Sheet 的「評比題目表」是否已填入資料。")