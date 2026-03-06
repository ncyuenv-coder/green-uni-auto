import streamlit as st
import pandas as pd
import gspread
import datetime
from google.oauth2.credentials import Credentials

# ==========================================
# 🌟 全域參數設定 & 資安防護罩
# ==========================================
if st.session_state.get("authentication_status") is not True:
    st.warning("⚠️ 請先至首頁登入系統！")
    st.stop()

st.set_page_config(page_title="嘉大綠色大學翻譯校正", page_icon="🌍", layout="wide")

# ==========================================
# 🎨 系統 UI 樣式設定 (極簡清爽風 + 編輯區文字放大)
# ==========================================
st.markdown("""
<style>
    /* 標題與外框樣式 */
    .morandi-title { background-color: #738A96; color: white; padding: 15px 20px; border-radius: 8px; font-weight: bold; font-size: 1.5em; margin-bottom: 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .morandi-subtitle { background-color: #9DAB86; color: white; padding: 10px 15px; border-radius: 6px; font-weight: bold; font-size: 1.2em; margin-bottom: 15px; margin-top: 10px; }
    
    /* 🌟 放大編輯區 (st.text_area) 內的文字，方便閱讀與修改 */
    div[data-baseweb="textarea"] textarea {
        font-size: 1.35em !important;
        line-height: 1.8 !important;
        color: #1A252C !important;
        padding: 15px !important;
    }
    
    /* 按鈕樣式 */
    div.stButton > button[kind="primary"] { border-radius: 8px !important; font-weight: bold !important; font-size: 1.3em !important; padding: 12px 30px !important; background-color: #D4A373 !important; border: none !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    div.stButton > button[kind="primary"]:hover { background-color: #C39262 !important; }
    div.stButton > button[kind="primary"] * { color: white !important; }
    
    /* 隱藏預設元件與調整選單大小 */
    div[data-baseweb="select"] { font-size: 1.15em !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 初始化及資料庫連線功能
# ==========================================
@st.cache_resource
def get_gcp_credentials():
    skey = st.secrets["gcp_oauth"].to_dict()
    return Credentials(
        token=None, refresh_token=skey.get("refresh_token"), 
        token_uri="https://oauth2.googleapis.com/token", 
        client_id=skey.get("client_id"), client_secret=skey.get("client_secret")
    )

@st.cache_data(ttl=60)
def load_questions_data():
    try:
        creds = get_gcp_credentials()
        gc = gspread.authorize(creds)
        sh = gc.open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8')
        ws = sh.worksheet("評比題目表")
        df = pd.DataFrame(ws.get_all_records())
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"讀取資料庫失敗: {e}")
        return pd.DataFrame()

# 🌟 核心：一鍵跨單位同步更新引擎 (同時更新翻譯內容與校正狀態)
def update_translation_across_units(q_id, new_translation):
    try:
        creds = get_gcp_credentials()
        gc = gspread.authorize(creds)
        sh = gc.open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8')
        ws = sh.worksheet("評比題目表")
        
        # 取得所有原始資料
        records = ws.get_all_values()
        headers = [str(h).strip() for h in records[0]]
        
        try:
            q_id_col_idx = headers.index('當年度題目')
            trans_col_idx = headers.index('2025參考文字_AI預留')
        except ValueError:
            st.error("⚠️ 找不到『當年度題目』或『2025參考文字_AI預留』欄位，請確認 Google Sheet 欄位名稱。")
            return False, 0
            
        # 檢查是否有建立「校正狀態」欄位
        try:
            status_col_idx = headers.index('校正狀態')
        except ValueError:
            st.error("⚠️ 找不到『校正狀態』欄位！請先至 Google Sheet 的評比題目表新增一個名為『校正狀態』的欄位。")
            return False, 0
            
        cells_to_update = []
        
        # 尋找所有該題號的列 (包含不同單位的同一題)
        for row_idx, row in enumerate(records):
            if row_idx == 0: continue # 略過標題列
            
            current_q_id = str(row[q_id_col_idx]).strip()
            if current_q_id == str(q_id).strip():
                # 更新翻譯文字
                cells_to_update.append(gspread.Cell(row=row_idx+1, col=trans_col_idx+1, value=str(new_translation)))
                # 更新校正狀態為「已校正」
                cells_to_update.append(gspread.Cell(row=row_idx+1, col=status_col_idx+1, value='已校正'))
                
        if cells_to_update:
            ws.update_cells(cells_to_update)
            return True, len(cells_to_update) // 2  # 回傳更新的列數
        return False, 0
        
    except Exception as e:
        st.error(f"寫入資料庫發生錯誤：{e}")
        return False, 0

# ==========================================
# 🚀 網頁主畫面與 UI 邏輯
# ==========================================
col_title, col_btn = st.columns([4, 1])
with col_title: 
    st.markdown("<div class='morandi-title'>🌍 翻譯校正區 (專注模式)</div>", unsafe_allow_html=True)
with col_btn:
    if st.button("🔄 同步最新資料", use_container_width=True):
        load_questions_data.clear()
        st.rerun()

with st.spinner("⏳ 正在載入題目庫..."):
    df_q = load_questions_data()

if not df_q.empty:
    # 檢查是否具備校正狀態欄位
    if '校正狀態' not in df_q.columns:
        st.warning("⚠️ 系統偵測到您的《評比題目表》中缺少「校正狀態」欄位。請先至 Google Sheet 新增該欄位，以便系統記錄您的存檔進度。")
        st.stop()

    # 🌟 整理出「唯一題號」清單 (濾除跨單位重複的題目)
    df_unique = df_q.drop_duplicates(subset=['當年度題目']).copy()
    df_unique['當年度題目'] = df_unique['當年度題目'].astype(str).str.strip()
    df_unique = df_unique[df_unique['當年度題目'] != '']
    
    # 組合下拉選單選項，並加入狀態提示
    q_options = []
    for _, row in df_unique.iterrows():
        q_id = row['當年度題目']
        q_title = row.get('中文標題', '無標題')
        # 🌟 進度改為：判斷「校正狀態」是否為「已校正」
        status_text = str(row.get('校正狀態', '')).strip()
        status_icon = "🟢" if status_text == '已校正' else "🔴"
        
        q_options.append(f"{status_icon} {q_id} - {q_title}")

    # 排序題號
    q_options.sort(key=lambda x: x.split(" ")[1])

    # ==========================================
    # 📍 題目選擇與進度條
    # ==========================================
    col_sel, col_stat = st.columns([3, 1])
    with col_sel:
        st.markdown("<div class='morandi-subtitle'>🔍 請選擇欲校正之題目</div>", unsafe_allow_html=True)
        selected_option = st.selectbox("", q_options, label_visibility="collapsed")
    
    with col_stat:
        total_q = len(q_options)
        done_q = sum(1 for opt in q_options if "🟢" in opt)
        st.markdown(f"""
        <div style='background-color: #FDF6E3; border: 1px solid #E6C27A; border-radius: 6px; padding: 15px; text-align: center; margin-top: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);'>
            <div style='font-size: 1.1em; color: #B85042; font-weight: bold;'>📝 編輯存檔進度</div>
            <div style='font-size: 1.6em; color: #2C3E50; font-weight: bold;'>{done_q} / {total_q}</div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("---")

    # ==========================================
    # 📍 翻譯校正專注編輯區 (滿版顯示)
    # ==========================================
    if selected_option:
        # 提取當前選擇的題號
        selected_q_id = selected_option.split(" - ")[0].replace("🟢 ", "").replace("🔴 ", "")
        
        # 取得該題的原始資料
        current_data = df_unique[df_unique['當年度題目'] == selected_q_id].iloc[0]
        
        # 直接顯示題號與標題，取消其他參考資訊
        st.markdown(f"<h3 style='color: #154360;'>📖 {selected_q_id}：{current_data.get('中文標題', '')}</h3>", unsafe_allow_html=True)
        st.markdown("<div class='morandi-subtitle'>✍️ 翻譯文字校正區</div>", unsafe_allow_html=True)
        
        # 取得目前的翻譯文字 (若為空則顯示空白)
        current_translation = str(current_data.get('2025參考文字_AI預留', '')).strip()
        if current_translation.lower() == 'nan': current_translation = ""
        
        # 使用 Session State 來管理輸入框內容，確保切換題目時更新
        input_key = f"trans_input_{selected_q_id}"
        
        # 🌟 大版面編輯區 (字體已透過 CSS 放大)
        new_translation = st.text_area(
            "編輯翻譯內容", 
            value=current_translation, 
            height=500, 
            key=input_key,
            label_visibility="collapsed",
            placeholder="請在此輸入或修改翻譯內容..."
        )
        
        # 操作按鈕區
        st.write("") 
        col_space1, col_save, col_space2 = st.columns([1, 2, 1])
        with col_save:
            if st.button("💾 儲存並標記為已校正", type="primary", use_container_width=True):
                with st.spinner("⏳ 正在將修改同步至所有相關單位，並更新進度..."):
                    success, updated_count = update_translation_across_units(selected_q_id, new_translation)
                    
                    if success:
                        # 清除快取，迫使重新抓取最新資料 (含校正狀態)
                        load_questions_data.clear()
                        st.session_state['last_saved_q'] = selected_q_id
                        st.session_state['last_saved_count'] = updated_count
                        st.rerun()

# 顯示成功儲存的提示訊息 (在重新整理後顯示)
if st.session_state.get('last_saved_q'):
    st.toast(f"🎉 題號 {st.session_state['last_saved_q']} 校正成功！已同步更新 {st.session_state['last_saved_count']} 筆跨單位紀錄。", icon="✅")
    st.session_state['last_saved_q'] = None # 顯示完就清除