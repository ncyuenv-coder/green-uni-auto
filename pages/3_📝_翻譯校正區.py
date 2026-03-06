import streamlit as st
import streamlit.components.v1 as components
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

# 初始化跳轉與進度記憶
if 'jump_to_qid' not in st.session_state:
    st.session_state['jump_to_qid'] = None

# 🌟 初始化「畫面置頂」觸發器
if 'scroll_to_top' not in st.session_state:
    st.session_state['scroll_to_top'] = False

# 執行畫面置頂 JavaScript (當設定為 True 時觸發)
if st.session_state['scroll_to_top']:
    components.html(
        """
        <script>
            // 尋找 Streamlit 的主要捲動容器並強制歸零
            var main = window.parent.document.querySelector('.main');
            if (main) main.scrollTop = 0;
            var app = window.parent.document.querySelector('[data-testid="stAppViewContainer"]');
            if (app) app.scrollTop = 0;
        </script>
        """,
        height=0
    )
    st.session_state['scroll_to_top'] = False # 觸發後立刻關閉，避免重複執行

# ==========================================
# 🎨 系統 UI 樣式設定
# ==========================================
st.markdown("""
<style>
    .morandi-title { background-color: #738A96; color: white; padding: 15px 20px; border-radius: 8px; font-weight: bold; font-size: 1.5em; margin-bottom: 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .morandi-subtitle { background-color: #9DAB86; color: white; padding: 10px 15px; border-radius: 6px; font-weight: bold; font-size: 1.2em; margin-bottom: 15px; margin-top: 10px; }
    
    /* 編輯區放大 */
    div[data-baseweb="textarea"] textarea { font-size: 1.35em !important; line-height: 1.8 !important; color: #1A252C !important; padding: 15px !important; }
    
    /* 按鈕樣式 */
    div.stButton > button[kind="primary"] { border-radius: 8px !important; font-weight: bold !important; font-size: 1.3em !important; padding: 12px 30px !important; background-color: #D4A373 !important; border: none !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    div.stButton > button[kind="primary"]:hover { background-color: #C39262 !important; }
    div.stButton > button[kind="primary"] * { color: white !important; }
    
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

def update_translation_across_units(q_id, new_translation):
    try:
        creds = get_gcp_credentials()
        gc = gspread.authorize(creds)
        sh = gc.open_by_key('1JNbpZoZHWZRrIzn0whcQFnCDkOZghZmMyFidLE7dxT8')
        ws = sh.worksheet("評比題目表")
        
        records = ws.get_all_values()
        headers = [str(h).strip() for h in records[0]]
        
        try:
            q_id_col_idx = headers.index('當年度題目')
            trans_col_idx = headers.index('2025參考文字_AI預留')
        except ValueError:
            st.error("⚠️ 找不到『當年度題目』或『2025參考文字_AI預留』欄位，請確認 Google Sheet 欄位名稱。")
            return False, 0
            
        try:
            status_col_idx = headers.index('校正狀態')
        except ValueError:
            st.error("⚠️ 找不到『校正狀態』欄位！請先至 Google Sheet 的評比題目表新增一個名為『校正狀態』的欄位。")
            return False, 0
            
        cells_to_update = []
        for row_idx, row in enumerate(records):
            if row_idx == 0: continue
            
            if str(row[q_id_col_idx]).strip() == str(q_id).strip():
                cells_to_update.append(gspread.Cell(row=row_idx+1, col=trans_col_idx+1, value=str(new_translation)))
                cells_to_update.append(gspread.Cell(row=row_idx+1, col=status_col_idx+1, value='已校正'))
                
        if cells_to_update:
            ws.update_cells(cells_to_update)
            return True, len(cells_to_update) // 2 
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
    if '校正狀態' not in df_q.columns:
        st.warning("⚠️ 系統偵測到您的《評比題目表》中缺少「校正狀態」欄位。請先至 Google Sheet 新增該欄位，以便系統記錄您的存檔進度。")
        st.stop()

    df_unique = df_q.drop_duplicates(subset=['當年度題目']).copy()
    df_unique['當年度題目'] = df_unique['當年度題目'].astype(str).str.strip()
    df_unique = df_unique[df_unique['當年度題目'] != '']
    
    q_options = []
    for _, row in df_unique.iterrows():
        q_id = row['當年度題目']
        q_title = row.get('中文標題', '無標題')
        status_text = str(row.get('校正狀態', '')).strip()
        status_icon = "🟢" if status_text == '已校正' else "🔴"
        q_options.append(f"{status_icon} {q_id} - {q_title}")

    q_options.sort(key=lambda x: x.split(" ")[1])

    # 🌟 智慧跳躍邏輯
    if st.session_state.get('jump_to_qid'):
        target_qid = st.session_state['jump_to_qid']
        for opt in q_options:
            if f" {target_qid} -" in opt:
                st.session_state['q_selector'] = opt
                break
        st.session_state['jump_to_qid'] = None # 用完即清空

    col_sel, col_stat = st.columns([3, 1])
    with col_sel:
        st.markdown("<div class='morandi-subtitle'>🔍 請選擇欲校正之題目</div>", unsafe_allow_html=True)
        selected_option = st.selectbox("", q_options, label_visibility="collapsed", key="q_selector")
    
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

    if selected_option:
        selected_q_id = selected_option.split(" ")[1]
        current_data = df_unique[df_unique['當年度題目'] == selected_q_id].iloc[0]
        
        st.markdown(f"<h3 style='color: #154360;'>📖 {selected_q_id}：{current_data.get('中文標題', '')}</h3>", unsafe_allow_html=True)
        st.markdown("<div class='morandi-subtitle'>✍️ 翻譯文字校正區</div>", unsafe_allow_html=True)
        
        current_translation = str(current_data.get('2025參考文字_AI預留', '')).strip()
        if current_translation.lower() == 'nan': current_translation = ""
        
        input_key = f"trans_input_{selected_q_id}"
        
        new_translation = st.text_area(
            "編輯翻譯內容", 
            value=current_translation, 
            height=500, 
            key=input_key,
            label_visibility="collapsed",
            placeholder="請在此輸入或修改翻譯內容..."
        )
        
        st.write("") 
        col_space1, col_save, col_space2 = st.columns([1, 2, 1])
        with col_save:
            if st.button("💾 儲存並標記為已校正", type="primary", use_container_width=True):
                with st.spinner("⏳ 正在將修改同步至所有相關單位，並更新進度..."):
                    success, updated_count = update_translation_across_units(selected_q_id, new_translation)
                    
                    if success:
                        load_questions_data.clear()
                        st.session_state['last_saved_q'] = selected_q_id
                        st.session_state['last_saved_count'] = updated_count
                        
                        # 🌟 啟動置頂開關
                        st.session_state['scroll_to_top'] = True
                        
                        # 尋找「下一題」
                        current_idx = q_options.index(selected_option)
                        next_opt = None
                        
                        # 優先往後找 🔴
                        for i in range(current_idx + 1, len(q_options)):
                            if "🔴" in q_options[i]:
                                next_opt = q_options[i]
                                break
                        # 若後面沒有，從頭找 🔴
                        if not next_opt:
                            for i in range(0, current_idx):
                                if "🔴" in q_options[i]:
                                    next_opt = q_options[i]
                                    break
                        
                        # 設定目標題號並刷新頁面
                        if next_opt:
                            st.session_state['jump_to_qid'] = next_opt.split(" ")[1]
                        else:
                            st.session_state['jump_to_qid'] = selected_q_id # 全綠則停在原地
                            
                        st.rerun()
                    else:
                        st.error("⚠️ 資料尚未成功送出，請重新確認並點選送出！")

# 顯示成功儲存的提示訊息
if st.session_state.get('last_saved_q'):
    st.toast(f"🎉 題號 {st.session_state['last_saved_q']} 校正成功！已自動為您切換至下一題。", icon="✅")
    st.session_state['last_saved_q'] = None