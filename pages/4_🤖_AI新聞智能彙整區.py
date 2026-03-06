import streamlit as st
import pandas as pd
import datetime
import io
import requests
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import qn

# ==========================================
# 🌟 全域設定 & UI 樣式
# ==========================================
st.set_page_config(page_title="嘉大 AI 新聞智能彙整", page_icon="🤖", layout="wide")

st.markdown("""
<style>
    button[data-baseweb="tab"] { background-color: #E6F0F9 !important; border-radius: 8px 8px 0px 0px !important; margin-right: 4px !important; padding: 10px 20px !important; border: 1px solid #AED6F1 !important; border-bottom: none !important; transition: all 0.3s ease; }
    button[data-baseweb="tab"] p { font-size: 1.4em !important; font-weight: bold !important; color: #2C3E50 !important; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #154360 !important; border: 1px solid #0B2331 !important; border-bottom: 3px solid #0B2331 !important; }
    button[data-baseweb="tab"][aria-selected="true"] p { color: #FFFFFF !important; }

    .morandi-title { background-color: #738A96; color: white; padding: 15px 20px; border-radius: 8px; font-weight: bold; font-size: 1.5em; margin-bottom: 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .morandi-dark-title { background-color: #5C6B73; color: white; padding: 12px 20px; border-radius: 6px; font-weight: bold; font-size: 1.3em; margin-bottom: 10px; margin-top: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    
    .raw-box { background-color: #FDF6E3; padding: 20px; border-left: 5px solid #E6C27A; border-radius: 6px; height: 400px; overflow-y: auto; line-height: 1.6; font-size: 1.1em; color: #555;}
    .ai-box { background-color: #E6F0F9; padding: 20px; border-left: 5px solid #8FAAB8; border-radius: 6px; height: 400px; overflow-y: auto; line-height: 1.8; font-size: 1.2em; color: #154360; font-weight: 500;}
    
    div.stButton > button[kind="primary"] { border-radius: 8px !important; font-weight: bold !important; font-size: 1.3em !important; padding: 12px 30px !important; background-color: #D4A373 !important; border: none !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔌 模擬資料庫與輔助工具 (待替換為真實 API)
# ==========================================
def mock_scrape_and_ai_process(start_date, end_date):
    """這是一個模擬器，展示未來 AI 抓完資料的長相"""
    import time
    time.sleep(2) # 模擬思考時間
    return [
        {
            "題號": "1.1", "中文標題": "校園節能減碳設施建置",
            "新聞日期": "2024-03-01", "新聞標題": "嘉義大學導入ISO50001並更新高耗能冷水機組",
            "原始內容": "國立嘉義大學為響應政府節能減碳政策，於本學期正式導入ISO50001能源管理系統。總務處表示，本次工程除了建置智慧電表外，更將圖書館老舊的高耗能冷水機組進行全面汰換。預計每年可節省約15%的用電量... (此處省略500字新聞稿)",
            "AI摘要": "國立嘉義大學廣泛運用資訊與通訊技術(ICT)支援節能計畫：\n1. ISO 50001 能源管理系統：採用該系統搭配智慧電表，精準追蹤校園用電。\n2. 汰換老舊設備：2024年更換圖書館空調系統與高耗能冷水機組，有效提升能源效率。\n【對應SDGs】：SDG 7 (可負擔的潔淨能源)、SDG 13 (氣候行動)",
            "照片清單": ["https://picsum.photos/800/600", "https://picsum.photos/600/800", "https://picsum.photos/1200/400"], # 模擬橫、直、全景圖
            "新聞連結": "https://www.ncyu.edu.tw/news/12345"
        }
    ]

def url_to_bytes(url):
    """將網路圖片 URL 轉為 Word 可用的 Bytes"""
    try:
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        return io.BytesIO(response.content)
    except:
        return None

# ==========================================
# 🌟 Word 產製引擎 (完全比照截圖排版)
# ==========================================
def set_run_font(run):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def generate_ai_word_report(q_id, q_title, news_records):
    doc = Document()
    
    # 標題
    doc.add_heading(f'{q_id} {q_title}', level=1)
    
    for idx, record in enumerate(news_records):
        if idx > 0:
            doc.add_paragraph("="*40) # 分隔線
        
        urls = record['照片清單']
        
        # 1. 照片表格 (置於最上方，比照截圖)
        if urls:
            table = doc.add_table(rows=0, cols=2)
            table.style = 'Table Grid'
            
            # 簡易排版邏輯：每兩張一列
            for i in range(0, len(urls), 2):
                row = table.add_row()
                # 第一張圖
                c1 = row.cells[0]
                c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p1 = c1.paragraphs[0]
                p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                img_bytes1 = url_to_bytes(urls[i])
                if img_bytes1:
                    p1.add_run().add_picture(img_bytes1, width=Cm(7.5))
                else:
                    p1.add_run("(圖片載入失敗)")
                    
                # 第二張圖 (如果有)
                c2 = row.cells[1]
                c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p2 = c2.paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if i + 1 < len(urls):
                    img_bytes2 = url_to_bytes(urls[i+1])
                    if img_bytes2:
                        p2.add_run().add_picture(img_bytes2, width=Cm(7.5))
                    else:
                        p2.add_run("(圖片載入失敗)")
                        
        # 2. 新聞標題與內容
        doc.add_heading(f"新聞標題：{record['新聞標題']}", level=2)
        
        p_date = doc.add_paragraph()
        p_date.add_run(f"發布日期：{record['新聞日期']}").italic = True
        
        # 處理摘要 (支援縮排)
        report_text_clean = re.sub(r'(^|\n)(\d+[\.\)]|[-•*])\s*\n\s*', r'\1\2 ', str(record['AI摘要']))
        for line in report_text_clean.replace('\r', '').split('\n'):
            line = line.strip()
            if not line: continue
                
            p = doc.add_paragraph()
            match = re.match(r'^([0-9a-zA-Z]+[\.、\)]|[\(（][0-9a-zA-Z一二三四五六七八九十]+[\)）]|[一二三四五六七八九十]+、|[-•*])\s*(.*)', line)
            if match:
                marker = match.group(1)
                content = match.group(2)
                p.paragraph_format.left_indent = Cm(0.8)
                p.paragraph_format.first_line_indent = Cm(-0.8)
                p.add_run(f"{marker} {content}")
            else:
                p.add_run(line)
        
        # 3. 原始新聞連結
        p_link = doc.add_paragraph()
        p_link.add_run("Additional evidence link: ").bold = True
        p_link.add_run(record['新聞連結'])
        
    # 設定全文件字體
    for p in doc.paragraphs:
        if not p.style.name.startswith('Heading'): p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        for r in p.runs: set_run_font(r)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs: set_run_font(r)
                    
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# ==========================================
# 🚀 網頁主畫面與 UI 邏輯
# ==========================================
st.markdown("<div class='morandi-title'>🤖 綠色大學 AI 新聞智能彙整中心</div>", unsafe_allow_html=True)

tab_ai, tab_view = st.tabs(["🤖 AI 智慧搜尋與生成區", "📥 檢視與下載彙整區"])

# ==========================================
# 🏷️ TAB 1: AI 智慧搜尋與生成區
# ==========================================
with tab_ai:
    st.markdown("<div class='morandi-dark-title'>⚙️ 批次處理任務設定</div>", unsafe_allow_html=True)
    c_start, c_end, c_btn = st.columns([2, 2, 2])
    with c_start: start_date = st.date_input("新聞起日", datetime.date(2023, 1, 1))
    with c_end: end_date = st.date_input("新聞迄日", datetime.date.today())
    
    st.info("💡 系統將讀取《原始新聞抓取》總表，根據題目與中文名稱，自動爬取此區間內的學校新聞，並由 Gemini 進行摘要與 SDGs 標註，最後存入《AI新聞資料庫》。")
    
    with c_btn:
        st.write("")
        if st.button("🚀 開始啟動 AI 批次彙整", type="primary", use_container_width=True):
            with st.spinner("🤖 爬蟲運作與 AI 摘要生成中，這可能需要幾分鐘的時間，請稍候..."):
                # 這裡目前呼叫 Mock Function，未來換成真實爬蟲
                results = mock_scrape_and_ai_process(start_date, end_date)
                st.session_state['temp_results'] = results
                st.success("✅ 批次處理完成！(目前為展示版，資料暫存於記憶體)")

    # 顯示處理結果對照 (管理員檢視用)
    if st.session_state.get('temp_results'):
        st.markdown("---")
        st.markdown("<div class='morandi-dark-title'>👀 最新 AI 處理成果預覽</div>", unsafe_allow_html=True)
        
        for res in st.session_state['temp_results']:
            st.markdown(f"<h3 style='color:#B85042;'>📌 {res['題號']}：{res['中文標題']}</h3>", unsafe_allow_html=True)
            
            c_raw, c_ai = st.columns([1, 1])
            with c_raw:
                st.markdown("**📄 爬蟲抓取原始新聞**")
                st.markdown(f"<div class='raw-box'><b>標題：{res['新聞標題']}</b><br>日期：{res['新聞日期']}<br><br>{res['原始內容']}<br><br>🔗 連結：{res['新聞連結']}</div>", unsafe_allow_html=True)
            with c_ai:
                st.markdown("**✨ Gemini 智能濃縮 (含 SDGs)**")
                # 簡單替換換行為 br 以利 HTML 顯示
                ai_html = res['AI摘要'].replace('\n', '<br>')
                st.markdown(f"<div class='ai-box'>{ai_html}</div>", unsafe_allow_html=True)
                
            st.markdown("**📸 抓取到之照片：**")
            cols = st.columns(min(len(res['照片清單']), 4))
            for i, img_url in enumerate(res['照片清單']):
                with cols[i % 4]:
                    st.image(img_url, use_container_width=True)
            st.markdown("---")

# ==========================================
# 🏷️ TAB 2: 檢視與下載彙整區
# ==========================================
with tab_view:
    st.markdown("<div class='morandi-dark-title'>📥 依題目檢視與下載</div>", unsafe_allow_html=True)
    
    # 這裡未來會從《AI新聞資料庫》讀取選項
    # 目前使用模擬選單
    mock_options = ["請選擇...", "1.1 - 校園節能減碳設施建置 (模擬資料)"]
    sel_q = st.selectbox("選擇題目", mock_options, label_visibility="collapsed")
    
    if sel_q != "請選擇...":
        # 載入模擬的單題資料
        res = mock_scrape_and_ai_process(None, None)[0]
        
        st.markdown("---")
        st.markdown(f"<h2 style='color:#154360;'>{res['題號']} {res['中文標題']}</h2>", unsafe_allow_html=True)
        
        # 1. 網頁版照片表格 (比照要求)
        if res['照片清單']:
            table_html = "<table style='width:100%; table-layout:fixed; border-collapse:collapse; border:1px solid #D9E0E3; background-color:white; margin-bottom: 20px;'>"
            urls = res['照片清單']
            for i in range(0, len(urls), 2):
                table_html += "<tr>"
                # 第一張
                table_html += f"<td style='border:1px solid #D9E0E3; padding:10px; text-align:center; width:50%; vertical-align:middle;'><img src='{urls[i]}' style='width:100%; max-height:300px; object-fit:contain;'></td>"
                # 第二張
                if i + 1 < len(urls):
                    table_html += f"<td style='border:1px solid #D9E0E3; padding:10px; text-align:center; width:50%; vertical-align:middle;'><img src='{urls[i+1]}' style='width:100%; max-height:300px; object-fit:contain;'></td>"
                else:
                    table_html += "<td style='border:1px solid #D9E0E3; width:50%;'></td>"
                table_html += "</tr>"
            table_html += "</table>"
            st.markdown(table_html, unsafe_allow_html=True)
            
        # 2. 濃縮精華新聞摘要
        st.markdown(f"#### 📰 新聞標題：{res['新聞標題']}")
        st.markdown(f"<div style='font-size:1.2em; line-height:1.8; color:#2C3E50; background-color:#F8FAFB; padding:20px; border-radius:8px;'>{res['AI摘要'].replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)
        
        # 3. 連結
        st.markdown(f"<br>**Additional evidence link:** [{res['新聞連結']}]({res['新聞連結']})", unsafe_allow_html=True)
        
        st.markdown("---")
        with st.spinner("⏳ 產製 Word 檔中..."):
            docx_bytes = generate_ai_word_report(res['題號'], res['中文標題'], [res])
            
        col1, col2, col3 = st.columns([3, 4, 3])
        with col2:
            st.download_button(
                label="📥 下載本題 AI 彙整報告 (Word檔)",
                data=docx_bytes,
                file_name=f"{res['題號']}_AI新聞彙整.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )