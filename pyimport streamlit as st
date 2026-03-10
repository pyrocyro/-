import streamlit as st
from pptx import Presentation
from pypinyin import pinyin, Style
import io

# 網頁設定
st.set_page_config(page_title="團隊專用-注音字體校對器", layout="wide")

st.title("🛡️ 源泉圓體注音校對工具 (團隊版)")
st.markdown("""
本工具專門抓出 **PPTX** 中「源泉圓體」可能出錯的破音字。
1. 在左側上傳從 Canva 匯出的 **PPTX**。
2. 根據下方的提示，回 Canva 修正對應頁面的文字。
""")

# 側邊欄上傳
st.sidebar.header("檔案上傳")
uploaded_file = st.sidebar.file_uploader("選擇 PPTX 檔案", type="pptx")

def get_errors(text):
    if not text.strip(): return []
    # heteronym=True 會列出所有可能的讀音
    res = pinyin(text, style=Style.BOPOMOFO, heteronym=True)
    return [{"char": c, "suggested": p[0], "all": "/".join(p)} for c, p in zip(text, res) if len(p) > 1]

if uploaded_file:
    prs = Presentation(io.BytesIO(uploaded_file.read()))
    
    # 建立頁面導覽
    tabs = st.tabs([f"第 {i+1} 頁" for i in range(len(prs.slides))])
    
    for i, slide in enumerate(prs.slides):
        with tabs[i]:
            # 抓取該頁所有文字
            text_run = []
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text_run.append(shape.text)
            
            full_text = " ".join(text_run)
            errors = get_errors(full_text)
            
            if errors:
                st.error(f"📍 本頁偵測到 {len(errors)} 個潛在錯誤")
                # 顯示對照表格
                st.table(errors)
                st.info("💡 請回 Canva 檢查上述文字，並確認注音是否正確。")
            else:
                st.success("✅ 這一頁沒有發現破音字！")
else:
    st.info("👈 請先在左側上傳檔案以開始分析。")

st.sidebar.markdown("---")
st.sidebar.write("製作：你的開發助理")