import streamlit as st
import os
import time
import sys
import types
from io import BytesIO
from openpyxl import load_workbook
from googletrans import Translator

# --- [서버 에러 해결을 위한 필수 패치] ---
# 파이썬 3.13+에서 사라진 cgi 모듈을 가짜로 만들어 에러를 방지합니다.
if "cgi" not in sys.modules:
    sys.modules["cgi"] = types.ModuleType("cgi")
# ---------------------------------------

# 1. 페이지 설정
st.set_page_config(page_title="SOP 전문 번역기", layout="wide")
st.title("🏭 제조 SOP 양식 유지 번역기")

# 2. API 키 및 번역기 설정
# Secrets에 등록된 키가 있으면 가져오고, 없으면 로컬 환경 변수를 확인합니다.
api_key = st.secrets.get("GOOGLE_API_KEY", os.getenv("API_KEY"))
translator = Translator()
lang_map = {"영어": "en", "중국어(간체)": "zh-cn", "베트남어": "vi", "일본어": "ja", "한국어": "ko"}

# 사이드바 설정
st.sidebar.header("🛠️ 시스템 상태")
if api_key:
    st.sidebar.success("✅ 시스템 연결 완료")
else:
    st.sidebar.warning("ℹ️ 기본 번역 모드 동작 중")

st.sidebar.header("🌐 번역 설정")
target_lang_nm = st.sidebar.selectbox("목표 언어 선택", list(lang_map.keys()))
target_lang_code = lang_map[target_lang_nm]

uploaded_file = st.file_uploader("원본 SOP 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    wb = load_workbook(uploaded_file)
    ws = wb.active 
    
    target_cells = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and len(str(cell.value).strip()) > 0:
                target_cells.append(cell)
    
    st.info(f"📋 총 {len(target_cells)}개의 텍스트 셀이 감지되었습니다.")

    if st.button("양식 유지 번역 시작", key="start_tr"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, cell in enumerate(target_cells):
            try:
                # 번역 실행
                translated = translator.translate(cell.value, dest=target_lang_code).text
                cell.value = translated 
                
                # 진행률 표시
                curr = (i + 1) / len(target_cells)
                progress_bar.progress(curr)
                status_text.text(f"⏳ 번역 중... ({i+1}/{len(target_cells)})")
                
                time.sleep(0.2) # 서버 차단 방지용 지연
            except Exception as e:
                # 오류 시 건너뛰고 계속 진행
                continue

        output = BytesIO()
        wb.save(output)
        
        st.success(f"✅ {target_lang_nm} 번역 완료!")
        st.download_button(
            label="📥 번역된 SOP 다운로드 (Excel)",
            data=output.getvalue(),
            file_name=f"Translated_SOP_{target_lang_nm}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
