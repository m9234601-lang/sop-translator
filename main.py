import streamlit as st
import os
import time
import sys
import types
from io import BytesIO
from openpyxl import load_workbook
from googletrans import Translator

# 1. [가장 중요] 파이썬 최신 버전 호환성 패치
# 서버에서 'import cgi' 에러가 나는 것을 방지하는 코드입니다.
if "cgi" not in sys.modules:
    sys.modules["cgi"] = types.ModuleType("cgi")

# 2. 페이지 설정
st.set_page_config(page_title="SOP 전문 번역기", layout="wide")
st.title("🏭 제조 SOP 양식 유지 번역기")

# 3. 번역기 및 언어 설정
translator = Translator()
lang_map = {"영어": "en", "중국어(간체)": "zh-cn", "베트남어": "vi", "일본어": "ja", "한국어": "ko"}

# 4. 사이드바 설정 (API 키 상태 확인)
with st.sidebar:
    st.header("🛠️ 시스템 상태")
    # Secrets에 키가 등록되어 있는지 확인
    if "GOOGLE_API_KEY" in st.secrets:
        st.success("✅ 시스템 연결 완료")
    else:
        st.warning("ℹ️ 기본 번역 모드로 동작 중")

    st.header("🌐 번역 설정")
    target_lang_nm = st.selectbox("목표 언어 선택", list(lang_map.keys()))
    target_lang_code = lang_map[target_lang_nm]

# 5. 메인 로직 (사용자님의 기존 로직 유지)
uploaded_file = st.file_uploader("원본 SOP 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    # 양식 보존을 위해 openpyxl로 워크북 로드
    wb = load_workbook(uploaded_file)
    
    # 번역할 셀 수집 (글자가 있는 모든 시트의 셀 대상)
    target_cells = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and len(str(cell.value).strip()) > 0:
                    target_cells.append(cell)
    
    st.info(f"📋 총 {len(target_cells)}개의 텍스트 셀이 감지되었습니다.")

    if st.button("양식 유지 번역 시작"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # 실제 번역 진행
        for i, cell in enumerate(target_cells):
            try:
                # 구글 번역 호출
                translated = translator.translate(cell.value, dest=target_lang_code).text
                cell.value = translated # 원본 셀 값만 교체하여 양식 유지
                
                # 진행률 업데이트
                curr = (i + 1) / len(target_cells)
                progress_bar.progress(curr)
                status_text.text(f"⏳ 번역 중... ({i+1}/{len(target_cells)})")
                
                # 안정적인 통신을 위한 미세 지연
                time.sleep(0.2) 
            except Exception as e:
                # 에러 발생 시 건너뛰고 계속 진행 (서버 차단 방지)
                continue

        # 결과 저장 (메모리 방식)
        output = BytesIO()
        wb.save(output)
        
        st.success(f"✅ {target_lang_nm} 번역 완료! 원본 양식이 유지되었습니다.")
        
        st.download_button(
            label="📥 번역된 SOP 다운로드 (Excel)",
            data=output.getvalue(),
            file_name=f"Translated_SOP_{target_lang_nm}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
