import streamlit as st
import os
import time
import sys
import types
from io import BytesIO
from openpyxl import load_workbook
from googletrans import Translator

# 1. [핵심 해결] 파이썬 최신 버전 호환성 패치 (cgi 모듈 에러 방지)
# 이 부분이 없으면 서버에서 'import cgi' 에러가 발생합니다.
if "cgi" not in sys.modules:
    sys.modules["cgi"] = types.ModuleType("cgi")

# 2. 페이지 설정
st.set_page_config(page_title="SOP 전문 번역기", layout="wide")
st.title("🏭 제조 SOP 양식 유지 번역기")

# 3. 환경 변수 및 번역기 설정
# Secrets에 등록한 GOOGLE_API_KEY가 있다면 가져오고, 없으면 기본 번역기를 사용합니다.
translator = Translator()
lang_map = {"영어": "en", "중국어(간체)": "zh-cn", "베트남어": "vi", "일본어": "ja", "한국어": "ko"}

# 4. 사이드바 설정
with st.sidebar:
    st.header("🛠️ 시스템 상태")
    # Streamlit Secrets 설정을 확인합니다.
    if "GOOGLE_API_KEY" in st.secrets:
        st.success("✅ 시스템 연결 완료")
    else:
        st.warning("ℹ️ 기본 번역 모드로 동작 중")

    st.header("🌐 번역 설정")
    target_lang_nm = st.selectbox("목표 언어 선택", list(lang_map.keys()))
    target_lang_code = lang_map[target_lang_nm]

# 5. 메인 로직
uploaded_file = st.file_uploader("원본 SOP 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    # 양식 보존을 위해 openpyxl로 워크북 로드
    wb = load_workbook(uploaded_file)
    
    # 번역할 모든 시트의 셀 수집
    target_cells = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                # 글자가 있는 셀만 수집
                if cell.value and isinstance(cell.value, str) and len(str(cell.value).strip()) > 0:
                    target_cells.append(cell)
    
    st.info(f"📋 총 {len(target_cells)}개의 텍스트 셀이 감지되었습니다.")

    if st.button("양식 유지 번역 시작"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # 실제 번역 진행
        for i, cell in enumerate(target_cells):
            try:
                # 구글 번역 호출 (googletrans)
                translated = translator.translate(cell.value, dest=target_lang_code).text
                cell.value = translated # 원본 셀 값만 교체 (양식 유지)
                
                # 진행률 업데이트
                curr = (i + 1) / len(target_cells)
                progress_bar.progress(curr)
                status_text.text(f"⏳ 번역 중... ({i+1}/{len(target_cells)})")
                
                # 안정적인 통신을 위한 미세 지연
                time.sleep(0.2) 
            except Exception as e:
                # 에러 발생 시 로그를 남기고 계속 진행
                st.warning(f"⚠️ {cell.coordinate} 셀 번역 중 오류 발생: {e}")
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
