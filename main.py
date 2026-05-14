import streamlit as st
import os
import time
from dotenv import load_dotenv
from openpyxl import load_workbook  # 양식 유지를 위한 라이브러리
from googletrans import Translator  # 실제 번역을 위한 라이브러리

# 1. 환경 변수 및 번역기 설정
load_dotenv()
api_key = os.getenv("API_KEY")
translator = Translator()
lang_map = {"영어": "en", "중국어(간체)": "zh-cn", "베트남어": "vi", "일본어": "ja"}

# 2. 페이지 설정
st.set_page_config(page_title="SOP 전문 번역기", layout="wide")
st.title("🏭 제조 SOP 양식 유지 번역기")

# 사이드바 설정
st.sidebar.header("🛠️ 시스템 상태")
st.sidebar.success("✅ API 로드 완료") if api_key else st.sidebar.error("❌ API 키 미등록")

st.sidebar.header("🌐 번역 설정")
target_lang_nm = st.sidebar.selectbox("목표 언어 선택", list(lang_map.keys()))
target_lang_code = lang_map[target_lang_nm]

uploaded_file = st.file_uploader("원본 SOP 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    # 양식 보존을 위해 openpyxl로 워크북 로드
    wb = load_workbook(uploaded_file)
    ws = wb.active # 첫 번째 시트 사용
    
    # 번역할 셀 수집 (글자가 있는 셀만)
    target_cells = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                target_cells.append(cell)
    
    st.info(f"📋 총 {len(target_cells)}개의 텍스트 셀이 감지되었습니다.")

    if st.button("양식 유지 번역 시작", key="start_tr"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # 실제 번역 진행
        for i, cell in enumerate(target_cells):
            try:
                # 구글 번역 호출
                translated = translator.translate(cell.value, dest=target_lang_code).text
                cell.value = translated # 원본 셀 값만 교체 (양식은 유지됨)
                
                # 진행률 업데이트
                curr = (i + 1) / len(target_cells)
                progress_bar.progress(curr)
                status_text.text(f"⏳ 번역 중... ({i+1}/{len(target_cells)})")
                
                time.sleep(0.1) # 안정적인 통신을 위한 미세 지연
            except Exception as e:
                st.warning(f"⚠️ {i+1}번 셀 번역 중 오류 발생: {e}")

        # 결과 저장 (메모리 방식)
        from io import BytesIO
        output = BytesIO()
        wb.save(output)
        
        st.success(f"✅ {target_lang_nm} 번역 완료! 원본 양식이 유지되었습니다.")
        
        st.download_button(
            label="📥 번역된 SOP 다운로드 (Excel)",
            data=output.getvalue(),
            file_name=f"Translated_SOP_{target_lang_nm}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )