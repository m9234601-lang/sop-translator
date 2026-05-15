import streamlit as st
import os
import time
from dotenv import load_dotenv
from openpyxl import load_workbook
import google.generativeai as genai
from io import BytesIO

# 1. 환경 변수 및 Gemini 설정
load_dotenv()
api_key = os.getenv("GOOGLE_API_KEY")

if api_key:
    genai.configure(api_key=api_key)
    # 404 에러 방지를 위해 가장 표준적인 모델 명칭 사용
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
    except Exception as e:
        st.error(f"모델 설정 오류: {e}")
else:
    st.sidebar.error("❌ GOOGLE_API_KEY가 설정되지 않았습니다.")

lang_map = {
    "영어": "English", 
    "중국어(간체)": "Simplified Chinese", 
    "베트남어": "Vietnamese", 
    "일본어": "Japanese"
}

# 2. 페이지 레이아웃
st.set_page_config(page_title="제조 SOP 양식 보존 번역기", layout="wide")
st.title("📑 제조 SOP 양식 완벽 보존 번역기 (Gemini)")

target_lang_nm = st.sidebar.selectbox("목표 언어 선택", list(lang_map.keys()))
target_lang_en = lang_map[target_lang_nm]

uploaded_file = st.file_uploader("원본 SOP 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file and api_key:
    # 양식 보존을 위해 data_only=False (수식/서식 유지)
    wb = load_workbook(uploaded_file, data_only=False)
    ws = wb.active 
    
    # 번역할 셀 수집
    target_cells = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and str(cell.value).strip():
                target_cells.append(cell)
    
    st.success(f"✅ 총 {len(target_cells)}개의 텍스트 셀을 발견했습니다.")

    if st.button("양식 보존 번역 시작"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, cell in enumerate(target_cells):
            try:
                # 제조 전문성 강화를 위한 프롬프트
                prompt = f"Translate the following manufacturing SOP text into {target_lang_en}. Keep technical terms professional: {cell.value}"
                
                # 가장 안정적인 응답 추출 방식
                response = model.generate_content(prompt)
                
                if response and response.text:
                    cell.value = response.text.strip()
                
                # 진행률 표시
                curr = (i + 1) / len(target_cells)
                progress_bar.progress(curr)
                status_text.text(f"⏳ 번역 중... ({i+1}/{len(target_cells)})")
                
                # 무료 API 속도 제한(RPM) 준수를 위한 지연
                time.sleep(1.2) 
                
            except Exception as e:
                st.warning(f"⚠️ {cell.coordinate} 셀 오류: {e}")
                time.sleep(2.0)

        # 엑셀 파일 생성
        output = BytesIO()
        wb.save(output)
        
        st.balloons()
        st.success("🎉 모든 번역이 완료되었습니다!")
        
        st.download_button(
            label="📥 번역된 SOP 엑셀 다운로드",
            data=output.getvalue(),
            file_name=f"Translated_SOP_{target_lang_nm}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
