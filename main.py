import streamlit as st
import google.generativeai as genai
from openpyxl import load_workbook
import io
import time

# 1. 페이지 설정
st.set_page_config(page_title="제조 SOP 양식 유지 번역기", layout="wide")
st.title("🏭 제조 SOP 양식 유지 번역기 (Gemini AI)")

# 2. 사이드바 설정 (API 키 및 언어 선택)
with st.sidebar:
    st.header("⚙️ 설정")
    # Secrets에 등록된 키를 자동으로 불러옵니다.
    api_key = st.secrets.get("GOOGLE_API_KEY", "")
    
    if api_key:
        genai.configure(api_key=api_key)
        st.success("✅ 시스템 연결 완료")
    else:
        st.error("❌ API 키 미등록 (Secrets 설정 필요)")
    
    st.divider()
    target_lang = st.selectbox("목표 언어 선택", ["영어", "한국어", "중국어", "베트남어", "일본어"])

# 3. 메인 로직 (사용자님의 기존 양식 유지 로직 적용)
uploaded_file = st.file_uploader("원본 SOP 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file and api_key:
    if st.button("양식 유지 번역 시작"):
        try:
            # 엑셀 로드 (양식 보존)
            wb = load_workbook(uploaded_file)
            model = genai.GenerativeModel('gemini-1.5-flash') # 속도가 빠른 모델 사용
            
            # 모든 시트에서 텍스트가 있는 셀만 수집
            all_cells = []
            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and len(str(cell.value).strip()) > 0:
                            all_cells.append(cell)
            
            progress_bar = st.progress(0)
            status_text = st.empty()

            # 실제 번역 진행 (값만 교체)
            for i, cell in enumerate(all_cells):
                status_text.text(f"⏳ 번역 중: {i+1}/{len(all_cells)} 셀 처리 중...")
                try:
                    # 전문 번역 프롬프트
                    prompt = f"Translate the following manufacturing SOP text into {target_lang}. Keep technical terms professional. Result only text: {cell.value}"
                    response = model.generate_content(prompt)
                    if response and response.text:
                        cell.value = response.text.strip()
                except:
                    continue # 오류 시 원문 유지
                
                progress_bar.progress((i + 1) / len(all_cells))
                time.sleep(0.1) # 안정적 호출을 위한 지연

            # 결과물 저장
            output = io.BytesIO()
            wb.save(output)
            
            st.success("🎉 모든 번역과 양식 보존 작업이 완료되었습니다!")
            st.download_button(
                label="📥 번역된 파일 다운로드",
                data=output.getvalue(),
                file_name=f"translated_SOP_{target_lang}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"오류 발생: {e}")
