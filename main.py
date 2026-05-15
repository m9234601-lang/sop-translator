import streamlit as st
import google.generativeai as genai
from openpyxl import load_workbook
import io
import time

# 1. 페이지 설정
st.set_page_config(page_title="SOP 양식 유지 번역기", layout="wide")
st.title("🏭 제조 SOP 양식 유지 번역기 (Gemini 엔진)")

# 2. 사이드바 설정
with st.sidebar:
    st.header("⚙️ 설정")
    # Secrets에 등록된 GOOGLE_API_KEY를 가져옵니다.
    api_key = st.secrets.get("GOOGLE_API_KEY", "")
    
    if api_key:
        genai.configure(api_key=api_key)
        st.success("✅ 시스템 정상 작동 중")
    else:
        st.error("❌ API 키를 Secrets에 등록해주세요.")
    
    st.divider()
    target_lang = st.selectbox("목표 언어", ["영어", "한국어", "중국어", "베트남어", "일본어"])

# 3. 번역 로직 (사용자님의 기존 양식 유지 방식 계승)
uploaded_file = st.file_uploader("원본 SOP 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file and api_key:
    if st.button("양식 유지 번역 시작"):
        try:
            wb = load_workbook(uploaded_file)
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            # 텍스트가 있는 모든 셀 수집
            all_cells = []
            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and len(str(cell.value).strip()) > 0:
                            all_cells.append(cell)
            
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, cell in enumerate(all_cells):
                status_text.text(f"⏳ 번역 중: {i+1}/{len(all_cells)} 셀 처리 중...")
                try:
                    # Gemini 번역 요청 (에러 없는 방식)
                    prompt = f"Translate the following manufacturing SOP text into {target_lang}. Keep technical terms professional. Result only text: {cell.value}"
                    response = model.generate_content(prompt)
                    if response.text:
                        cell.value = response.text.strip() # 셀 값만 교체 (양식 유지)
                except:
                    continue 
                
                progress_bar.progress((i + 1) / len(all_cells))
                time.sleep(0.1)

            # 결과물 저장
            output = io.BytesIO()
            wb.save(output)
            
            st.success("🎉 번역이 완료되었습니다!")
            st.download_button(
                label="📥 번역된 파일 다운로드",
                data=output.getvalue(),
                file_name=f"translated_SOP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"오류 발생: {e}")
