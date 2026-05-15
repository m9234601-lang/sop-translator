import streamlit as st
import openpyxl
import google.generativeai as genai
import sys
import types
import io
from copy import copy

# 1. 파이썬 버전 호환성 패치
if "cgi" not in sys.modules:
    sys.modules["cgi"] = types.ModuleType("cgi")

# 2. 페이지 설정
st.set_page_config(page_title="제조 SOP 양식 유지 번역기", layout="wide")

# 3. 사이드바 설정
with st.sidebar:
    st.title("⚙️ 설정")
    # Secrets 우선 확인, 없으면 입력창 표시
    api_key = st.secrets.get("GOOGLE_API_KEY", "")
    if not api_key:
        api_key = st.text_input("Google API Key를 입력하세요", type="password")
    
    if api_key:
        genai.configure(api_key=api_key)
        st.success("✅ API 키 등록 완료")
    else:
        st.error("❌ API 키 미등록")

    st.divider()
    st.header("🌐 번역 설정")
    target_lang = st.selectbox(
        "목표 언어 선택",
        ["한국어", "영어", "중국어", "일본어", "베트남어"],
        index=0
    )

# 4. 메인 화면
st.title("📄 제조 SOP 양식 유지 번역기")
st.info(f"엑셀의 **글꼴, 크기, 색상, 병합**을 그대로 유지하며 내용을 **{target_lang}**로 번역합니다.")

uploaded_file = st.file_uploader("SOP 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file and api_key:
    if st.button(f"{target_lang}로 번역 시작"):
        try:
            # 엑셀 로드
            wb = openpyxl.load_workbook(uploaded_file)
            model = genai.GenerativeModel('gemini-1.5-flash') # 속도가 빠른 모델로 변경
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for sheet in wb.worksheets:
                # 데이터가 있는 셀만 추출하여 번역 효율성 높임
                cells_to_translate = [cell for row in sheet.iter_rows() for cell in row if cell.value and isinstance(cell.value, str)]
                total_cells = len(cells_to_translate)
                
                for i, cell in enumerate(cells_to_translate):
                    status_text.text(f"번역 중: {cell.coordinate} 셀 처리 중... ({i+1}/{total_cells})")
                    
                    try:
                        # 전문적인 제조 기술 용어 유지를 위한 프롬프트
                        prompt = f"Translate the following manufacturing SOP text into {target_lang}. Keep technical terms if they are standard in the industry. Output only the translated text. Text: {cell.value}"
                        response = model.generate_content(prompt)
                        
                        if response and response.text:
                            # 기존 스타일(글꼴, 색상 등)을 유지하며 값만 변경
                            cell.value = response.text.strip()
                    except Exception as e:
                        # 에러 발생 시 로그만 찍고 원문 유지
                        print(f"Error at {cell.coordinate}: {e}")
                    
                    # 진행률 업데이트
                    progress_bar.progress((i + 1) / total_cells)

            # 파일 저장
            output = io.BytesIO()
            wb.save(output)
            processed_data = output.getvalue()

            st.success("✅ 번역이 완료되었습니다!")
            st.download_button(
                label="📂 번역된 엑셀 다운로드",
                data=processed_data,
                file_name=f"translated_{target_lang}_SOP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"오류 발생: {e}")

elif not api_key:
    st.warning("왼쪽 사이드바에 API 키를 등록해 주세요.")
