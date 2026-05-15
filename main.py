import streamlit as st
import openpyxl
import google.generativeai as genai
import sys
import types
import io

# 1. 시스템 호환성 패치
if "cgi" not in sys.modules:
    sys.modules["cgi"] = types.ModuleType("cgi")

# 2. 페이지 설정
st.set_page_config(page_title="제조 SOP 양식 유지 번역기", layout="wide")

# 3. 사이드바 설정 (Secrets 연동 강화)
with st.sidebar:
    st.title("⚙️ 설정")
    # Secrets에서 키를 가져오되, 없으면 입력창 표시
    api_key = st.secrets.get("GOOGLE_API_KEY", "")
    if not api_key:
        api_key = st.text_input("Google API Key를 입력하세요", type="password")
    
    if api_key:
        genai.configure(api_key=api_key)
        st.success("✅ API 키 등록 확인됨")
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
st.info(f"엑셀의 모든 양식을 보존하며 내용을 **{target_lang}**로 번역합니다.")

uploaded_file = st.file_uploader("SOP 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file and api_key:
    if st.button(f"{target_lang}로 번역 시작"):
        try:
            # 엑셀 로드
            wb = openpyxl.load_workbook(uploaded_file)
            # 가장 범용적인 gemini-1.5-flash 모델 사용
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # 번역 대상 셀 수집 (빈 셀 제외, 문자열만)
            all_worksheets = wb.worksheets
            for sheet in all_worksheets:
                cells = []
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and len(str(cell.value).strip()) > 0:
                            cells.append(cell)
                
                total = len(cells)
                for i, cell in enumerate(cells):
                    status_text.text(f"번역 중: {sheet.title} 시트 - {cell.coordinate} ({i+1}/{total})")
                    
                    try:
                        # 번역 시도 (안전성 필터 완화 및 명확한 지시)
                        prompt = f"You are a professional translator. Translate the following manufacturing SOP text into {target_lang}. Preserve technical terms if appropriate. Provide ONLY the translation without any preamble. Text: {cell.value}"
                        
                        response = model.generate_content(prompt)
                        
                        # 응답 검증 후 셀 값 교체
                        if response and response.candidates:
                            translated_text = response.text.strip()
                            if translated_text:
                                cell.value = translated_text
                    except Exception as e:
                        # 에러 발생 시 로그 출력 (사용자는 계속 진행 가능)
                        st.warning(f"{cell.coordinate} 셀 번역 실패: {str(e)}")
                    
                    progress_bar.progress((i + 1) / total)

            # 파일 저장
            output = io.BytesIO()
            wb.save(output)
            processed_data = output.getvalue()

            st.success("🎉 모든 번역 작업이 완료되었습니다!")
            st.download_button(
                label="📂 번역된 엑셀 파일 받기",
                data=processed_data,
                file_name=f"translated_{target_lang}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"중대 오류 발생: {e}")

elif not api_key:
    st.warning("먼저 API 키를 설정해 주세요.")
