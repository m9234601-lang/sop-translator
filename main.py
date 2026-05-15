import streamlit as st
import openpyxl
import google.generativeai as genai
import sys
import types
import io

# 1. 최신 파이썬 호환성 패치
if "cgi" not in sys.modules:
    sys.modules["cgi"] = types.ModuleType("cgi")

# 2. 페이지 설정
st.set_page_config(page_title="제조 SOP 양식 유지 번역기", layout="wide")

# 3. 사이드바 설정
with st.sidebar:
    st.title("⚙️ 설정")
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
st.info(f"엑셀의 **테두리, 색상, 병합 상태**를 그대로 유지하며 내용만 **{target_lang}**로 번역합니다.")

uploaded_file = st.file_uploader("SOP 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file and api_key:
    if st.button(f"{target_lang}로 양식 유지 번역 시작"):
        try:
            # 엑셀 로드 (data_only=False로 설정해야 수식/양식이 유지됨)
            wb = openpyxl.load_workbook(uploaded_file)
            model = genai.GenerativeModel('gemini-pro')
            
            # 번역 진행 상황 표시
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # 모든 시트 순회
            for sheet in wb.worksheets:
                total_cells = sheet.max_row * sheet.max_column
                processed_count = 0
                
                for row in sheet.iter_rows():
                    for cell in row:
                        processed_count += 1
                        # 텍스트가 있는 셀만 번역
                        if cell.value and isinstance(cell.value, str):
                            status_text.text(f"번역 중: {cell.coordinate} 셀 처리 중...")
                            try:
                                prompt = f"Translate this manufacturing SOP text into {target_lang}. Keep technical terms if necessary and maintain a professional tone. Text: {cell.value}"
                                response = model.generate_content(prompt)
                                cell.value = response.text
                            except:
                                pass # 에러 시 원문 유지
                        
                        # 진행바 업데이트 (너무 자주 하면 느려지므로 적절히 조절)
                        if processed_count % 10 == 0:
                            progress_bar.progress(processed_count / total_cells)

            # 결과물 저장
            output = io.BytesIO()
            wb.save(output)
            processed_data = output.getvalue()

            st.success("✅ 번역이 완료되었습니다!")
            
            st.download_button(
                label="📂 번역된 엑셀 다운로드 (양식 유지됨)",
                data=processed_data,
                file_name=f"translated_{target_lang}_SOP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"오류 발생: {e}")
            st.info("엑셀 파일이 암호화되어 있거나 손상되었는지 확인해 주세요.")

elif not api_key:
    st.warning("왼쪽 사이드바에 API 키를 등록해 주세요.")
