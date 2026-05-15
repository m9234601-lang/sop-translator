import streamlit as st
import pandas as pd
import google.generativeai as genai
import os
from dotenv import load_dotenv
import sys
import types

# 1. 파이썬 최신 버전 호환성 패치 (cgi 모듈 에러 방지)
if "cgi" not in sys.modules:
    sys.modules["cgi"] = types.ModuleType("cgi")

# 2. 환경 변수 로드 (로설 실행용)
load_dotenv()

# 3. 페이지 설정
st.set_page_config(page_title="제조 SOP 양식 유지 번역기", layout="wide")

# 4. 사이드바 - API 키 설정
with st.sidebar:
    st.title("⚙️ 설정")
    # Streamlit Secrets 또는 직접 입력에서 키를 가져옴
    api_key = st.text_input("Google API Key를 입력하세요", 
                            value=st.secrets.get("GOOGLE_API_KEY", ""), 
                            type="password")
    
    if api_key:
        genai.configure(api_key=api_key)
        st.success("✅ API 키 등록 완료")
    else:
        st.error("❌ API 키 미등록 (번역 불가)")

# 5. 메인 화면
st.title("📄 제조 SOP 양식 유지 번역기")
st.info("엑셀 파일을 업로드하면 양식을 유지한 채 내용만 번역합니다.")

uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file and api_key:
    try:
        # 데이터 로드
        df = pd.read_excel(uploaded_file)
        
        st.subheader("업로드된 데이터 미리보기")
        st.dataframe(df.head())

        if st.button("번역 시작"):
            with st.spinner("Gemini AI가 번역 중입니다..."):
                # 번역 모델 설정
                model = genai.GenerativeModel('gemini-pro')
                
                def translate_text(text):
                    if pd.isna(text) or str(text).strip() == "":
                        return text
                    try:
                        prompt = f"Translate the following manufacturing SOP text into Korean, keeping the professional tone: {text}"
                        response = model.generate_content(prompt)
                        return response.text
                    except Exception as e:
                        return f"Error: {str(e)}"

                # 모든 셀에 대해 번역 수행 (양식 유지를 위해 전체 적용)
                # 실제 데이터 열이 있는 범위만 지정하거나 전체 적용
                translated_df = df.copy()
                for col in translated_df.columns:
                    translated_df[col] = translated_df[col].apply(translate_text)

                st.success("✅ 번역 완료!")
                st.subheader("번역 결과")
                st.dataframe(translated_df.head())

                # 다운로드 버튼
                output = pd.ExcelWriter("translated_sop.xlsx", engine="openpyxl")
                translated_df.to_excel(output, index=False)
                output.close()
                
                with open("translated_sop.xlsx", "rb") as f:
                    st.download_button(
                        label="번역된 엑셀 다운로드",
                        data=f,
                        file_name="translated_sop.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    except Exception as e:
        st.error(f"파일 처리 중 오류 발생: {e}")
        st.info("팁: 엑셀 파일의 인덱스나 열 형식이 올바른지 확인해 주세요.")

elif not api_key:
    st.warning("왼쪽 사이드바에 API 키를 먼저 등록해 주세요.")
