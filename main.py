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

# 2. 페이지 설정
st.set_page_config(page_title="제조 SOP 양식 유지 번역기", layout="wide")

# 3. 사이드바 설정
with st.sidebar:
    st.title("⚙️ 설정")
    
    # API 키 확인 (Secrets에서 가져오거나 직접 입력)
    api_key = st.secrets.get("GOOGLE_API_KEY", "")
    if not api_key:
        api_key = st.text_input("Google API Key를 입력하세요", type="password")
    
    if api_key:
        genai.configure(api_key=api_key)
        st.success("✅ API 키 등록 완료")
    else:
        st.error("❌ API 키 미등록 (번역 불가)")

    st.divider()
    
    # [언어 선택 기능 추가]
    st.header("🌐 번역 설정")
    target_lang = st.selectbox(
        "목표 언어 선택",
        ["한국어", "영어", "중국어", "일본어", "베트남어"],
        index=0
    )

# 4. 메인 화면
st.title("📄 제조 SOP 양식 유지 번역기")
st.info(f"엑셀 파일을 업로드하면 양식을 유지한 채 **{target_lang}**로 번역합니다.")

uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file and api_key:
    try:
        # 데이터 로드
        df = pd.read_excel(uploaded_file)
        
        st.subheader("업로드된 데이터 미리보기")
        st.dataframe(df.head())

        if st.button(f"{target_lang}로 번역 시작"):
            with st.spinner(f"Gemini AI가 {target_lang}로 번역 중입니다..."):
                # 번역 모델 설정
                model = genai.GenerativeModel('gemini-pro')
                
                def translate_text(text):
                    if pd.isna(text) or str(text).strip() == "" or isinstance(text, (int, float)):
                        return text
                    try:
                        # 프롬프트에 선택한 언어 반영
                        prompt = f"Translate the following manufacturing SOP text into {target_lang}, keeping the professional tone and formatting: {text}"
                        response = model.generate_content(prompt)
                        return response.text
                    except Exception:
                        return text # 에러 발생 시 원문 유지

                # 전체 데이터 번역 수행
                translated_df = df.copy()
                for col in translated_df.columns:
                    translated_df[col] = translated_df[col].apply(translate_text)

                st.success("✅ 번역 완료!")
                st.subheader("번역 결과 확인")
                st.dataframe(translated_df)

                # 다운로드 버튼 생성
                import io
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    translated_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📂 번역된 엑셀 파일 다운로드",
                    data=buffer.getvalue(),
                    file_name=f"translated_sop_{target_lang}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"파일 처리 중 오류 발생: {e}")

elif not api_key:
    st.warning("왼쪽 사이드바에 API 키를 등록하거나 입력해 주세요.")
