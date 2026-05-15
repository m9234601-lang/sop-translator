import streamlit as st
import os
import time
from dotenv import load_dotenv
from openpyxl import load_workbook
import google.generativeai as genai
from io import BytesIO
import copy

# 1. 초기 설정
load_dotenv()
api_key = os.getenv("GOOGLE_API_KEY")

if api_key:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('models/gemini-1.5-flash')
else:
    st.sidebar.error("❌ GOOGLE_API_KEY를 설정해주세요.")

lang_map = {"영어": "English", "중국어(간체)": "Simplified Chinese", "베트남어": "Vietnamese", "일본어": "Japanese"}

st.set_page_config(page_title="SOP 양식 유지 번역기", layout="wide")
st.title("📑 제조 SOP 양식 완벽 보존 번역기")

target_lang = st.sidebar.selectbox("목표 언어", list(lang_map.keys()))
uploaded_file = st.file_uploader("SOP 엑셀 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file and api_key:
    # data_only=False가 핵심입니다. 수식과 양식을 그대로 가져옵니다.
    wb = load_workbook(uploaded_file, data_only=False)
    ws = wb.active 
    
    # 번역할 대상 추출 (데이터가 있는 셀만)
    target_cells = []
    for row in ws.iter_rows():
        for cell in row:
            # 문자열이고, 공백이 아닌 경우에만 번역 대상으로 선정
            if cell.value and isinstance(cell.value, str) and str(cell.value).strip():
                # 스타일 보존을 위해 셀 객체 자체를 저장
                target_cells.append(cell)
    
    st.info(f"✅ 양식 유지 모드: 총 {len(target_cells)}개의 텍스트 셀을 발견했습니다.")

    if st.button("양식 보존 번역 시작"):
        progress_bar = st.progress(0)
        
        for i, cell in enumerate(target_cells):
            try:
                # 제조 용어 보존을 위한 프롬프트 강화
                prompt = f"""
                You are a professional manufacturing engineer. 
                Translate the following SOP text into {lang_map[target_lang]}.
                - Maintain technical terms (e.g., Purging, Injection, Shot).
                - Keep the tone professional and concise.
                Text: {cell.value}
                """
                
                response = model.generate_content(prompt)
                translated = response.text.strip()
                
                # 핵심: .value만 변경하면 openpyxl이 기존 cell.font, cell.border, cell.fill 등을 그대로 유지합니다.
                cell.value = translated
                
                # 진행도 표시
                curr = (i + 1) / len(target_cells)
                progress_bar.progress(curr)
                time.sleep(0.5) # API 안정성을 위한 간격
                
            except Exception as e:
                st.error(f"Error at {cell.coordinate}: {e}")
                continue

        # 결과 저장
        output = BytesIO()
        wb.save(output)
        processed_data = output.getvalue()
        
        st.success("🎉 양식 훼손 없이 번역이 완료되었습니다!")
        st.download_button(
            label="📥 번역된 엑셀 다운로드",
            data=processed_data,
            file_name=f"Fixed_Format_{target_lang}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
