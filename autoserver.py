import streamlit as st
import pandas as pd

st.title("데이터 불러오기")


pdf_file = 'asdsad'

with open(pdf_file, 'rb') as f:
        pdf_bytes = f.read()
    
st.download_button(label="수료증 다운로드", data=pdf_bytes, file_name=pdf_file)