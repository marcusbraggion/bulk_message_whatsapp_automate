import streamlit as st
import pandas as pd


uploaded_file = st.file_uploader("Upload a file")
if uploaded_file is not None:
    
    df = pd.read_excel(uploaded_file, dtype={"TELEFONE": str})
    
    st.write(df)


