import streamlit as st
import pandas as pd
from whatsapp_bulk_messaging import send_whatsapp


def main():
    st.title("Programa para enviar mensagens via WhatsApp Web")
    uploaded_file = st.file_uploader("Upload a file")
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, dtype={"TELEFONE": str})

    send_whatsapp(df, msg="Olá, Marcus")

if __name__ == "__main__":
    main()
    

