import streamlit as st

st.set_page_config(page_title="Novo Projeto SIARCON")

st.title("🚀 Novo Projeto Iniciado")
st.write("O ambiente está configurado e rodando!")

# Um teste de interação simples
nome = st.text_input("Qual o nome deste módulo?")
if nome:
    st.success(f"O módulo {nome} foi criado com sucesso.")
