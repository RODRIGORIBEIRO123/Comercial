import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador Propostas SIARCON", layout="wide", page_icon="📄")
st.title("📄 Gerador de Propostas Automático - SIARCON")

# === NOME EXATO DA PLANILHA DO BANCO DE DADOS ===
PLANILHA_NOME = "DB_Propostas_Siarcon" 

# --- CONEXÃO SEGURA ---
@st.cache_resource
def conectar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open(PLANILHA_NOME)
    except Exception as e:
        st.error(f"Erro na conexão com Google Sheets: {e}")
        st.stop()

# --- CARREGAR DADOS ---
def carregar_dados():
    sh = conectar_google_sheets()
    def ler_aba(nome):
        try: return pd.DataFrame(sh.worksheet(nome).get_all_records())
        except: return pd.DataFrame()

    return ler_aba("Exclusoes"), ler_aba("Responsabilidades"), ler_aba("Clientes"), ler_aba("Coberturas")

def salvar_no_banco(aba, dados_lista):
    try:
        sh = conectar_google_sheets()
        sh.worksheet(aba).append_row(dados_lista)
        st.cache_data.clear()
        st.toast(f"✅ Salvo em {aba} com sucesso!", icon="💾")
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")

# Carrega Banco de Dados
try:
    df_exclusoes, df_resp, df_clientes, df_coberturas = carregar_dados()
except Exception as e:
    st.error("Erro ao ler banco de dados. Verifique a planilha.")
    st.stop()

def formatar_data_portugues(dt):
    meses = {1:'Janeiro', 2:'Fevereiro', 3:'Março', 4:'Abril', 5:'Maio', 6:'Junho',
             7:'Julho', 8:'Agosto', 9:'Setembro', 10:'Outubro', 11:'Novembro', 12:'Dezembro'}
    return f"Limeira, {dt.day} de {meses[dt.month]} de {dt.year}"

# ==============================================================================
# 1. DADOS DO PROJETO E CLIENTE
# ==============================================================================
st.header("1. Cliente e Projeto")

with st.expander("➕ Cadastrar NOVO Cliente"):
    with st.form("form_cliente"):
        c_emp = st.text_input("Nome da Empresa")
        c_cont, c_fone, c_email, c_cid = st.text_input("Contato"), st.text_input("Telefone"), st.text_input("Email"), st.text_input("Cidade/Estado
