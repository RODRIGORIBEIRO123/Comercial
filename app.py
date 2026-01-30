import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Gerador Propostas SIARCON", layout="wide", page_icon="📄")
st.title("📄 Gerador de Propostas - SIARCON")

# Nome EXATO da sua planilha no Google
PLANILHA_NOME = "DB_Propostas_Siarcon"

# --- CONEXÃO SEGURA ---
@st.cache_resource
def conectar_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # Pega as credenciais do arquivo secrets.toml
    creds_dict = st.secrets["gcp_service_account"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client.open(PLANILHA_NOME)

def carregar_dados():
    sh = conectar_google_sheets()
    # Pega todos os dados como lista de dicionários
    d_esc = sh.worksheet("Escopos").get_all_records()
    d_exc = sh.worksheet("Exclusoes").get_all_records()
    d_resp = sh.worksheet("Responsabilidades").get_all_records()
    return pd.DataFrame(d_esc), pd.DataFrame(d_exc), pd.DataFrame(d_resp)

def salvar_no_banco(aba, dados_lista):
    sh = conectar_google_sheets()
    ws = sh.worksheet(aba)
    ws.append_row(dados_lista)
    st.cache_data.clear()
    st.toast(f"✅ Salvo em {aba}!")

# Carrega Dados Iniciais
try:
    df_escopos, df_exclusoes, df_resp = carregar_dados()
except Exception as e:
    st.error(f"Erro de Conexão: {e}")
    st.stop()

# --- INTERFACE ---

# 1. DADOS DO PROJETO
st.header("1. Dados do Projeto")
c1, c2 = st.columns(2)
meses = {1:'Janeiro', 2:'Fevereiro', 3:'Março', 4:'Abril', 5:'Maio', 6:'Junho',
         7:'Julho', 8:'Agosto', 9:'Setembro', 10:'Outubro', 11:'Novembro', 12:'Dezembro'}
hoje = date.today()
data_txt = f"Limeira, {hoje.day} de {meses[hoje.month]} de {hoje.year}"

with c1:
    st.info(f"📅 {data_txt}")
    nome_contato = st.text_input("Contato")
    fone = st.text_input("Telefone")
    email = st.text_input("Email")
with c2:
    cliente = st.text_input("Cliente")
    projeto = st.text_input("Nome do Projeto")
    local = st.text_input("Cidade/Estado")
    num_prop = st.text_input("Nº Proposta", value=f"P-{hoje.year}-XXX")

# 2. COBERTURA
st.markdown("---")
st.header("2. Cobertura")
texto_cob = st.text_area("Texto de Cobertura", 
    value="Os custos aqui apresentados compreendem: instalação com fornecimento de equipamentos, materiais e mão-de-obra...", 
    height=80)
tem_docs = st.checkbox("Incluir Documentos de Referência?", value=True)
lista_docs = st.text_area("Lista de Docs:") if tem_docs else ""

# 3. RESPONSABILIDADES (LEITURA E ESCRITA)
st.markdown("---")
st.header("3. Responsabilidades do Cliente")

with st.expander("➕ Cadastrar Nova Responsabilidade"):
    with st.form("nova_resp"):
        nr_curto = st.text_input("Título Curto")
        nr_longo = st.text_input("Texto Completo")
        if st.form_submit_button("Salvar"):
            salvar_no_banco("Responsabilidades", [nr_curto, nr_longo])
            st.rerun()

dict_resp = dict(zip(df_resp['Titulo_Curto'], df_resp['Texto_Completo']))
sel_resp = st.multiselect("Selecione:", list(dict_resp.keys()), default=list(dict_resp.keys()))
resp_final = [dict_resp[k] for k in sel_resp if k in dict_resp]

# 4. ESCOPO TÉCNICO (LEITURA E ESCRITA)
st.markdown("---")
st.header("4. Escopo Técnico")
intro = st.text_area("Introdução do Escopo")

with st.expander("➕ Cadastrar Novo Item de Escopo"):
    with st.form("novo_esc"):
        cats = df_escopos['Categoria'].unique().tolist()
        ne_cat = st.selectbox("Categoria", cats) # Poderia ser text_input para nova
        ne_tit = st.text_input("Nome Equipamento")
        ne_txt = st.text_input("Descrição Técnica")
        if st.form_submit_button("Salvar"):
            salvar_no_banco("Escopos", [ne_cat, ne_tit, ne_txt])
            st.rerun()

escopo_final = []
contador = 1
for cat in df_escopos['Categoria'].unique():
    if st.checkbox(f"📁 {cat}", key=cat):
        df_c = df_escopos[df_escopos['Categoria'] == cat]
        itens = st.multiselect(f"Itens de {cat}", df_c['Titulo_Curto'].tolist())
        
        lista_textos = []
        if itens:
            st.caption("Detalhamento:")
            for item in itens:
                col_q, col_c = st.columns([0.2, 0.8])
                qtd = col_q.number_input(f"Qtd {item}", 1, key=f"q{item}")
                comp = col_c.text_input(f"Complemento {item}", key=f"c{item}")
                texto = f"Fornecimento de {int(qtd)} {item}"
                if comp: texto += f", {comp}"
                texto += "."
                lista_textos.append(texto)
            
            escopo_final.append({'indice': f"1.{contador}", 'nome': cat.upper(), 'itens': lista_textos})
            contador += 1

# 5. EXCLUSÕES (LEITURA E ESCRITA)
st.markdown("---")
st.header("5. Exclusões")
with st.expander("➕ Cadastrar Nova Exclusão"):
    with st.form("nova_exc"):
        nex_c = st.text_input("Título")
        nex_l = st.text_input("Texto")
        if st.form_submit_button("Salvar"):
            salvar_no_banco("Exclusoes", [nex_c, nex_l])
            st.rerun()

dict_exc = dict(zip(df_exclusoes['Titulo_Curto'], df_exclusoes['Texto_Completo']))
sel_exc = st.multiselect("Exclusões:", list(dict_exc.keys()), default=list(dict_exc.keys()))
exc_final = [dict_exc[k] for k in sel_exc if k in dict_exc]

# 6. COMERCIAL
st.markdown("---")
st.header("6. Comercial")
c_v, c_m = st.columns(2)
valor = c_v.text_input("Valor Total")
mes = c_m.text_input("Mês/Ano", value=f"{hoje.month}/{hoje.year}")

# GERAR
st.markdown("---")
if st.button("🚀 GERAR PROPOSTA", type="primary"):
    contexto = {
        'data_formatada': data_txt,
        'nome_contato': nome_contato, 'fone': fone, 'email': email,
        'nome_cliente': cliente, 'nome_projeto': projeto, 'cidade_estado': local,
        'numero_proposta': num_prop,
        'texto_cobertura': texto_cob,
        'tem_docs': tem_docs, 'docs_referencia': lista_docs,
        'lista_resp_cliente': resp_final,
        'escopo_estruturado': escopo_final,
        'lista_exclusoes': exc_final,
        'intro_servico': intro,
        'mes_base': mes, 'valor_total': valor,
        'revisao': "R-00"
    }
    try:
        doc = DocxTemplate("Template_Siarcon.docx")
        doc.render(contexto)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success("Sucesso!")
        st.download_button("Baixar", bio, f"Proposta_{num_prop}.docx")
    except Exception as e:
        st.error(f"Erro: {e}")
