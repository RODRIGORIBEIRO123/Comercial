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
PLANILHA = "DB_Propostas_Siarcon"

# --- CONEXÃO E FUNÇÕES ÚTEIS ---
@st.cache_resource
def conectar():
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
        return gspread.authorize(creds).open(PLANILHA)
    except Exception as e:
        st.error(f"Erro Conexão: {e}"); st.stop()

def carregar_abas(lista_abas):
    sh = conectar()
    dados = {}
    for aba in lista_abas:
        try: dados[aba] = pd.DataFrame(sh.worksheet(aba).get_all_records())
        except: dados[aba] = pd.DataFrame()
    return dados

def salvar(aba, lista_valores):
    conectar().worksheet(aba).append_row(lista_valores)
    st.cache_data.clear(); st.rerun()

# Função que cria os formulários de "Cadastrar Novo" automaticamente
def ui_cadastro(titulo, aba, labels_campos):
    with st.expander(f"➕ {titulo}"):
        with st.form(f"form_{aba}"):
            cols = st.columns(len(labels_campos))
            valores = [cols[i].text_input(label) for i, label in enumerate(labels_campos)]
            if st.form_submit_button("💾 Salvar"):
                if all(valores): salvar(aba, valores)
                else: st.warning("Preencha tudo!")

# --- CARREGAMENTO ---
dfs = carregar_abas(["Escopos", "Exclusoes", "Responsabilidades", "Clientes", "Coberturas"])

# --- 1. CLIENTE E PROJETO ---
st.header("1. Cliente e Projeto")
ui_cadastro("Novo Cliente", "Clientes", ["Empresa", "Contato", "Fone", "Email", "Cidade/Estado"])

lista_cli = dfs["Clientes"]['Empresa'].tolist() if not dfs["Clientes"].empty else []
cli_sel = st.selectbox("Cliente:", ["Novo..."] + lista_cli)

# Preenchimento Automático
p_dados = ["", "", "", "", ""]
if cli_sel != "Novo...":
    p_dados = dfs["Clientes"][dfs["Clientes"]['Empresa'] == cli_sel].iloc[0].values.tolist()

c1, c2 = st.columns(2)
with c1:
    contato = st.text_input("Contato", value=p_dados[1])
    fone = st.text_input("Fone", value=str(p_dados[2]))
    email = st.text_input("Email", value=p_dados[3])
with c2:
    cliente = st.text_input("Empresa", value=p_dados[0])
    local = st.text_input("Local", value=p_dados[4])
    projeto = st.text_input("Nome do Projeto")
    num_prop = st.text_input("Nº Proposta", value=f"P-{date.today().year}-XXX")

# --- 2. COBERTURA ---
st.header("2. Cobertura")
ui_cadastro("Nova Cobertura", "Coberturas", ["Texto Completo"])
lista_cob = dfs["Coberturas"]['Texto_Completo'].tolist() if not dfs["Coberturas"].empty else ["Texto Padrão..."]
txt_cob = st.text_area("Texto:", value=st.selectbox("Modelo:", lista_cob), height=100)
docs = st.text_area("Docs Referência:") if st.checkbox("Incluir Docs?", value=True) else ""

# --- 3. RESPONSABILIDADES ---
st.header("3. Responsabilidades")
ui_cadastro("Nova Resp.", "Responsabilidades", ["Título Curto", "Texto Completo"])
dict_resp = dict(zip(dfs["Responsabilidades"].iloc[:,0], dfs["Responsabilidades"].iloc[:,1])) if not dfs["Responsabilidades"].empty else {}
sel_resp = st.multiselect("Selecione:", list(dict_resp.keys()), default=list(dict_resp.keys()))
resp_final = [dict_resp[k] for k in sel_resp]

# --- 4. ESCOPO TÉCNICO ---
st.header("4. Escopo Técnico")
intro = st.text_area("Introdução Escopo")
ui_cadastro("Novo Item Escopo", "Escopos", ["Categoria", "Título Curto", "Texto Completo"])

escopo_final = []
if not dfs["Escopos"].empty and 'Categoria' in dfs["Escopos"].columns:
    for cat in sorted(dfs["Escopos"]['Categoria'].unique()):
        with st.expander(f"📁 {cat}", expanded=True):
            df_c = dfs["Escopos"][dfs["Escopos"]['Categoria'] == cat]
            dict_cat = dict(zip(df_c['Titulo_Curto'], df_c['Texto_Completo']))
            
            selecionados = st.multiselect(f"Itens de {cat}:", list(dict_cat.keys()))
            itens_processados = []
            
            if selecionados:
                for item in selecionados:
                    c_qtd, c_comp = st.columns([0.2, 0.8])
                    qtd = c_qtd.number_input(f"Qtd", 1, key=f"q_{item}")
                    comp = c_comp.text_input(f"Comp. {item}", key=f"c_{item}")
                    
                    # Monta Texto: TextoBanco — Qtd: X. Complemento.
                    txt = dict_cat[item]
                    extras = [f"Qtd: {qtd}"] if qtd > 1 else []
                    if comp: extras.append(comp)
                    if extras: txt += f" — {'. '.join(extras)}."
                    itens_processados.append(txt)
                
                escopo_final.append({'nome': cat.upper(), 'itens': itens_processados})

# --- 5. EXCLUSÕES ---
st.header("5. Exclusões")
ui_cadastro("Nova Exclusão", "Exclusoes", ["Título Curto", "Texto Completo"])
dict_exc = dict(zip(dfs["Exclusoes"].iloc[:,0], dfs["Exclusoes"].iloc[:,1])) if not dfs["Exclusoes"].empty else {}
exc_final = [dict_exc[k] for k in st.multiselect("Itens:", list(dict_exc.keys()), default=list(dict_exc.keys()))]

# --- 6. COMERCIAL E GERAÇÃO ---
st.header("6. Comercial")
c1, c2 = st.columns(2)
valor = c1.text_input("Valor R$")
mes = c2.text_input("Mês/Ano", value=f"{date.today().month}/{date.today().year}")

st.divider()
if st.button("🚀 GERAR PROPOSTA", type="primary"):
    meses = {1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'}
    dt_hoje = date.today()
    
    ctx = {
        'data_formatada': f"Limeira, {dt_hoje.day} de {meses[dt_hoje.month]} de {dt_hoje.year}",
        'nome_contato': contato, 'fone': fone, 'email': email,
        'nome_cliente': cliente, 'nome_projeto': projeto, 'cidade_estado': local, 'numero_proposta': num_prop,
        'texto_cobertura': txt_cob, 'tem_docs': bool(docs), 'docs_referencia': docs,
        'lista_resp_cliente': resp_final, 'escopo_estruturado': escopo_final,
        'lista_exclusoes': exc_final, 'intro_servico': intro,
        'mes_base': mes, 'valor_total': valor, 'revisao': "R-00"
    }

    try:
        doc = DocxTemplate("Template_Siarcon.docx"); doc.render(ctx)
        bio = io.BytesIO(); doc.save(bio); bio.seek(0)
        st.download_button("📥 Baixar Word", bio, f"Proposta_{num_prop}.docx")
    except Exception as e: st.error(f"Erro: {e}")
