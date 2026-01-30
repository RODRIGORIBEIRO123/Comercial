import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador Propostas SIARCON", layout="wide", page_icon="📄")
st.title("📄 Gerador de Propostas - SIARCON")

# === NOME EXATO DA PLANILHA ===
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
        try:
            return pd.DataFrame(sh.worksheet(nome).get_all_records())
        except:
            return pd.DataFrame()

    df_esc = ler_aba("Escopos")
    df_exc = ler_aba("Exclusoes")
    df_resp = ler_aba("Responsabilidades")
    df_cli = ler_aba("Clientes")
    df_cob = ler_aba("Coberturas")
    
    return df_esc, df_exc, df_resp, df_cli, df_cob

# --- SALVAR NO BANCO ---
def salvar_no_banco(aba, dados_lista):
    try:
        sh = conectar_google_sheets()
        ws = sh.worksheet(aba)
        ws.append_row(dados_lista)
        st.cache_data.clear()
        st.toast(f"✅ Salvo em {aba} com sucesso!", icon="💾")
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")

# Carrega tudo
try:
    df_escopos, df_exclusoes, df_resp, df_clientes, df_coberturas = carregar_dados()
except Exception as e:
    st.error("Erro ao ler dados. Verifique a planilha.")
    st.stop()

# --- FUNÇÃO DATA PT-BR ---
def formatar_data_portugues(dt):
    meses = {1:'Janeiro', 2:'Fevereiro', 3:'Março', 4:'Abril', 5:'Maio', 6:'Junho',
             7:'Julho', 8:'Agosto', 9:'Setembro', 10:'Outubro', 11:'Novembro', 12:'Dezembro'}
    return f"Limeira, {dt.day} de {meses[dt.month]} de {dt.year}"

# ==============================================================================
# 1. DADOS DO PROJETO E CLIENTE
# ==============================================================================
st.header("1. Cliente e Projeto")

# Cadastro Cliente
with st.expander("➕ Cadastrar NOVO Cliente"):
    with st.form("form_cliente"):
        c_emp = st.text_input("Nome da Empresa")
        c_cont = st.text_input("Nome do Contato")
        c_fone = st.text_input("Telefone")
        c_email = st.text_input("Email")
        c_cid = st.text_input("Cidade/Estado")
        if st.form_submit_button("💾 Salvar Cliente"):
            if c_emp:
                salvar_no_banco("Clientes", [c_emp, c_cont, c_fone, c_email, c_cid])
                st.rerun()

# Seleção Cliente
lista_clientes = df_clientes['Empresa'].tolist() if not df_clientes.empty else []
cliente_selecionado = st.selectbox("Selecione Cliente Existente:", ["Novo / Digitar Manualmente"] + lista_clientes)

p_empresa, p_contato, p_fone, p_email, p_cidade = "", "", "", "", ""
if cliente_selecionado != "Novo / Digitar Manualmente":
    dados_cli = df_clientes[df_clientes['Empresa'] == cliente_selecionado].iloc[0]
    p_empresa = dados_cli['Empresa']
    p_contato = dados_cli['Nome_Contato']
    p_fone = str(dados_cli['Telefone'])
    p_email = dados_cli['Email']
    p_cidade = dados_cli['Cidade_Estado']

c1, c2 = st.columns(2)
hoje = date.today()
data_txt = formatar_data_portugues(hoje)

with c1:
    st.info(f"📅 {data_txt}")
    nome_contato = st.text_input("Nome do Contato", value=p_contato)
    fone = st.text_input("Telefone", value=p_fone)
    email = st.text_input("Email", value=p_email)

with c2:
    nome_cliente = st.text_input("Empresa (Cliente)", value=p_empresa)
    cidade_estado = st.text_input("Cidade/Estado", value=p_cidade)
    nome_projeto = st.text_input("Nome do Projeto")
    num_prop = st.text_input("Nº Proposta", value=f"P-{hoje.year}-XXX")

# ==============================================================================
# 2. COBERTURA
# ==============================================================================
st.markdown("---")
st.header("2. Cobertura")

# Cadastro Cobertura
with st.expander("➕ Cadastrar NOVA Cobertura"):
    with st.form("form_cob"):
        nova_cob_txt = st.text_area("Texto da Cobertura")
        if st.form_submit_button("💾 Salvar"):
            if nova_cob_txt:
                salvar_no_banco("Coberturas", [nova_cob_txt])
                st.rerun()

lista_cob = df_coberturas['Texto_Completo'].tolist() if not df_coberturas.empty else ["Os custos aqui apresentados compreendem..."]
texto_escolhido = st.selectbox("Modelo de Texto:", lista_cob)
texto_cob_final = st.text_area("Texto Final:", value=texto_escolhido, height=100)

tem_docs = st.checkbox("Incluir Documentos de Referência?", value=True)
lista_docs = st.text_area("Lista de Documentos:") if tem_docs else ""

# ==============================================================================
# 3. RESPONSABILIDADES DO CLIENTE (SIMPLIFICADO)
# ==============================================================================
st.markdown("---")
st.header("3. Responsabilidades do Cliente")

with st.expander("➕ Cadastrar Responsabilidade"):
    with st.form("nova_resp"):
        nr_curto = st.text_input("Título Curto (Menu)")
        nr_longo = st.text_input("Texto Completo (Proposta)")
        if st.form_submit_button("💾 Salvar"):
            if nr_curto and nr_longo:
                # Agora salva apenas 2 campos (igual às outras abas)
                salvar_no_banco("Responsabilidades", [nr_curto, nr_longo])
                st.rerun()

dict_resp = dict(zip(df_resp['Titulo_Curto'], df_resp['Texto_Completo'])) if not df_resp.empty else {}
sel_resp = st.multiselect("Selecione:", list(dict_resp.keys()), default=list(dict_resp.keys()))
resp_final = [dict_resp[k] for k in sel_resp if k in dict_resp]

# ==============================================================================
# 4. ESCOPO TÉCNICO (SIMPLIFICADO - SEM CATEGORIA)
# ==============================================================================
st.markdown("---")
st.header("4. Escopo Técnico")
intro = st.text_area("Introdução do Escopo")

# --- CADASTRO SIMPLIFICADO ---
with st.expander("➕ Cadastrar NOVO Item de Escopo"):
    with st.form("novo_esc"):
        ne_tit = st.text_input("Título Curto (Para aparecer no Menu)")
        ne_txt = st.text_input("Texto Completo (Para sair na Proposta)")
        if st.form_submit_button("💾 Salvar Item"):
            if ne_tit and ne_txt:
                salvar_no_banco("Escopos", [ne_tit, ne_txt])
                st.rerun()

# --- SELEÇÃO DE ESCOPO ---
dict_escopos = dict(zip(df_escopos['Titulo_Curto'], df_escopos['Texto_Completo'])) if not df_escopos.empty else {}

itens_selecionados = st.multiselect("Selecione os Itens do Escopo:", options=list(dict_escopos.keys()))

escopo_final_word = []

if itens_selecionados:
    st.markdown("### 📝 Detalhamento dos Itens")
    st.info("Ajuste quantidades e complementos. O texto final será: 'Texto do Banco' + Complementos.")
    
    for item_curto in itens_selecionados:
        texto_longo_banco = dict_escopos[item_curto]
        
        st.markdown(f"**Item:** {item_curto}")
        
        c_qtd, c_comp = st.columns([0.2, 0.8])
        qtd = c_qtd.number_input(f"Qtd", min_value=1, value=1, key=f"q_{item_curto}")
        comp = c_comp.text_input(f"Complemento", placeholder="Marca, Modelo...", key=f"c_{item_curto}")
        
        # Montagem do texto
        texto_para_doc = texto_longo_banco
        
        detalhes_extras = []
        if qtd > 1:
            detalhes_extras.append(f"Quantidade: {qtd}")
        if comp:
            detalhes_extras.append(comp)
            
        if detalhes_extras:
            texto_para_doc += f" — {', '.join(detalhes_extras)}."
            
        escopo_final_word.append(texto_para_doc)
        st.divider()

# ==============================================================================
# 5. EXCLUSÕES
# ==============================================================================
st.markdown("---")
st.header("5. Exclusões")

with st.expander("➕ Cadastrar Exclusão"):
    with st.form("nova_exc"):
        nex_c = st.text_input("Título Curto")
        nex_l = st.text_input("Texto Completo")
        if st.form_submit_button("💾 Salvar"):
            if nex_c:
                salvar_no_banco("Exclusoes", [nex_c, nex_l])
                st.rerun()

dict_exc = dict(zip(df_exclusoes['Titulo_Curto'], df_exclusoes['Texto_Completo'])) if not df_exclusoes.empty else {}
sel_exc = st.multiselect("Exclusões:", list(dict_exc.keys()), default=list(dict_exc.keys()))
exc_final = [dict_exc[k] for k in sel_exc if k in dict_exc]

# ==============================================================================
# 6. COMERCIAL
# ==============================================================================
st.markdown("---")
st.header("6. Comercial")
c_v, c_m = st.columns(2)
valor = c_v.text_input("Valor Total (R$)")
mes = c_m.text_input("Mês/Ano Base", value=f"{hoje.month}/{hoje.year}")

# ==============================================================================
# BOTÃO GERAR
# ==============================================================================
st.markdown("---")
if st.button("🚀 GERAR PROPOSTA (.DOCX)", type="primary"):
    
    contexto = {
        'data_formatada': data_txt,
        'nome_contato': nome_contato, 'fone': fone, 'email': email,
        'nome_cliente': nome_cliente, 'nome_projeto': nome_projeto, 'cidade_estado': cidade_estado,
        'numero_proposta': num_prop,
        'texto_cobertura': texto_cob_final,
        'tem_docs': tem_docs, 'docs_referencia': lista_docs,
        'lista_resp_cliente': resp_final,
        'escopo_simples': escopo_final_word, 
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
        st.success("✅ Proposta Gerada!")
        st.download_button("📥 Baixar Arquivo", bio, f"Proposta_{num_prop}.docx")
    except Exception as e:
        st.error(f"Erro ao gerar: {e}")
