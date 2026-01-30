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

# --- CARREGAR DADOS (AGORA COM CLIENTES E COBERTURAS) ---
def carregar_dados():
    sh = conectar_google_sheets()
    
    # Função auxiliar para ler aba com segurança
    def ler_aba(nome):
        try:
            return pd.DataFrame(sh.worksheet(nome).get_all_records())
        except:
            return pd.DataFrame() # Retorna vazio se a aba não existir ainda

    df_esc = ler_aba("Escopos")
    df_exc = ler_aba("Exclusoes")
    df_resp = ler_aba("Responsabilidades")
    df_cli = ler_aba("Clientes")    # NOVA
    df_cob = ler_aba("Coberturas")  # NOVA
    
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
    st.error("Erro ao ler dados. Verifique se criou as abas 'Clientes' e 'Coberturas' na planilha.")
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

# --- CADASTRO DE NOVO CLIENTE ---
with st.expander("➕ Cadastrar NOVO Cliente no Banco"):
    with st.form("form_cliente"):
        c_emp = st.text_input("Nome da Empresa")
        c_cont = st.text_input("Nome do Contato")
        c_fone = st.text_input("Telefone")
        c_email = st.text_input("Email")
        c_cid = st.text_input("Cidade/Estado")
        
        if st.form_submit_button("💾 Salvar Cliente"):
            if c_emp:
                # Ordem: Empresa, Nome_Contato, Telefone, Email, Cidade_Estado
                salvar_no_banco("Clientes", [c_emp, c_cont, c_fone, c_email, c_cid])
                st.rerun()
            else:
                st.warning("Nome da empresa é obrigatório.")

# --- SELEÇÃO DE CLIENTE ---
lista_clientes = df_clientes['Empresa'].tolist() if not df_clientes.empty else []
cliente_selecionado = st.selectbox("Selecione um Cliente Existente:", ["Novo / Digitar Manualmente"] + lista_clientes)

# Variáveis padrão (vazias)
p_empresa = ""
p_contato = ""
p_fone = ""
p_email = ""
p_cidade = ""

# Se selecionou alguém do banco, preenche as variáveis
if cliente_selecionado != "Novo / Digitar Manualmente":
    dados_cli = df_clientes[df_clientes['Empresa'] == cliente_selecionado].iloc[0]
    p_empresa = dados_cli['Empresa']
    p_contato = dados_cli['Nome_Contato']
    p_fone = str(dados_cli['Telefone'])
    p_email = dados_cli['Email']
    p_cidade = dados_cli['Cidade_Estado']

# CAMPOS DE TEXTO (Preenchidos automaticamente ou manual)
c1, c2 = st.columns(2)
hoje = date.today()
data_txt = formatar_data_portugues(hoje)

with c1:
    st.info(f"📅 {data_txt}")
    # O valor padrão (value) vem do banco se selecionado
    nome_contato = st.text_input("Nome do Contato", value=p_contato)
    fone = st.text_input("Telefone", value=p_fone)
    email = st.text_input("Email", value=p_email)

with c2:
    nome_cliente = st.text_input("Empresa (Cliente)", value=p_empresa)
    cidade_estado = st.text_input("Cidade/Estado", value=p_cidade)
    nome_projeto = st.text_input("Nome do Projeto (Ref.)")
    num_prop = st.text_input("Nº Proposta", value=f"P-{hoje.year}-XXX")

# ==============================================================================
# 2. COBERTURA DO FORNECIMENTO
# ==============================================================================
st.markdown("---")
st.header("2. Cobertura")

# --- CADASTRO DE NOVA COBERTURA ---
with st.expander("➕ Cadastrar NOVA Cobertura Padrão"):
    with st.form("form_cob"):
        nova_cob_txt = st.text_area("Texto da Cobertura")
        if st.form_submit_button("💾 Salvar Cobertura"):
            if nova_cob_txt:
                salvar_no_banco("Coberturas", [nova_cob_txt])
                st.rerun()

# Seleção
lista_cob = df_coberturas['Texto_Completo'].tolist() if not df_coberturas.empty else []
# Adiciona uma opção padrão caso o banco esteja vazio
if not lista_cob:
    lista_cob = ["Os custos aqui apresentados compreendem: instalação com fornecimento de equipamentos, materiais e mão-de-obra..."]

texto_escolhido = st.selectbox("Escolha o Modelo de Texto:", lista_cob)
texto_cob_final = st.text_area("Texto Final (Editável):", value=texto_escolhido, height=100)

tem_docs = st.checkbox("Incluir Documentos de Referência?", value=True)
lista_docs = st.text_area("Lista de Documentos:") if tem_docs else ""

# ==============================================================================
# 3. RESPONSABILIDADES DO CLIENTE
# ==============================================================================
st.markdown("---")
st.header("3. Responsabilidades do Cliente")

with st.expander("➕ Cadastrar Responsabilidade"):
    with st.form("nova_resp"):
        nr_curto = st.text_input("Título Curto")
        nr_longo = st.text_input("Texto Completo")
        if st.form_submit_button("💾 Salvar"):
            salvar_no_banco("Responsabilidades", [nr_curto, nr_longo])
            st.rerun()

dict_resp = dict(zip(df_resp['Titulo_Curto'], df_resp['Texto_Completo'])) if not df_resp.empty else {}
sel_resp = st.multiselect("Selecione:", list(dict_resp.keys()), default=list(dict_resp.keys()))
resp_final = [dict_resp[k] for k in sel_resp if k in dict_resp]

# ==============================================================================
# 4. ESCOPO TÉCNICO
# ==============================================================================
st.markdown("---")
st.header("4. Escopo Técnico")
intro = st.text_area("Introdução do Escopo")

with st.expander("➕ Cadastrar Item de Escopo"):
    with st.form("novo_esc"):
        cats = df_escopos['Categoria'].unique().tolist() if not df_escopos.empty else []
        c_cat, c_tit, c_txt = st.columns([0.3, 0.3, 0.4])
        
        # Opção de Nova Categoria
        cat_sel = c_cat.selectbox("Categoria", ["Nova..."] + cats)
        if cat_sel == "Nova...":
            cat_final = c_cat.text_input("Nome da Nova Categoria")
        else:
            cat_final = cat_sel
            
        ne_tit = c_tit.text_input("Nome Equipamento")
        ne_txt = c_txt.text_input("Descrição Técnica")
        
        if st.form_submit_button("💾 Salvar"):
            if cat_final and ne_tit and ne_txt:
                salvar_no_banco("Escopos", [cat_final, ne_tit, ne_txt])
                st.rerun()

escopo_final = []
contador = 1
st.subheader("Seleção por Disciplina")

if not df_escopos.empty:
    for cat in df_escopos['Categoria'].unique():
        if st.checkbox(f"📁 {cat}", key=cat):
            df_c = df_escopos[df_escopos['Categoria'] == cat]
            itens = st.multiselect(f"Itens de {cat}", df_c['Titulo_Curto'].tolist(), key=f"sel_{cat}")
            
            lista_textos = []
            if itens:
                st.caption(f"Detalhamento - {cat}:")
                for item in itens:
                    col_q, col_c = st.columns([0.2, 0.8])
                    qtd = col_q.number_input(f"Qtd ({item})", min_value=1, value=1, key=f"q_{cat}_{item}")
                    comp = col_c.text_input(f"Complemento ({item})", key=f"c_{cat}_{item}")
                    
                    texto = f"Fornecimento de {int(qtd)} {item}"
                    if comp: texto += f", {comp}"
                    texto += "."
                    lista_textos.append(texto)
                
                escopo_final.append({'indice': f"1.{contador}", 'nome': cat.upper(), 'itens': lista_textos})
                contador += 1
                st.markdown("---")

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
        st.success("✅ Proposta Gerada!")
        st.download_button("📥 Baixar Arquivo", bio, f"Proposta_{num_prop}.docx")
    except Exception as e:
        st.error(f"Erro ao gerar: {e}")
