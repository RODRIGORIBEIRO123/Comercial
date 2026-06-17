import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
import math
import re
import uuid
from datetime import date, datetime, timezone, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO DA TELA ---
st.set_page_config(page_title="App SIARCON - Propostas e Custos", layout="wide", page_icon="📄")

def buscar_logo():
    nomes_possiveis = ["SIARCON.png", "SIARCON .png", "siarcon.png", "Siarcon.png", "logo.png"]
    for nome in nomes_possiveis:
        if os.path.exists(nome):
            return nome
    return None

ARQUIVO_LOGO = buscar_logo()

# ==========================================
# 🟢 CONEXÃO COM O GOOGLE SHEETS
# ==========================================
PLANILHA_URL = "https://docs.google.com/spreadsheets/d/1DgBxNqwUepO2RW6GdRwnFHxg7dLlWiRGZjdglkQ8Ls0/edit?gid=1169331401#gid=1169331401"

@st.cache_resource
def conectar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open_by_url(PLANILHA_URL)
    except Exception as e:
        st.error(f"Erro na conexão com Google Sheets: {e}. Verifique o link.")
        st.stop()

fuso_br = timezone(timedelta(hours=-3))

# ==========================================
# 🔐 INICIALIZAÇÃO SEGURA DE VARIÁVEIS GLOBAIS
# ==========================================
if "usuario_logado" not in st.session_state: st.session_state.usuario_logado = None
if "nome_exibicao" not in st.session_state: st.session_state.nome_exibicao = ""
if "menu_selecionado" not in st.session_state: st.session_state.menu_selecionado = "🏠 Tela Inicial"
if "orcamento" not in st.session_state: st.session_state.orcamento = []
if "historico_precos" not in st.session_state: st.session_state.historico_precos = []
if 'nome_projeto_orcamento' not in st.session_state: st.session_state.nome_projeto_orcamento = ""
if 'projeto_para_abrir' not in st.session_state: st.session_state.projeto_para_abrir = None
if 'dados_projeto_abrir' not in st.session_state: st.session_state.dados_projeto_abrir = {}
if 'wizard_ativo' not in st.session_state: st.session_state.wizard_ativo = False
if 'paineis_auto' not in st.session_state: st.session_state.paineis_auto = []
if 'confirmar_limpar' not in st.session_state: st.session_state.confirmar_limpar = False

# ==========================================
# TELA DE LOGIN
# ==========================================
if st.session_state.usuario_logado is None:
    st.markdown("""
        <style>
        .block-container { padding-top: 0rem !important; margin-top: -2rem !important; }
        header {display: none !important;}
        [data-testid="collapsedControl"] {display: none !important;}
        .stApp { background: linear-gradient(135deg, #1C8590 0%, #8FD3B5 100%) !important; }
        [data-testid="stForm"] { background-color: white; border-radius: 12px; padding: 30px; box-shadow: 0 8px 24px rgba(0,0,0,0.15); border: none; position: relative; z-index: 100; }
        [data-testid="stFormSubmitButton"] button { background-color: #2b7bc4 !important; color: white !important; font-weight: bold !important; border-radius: 6px !important; height: 45px !important; border: none !important; margin-top: 15px !important; }
        [data-testid="stFormSubmitButton"] button:hover { background-color: #1a5c96 !important; }
        input { border-bottom: 2px solid #ccc !important; border-top: none !important; border-left: none !important; border-right: none !important; border-radius: 0 !important; background-color: transparent !important; box-shadow: none !important; padding-left: 0px !important; }
        input:focus { border-bottom: 2px solid #1C8590 !important; }
        </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.write("")
        col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
        with col_img2:
            if ARQUIVO_LOGO: st.image(ARQUIVO_LOGO, use_container_width=True)
            else: st.markdown("<h2 style='text-align: center; color: white; margin-bottom:0;'>SIARCON</h2>", unsafe_allow_html=True)
        
        with st.form("form_login"):
            st.markdown("""
                <div style="text-align: center; margin-bottom: 15px;">
                    <div style="width: 60px; height: 60px; background-color: #4A5568; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; margin-bottom: 10px;">
                        <svg viewBox="0 0 24 24" width="35" height="35" fill="white">
                            <path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/>
                        </svg>
                    </div>
                    <h3 style="color: #333; margin: 0; font-size: 18px;">Bem-Vindo a plataforma comercial da SIARCON</h3>
                </div>
            """, unsafe_allow_html=True)
            
            c_user = st.text_input("Usuário:", placeholder="Ex: rodrigo.ribeiro")
            c_pass = st.text_input("Senha:", type="password", placeholder="••••")
            submit_login = st.form_submit_button("Entrar no Sistema", use_container_width=True)
            
            if submit_login:
                usuarios_validos = {
                    "giovanna.ribeiro": "1234", "aline.ferraz": "1234", "janaina.dias": "1234",
                    "victor.hugo": "1234", "rodrigo.ribeiro": "1234", "rodrigo": "1234", "engenharia": "1234",
                    "suprimentos": "1234", "obras": "1234"
                }
                user_limpo = c_user.lower().strip()
                if user_limpo in usuarios_validos and c_pass == usuarios_validos[user_limpo]:
                    st.session_state.usuario_logado = user_limpo
                    st.session_state.nome_exibicao = user_limpo.split('.')[0].capitalize()
                    st.session_state.paineis_auto = []
                    st.session_state.nome_projeto_orcamento = ""
                    st.session_state.wizard_ativo = False
                    st.session_state.menu_selecionado = "🏠 Tela Inicial"
                    st.rerun()
                else: st.error("❌ Usuário ou senha incorretos.")
    st.stop()

st.markdown("""
    <style>
    .block-container { padding-top: 3rem !important; }
    header {display: flex !important;}
    [data-testid="collapsedControl"] {display: flex !important;}
    </style>
""", unsafe_allow_html=True)

# === MENU LATERAL ===
if ARQUIVO_LOGO: st.sidebar.image(ARQUIVO_LOGO, use_container_width=True)
else: st.sidebar.markdown("### SIARCON")
    
st.sidebar.title("Navegação Principal")
st.sidebar.markdown(f"👤 Logado como: **{st.session_state.nome_exibicao}**")
if st.sidebar.button("🚪 Sair do Perfil", type="secondary"):
    st.session_state.usuario_logado = None
    st.session_state.nome_exibicao = ""
    st.session_state.paineis_auto = []
    st.session_state.nome_projeto_orcamento = ""
    st.session_state.menu_selecionado = "🏠 Tela Inicial"
    st.rerun()

st.sidebar.markdown("---")
opcoes_menu = ["🏠 Tela Inicial", "📄 Gerador de Propostas", "🔌 Levantamento de Automação"]
menu_ui = st.sidebar.radio("Módulos do Sistema", opcoes_menu, index=opcoes_menu.index(st.session_state.menu_selecionado))
st.sidebar.markdown("---")

if menu_ui != st.session_state.menu_selecionado:
    st.session_state.menu_selecionado = menu_ui
    st.rerun()

# ==============================================================================
# TELA 0: HOME
# ==============================================================================
if st.session_state.menu_selecionado == "🏠 Tela Inicial":
    st.write("")
    st.markdown(f"<h1 style='text-align: center; color: #178B96;'>Bem-vindo(a), {st.session_state.nome_exibicao}!</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px; color: #666;'>Portal Comercial e de Engenharia SIARCON. Selecione o módulo desejado para iniciar:</p>", unsafe_allow_html=True)
    st.write("")
    st.write("")
    col_vazia_esq, col_card1, col_vazia_meio, col_card2, col_vazia_dir = st.columns([1, 2.5, 0.5, 2.5, 1])
    
    with col_card1:
        st.markdown("""
        <div style='text-align: center; padding: 30px; background: white; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border-top: 5px solid #1C8590;'>
            <h1 style='font-size: 50px; margin-bottom: 10px;'>📄</h1>
            <h3 style='color: #333;'>Gerador de Propostas</h3>
            <p style='color: #666; font-size: 14px; height: 40px;'>Criação rápida e padronizada de escopos técnicos e comerciais em Word.</p>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Acessar Módulo ➔", key="btn_home_prop", type="primary", use_container_width=True):
            st.session_state.menu_selecionado = "📄 Gerador de Propostas"
            st.rerun()

    with col_card2:
        st.markdown("""
        <div style='text-align: center; padding: 30px; background: white; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border-top: 5px solid #1C8590;'>
            <h1 style='font-size: 50px; margin-bottom: 10px;'>🔌</h1>
            <h3 style='color: #333;'>Levantamento de Automação</h3>
            <p style='color: #666; font-size: 14px; height: 40px;'>Dimensionamento estrutural e financeiro de hardware, infraestrutura e supervisório.</p>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Acessar Módulo ➔", key="btn_home_auto", type="primary", use_container_width=True):
            st.session_state.menu_selecionado = "🔌 Levantamento de Automação"
            st.rerun()

# ==============================================================================
# MÓDULO 1: GERADOR DE PROPOSTAS
# ==============================================================================
elif st.session_state.menu_selecionado == "📄 Gerador de Propostas":
    st.sidebar.subheader("Opções da Proposta")
    modo_preenchimento = st.sidebar.radio("Como deseja preencher o Escopo Técnico?", ["📋 Preenchimento Manual", "📊 Automático (Excel)"])
    st.title("📄 Gerador de Propostas - SIARCON")

    def carregar_dados_propostas():
        sh = conectar_google_sheets()
        def ler_aba(nome):
            try: return pd.DataFrame(sh.worksheet(nome).get_all_records())
            except: return pd.DataFrame()
        return ler_aba("Escopos"), ler_aba("Exclusoes"), ler_aba("Responsabilidades"), ler_aba("Clientes"), ler_aba("Coberturas")

    def salvar_no_banco(aba, dados_lista):
        try:
            sh = conectar_google_sheets()
            sh.worksheet(aba).append_row(dados_lista)
            st.cache_data.clear()
            st.toast(f"✅ Salvo em {aba} com sucesso!", icon="💾")
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

    try: df_escopos, df_exclusoes, df_resp, df_clientes, df_coberturas = carregar_dados_propostas()
    except Exception as e:
        st.error("Erro ao ler banco de dados. Verifique a planilha.")
        st.stop()

    def formatar_data_portugues(dt):
        meses = {1:'Janeiro', 2:'Fevereiro', 3:'Março', 4:'Abril', 5:'Maio', 6:'Junho', 7:'Julho', 8:'Agosto', 9:'Setembro', 10:'Outubro', 11:'Novembro', 12:'Dezembro'}
        return f"Limeira, {dt.day} de {meses[dt.month]} de {dt.year}"

    st.header("1. Cliente e Projeto")

    with st.expander("➕ Cadastrar NOVO Cliente"):
        with st.form("form_cliente"):
            c_emp = st.text_input("Nome da Empresa")
            c_cont, c_fone, c_email, c_cid = st.text_input("Contato"), st.text_input("Telefone"), st.text_input("Email"), st.text_input("Cidade/Estado")
            if st.form_submit_button("💾 Salvar Cliente") and c_emp:
                salvar_no_banco("Clientes", [c_emp, c_cont, c_fone, c_email, c_cid])
                st.rerun()

    lista_clientes = df_clientes['Empresa'].tolist() if not df_clientes.empty else []
    cliente_selecionado = st.selectbox("Selecione Cliente Existente:", ["Novo / Digitar Manualmente"] + lista_clientes)

    p_emp, p_cont, p_fone, p_email, p_cid = "", "", "", "", ""
    if cliente_selecionado != "Novo / Digitar Manualmente":
        d_cli = df_clientes[df_clientes['Empresa'] == cliente_selecionado].iloc[0]
        p_emp, p_cont, p_fone, p_email, p_cid = d_cli['Empresa'], d_cli['Nome_Contato'], str(d_cli['Telefone']), d_cli['Email'], d_cli['Cidade_Estado']

    c1, c2 = st.columns(2)
    hoje = date.today()
    data_txt = formatar_data_portugues(hoje)

    with c1:
        st.info(f"📅 {data_txt}")
        nome_contato = st.text_input("Nome do Contato", value=p_cont)
        fone = st.text_input("Telefone", value=p_fone)
        email = st.text_input("Email", value=p_email)

    with c2:
        nome_cliente = st.text_input("Empresa (Cliente)", value=p_emp)
        cidade_estado = st.text_input("Cidade/Estado", value=p_cid)
        nome_projeto = st.text_input("Nome do Projeto")
        num_prop = st.text_input("Nº Proposta", value=f"P-{hoje.year}-XXX")

    st.markdown("---")
    st.header("2. COBERTURA")
    with st.expander("➕ Cadastrar NOVA Cobertura"):
        with st.form("form_cob"):
            nova_cob_txt = st.text_area("Texto da Cobertura")
            if st.form_submit_button("💾 Salvar") and nova_cob_txt:
                salvar_no_banco("Coberturas", [nova_cob_txt])
                st.rerun()

    lista_cob = df_coberturas['Texto_Completo'].tolist() if not df_coberturas.empty else ["Os custos aqui apresentados compreendem..."]
    texto_cob_final = st.text_area("Texto Final:", value=st.selectbox("Modelo de Texto:", lista_cob), height=100)
    tem_docs = st.checkbox("Incluir Documentos de Referência?", value=True)
    lista_docs = st.text_area("Lista de Documentos:") if tem_docs else ""

    st.markdown("---")
    st.header("3. Responsabilidades do Cliente")
    dict_resp = dict(zip(df_resp['Titulo_Curto'], df_resp['Texto_Completo'])) if not df_resp.empty else {}
    sel_resp = st.multiselect("Selecione:", list(dict_resp.keys()), default=list(dict_resp.keys()))
    resp_final = [dict_resp[k] for k in sel_resp if k in dict_resp]

    st.markdown("---")
    intro = st.text_area("Introdução do Escopo", value="Trata-se do fornecimento de materiais e mão de obra conforme itens abaixo:")

    escopo_estruturado = [] 
    eap_estruturada = []    
    valor_total_calculado = 0.0

    if modo_preenchimento == "📋 Preenchimento Manual":
        st.header("4. Escopo Técnico (Modo Manual)")
        if 'Categoria' in df_escopos.columns:
            categorias = sorted(df_escopos['Categoria'].unique())
            contador_cat = 1
            for cat in categorias:
                with st.expander(f"📁 {cat}", expanded=True):
                    df_cat = df_escopos[df_escopos['Categoria'] == cat]
                    dict_cat = dict(zip(df_cat['Titulo_Curto'], df_cat['Texto_Completo']))
                    itens_selecionados = st.multiselect(f"Itens de {cat}:", options=list(dict_cat.keys()), key=f"sel_{cat}")
                    lista_detalhada = []
                    lista_eap = []
                    contador_item = 1
                    if itens_selecionados:
                        for item_curto in itens_selecionados:
                            texto_base = dict_cat[item_curto]
                            col_q, col_nome = st.columns([0.15, 0.85])
                            qtd = col_q.number_input(f"Qtd", min_value=1, value=1, key=f"q_{cat}_{item_curto}", label_visibility="visible")
                            col_nome.write(f"**{item_curto}**")
                            if "AMORTECEDOR" not in item_curto.upper():
                                lista_eap.append({'indice': f"{contador_cat}.{contador_item}", 'nome': item_curto})
                                contador_item += 1 
                            texto_final = texto_base
                            if qtd > 1: texto_final += f" — Qtd: {qtd}."
                            lista_detalhada.append(texto_final)
                        eap_estruturada.append({'indice': str(contador_cat), 'categoria': cat.upper(), 'itens': lista_eap})
                        escopo_estruturado.append({'nome': f"{contador_cat}. {cat.upper()}", 'itens': lista_detalhada})
                        contador_cat += 1
    else:
        st.header("4. Escopo Técnico (Upload da Planilha)")
        arquivo_excel = st.file_uploader("📂 Faça o upload da Planilha Orçamentária (.xlsx)", type=["xlsx"])
        if arquivo_excel is not None:
            try:
                xls = pd.ExcelFile(arquivo_excel)
                nome_aba = st.selectbox("Selecione a Aba (Planilha) onde está o Orçamento:", xls.sheet_names)
                df_orc = pd.read_excel(xls, sheet_name=nome_aba, header=None)
                if df_orc.shape[1] < 14: df_orc = df_orc.reindex(columns=list(range(14)))

                for idx, row in df_orc.iterrows():
                    texto_col_e = str(row[4]).upper().strip() if 4 < len(row) else ""
                    if "PREÇO VENDA" in texto_col_e or "PRECO VENDA" in texto_col_e:
                        val_bruto = row[9] 
                        if pd.notna(val_bruto):
                            if isinstance(val_bruto, (int, float)): valor_total_calculado = float(val_bruto)
                            else:
                                try: valor_total_calculado = float(str(val_bruto).upper().replace("R$", "").replace(".", "").replace(",", ".").strip())
                                except: pass
                        break

                categoria_atual_nome = "ESCOPO GERAL"
                categoria_atual_indice = ""
                itens_detalhados = []
                itens_eap = []
                contador_item = 1
                
                for index, row in df_orc.iterrows():
                    descricao = str(row[2]).strip() if 2 < len(row) else ""
                    if "CUSTO INDIRETO" in descricao.upper(): break 
                    
                    col_b = str(row[1]).strip() if 1 < len(row) else ""
                    unidade = str(row[3]).strip() if 3 < len(row) else ""
                    quantidade = row[4] if 4 < len(row) else pd.NA
                    
                    if pd.isna(descricao) or descricao == "" or descricao.upper() in ["NAN", "DESCRIÇÃO DOS MATERIAIS", "ITEM"]: continue
                        
                    is_header = False
                    if col_b and col_b.lower() != 'nan' and col_b != '-':
                        if col_b[0].isdigit(): is_header = True
                            
                    if is_header:
                        if categoria_atual_nome != "ESCOPO GERAL" and len(itens_detalhados) > 0:
                            escopo_estruturado.append({'nome': f"{categoria_atual_indice} - {categoria_atual_nome}".strip(' -'), 'itens': itens_detalhados})
                            eap_estruturada.append({'indice': categoria_atual_indice, 'categoria': categoria_atual_nome.upper(), 'itens': itens_eap})
                        categoria_atual_indice = col_b
                        categoria_atual_nome = descricao
                        itens_detalhados = []
                        itens_eap = []
                        contador_item = 1
                    else:
                        has_qty = not pd.isna(quantidade) and str(quantidade).strip() not in ["", "nan", "-"]
                        if has_qty:
                            try: qtd_fmt = int(float(str(quantidade).replace(",", "."))) if float(str(quantidade).replace(",", ".")).is_integer() else round(float(str(quantidade).replace(",", ".")), 2)
                            except: qtd_fmt = quantidade
                            uni_fmt = f" {unidade}" if unidade.lower() not in ["nan", "", "-"] else ""
                            nome_resumido = descricao.split('|')[0].strip()[:80] + "..." if len(descricao.split('|')[0].strip()) > 80 else descricao.split('|')[0].strip()
                            
                            if "AMORTECEDOR" not in descricao.upper():
                                itens_eap.append({'indice': f"{categoria_atual_indice}.{contador_item}" if categoria_atual_indice else str(contador_item), 'nome': nome_resumido})
                                contador_item += 1 
                            itens_detalhados.append(f"Fornecimento / Instalação de {qtd_fmt}{uni_fmt} - {descricao}.")
                        else:
                            if len(itens_detalhados) > 0: itens_detalhados[-1] += f"\n\n{descricao}"
                
                if categoria_atual_nome != "ESCOPO GERAL" and len(itens_detalhados) > 0:
                    escopo_estruturado.append({'nome': f"{categoria_atual_indice} - {categoria_atual_nome}".strip(' -'), 'itens': itens_detalhados})
                    eap_estruturada.append({'indice': categoria_atual_indice, 'categoria': categoria_atual_nome.upper(), 'itens': itens_eap})
                if len(escopo_estruturado) > 0: st.success(f"✅ Planilha processada com sucesso.")
            except Exception as e: st.error(f"Erro ao processar: {e}")

    st.markdown("---")
    st.header("5. Exclusões")
    dict_exc = dict(zip(df_exclusoes['Titulo_Curto'], df_exclusoes['Texto_Completo'])) if not df_exclusoes.empty else {}
    sel_exc = st.multiselect("Exclusões:", list(dict_exc.keys()), default=list(dict_exc.keys()))
    exc_final = [dict_exc[k] for k in sel_exc if k in dict_exc]

    st.markdown("---")
    st.header("6. Comercial")
    valor_formatado_sugerido = f"R$ {valor_total_calculado:_.2f}".replace('.', ',').replace('_', '.')
    c_v, c_m = st.columns(2)
    valor = c_v.text_input("Valor Total (R$):", value=valor_formatado_sugerido if valor_total_calculado > 0 else "")
    mes = c_m.text_input("Mês/Ano Base", value=f"{hoje.month}/{hoje.year}")

    st.markdown("---")
    if st.button("🚀 GERAR PROPOSTA (.DOCX)", type="primary"):
        if len(escopo_estruturado) == 0: st.warning("⚠️ O escopo técnico está vazio.")
        contexto = {
            'data_formatada': data_txt, 'nome_contato': nome_contato, 'fone': fone, 'email': email,
            'nome_cliente': nome_cliente, 'nome_projeto': nome_projeto, 'cidade_estado': cidade_estado,
            'numero_proposta': num_prop, 'texto_cobertura': texto_cob_final, 'tem_docs': tem_docs, 
            'docs_referencia': lista_docs, 'lista_resp_cliente': resp_final, 'eap_estruturada': eap_estruturada, 
            'escopo_estruturado': escopo_estruturado, 'lista_exclusoes': exc_final, 'intro_servico': intro,
            'mes_base': mes, 'valor_total': valor, 'revisao': "R-00"
        }
        try:
            doc = DocxTemplate("Template_Siarcon.docx") 
            doc.render(contexto)
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            st.success("✅ Proposta Gerada!")
            st.download_button("📥 Baixar Arquivo Word", bio, f"Proposta_{num_prop}.docx")
        except Exception as e: st.error(f"Erro: {e}")

# ==============================================================================
# MÓDULO 2: LEVANTAMENTO DE AUTOMAÇÃO
# ==============================================================================
elif st.session_state.menu_selecionado == "🔌 Levantamento de Automação":
    st.title("🔌 Engenharia e Custos - Automação e Infra")
    st.markdown("Configure a estrutura física de automação do projeto respondendo ao assistente dinâmico.")
    
    c_proj1, c_proj2 = st.columns([3, 1])
    nome_proj = c_proj1.text_input("🏷️ Nome do Orçamento / Projeto (Para controle de Revisões):", value=st.session_state.nome_projeto_orcamento)
    rev_proj = c_proj2.text_input("Revisão", value="R-00")
    st.session_state.nome_projeto_orcamento = nome_proj
    st.markdown("---")

    REGRA_IO = {
        "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": {"AI": 1, "AO": 1, "DI": 1, "DO": 1},
        "Transmissor de temperatura e umidade para duto (TT/MT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura para duto (TT)": {"AI": 1, "AO": 1, "DI": 0, "DO": 0},
        "Válvula de controle proporcional com atuador (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula de controle de água gelada proporcional (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula de controle de água quente proporcional (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula de controle de vapor proporcional (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Relé de Corrente - Status Compressor (TC)": {"AI": 0, "AO": 0, "DI": 1, "DO": 2},
        "Termostato de segurança (TSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Pressostato diferencial para ar (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Resistência de aquecimento (Equipamento) (RAQ)": {"AI": 0, "AO": 1, "DI": 2, "DO": 1},
        "Resistência de aquecimento (Duto) (RAQ)": {"AI": 0, "AO": 1, "DI": 2, "DO": 1},
        "Válvula motorizada Bypass Proporcional (até 2.1/2\") (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula motorizada Bypass Proporcional (3\" ou 4\") (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula motorizada Bypass Proporcional (5\") (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula motorizada Bypass Proporcional (6\") (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula motorizada Bypass Proporcional (8\") (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Transmissor de pressão para água (PIT)": {"AI": 1, "AO": 2, "DI": 0, "DO": 0},
        "Transmissor de vazão para água (FIT)": {"AI": 1, "AO": 1, "DI": 0, "DO": 0},
        "Válvula bloqueio motorizada (XV)": {"AI": 0, "AO": 0, "DI": 0, "DO": 1},
        "Chave de fluxo (FS)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Bombas (I/O para controlador)": {"AI": 0, "AO": 1, "DI": 1, "DO": 1},
        "Tanques (I/O para controlador)": {"AI": 1, "AO": 0, "DI": 1, "DO": 1},
        "Pressostato para monitorar os filtros G4 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Pressostato para monitorar os filtros F9 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Pressostato para monitorar os filtros H13/H14 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Transmissor de pressão diferencial (monitorar os filtros G4) (PDIT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de pressão diferencial (monitorar os filtros F9) (PDIT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de pressão diferencial (monitorar os filtros H13) (PDIT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de pressão diferencial entre salas (PDT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de pressão diferencial entre salas com display (PDIT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura Ambiente (TT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura ambiente com display (TIT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura e umidade ambiente (TT/MT)": {"AI": 2, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura e umidade ambiente com display (TIT/MIT)": {"AI": 2, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de CO2 ambiente (AT/AIT)": {"AI": 1, "AO": 1, "DI": 0, "DO": 1},
        "Transmissor de temperatura de imersão (TT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura de imersão com display (TIT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0}
    }

    banco_schneider_comum = {
        "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1490.00,
        "Transmissor de temperatura e umidade para duto (TT/MT)": 2050.00,
        "Transmissor de temperatura para duto (TT)": 800.00,
        "Válvula de controle proporcional com atuador (TCV)": 0.00,
        "Válvula de controle de água gelada proporcional (TCV)": 0.00,
        "Válvula de controle de água quente proporcional (TCV)": 0.00,
        "Válvula de controle de vapor proporcional (TCV)": 0.00,
        "Relé de Corrente - Status Compressor (TC)": 150.00,
        "Termostato de segurança (TSH)": 250.00,
        "Pressostato diferencial para ar (PSH)": 349.00,
        "Resistência de aquecimento (Equipamento) (RAQ)": 0.00,
        "Resistência de aquecimento (Duto) (RAQ)": 0.00,
        "Válvula motorizada Bypass Proporcional (até 2.1/2\") (TCV)": 2690.00,
        "Válvula motorizada Bypass Proporcional (3\" ou 4\") (TCV)": 4950.00,
        "Válvula motorizada Bypass Proporcional (5\") (TCV)": 6450.00,
        "Válvula motorizada Bypass Proporcional (6\") (TCV)": 7900.00,
        "Válvula motorizada Bypass Proporcional (8\") (TCV)": 9200.00,
        "Transmissor de pressão para água (PIT)": 1359.00,
        "Transmissor de vazão para água (FIT)": 3550.00,
        "Válvula bloqueio motorizada (XV)": 0.00,
        "Chave de fluxo (FS)": 349.00,
        "Bombas (I/O para controlador)": 0.00,
        "Tanques (I/O para controlador)": 0.00,
        "Pressostato para monitorar os filtros G4 (PSH)": 349.00,
        "Pressostato para monitorar os filtros F9 (PSH)": 349.00,
        "Pressostato para monitorar os filtros H13/H14 (PSH)": 349.00,
        "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": 349.00,
        "Transmissor de pressão diferencial (monitorar os filtros G4) (PDIT)": 1490.00,
        "Transmissor de pressão diferencial (monitorar os filtros F9) (PDIT)": 1490.00,
        "Transmissor de pressão diferencial (monitorar os filtros H13) (PDIT)": 1490.00,
        "Transmissor de pressão diferencial entre salas (PDT)": 1490.00,
        "Transmissor de pressão diferencial entre salas com display (PDIT)": 2110.00,
        "Transmissor de temperatura Ambiente (TT)": 2050.00,
        "Transmissor de temperatura ambiente com display (TIT)": 2650.00,
        "Transmissor de temperatura e umidade ambiente (TT/MT)": 2050.00,
        "Transmissor de temperatura e umidade ambiente com display (TIT/MIT)": 2650.00,
        "Transmissor de CO2 ambiente (AT/AIT)": 0.00,
        "Transmissor de temperatura de imersão (TT)": 0.00,
        "Transmissor de temperatura de imersão com display (TIT)": 0.00,
        "Custo AI/AO": 565.00, "Custo DI/DO": 120.00,
        "Licença Supervisório - SEM CFR-21 (Base)": 23000.00, "Licença Supervisório - SEM CFR-21 (Por Ponto I/O)": 100.00,
        "Licença Supervisório - COM CFR-21 (Base)": 23000.00, "Licença Supervisório - COM CFR-21 (Por Ponto I/O)": 285.00,
        "Licença Supervisório - Schneider EBO (Base)": 13000.00, "Licença Supervisório - Schneider EBO (Por Ponto I/O)": 110.00,
        "MP-C-15A": 4649.49, "MP-C-18A": 5185.54, "MP-C-24A": 7290.75, "MP-C-36A": 9459.08,
        "Schneider - Sensor de Temperatura NTC (Duto)": 120.00,
        "Schneider - Sensor de Temperatura NTC (Ambiente)": 85.00,
        "Schneider - Servidor de Automação (SpaceLogic AS-P/AS-B)": 9500.00
    }

    banco_siemens = {
        "Siemens - CPU 1214C DC/DC/DC": 2500.00, "Siemens - CPU 1215C DC/DC/DC": 3200.00,
        "Siemens - SM 1231 AI 8x13Bit": 1900.00, "Siemens - SM 1232 AQ 4x14Bit": 2100.00,
        "Siemens - SM 1221 DI 16x24VDC": 1200.00, "Siemens - SM 1222 DQ 16x24VDC": 1300.00,
        "Siemens - Fonte 24VDC 2.5A": 800.00, "Siemens - Cartão de Memória 4MB": 400.00,
        "Siemens - CPU 1511-1 PN": 5500.00, "Siemens - AI 8xU/I HS": 2800.00,
        "Siemens - AQ 4xU/I ST": 3100.00, "Siemens - DI 16x24VDC HF": 1800.00,
        "Siemens - DQ 16x24VDC/0.5A": 1900.00, "Siemens - Fonte PM 1507 24VDC 8A": 1500.00,
        "Siemens - Cartão de Memória 12MB": 900.00, "Siemens - Serviço Custo AI/AO": 750.00,
        "Siemens - Serviço Custo DI/DO": 180.00
    }

    banco_mercato = {
        "Mercato - Controlador MCP-4 (2AO, 3DO, 4UI)": 950.00,
        "Mercato - Controlador MCP-8 (4AO, 5DO, 8UI, 2DI)": 1400.00,
        "Mercato - Controlador MCP-12 (4AO, 8DO, 12UI, 4DI)": 1850.00,
        "Mercato - Sensor de Temperatura NTC (Duto)": 120.00,
        "Mercato - Sensor de Temperatura NTC (Ambiente)": 85.00,
        "Mercato - Sensor de Temperatura NTC com Display (Ambiente)": 350.00,
        "Mercato - Serviço Parametrização por Ponto": 80.00,
        "Mercato - IHM Básica 4.3\"": 1700.00
    }
    
    banco_cfr_servicos = {
        "CFR21 Qualificável - Até 100 pts": 70.00,
        "CFR21 Qualificável - 101 a 250 pts": 50.00,
        "CFR21 Qualificável - Acima de 250 pts": 30.00,
        "CFR21 Qualificado - Até 30 pts": 400.00,
        "CFR21 Qualificado - 31 a 60 pts": 350.00,
        "CFR21 Qualificado - 61 a 99 pts": 320.00,
        "CFR21 Qualificado - 100 a 150 pts": 290.00,
        "CFR21 Qualificado - 151 a 200 pts": 250.00,
        "CFR21 Qualificado - 201 a 250 pts": 220.00,
        "CFR21 Qualificado - Acima de 250 pts": 200.00
    }

    banco_ihm = { "IHM Padrão 7\"": 3400.00, "IHM Premium 10\"": 8500.00, "Sem Interface (Cego)": 0.00 }

    banco_padrao_precos = {**banco_schneider_comum, **banco_siemens, **banco_mercato, **banco_ihm, **banco_cfr_servicos}

    if 'precos_banco' not in st.session_state:
        st.session_state.precos_banco = banco_padrao_precos.copy()
        try:
            sh = conectar_google_sheets()
            try:
                aba_p = sh.worksheet("Precos").get_all_values()
                if len(aba_p) > 1:
                    precos_bd = {linha[0]: float(linha[1]) for linha in aba_p[1:] if len(linha) > 1}
                    st.session_state.precos_banco.update(precos_bd)
            except: pass
        except: pass

    for k_n, v_n in banco_padrao_precos.items():
        if k_n not in st.session_state.precos_banco: st.session_state.precos_banco[k_n] = v_n

    GRUPOS_INSTRUMENTOS = {
        "🔹 Controle (HVAC e Máquinas)": [
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)",
            "Transmissor de temperatura e umidade para duto (TT/MT)", "Transmissor de temperatura para duto (TT)",
            "Válvula de controle proporcional com atuador (TCV)", "Válvula de controle de água gelada proporcional (TCV)",
            "Válvula de controle de água quente proporcional (TCV)", "Válvula de controle de vapor proporcional (TCV)",
            "Relé de Corrente - Status Compressor (TC)", "Termostato de segurança (TSH)",
            "Pressostato diferencial para ar (PSH)", "Resistência de aquecimento (Equipamento) (RAQ)", "Resistência de aquecimento (Duto) (RAQ)"
        ],
        "💧 Controle (Central de Água Gelada - CAG)": [
            "Válvula motorizada Bypass Proporcional (até 2.1/2\") (TCV)", "Válvula motorizada Bypass Proporcional (3\" ou 4\") (TCV)",
            "Válvula motorizada Bypass Proporcional (5\") (TCV)", "Válvula motorizada Bypass Proporcional (6\") (TCV)",
            "Válvula motorizada Bypass Proporcional (8\") (TCV)", "Transmissor de pressão para água (PIT)",
            "Transmissor de vazão para água (FIT)", "Transmissor de temperatura de imersão (TT)", "Transmissor de temperatura de imersão com display (TIT)",
            "Válvula bloqueio motorizada (XV)", "Chave de fluxo (FS)", "Bombas (I/O para controlador)", "Tanques (I/O para controlador)"
        ],
        "🔸 Monitoramento (Filtros e Status)": [
            "Pressostato para monitorar os filtros G4 (PSH)", "Pressostato para monitorar os filtros F9 (PSH)", "Pressostato para monitorar os filtros H13/H14 (PSH)",
            "Status funcionamento ventilador ou exaustor (partida direta) (PSH)", "Transmissor de pressão diferencial (monitorar os filtros G4) (PDIT)",
            "Transmissor de pressão diferencial (monitorar os filtros F9) (PDIT)", "Transmissor de pressão diferencial (monitorar os filtros H13) (PDIT)"
        ],
        "🟢 Monitoramento e Controle de Ambientes": [
            "Transmissor de pressão diferencial entre salas (PDT)", "Transmissor de pressão diferencial entre salas com display (PDIT)",
            "Transmissor de temperatura Ambiente (TT)", "Transmissor de temperatura ambiente com display (TIT)",
            "Transmissor de temperatura e umidade ambiente (TT/MT)", "Transmissor de temperatura e umidade ambiente com display (TIT/MIT)",
            "Transmissor de CO2 ambiente (AT/AIT)", "Resistência de aquecimento (Duto) (RAQ)"
        ]
    }
    
    KITS_PADRAO = {
        "❄️ UTA Padrão - Água Gelada": {
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1,
            "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Válvula de controle de água gelada proporcional (TCV)": 1,
            "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1, "Pressostato para monitorar os filtros H13/H14 (PSH)": 1
        },
        "🌬️ UTA Padrão - Expansão Direta": {
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1,
            "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Relé de Corrente - Status Compressor (TC)": 2,
            "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1, "Pressostato para monitorar os filtros H13/H14 (PSH)": 1
        },
        "🔥 UTA Padrão - Água Gelada + Resistência": {
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1,
            "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Válvula de controle de água gelada proporcional (TCV)": 1,
            "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1, "Pressostato para monitorar os filtros H13/H14 (PSH)": 1,
            "Termostato de segurança (TSH)": 1, "Pressostato diferencial para ar (PSH)": 1, "Resistência de aquecimento (Equipamento) (RAQ)": 1
        },
        "ENTREGÁVEL EXP. DIRET + RESISTÊNCIA": {
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1,
            "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Relé de Corrente - Status Compressor (TC)": 2,
            "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1, "Pressostato para monitorar os filtros H13/H14 (PSH)": 1,
            "Termostato de segurança (TSH)": 1, "Pressostato diferencial para ar (PSH)": 1, "Resistência de aquecimento (Equipamento) (RAQ)": 1
        },
        "💨 Adicional: Ventilador/Exaustor (Inversor)": { "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1 },
        "⚙️ Adicional: Ventilador/Exaustor (Partida Direta)": { "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": 1 }
    }

    def calcular_painel_fisico(qtd_controladores):
        if qtd_controladores == 0: return "Sem Painel", 0.0
        elif qtd_controladores <= 4: return "Quadro 600x400mm", 4500.00
        elif qtd_controladores <= 10: return "Quadro 800x600mm", 5900.00
        else: return "Armário 1200x800mm", 9250.00

    def dimensionar_controladores(total_io):
        c36 = c24 = c18 = c15 = 0
        rem = total_io
        while rem > 0:
            if rem > 24: c36 += 1; rem -= 36
            elif rem > 18: c24 += 1; rem -= 24
            elif rem > 15: c18 += 1; rem -= 18
            else: c15 += 1; rem -= 15
        return c36, c24, c18, c15

    def dimensionar_siemens_1200(ai, ao, di, do):
        hw = {}
        rem_ai = max(0, ai - 2); rem_ao = max(0, ao - 0)
        rem_di = max(0, di - 14); rem_do = max(0, do - 10)
        if ai>0 or ao>0 or di>0 or do>0:
            hw["Siemens - CPU 1214C DC/DC/DC"] = 1
            hw["Siemens - Fonte 24VDC 2.5A"] = 1
            hw["Siemens - Cartão de Memória 4MB"] = 1
        if rem_ai > 0: hw["Siemens - SM 1231 AI 8x13Bit"] = (rem_ai + 7) // 8
        if rem_ao > 0: hw["Siemens - SM 1232 AQ 4x14Bit"] = (rem_ao + 3) // 4
        if rem_di > 0: hw["Siemens - SM 1221 DI 16x24VDC"] = (rem_di + 15) // 16
        if rem_do > 0: hw["Siemens - SM 1222 DQ 16x24VDC"] = (rem_do + 15) // 16
        return hw

    def dimensionar_siemens_1500(ai, ao, di, do):
        hw = {}
        rem_ai = max(0, ai - 0); rem_ao = max(0, ao - 0)
        rem_di = max(0, di - 0); rem_do = max(0, do - 0)
        if ai>0 or ao>0 or di>0 or do>0:
            hw["Siemens - CPU 1511-1 PN"] = 1
            hw["Siemens - Fonte PM 1507 24VDC 8A"] = 1
            hw["Siemens - Cartão de Memória 12MB"] = 1
        if rem_ai > 0: hw["Siemens - AI 8xU/I HS"] = (rem_ai + 7) // 8
        if rem_ao > 0: hw["Siemens - AQ 4xU/I ST"] = (rem_ao + 3) // 4
        if rem_di > 0: hw["Siemens - DI 16x24VDC HF"] = (rem_di + 15) // 16
        if rem_do > 0: hw["Siemens - DQ 16x24VDC/0.5A"] = (rem_do + 15) // 16
        return hw
        
    def dimensionar_mercato(ui, ao, do):
        if ui <= 4 and ao <= 2 and do <= 3: return "Mercato - Controlador MCP-4 (2AO, 3DO, 4UI)"
        elif ui <= 8 and ao <= 4 and do <= 5: return "Mercato - Controlador MCP-8 (4AO, 5DO, 8UI, 2DI)"
        elif ui <= 12 and ao <= 4 and do <= 8: return "Mercato - Controlador MCP-12 (4AO, 8DO, 12UI, 4DI)"
        return None

    aba_auto, aba_infra, aba_precos, aba_resumo = st.tabs([
        "🚀 Dimensionamento de Automação", "🔌 Infraestrutura Lançamento", "💲 Base de Preços", "📊 Orçamento Final"
    ])

    with aba_auto:
        if not st.session_state.wizard_ativo:
            if st.button("➕ Criar Novo Quadro de Automação", type="primary"):
                st.session_state.wizard_ativo = True
                st.rerun()

        if st.session_state.wizard_ativo:
            with st.container(border=True):
                st.markdown("### 🧙‍♂️ Assistente de Configuração de Quadro")
                
                arquitetura_opt = st.radio("1. Selecione a Arquitetura do Hardware do Quadro:", ["SpaceLogic (Schneider)", "S7-1200 (Siemens)", "S7-1500 (Siemens)", "MCP Parametrizável (Mercato - Linha mais econômica)"], horizontal=True)
                is_mercato_arch = "Mercato" in arquitetura_opt
                is_siemens_arch = "Siemens" in arquitetura_opt
                
                if is_mercato_arch: opcoes_ihm_wizard = ["Sem Interface (Cego)", "Mercato - IHM Básica 4.3\""]
                else: opcoes_ihm_wizard = ["Sem Interface (Cego)", "IHM Padrão 7\"", "IHM Premium 10\""]
                ihm_selecionada = st.radio("2. O quadro possuirá IHM local?", opcoes_ihm_wizard, horizontal=True)
                
                tipo_q = st.radio("3. Selecione o Tipo do Quadro:", ["Controle (HVAC/Máquinas)", "CAG (Central de Água Gelada)"], horizontal=True)
                
                tipo_cfr_wizard = "Não Aplicável"
                if is_mercato_arch:
                    sup_opt = "Não"
                    soft_sel = "Sem Supervisório"
                    st.info("ℹ️ A arquitetura Parametrizável Mercato não contempla sistema de integração em rede / supervisório nativo nesta configuração padrão.")
                else:
                    sup_opt = st.radio("4. Este quadro fará parte de um Sistema de Supervisório?", ["Não", "Sim"], horizontal=True)
                    soft_sel = "Sem Supervisório"
                    if sup_opt == "Sim":
                        if is_siemens_arch: opcoes_soft = ["Sistema supervisório SEM certificação CFR-21", "Sistema supervisório COM certificação CFR-21"]
                        else: opcoes_soft = ["Sistema supervisório SEM certificação CFR-21", "Sistema supervisório COM certificação CFR-21", "Sistema de monitoramento Schneider EBO"]
                        soft_sel = st.selectbox("Selecione o Software de Supervisão:", opcoes_soft)
                        
                        if "COM certificação" in soft_sel:
                            tipo_cfr_wizard = st.radio(
                                "Selecione a Modalidade do CFR-21 Part 11:",
                                ["CFR21 Part 11 - Qualificável", "CFR21 Part 11 - Qualificado"],
                                help="**Qualificável:** Onde a SIARCON após receber todos os requisitos de usuário, irá fornecer o sistema de automação com todos os pré requisitos para que o cliente possa realizar a qualificação.\n\n**Qualificado:** A SIARCON entregará todos os protocolos e testes que qualifiquem o sistema supervisório."
                            )
                
                tag_q = st.text_input("5. Insira a TAG / Identificação do Quadro (Ex: QTA-01, QD-CAG):")
                config_opt = st.radio("6. Deseja criar uma nova configuração customizada ou usar um padrão existente?", ["Usar Padrão Existente (Kits)", "Criar Nova Configuração Customizada (Em Branco)"], horizontal=True)
                
                kit_final_selecionado = None
                if config_opt == "Usar Padrão Existente (Kits)":
                    opcoes_kits_filtrados = list(KITS_PADRAO.keys())
                    if "CAG" in tipo_q: opcoes_kits_filtrados = [k for k in KITS_PADRAO.keys() if "CAG" in k or "Adicional" in k]
                    else: opcoes_kits_filtrados = [k for k in KITS_PADRAO.keys() if "CAG" not in k]
                    kit_final_selecionado = st.selectbox("Selecione o Modelo Padrão SIARCON:", ["Selecione..."] + opcoes_kits_filtrados)
                
                sobra_opt = st.radio("7. Deseja considerar 20% de sobra nas I/O (Reserva Técnica)?", ["Não", "Sim"], horizontal=True)
                
                c_conf, c_canc = st.columns(2)
                if c_conf.button("🚀 Confirmar e Montar Quadro", use_container_width=True):
                    if not tag_q: st.warning("⚠️ Insira uma TAG válida para identificar o quadro.")
                    elif config_opt == "Usar Padrão Existente (Kits)" and kit_final_selecionado == "Selecione...": st.warning("⚠️ Selecione um kit padrão.")
                    else:
                        novos_instrumentos = {k: 0 for k in REGRA_IO.keys()}
                        grupos_equip = []
                        if config_opt == "Usar Padrão Existente (Kits)":
                            for item_nome, qtd_padrao in KITS_PADRAO[kit_final_selecionado].items():
                                if item_nome in novos_instrumentos: novos_instrumentos[item_nome] = qtd_padrao
                            nome_limpo = kit_final_selecionado.split(" ", 1)[1] if " " in kit_final_selecionado else kit_final_selecionado
                            grupos_equip.append({"nome_grupo": f"{nome_limpo}", "multiplicador": 1, "instrumentos": novos_instrumentos, "tags_lista": [""]})
                        else:
                            grupos_equip.append({"nome_grupo": "Equipamento Novo", "multiplicador": 1, "instrumentos": novos_instrumentos, "tags_lista": [""]})

                        novo_quadro = {
                            "id": str(uuid.uuid4()),
                            "nome": tag_q, "tipo": tipo_q, "supervisorio": soft_sel, "arquitetura": arquitetura_opt,
                            "tipo_cfr": tipo_cfr_wizard,
                            "modo_config": config_opt, "ihm": ihm_selecionada, "sobra_20": sobra_opt, "grupos_equipamentos": grupos_equip
                        }
                        
                        # Insere invertido (o mais novo fica no topo da tela)
                        st.session_state.paineis_auto.insert(0, novo_quadro)
                        st.session_state.wizard_ativo = False
                        st.rerun()
                        
                if c_canc.button("❌ Cancelar", use_container_width=True):
                    st.session_state.wizard_ativo = False
                    st.rerun()

        st.write("")

        # Tratamento da exclusão/limpeza total
        if st.session_state.confirmar_limpar:
            st.warning("⚠️ Tem certeza que deseja sair e PERDER todo o preenchimento não salvo nesta tela?")
            c_sim, c_nao = st.columns(2)
            if c_sim.button("✔️ Sim, Apagar Tudo e Sair"):
                st.session_state.paineis_auto = []
                st.session_state.nome_projeto_orcamento = ""
                st.session_state.confirmar_limpar = False
                st.rerun()
            if c_nao.button("❌ Não, Cancelar e Voltar"):
                st.session_state.confirmar_limpar = False
                st.rerun()

        for p_idx, p_data in enumerate(st.session_state.paineis_auto):
            is_mercato_quadro = ('Mercato' in p_data.get('arquitetura', ''))
            is_schneider_quadro = ('Schneider' in p_data.get('arquitetura', ''))
            tem_sobra_20 = (p_data.get('sobra_20', 'Não') == 'Sim')
            
            # Utilizar st.expander em vez de st.container. O primeiro (p_idx == 0) fica expandido, os antigos ficam recolhidos
            with st.expander(f"🎛️ Quadro: {p_data['nome']} - {p_data.get('arquitetura', '')}", expanded=(p_idx == 0)):
                c_icone, c_nome_painel, c_ihm_painel = st.columns([0.5, 4, 2])
                c_icone.markdown("## 🎛️")
                p_data['nome'] = c_nome_painel.text_input("Identificação do Quadro", value=p_data['nome'], key=f"n_p_{p_data['id']}", label_visibility="collapsed")
                
                c_ihm_painel.markdown(f"<div style='padding-top:10px; color:#555;'><b>IHM:</b> {p_data.get('ihm', 'Sem Interface (Cego)')}</div>", unsafe_allow_html=True)
                st.caption(f"**Arquitetura:** {p_data.get('arquitetura', 'SpaceLogic (Schneider)')} | **Supervisão:** {p_data.get('supervisorio', 'Sem Supervisório')} | **CFR-21:** {p_data.get('tipo_cfr', 'Não Aplicável')} | **Reserva 20%:** {p_data.get('sobra_20', 'Não')}")
                
                with st.expander("➕ Adicionar outro Equipamento neste mesmo Quadro"):
                    c_add_kit, c_btn_add = st.columns([3, 1])
                    sub_kit = c_add_kit.selectbox("Escolha o Equipamento / Kit:", ["Selecione...", "Equipamento Novo (Em Branco)"] + list(KITS_PADRAO.keys()), key=f"sub_kit_{p_data['id']}")
                    if c_btn_add.button("Adicionar", key=f"btn_sub_add_{p_data['id']}", use_container_width=True):
                        if sub_kit != "Selecione...":
                            novos_inst = {k: 0 for k in REGRA_IO.keys()}
                            
                            if sub_kit == "Equipamento Novo (Em Branco)":
                                p_data['grupos_equipamentos'].append({"nome_grupo": "Equipamento Novo", "multiplicador": 1, "instrumentos": novos_inst, "tags_lista": [""]})
                            else:
                                for item_nome, qtd_padrao in KITS_PADRAO[sub_kit].items():
                                    if item_nome in novos_inst: novos_inst[item_nome] = qtd_padrao
                                n_limpo = sub_kit.split(" ", 1)[1] if " " in sub_kit else sub_kit
                                p_data['grupos_equipamentos'].append({"nome_grupo": f"{n_limpo}", "multiplicador": 1, "instrumentos": novos_inst, "tags_lista": [""]})
                            st.rerun()

                raw_ai_painel = raw_ao_painel = raw_di_painel = raw_do_painel = 0

                for g_idx, g_data in enumerate(p_data['grupos_equipamentos']):
                    total_ai_g_single = total_ao_g_single = total_di_g_single = total_do_g_single = 0
                    
                    open_p_hvac = False 
                    with st.expander(f"📦 {g_data['nome_grupo']}", expanded=open_p_hvac):
                        qtd_key = f"m_g_{p_data['id']}_{g_idx}"
                        qtd_atual = st.session_state.get(qtd_key, g_data.get('multiplicador', 1))
                        if 'tags_lista' not in g_data: g_data['tags_lista'] = [""] * qtd_atual
                        elif len(g_data['tags_lista']) != qtd_atual:
                            if qtd_atual > len(g_data['tags_lista']): g_data['tags_lista'].extend([""] * (qtd_atual - len(g_data['tags_lista'])))
                            else: g_data['tags_lista'] = g_data['tags_lista'][:qtd_atual]

                        render_qtd = min(qtd_atual, 5) 
                        col_ratios = [3] + [1.5] * render_qtd + [1]
                        cols = st.columns(col_ratios)
                        g_data['nome_grupo'] = cols[0].text_input("Equipamento", value=g_data['nome_grupo'], key=f"n_g_{p_data['id']}_{g_idx}")
                        for i in range(render_qtd): g_data['tags_lista'][i] = cols[i+1].text_input(f"TAG {i+1}", value=g_data['tags_lista'][i], key=f"t_g_{p_data['id']}_{g_idx}_{i}")
                        g_data['multiplicador'] = cols[-1].number_input("Qtd", min_value=1, value=qtd_atual, key=qtd_key)
                        
                        if qtd_atual > 5: st.caption("⚠️ Para mais de 5 equipamentos, as TAGs extras podem ser inseridas como anotações no final do projeto.")

                        with st.container():
                            for inst, q in g_data['instrumentos'].items():
                                io_vals = REGRA_IO.get(inst, {"AI": 0, "AO": 0, "DI": 0, "DO": 0})
                                total_ai_g_single += q * io_vals["AI"]
                                total_ao_g_single += q * io_vals["AO"]
                                total_di_g_single += q * io_vals["DI"]
                                total_do_g_single += q * io_vals["DO"]
                                
                            # +2 DI's da Seletora de Painel por equipamento contabilizado
                            total_di_g_single += 2
                                
                            if is_mercato_quadro:
                                ui_nec = total_ai_g_single + total_di_g_single
                                ui_check = math.ceil(ui_nec * 1.2) if tem_sobra_20 else ui_nec
                                ao_check = math.ceil(total_ao_g_single * 1.2) if tem_sobra_20 else total_ao_g_single
                                do_check = math.ceil(total_do_g_single * 1.2) if tem_sobra_20 else total_do_g_single
                                
                                modelo_mcp = dimensionar_mercato(ui_check, ao_check, do_check)
                                if not modelo_mcp:
                                    st.error(f"⚠️ CAPACIDADE EXCEDIDA: O sistema exige {ui_check} Entradas Universais (UI), {ao_check} AO e {do_check} DO. Isso ultrapassa a capacidade máxima do maior controlador da linha MCP Mercato. Configuração não recomendada, divida o sistema ou use arquiteturas modulares (Schneider/Siemens).")
                                else:
                                    st.success(f"✅ OK! Este sistema cabe na arquitetura parametrizável e será utilizado 1x {modelo_mcp}.")

                        with st.expander("⚙️ Ajuste Fino de Instrumentos (Engenharia)"):
                            for grupo_nome, lista_itens in GRUPOS_INSTRUMENTOS.items():
                                open_p_eng = False
                                with st.expander(grupo_nome, expanded=open_p_eng):
                                    cols_inst = st.columns(2)
                                    for i, inst in enumerate(lista_itens):
                                        if inst not in g_data['instrumentos']: g_data['instrumentos'][inst] = 0
                                        with cols_inst[i % 2]:
                                            g_data['instrumentos'][inst] = st.number_input(inst, min_value=0, step=1, value=g_data['instrumentos'][inst], key=f"inst_{p_data['id']}_{g_idx}_{inst}")
                        if st.button("🗑️ Remover Máquina", key=f"del_{p_data['id']}_{g_idx}"):
                            p_data['grupos_equipamentos'].pop(g_idx)
                            st.rerun()

                for g_data in p_data['grupos_equipamentos']:
                    qtd_atual_calc = g_data.get('multiplicador', 1)
                    for inst, q in g_data['instrumentos'].items():
                        io_vals = REGRA_IO.get(inst, {"AI": 0, "AO": 0, "DI": 0, "DO": 0})
                        raw_ai_painel += q * io_vals["AI"] * qtd_atual_calc
                        raw_ao_painel += q * io_vals["AO"] * qtd_atual_calc
                        raw_di_painel += q * io_vals["DI"] * qtd_atual_calc
                        raw_do_painel += q * io_vals["DO"] * qtd_atual_calc

                    # Adiciona as DIs invisíveis de cada máquina no painel global
                    raw_di_painel += (2 * qtd_atual_calc)

                reserva_ai = math.ceil(raw_ai_painel * 0.2) if tem_sobra_20 else 0
                reserva_ao = math.ceil(raw_ao_painel * 0.2) if tem_sobra_20 else 0
                reserva_di = math.ceil(raw_di_painel * 0.2) if tem_sobra_20 else 0
                reserva_do = math.ceil(raw_do_painel * 0.2) if tem_sobra_20 else 0
                
                hw_ai_painel = raw_ai_painel + reserva_ai
                hw_ao_painel = raw_ao_painel + reserva_ao
                hw_di_painel = raw_di_painel + reserva_di
                hw_do_painel = raw_do_painel + reserva_do

                total_io_pontos_hw = hw_ai_painel + hw_ao_painel + hw_di_painel + hw_do_painel
                
                st.markdown("##### 🧠 Estrutura Total de I/O do Quadro")
                m1, m2, m3, m4, m5 = st.columns(5)
                titulo_total = "Total I/O Físico (Com Reserva)" if tem_sobra_20 else "Total I/O Físico"
                m1.metric(titulo_total, str(total_io_pontos_hw)) 
                m2.metric("AI / UI", hw_ai_painel)
                m3.metric("AO", hw_ao_painel)
                m4.metric("DI / UI", hw_di_painel)
                m5.metric("DO", hw_do_painel)
                
                if st.button("🗑️ Deletar Todo este Quadro", key=f"del_quadro_{p_data['id']}"):
                    st.session_state.paineis_auto.pop(p_idx)
                    st.rerun()
        
        if st.session_state.paineis_auto:
            st.markdown("---")
            
            c_bot_salvar, c_bot_sair = st.columns(2)
            
            if c_bot_salvar.button("💾 Salvar Rascunho e Sair (Retomar depois)", type="primary", use_container_width=True):
                if not st.session_state.nome_projeto_orcamento: st.warning("⚠️ Atenção: Preencha o 'Nome do Orçamento / Projeto' no topo da página antes de salvar o rascunho.")
                else:
                    try:
                        sh = conectar_google_sheets()
                        try: ws_hist_orc = sh.worksheet("Historico_Orcamentos")
                        except:
                            ws_hist_orc = sh.add_worksheet(title="Historico_Orcamentos", rows="1000", cols="8")
                            ws_hist_orc.append_row(["Data/Hora", "Nome do Projeto", "Revisão", "Subtotal Hardware", "Serviços de Lógica", "Custo Total Estimado", "Configuracao_JSON", "Usuário"])
                        todas_linhas = ws_hist_orc.get_all_values()
                        contagem_revisoes = sum(1 for r in todas_linhas[1:] if r[1].strip().upper() == st.session_state.nome_projeto_orcamento.strip().upper())
                        revisao_atual = f"R-{contagem_revisoes:02d}"
                        nova_linha = [datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M:%S"), st.session_state.nome_projeto_orcamento, revisao_atual, "R$ 0,00", "R$ 0,00", "R$ 0,00 (Rascunho)", json.dumps(st.session_state.paineis_auto), st.session_state.usuario_logado]
                        ws_hist_orc.append_row(nova_linha)
                        st.session_state.paineis_auto = []
                        st.session_state.nome_projeto_orcamento = ""
                        st.cache_data.clear()
                        st.toast("📝 Rascunho salvo na nuvem com sucesso! Tela limpa.", icon="💾")
                        st.rerun()
                    except Exception as e: st.error(f"Erro ao salvar: {e}")

            if c_bot_sair.button("🗑️ Somente Sair (Limpar Tela)", type="secondary", use_container_width=True):
                st.session_state.confirmar_limpar = True
                st.rerun()

        st.markdown("---")
        with st.expander(f"📂 Abrir Orçamento Existente (Histórico de {st.session_state.nome_exibicao})"):
            try:
                sh = conectar_google_sheets()
                todas_linhas = sh.worksheet("Historico_Orcamentos").get_all_values()
                if len(todas_linhas) > 1:
                    for idx_rev, linha in enumerate(todas_linhas[1:][::-1]):
                        idx_real = len(todas_linhas) - 2 - idx_rev
                        usuario_registro = linha[7] if len(linha) > 7 else "rodrigo.ribeiro"
                        if usuario_registro.strip().lower() != st.session_state.usuario_logado.strip().lower(): continue
                        with st.container():
                            c1, c2, c3, c4, c5 = st.columns([1.5, 3, 1.5, 1, 1])
                            c1.write(f"📅 {linha[0]}")
                            c2.write(f"**{linha[1]}** `({linha[2] if len(linha)>=7 else 'R-00'})`")
                            c3.write(linha[5] if len(linha)>=7 else linha[4])
                            if c4.button("📂 Carregar", key=f"btn_abrir_{idx_real}"):
                                st.session_state.projeto_para_abrir = idx_real
                                st.session_state.dados_projeto_abrir = {'nome': linha[1], 'json': linha[6] if len(linha)>=7 else linha[5]}
                            if c5.button("🗑️ Excluir", key=f"btn_del_hist_{idx_real}", type="secondary"):
                                try:
                                    ws_hist_orc = sh.worksheet("Historico_Orcamentos")
                                    ws_hist_orc.delete_rows(idx_real + 2)
                                    st.cache_data.clear()
                                    st.toast("🗑️ Orçamento excluído do histórico com sucesso!", icon="✅")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Erro ao excluir: {e}")
                        st.markdown("---")
                            
                    if st.session_state.get('projeto_para_abrir') is not None:
                        d_a = st.session_state.get('dados_projeto_abrir', {})
                        st.warning(f"⚠️ Carregar os dados de '{d_a['nome']}' irá substituir as configurações atuais.")
                        c_sim, c_nao = st.columns(2)
                        if c_sim.button("✔️ Sim, substituir tela", use_container_width=True):
                            st.session_state.paineis_auto = json.loads(d_a['json'])
                            # Previnir quadros antigos sem id:
                            for p in st.session_state.paineis_auto: 
                                if 'id' not in p: p['id'] = str(uuid.uuid4())
                            st.session_state.nome_projeto_orcamento = d_a['nome']
                            st.session_state.projeto_para_abrir = None
                            st.rerun()
                        if c_nao.button("❌ Não, cancelar", use_container_width=True):
                            st.session_state.projeto_para_abrir = None
                            st.rerun()
            except Exception as e: st.write("A aba 'Historico_Orcamentos' está vazia.")

    with aba_infra:
        st.header("Cálculo de Infraestrutura Lançamento Manual")
        st.write("Insira lançamentos avulsos de infraestrutura caso não estejam contemplados nos quadros:")
        with st.form("form_infra_manual"):
            col_i1, col_i2, col_i3 = st.columns([2,1,1])
            desc_infra = col_i1.text_input("Descrição do Item de Infraestrutura")
            qtd_infra = col_i2.number_input("Quantidade (m/un)", min_value=0.0, step=1.0)
            valor_infra = col_i3.number_input("Preço Unitário (R$)", min_value=0.0, step=0.1)
            if st.form_submit_button("➕ Adicionar à Infra"):
                if qtd_infra > 0 and valor_infra > 0:
                    st.session_state.orcamento.append({"Categoria": "Infraestrutura (Avulsa)", "Item": desc_infra, "Quantidade": qtd_infra, "Custo_Total": qtd_infra * valor_infra})
                    st.success("Adicionado com sucesso.")
                    
    with aba_precos:
        st.header("Gestão da Base de Preços")
        st.write("Altere os valores e salve na nuvem para manter a equipe comercial sincronizada.")
        
        st.subheader("Base Geral e Schneider")
        lista_schneider = list(banco_schneider_comum.keys())
        lista_schneider.extend(["IHM Padrão 7\"", "IHM Premium 10\""])
        df_geral = pd.DataFrame([{"Item / Equipamento": k, "Valor Atual (R$)": st.session_state.precos_banco.get(k, 0.0)} for k in lista_schneider if k in st.session_state.precos_banco])
        edited_geral = st.data_editor(df_geral, use_container_width=True, hide_index=True, key="ed_geral")
        
        st.subheader("Base Siemens")
        lista_siemens = list(banco_siemens.keys())
        df_siemens = pd.DataFrame([{"Item / Equipamento": k, "Valor Atual (R$)": st.session_state.precos_banco.get(k, 0.0)} for k in lista_siemens if k in st.session_state.precos_banco])
        edited_siemens = st.data_editor(df_siemens, use_container_width=True, hide_index=True, key="ed_siem")

        st.subheader("Base Mercato e NTC")
        lista_mercato = list(banco_mercato.keys())
        df_mercato = pd.DataFrame([{"Item / Equipamento": k, "Valor Atual (R$)": st.session_state.precos_banco.get(k, 0.0)} for k in lista_mercato if k in st.session_state.precos_banco])
        edited_mercato = st.data_editor(df_mercato, use_container_width=True, hide_index=True, key="ed_merc")
        
        st.subheader("Serviços CFR-21 e Qualificação")
        lista_cfr = list(banco_cfr_servicos.keys())
        df_cfr = pd.DataFrame([{"Item / Equipamento": k, "Valor Atual (R$)": st.session_state.precos_banco.get(k, 0.0)} for k in lista_cfr if k in st.session_state.precos_banco])
        edited_cfr = st.data_editor(df_cfr, use_container_width=True, hide_index=True, key="ed_cfr")
        
        if st.button("💾 Salvar Novos Preços no Banco de Dados", type="primary"):
            alterou_algo = False
            novos_historicos = []
            edited_total = pd.concat([edited_geral, edited_siemens, edited_mercato, edited_cfr], ignore_index=True)
            for idx, row in edited_total.iterrows():
                item = row['Item / Equipamento']
                novo_valor = row['Valor Atual (R$)']
                antigo_valor = st.session_state.precos_banco.get(item, 0.0)
                if novo_valor != antigo_valor:
                    novo_hist = {"Data/Hora": datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M:%S"), "Item Alterado": item, "Valor Antigo": f"R$ {antigo_valor:.2f}", "Novo Valor": f"R$ {novo_valor:.2f}"}
                    st.session_state.historico_precos.append(novo_hist)
                    novos_historicos.append(novo_hist)
                    st.session_state.precos_banco[item] = novo_valor
                    alterou_algo = True
            if alterou_algo:
                try:
                    sh = conectar_google_sheets()
                    try: ws_precos = sh.worksheet("Precos")
                    except:
                        ws_precos = sh.add_worksheet(title="Precos", rows="100", cols="2")
                        ws_precos.append_row(["Item", "Valor"])
                    ws_precos.clear()
                    ws_precos.append_rows([["Item", "Valor"]] + [[k, v] for k, v in st.session_state.precos_banco.items()])
                    try: ws_hist = sh.worksheet("Historico_Precos")
                    except:
                        ws_hist = sh.add_worksheet(title="Historico_Precos", rows="1000", cols="4")
                        ws_hist.append_row(["Data/Hora", "Item Alterado", "Valor Antigo", "Novo Valor"])
                    if novos_historicos: ws_hist.append_rows([[h["Data/Hora"], h["Item Alterado"], h["Valor Antigo"], h["Novo Valor"]] for h in novos_historicos])
                    st.cache_data.clear()
                    st.success("✅ Preços atualizados na nuvem!")
                except Exception as e: st.error(f"Erro ao salvar: {e}")

    with aba_resumo:
        st.header("Consolidação Financeira do Orçamento")
        linhas_inst_campo = []
        linhas_hardware = []
        linhas_software = []
        linhas_pontos = []
        linhas_servicos = []
        
        softwares_incluidos = {}
        
        total_ai_schneider = total_ao_schneider = total_di_schneider = total_do_schneider = 0
        total_ai_siemens = total_ao_siemens = total_di_siemens = total_do_siemens = 0
        total_io_mercato = 0
        
        custo_base_schneider = 0.0
        custo_base_siemens = 0.0
        custo_base_mercato = 0.0

        descritivo_linhas_excel = ["DESCRIÇÃO TÉCNICA DO ESCOPO CONTEMPLADO:\n"]
        descritivo_comercial_linhas = []

        for p in st.session_state.paineis_auto:
            arquitetura_atual = p.get('arquitetura', 'SpaceLogic (Schneider)')
            is_siemens_1200 = (arquitetura_atual == 'S7-1200 (Siemens)')
            is_siemens_1500 = (arquitetura_atual == 'S7-1500 (Siemens)')
            is_siemens = is_siemens_1200 or is_siemens_1500
            is_mercato = ('Mercato' in arquitetura_atual)
            is_schneider = ('Schneider' in arquitetura_atual)
            tem_sobra_20 = (p.get('sobra_20', 'Não') == 'Sim')
            tipo_cfr_painel = p.get('tipo_cfr', 'Não Aplicável')
            
            raw_ai_painel = raw_ao_painel = raw_di_painel = raw_do_painel = 0
            qtd_equipamentos_painel = 0
            
            lista_equip_nomes = []
            tem_resistencia = False
            tem_filtro_pdt = False
            tem_filtro_psh = False
            
            lista_instrumentos_nomes = set()
            lista_instrumentos_detalhados = []
            controladores_desc_lista = []
            
            for g in p['grupos_equipamentos']:
                mult = g.get('multiplicador', 1)
                qtd_equipamentos_painel += mult
                
                lista_tags = [t for t in g.get('tags_lista', []) if t.strip() != ""]
                str_tags = f" [TAGs: {', '.join(lista_tags)}]" if len(lista_tags) > 0 else ""
                
                nome_limpo_grupo = g['nome_grupo'].replace("Equipamento Novo", "").replace("Equipamento Customizado", "").strip()
                if not nome_limpo_grupo:
                    lista_equip_nomes.append(f"{mult}x Equipamento{str_tags}")
                    nome_equip = f"Equipamento{str_tags}"
                else:
                    lista_equip_nomes.append(f"{mult}x {nome_limpo_grupo}{str_tags}")
                    nome_equip = f"{nome_limpo_grupo}{str_tags}"
                
                raw_ai_g_single = raw_ao_g_single = raw_di_g_single = raw_do_g_single = 0
                
                for inst, qtd in g['instrumentos'].items():
                    if qtd > 0:
                        qtd_final = qtd * mult
                        item_nome_real = inst
                        
                        if is_mercato:
                            if "Transmissor de temperatura para duto (TT)" in inst: item_nome_real = "Mercato - Sensor de Temperatura NTC (Duto)"
                            elif "Transmissor de temperatura Ambiente (TT)" in inst: item_nome_real = "Mercato - Sensor de Temperatura NTC (Ambiente)"
                            elif "Transmissor de temperatura ambiente com display" in inst: item_nome_real = "Mercato - Sensor de Temperatura NTC com Display (Ambiente)"
                        elif is_schneider:
                            if "Transmissor de temperatura para duto (TT)" in inst: item_nome_real = "Schneider - Sensor de Temperatura NTC (Duto)"
                            elif "Transmissor de temperatura Ambiente (TT)" in inst: item_nome_real = "Schneider - Sensor de Temperatura NTC (Ambiente)"
                        
                        nome_curto_inst = item_nome_real.replace("Mercato - ", "").replace("Schneider - ", "").replace("Siemens - ", "")
                        nome_curto_inst = re.sub(r'\s*\([A-Z/]+\)$', '', nome_curto_inst)
                        lista_instrumentos_nomes.add(nome_curto_inst)
                        
                        if "resistência" in inst.lower() or "raq" in inst.lower() or "tsh" in inst.lower():
                            tem_resistencia = True
                        
                        if "filtro" in inst.lower():
                            if "pdit" in inst.lower() or "pdt" in inst.lower(): tem_filtro_pdt = True
                            if "psh" in inst.lower() or "pressostato" in inst.lower(): tem_filtro_psh = True

                        func_inst = "Medição Genérica"
                        if "pressão dif. para ar" in inst.lower(): func_inst = "Medição da Vazão de Ar"
                        elif "temperatura e umidade" in inst.lower(): func_inst = "Medição de Temperatura e Umidade"
                        elif "temperatura" in inst.lower(): func_inst = "Medição de Temperatura"
                        elif "proporcional" in inst.lower(): func_inst = "Controle Proporcional Térmico"
                        elif "compressor" in inst.lower(): func_inst = "Status de Operação do Compressor"
                        elif "termostato" in inst.lower(): func_inst = "Segurança de Sobreaquecimento"
                        elif "resistência" in inst.lower() and "pressostato" not in inst.lower(): func_inst = "Aquecimento"
                        elif "pressostato diferencial para ar" in inst.lower() and "resistência" in inst.lower(): func_inst = "Segurança da Resistência por Fluxo de Ar"
                        elif "pressostato para monitorar" in inst.lower(): func_inst = "Alarme de Saturação de Filtro"
                        elif "transmissor de pressão diferencial (monitorar" in inst.lower(): func_inst = "Monitoramento da Saturação do Filtro"
                        elif "funcionamento ventilador" in inst.lower(): func_inst = "Status de Operação do Ventilador"
                        elif "co2" in inst.lower(): func_inst = "Medição da Qualidade do Ar (CO2)"
                        
                        lista_instrumentos_detalhados.append((nome_curto_inst, qtd_final, func_inst))

                        preco_item = st.session_state.precos_banco.get(item_nome_real, st.session_state.precos_banco.get(inst, 0.0))
                        io_vals = REGRA_IO.get(inst, {"AI": 0, "AO": 0, "DI": 0, "DO": 0})
                        
                        raw_ai_painel += qtd_final * io_vals["AI"]
                        raw_ao_painel += qtd_final * io_vals["AO"]
                        raw_di_painel += qtd_final * io_vals["DI"]
                        raw_do_painel += qtd_final * io_vals["DO"]
                        
                        raw_ai_g_single += qtd * io_vals["AI"]
                        raw_ao_g_single += qtd * io_vals["AO"]
                        raw_di_g_single += qtd * io_vals["DI"]
                        raw_do_g_single += qtd * io_vals["DO"]
                        
                        custo_tot_inst = qtd_final * preco_item
                        linhas_inst_campo.append({"Categoria": "Instrumentação de Campo", "Item": f"{item_nome_real} ({nome_equip} - {p['nome']})", "Preço Unit.": preco_item, "Qtd": qtd_final, "Custo Total": custo_tot_inst})
                        linhas_pontos.append({"Painel": p['nome'], "Grupo/Equipamento": nome_equip, "Instrumento": item_nome_real, "Quantidade Total": qtd_final, "Entrada Digital (DI)": qtd_final * io_vals["DI"], "Saída Digital (DO)": qtd_final * io_vals["DO"], "Entrada Analógica (AI)": qtd_final * io_vals["AI"], "Saída Analógica (AO)": qtd_final * io_vals["AO"]})
                        
                        if is_siemens: custo_base_siemens += custo_tot_inst
                        elif is_mercato: custo_base_mercato += custo_tot_inst
                        else: custo_base_schneider += custo_tot_inst
                
                # ADIÇÃO DA CHAVE SELETORA AUTO/MANUAL POR EQUIPAMENTO (2 DIs)
                raw_di_painel += (2 * mult)
                raw_di_g_single += 2
                linhas_pontos.append({"Painel": p['nome'], "Grupo/Equipamento": nome_equip, "Instrumento": "Chave Seletora Auto/Manual (Painel Elétrico)", "Quantidade Total": mult, "Entrada Digital (DI)": 2 * mult, "Saída Digital (DO)": 0, "Entrada Analógica (AI)": 0, "Saída Analógica (AO)": 0})

                if is_mercato:
                    reserva_g_ui = math.ceil((raw_ai_g_single + raw_di_g_single) * 0.2) if tem_sobra_20 else 0
                    reserva_g_ao = math.ceil(raw_ao_g_single * 0.2) if tem_sobra_20 else 0
                    reserva_g_do = math.ceil(raw_do_g_single * 0.2) if tem_sobra_20 else 0
                    ui_nec = raw_ai_g_single + raw_di_g_single + reserva_g_ui
                    ao_nec = raw_ao_g_single + reserva_g_ao
                    do_nec = raw_do_g_single + reserva_g_do
                    
                    modelo_mcp = dimensionar_mercato(ui_nec, ao_nec, do_nec)
                    if modelo_mcp:
                        controladores_desc_lista.append(f"{mult}x {modelo_mcp.replace('Mercato - ', '')}")
                        p_hw = st.session_state.precos_banco.get(modelo_mcp, 1650.0)
                        linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"{modelo_mcp} ({g['nome_grupo']} - {p['nome']})", "Preço Unit.": p_hw, "Qtd": mult, "Custo Total": mult * p_hw})
                        custo_base_mercato += (mult * p_hw)
                    else:
                        linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"⚠️ ALERTA: Capacidade MCP Excedida ({g['nome_grupo']} - {p['nome']})", "Preço Unit.": 0.0, "Qtd": mult, "Custo Total": 0.0})

            reserva_ai_painel = math.ceil(raw_ai_painel * 0.2) if tem_sobra_20 else 0
            reserva_ao_painel = math.ceil(raw_ao_painel * 0.2) if tem_sobra_20 else 0
            reserva_di_painel = math.ceil(raw_di_painel * 0.2) if tem_sobra_20 else 0
            reserva_do_painel = math.ceil(raw_do_painel * 0.2) if tem_sobra_20 else 0
            
            hw_ai_painel = raw_ai_painel + reserva_ai_painel
            hw_ao_painel = raw_ao_painel + reserva_ao_painel
            hw_di_painel = raw_di_painel + reserva_di_painel
            hw_do_painel = raw_do_painel + reserva_do_painel
            
            tot_io_painel_hw = hw_ai_painel + hw_ao_painel + hw_di_painel + hw_do_painel
            
            if reserva_ai_painel > 0 or reserva_ao_painel > 0 or reserva_di_painel > 0 or reserva_do_painel > 0:
                linhas_pontos.append({"Painel": p['nome'], "Grupo/Equipamento": "Reserva Técnica (20%)", "Instrumento": "Pontos de Sobra Física do Quadro", "Quantidade Total": "-", "Entrada Digital (DI)": reserva_di_painel, "Saída Digital (DO)": reserva_do_painel, "Entrada Analógica (AI)": reserva_ai_painel, "Saída Analógica (AO)": reserva_ao_painel})

            if is_siemens:
                total_ai_siemens += raw_ai_painel; total_ao_siemens += raw_ao_painel; total_di_siemens += raw_di_painel; total_do_siemens += raw_do_painel
            elif is_mercato:
                total_io_mercato += (raw_ai_painel + raw_ao_painel + raw_di_painel + raw_do_painel)
            else:
                total_ai_schneider += raw_ai_painel; total_ao_schneider += raw_ao_painel; total_di_schneider += raw_di_painel; total_do_schneider += raw_do_painel
            
            if tot_io_painel_hw > 0:
                if is_mercato:
                    nome_caixa, preco_caixa = calcular_painel_fisico(qtd_equipamentos_painel)
                    linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"Estrutura Física: {nome_caixa} ({p['nome']})", "Preço Unit.": preco_caixa, "Qtd": 1, "Custo Total": preco_caixa})
                    custo_base_mercato += preco_caixa
                        
                else:
                    nome_caixa, preco_caixa = calcular_painel_fisico(tot_io_painel_hw/15)
                    linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"Estrutura Física: {nome_caixa} ({p['nome']})", "Preço Unit.": preco_caixa, "Qtd": 1, "Custo Total": preco_caixa})
                    
                    if is_siemens:
                        custo_base_siemens += preco_caixa
                        if is_siemens_1200: hw_s = dimensionar_siemens_1200(hw_ai_painel, hw_ao_painel, hw_di_painel, hw_do_painel)
                        else: hw_s = dimensionar_siemens_1500(hw_ai_painel, hw_ao_painel, hw_di_painel, hw_do_painel)
                        for i_hw, q_hw in hw_s.items():
                            if q_hw > 0:
                                controladores_desc_lista.append(f"{q_hw}x {i_hw.replace('Siemens - ', '')}")
                                p_hw = st.session_state.precos_banco.get(i_hw, 0.0)
                                linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"{i_hw} ({p['nome']})", "Preço Unit.": p_hw, "Qtd": q_hw, "Custo Total": q_hw * p_hw})
                                custo_base_siemens += (q_hw * p_hw)
                    else:
                        custo_base_schneider += preco_caixa
                        c36, c24, c18, c15 = dimensionar_controladores(tot_io_painel_hw)
                        if c36 > 0: 
                            controladores_desc_lista.append(f"{c36}x Controlador MP-C-36A")
                            linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"Controlador MP-C-36A ({p['nome']})", "Preço Unit.": st.session_state.precos_banco.get("MP-C-36A", 9459.0), "Qtd": c36, "Custo Total": c36 * st.session_state.precos_banco.get("MP-C-36A", 9459.0)})
                            custo_base_schneider += (c36 * st.session_state.precos_banco.get("MP-C-36A", 9459.0))
                        if c24 > 0: 
                            controladores_desc_lista.append(f"{c24}x Controlador MP-C-24A")
                            linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"Controlador MP-C-24A ({p['nome']})", "Preço Unit.": st.session_state.precos_banco.get("MP-C-24A", 7290.0), "Qtd": c24, "Custo Total": c24 * st.session_state.precos_banco.get("MP-C-24A", 7290.0)})
                            custo_base_schneider += (c24 * st.session_state.precos_banco.get("MP-C-24A", 7290.0))
                        if c18 > 0: 
                            controladores_desc_lista.append(f"{c18}x Controlador MP-C-18A")
                            linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"Controlador MP-C-18A ({p['nome']})", "Preço Unit.": st.session_state.precos_banco.get("MP-C-18A", 5185.0), "Qtd": c18, "Custo Total": c18 * st.session_state.precos_banco.get("MP-C-18A", 5185.0)})
                            custo_base_schneider += (c18 * st.session_state.precos_banco.get("MP-C-18A", 5185.0))
                        if c15 > 0: 
                            controladores_desc_lista.append(f"{c15}x Controlador MP-C-15A")
                            linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"Controlador MP-C-15A ({p['nome']})", "Preço Unit.": st.session_state.precos_banco.get("MP-C-15A", 4649.0), "Qtd": c15, "Custo Total": c15 * st.session_state.precos_banco.get("MP-C-15A", 4649.0)})
                            custo_base_schneider += (c15 * st.session_state.precos_banco.get("MP-C-15A", 4649.0))
                
                if p.get('ihm') and "Cego" not in p['ihm']:
                    preco_ihm = st.session_state.precos_banco.get(p['ihm'], 0.0)
                    if preco_ihm > 0: 
                        linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"Interface: {p['ihm']} ({p['nome']})", "Preço Unit.": preco_ihm, "Qtd": 1, "Custo Total": preco_ihm})
                        if is_mercato: custo_base_mercato += preco_ihm
                        elif is_siemens: custo_base_siemens += preco_ihm
                        else: custo_base_schneider += preco_ihm

                s_type = p.get('supervisorio', "Sem Supervisório")
                if s_type != "Sem Supervisório":
                    if is_schneider:
                        pr_as = st.session_state.precos_banco.get("Schneider - Servidor de Automação (SpaceLogic AS-P/AS-B)", 9500.0)
                        controladores_desc_lista.append("1x Servidor de Automação AS-P/AS-B")
                        linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"Servidor de Automação AS-P/AS-B ({p['nome']})", "Preço Unit.": pr_as, "Qtd": 1, "Custo Total": pr_as})
                        custo_base_schneider += pr_as
                        
                    chave_soft = (s_type, tipo_cfr_painel)
                    if chave_soft not in softwares_incluidos: softwares_incluidos[chave_soft] = 0
                    softwares_incluidos[chave_soft] += (raw_ai_painel + raw_ao_painel + raw_di_painel + raw_do_painel)

            # --- CONSTRUÇÃO DOS TEXTOS E DESCRITIVOS (RESUMIDO E COMERCIAL) ---
            ihm_desc = f"com IHM instalada na porta, com display de {p['ihm'].replace('Mercato - ', '').replace('IHM Padrão ', '').replace('IHM Premium ', '').replace('IHM Básica ', '')}" if "Cego" not in p['ihm'] else "sem interface IHM instalada"
            
            if "Sem" in p['supervisorio']: sup_desc = "Stand-alone (sem supervisório)"
            elif "EBO" in p['supervisorio']: sup_desc = "integrado ao sistema supervisório EBO"
            else: sup_desc = "integrado ao sistema supervisório"
            
            eq_desc = ", ".join(lista_equip_nomes)
            nome_arquitetura = arquitetura_atual.replace(" - Linha mais econômica", "")
            ctrl_desc = ", ".join(controladores_desc_lista) if controladores_desc_lista else "Controlador a definir"
            
            texto_filtro = ""
            bullet_filtros = ""
            if tem_filtro_pdt:
                texto_filtro = "monitoramento contínuo da saturação dos filtros"
                bullet_filtros = "• Monitoramento contínuo e alarmes de saturação de filtros.\n"
            elif tem_filtro_psh:
                texto_filtro = "monitoramento para alarme devido à saturação dos filtros"
                bullet_filtros = "• Monitoramento para alarme devido à saturação de filtros.\n"
                
            res_desc_intro = "controle da resistência elétrica de aquecimento" if tem_resistencia else ""
            
            componentes_intro = []
            if texto_filtro: componentes_intro.append(texto_filtro)
            if res_desc_intro: componentes_intro.append(res_desc_intro)
            texto_intro_extra = ", incluindo " + " e ".join(componentes_intro) if componentes_intro else ""

            # Descritivo Resumido (Excel)
            texto_p = (
                f"Sistema de automação dedicado para controle de {eq_desc}{texto_intro_extra}.\n\n"
                f"O sistema contempla quadro de automação [TAG: {p['nome']}] {ihm_desc}, "
                f"baseado na tecnologia {nome_arquitetura} ({ctrl_desc}), operando no modo {sup_desc}, "
                f"permitindo a visualização em tempo real e o controle dos seguintes parâmetros operacionais gerais:\n\n"
                f"• Status de operação dos equipamentos.\n"
                f"{bullet_filtros}"
            )
            if tem_resistencia: texto_p += "• Status e acionamento da resistência elétrica.\n"
            texto_p += (
                f"• Leitura de instrumentos de campo diversos ({', '.join(list(lista_instrumentos_nomes))}).\n"
                f"• Condições gerais de funcionamento.\n\n"
                f"A solução proporciona maior confiabilidade operacional, facilidade de manutenção e gestão eficiente dos ativos térmicos e de controle de ar."
            )
            
            if tipo_cfr_painel == 'CFR21 Part 11 - Qualificável':
                texto_p += "\n\nO sistema será fornecido de forma Qualificável conforme normas CFR 21 Part 11, atendendo a todos os requisitos técnicos e de software para que o cliente realize a qualificação posterior."
            elif tipo_cfr_painel == 'CFR21 Part 11 - Qualificado':
                texto_p += "\n\nO sistema será integralmente Qualificado conforme normas CFR 21 Part 11, com a entrega de todos os protocolos pela equipe especializada da SIARCON."

            descritivo_linhas_excel.append(texto_p)
            
            # Descritivo Detalhado (Proposta Comercial)
            txt_com = f"**Sistema completo de automação [TAG: {p['nome']}]**, construído com arquitetura de controladores **{nome_arquitetura}** ({ctrl_desc}). O sistema operará de forma **{sup_desc}**, {ihm_desc}.\n\n"
            
            if tipo_cfr_painel == 'CFR21 Part 11 - Qualificável':
                txt_com += "O sistema fornecido possuirá as licenças e os parâmetros necessários para ser totalmente **Qualificável (CFR 21 Part 11)**. A SIARCON garantirá todos os requisitos técnicos de software, deixando o ambiente pronto para que a qualificação final seja realizada por empresa à escolha do cliente.\n\n"
            elif tipo_cfr_painel == 'CFR21 Part 11 - Qualificado':
                txt_com += "O sistema contemplado será integralmente **Qualificado (CFR 21 Part 11)**. A SIARCON executará e entregará toda a documentação comprobatória e a execução dos protocolos e validações pertinentes, garantindo a certificação total do ambiente de supervisão.\n\n"
            
            txt_com += "O quadro de automação será responsável pela aquisição de dados e controle da seguinte instrumentação de campo:\n\n"
            
            for desc_inst, qt_inst, func_inst in lista_instrumentos_detalhados:
                txt_com += f"• **{qt_inst}x {func_inst}:** {desc_inst}\n"
                
            txt_com += "\n**Lógica de Operação do Sistema:**\n"
            txt_com += "O sistema realizará o controle da vazão de ar de forma constante, efetuando os ajustes necessários no inversor para que a vazão volumétrica seja mantida, independentemente do nível de saturação dos filtros no tempo. O controlador modulará proporcionalmente a válvula da serpentina (ou estágios do compressor) para atingir os parâmetros exatos de setpoint térmico demandados pelo ambiente."
            
            if tem_resistencia:
                txt_com += " A resistência elétrica de aquecimento será acionada por malha de controle PID dedicada, permitindo ajuste fino de temperatura e desumidificação, possuindo intertravamento de segurança via termostato mecânico de proteção e confirmação de fluxo de ar."
                
            descritivo_comercial_linhas.append(txt_com)

        texto_descritivo_final = "\n\n----------------------------------------------------\n\n".join(descritivo_linhas_excel)

        # Trata o Software e as Regras de Cobrança do CFR-21
        for (s_name, t_cfr), pts_total in softwares_incluidos.items():
            b_k, p_k = "", ""
            if "SEM certificação" in s_name: b_k, p_k = "Licença Supervisório - SEM CFR-21 (Base)", "Licença Supervisório - SEM CFR-21 (Por Ponto I/O)"
            elif "COM certificação" in s_name: b_k, p_k = "Licença Supervisório - COM CFR-21 (Base)", "Licença Supervisório - COM CFR-21 (Por Ponto I/O)"
            else: b_k, p_k = "Licença Supervisório - Schneider EBO (Base)", "Licença Supervisório - Schneider EBO (Por Ponto I/O)"
            
            p_base = st.session_state.precos_banco.get(b_k, 23000.0)
            p_pto = st.session_state.precos_banco.get(p_k, 100.0)
            linhas_software.append({"Categoria": "Software de Supervisão", "Item": f"Licença Base: {s_name}", "Preço Unit.": p_base, "Qtd": 1, "Custo Total": p_base})
            if p_pto > 0 and pts_total > 0:
                linhas_software.append({"Categoria": "Software de Supervisão", "Item": f"Pontos Licenciados no Software ({pts_total} canais ativos)", "Preço Unit.": p_pto, "Qtd": pts_total, "Custo Total": pts_total * p_pto})

            # Lógica de Cobrança do Serviço CFR-21 (Registrado em Serviços de Lógica)
            if "COM certificação" in s_name:
                custo_cfr_unit = 0.0
                if t_cfr == "CFR21 Part 11 - Qualificável":
                    if pts_total <= 100: custo_cfr_unit = st.session_state.precos_banco.get("CFR21 Qualificável - Até 100 pts", 70.0)
                    elif pts_total <= 250: custo_cfr_unit = st.session_state.precos_banco.get("CFR21 Qualificável - 101 a 250 pts", 50.0)
                    else: custo_cfr_unit = st.session_state.precos_banco.get("CFR21 Qualificável - Acima de 250 pts", 30.0)
                    
                    linhas_servicos.append({"Categoria": "Serviços de Lógica", "Item": f"Preparação do Sistema Qualificável (CFR21) - Por Ponto", "Preço Unit.": custo_cfr_unit, "Qtd": pts_total, "Custo Total": pts_total * custo_cfr_unit})
                    
                elif t_cfr == "CFR21 Part 11 - Qualificado":
                    if pts_total <= 30: custo_cfr_unit = st.session_state.precos_banco.get("CFR21 Qualificado - Até 30 pts", 400.0)
                    elif pts_total <= 60: custo_cfr_unit = st.session_state.precos_banco.get("CFR21 Qualificado - 31 a 60 pts", 350.0)
                    elif pts_total <= 99: custo_cfr_unit = st.session_state.precos_banco.get("CFR21 Qualificado - 61 a 99 pts", 320.0)
                    elif pts_total <= 150: custo_cfr_unit = st.session_state.precos_banco.get("CFR21 Qualificado - 100 a 150 pts", 290.0)
                    elif pts_total <= 200: custo_cfr_unit = st.session_state.precos_banco.get("CFR21 Qualificado - 151 a 200 pts", 250.0)
                    elif pts_total <= 250: custo_cfr_unit = st.session_state.precos_banco.get("CFR21 Qualificado - 201 a 250 pts", 220.0)
                    else: custo_cfr_unit = st.session_state.precos_banco.get("CFR21 Qualificado - Acima de 250 pts", 200.0)

                    linhas_servicos.append({"Categoria": "Serviços de Lógica", "Item": f"Execução de Qualificação Integral (CFR21) - Por Ponto", "Preço Unit.": custo_cfr_unit, "Qtd": pts_total, "Custo Total": pts_total * custo_cfr_unit})

        for infra_avulsa in st.session_state.orcamento:
            linhas_inst_campo.append({"Categoria": "Instrumentação de Campo", "Item": infra_avulsa['Item'], "Preço Unit.": infra_avulsa['Custo_Total']/infra_avulsa['Quantidade'] if infra_avulsa['Quantidade'] > 0 else 0, "Qtd": infra_avulsa['Quantidade'], "Custo Total": infra_avulsa['Custo_Total']})

        df_inst = pd.DataFrame(linhas_inst_campo)
        df_hw = pd.DataFrame(linhas_hardware)
        df_sw = pd.DataFrame(linhas_software)
        
        subtotal_inst = df_inst['Custo Total'].sum() if not df_inst.empty else 0
        subtotal_hw = df_hw['Custo Total'].sum() if not df_hw.empty else 0
        subtotal_sw = df_sw['Custo Total'].sum() if not df_sw.empty else 0
        
        if (total_ai_schneider + total_ao_schneider) > 0:
            pr_ai_sch = st.session_state.precos_banco.get("Custo AI/AO", 565.0)
            linhas_servicos.append({"Categoria": "Serviços de Lógica", "Item": "Serviços de lógica: Pontos Analógicos", "Preço Unit.": pr_ai_sch, "Qtd": (total_ai_schneider + total_ao_schneider), "Custo Total": (total_ai_schneider + total_ao_schneider) * pr_ai_sch})
        if (total_di_schneider + total_do_schneider) > 0:
            pr_di_sch = st.session_state.precos_banco.get("Custo DI/DO", 120.0)
            linhas_servicos.append({"Categoria": "Serviços de Lógica", "Item": "Serviços de lógica: Pontos Digitais", "Preço Unit.": pr_di_sch, "Qtd": (total_di_schneider + total_do_schneider), "Custo Total": (total_di_schneider + total_do_schneider) * pr_di_sch})
        if (total_ai_siemens + total_ao_siemens) > 0:
            pr_ai_siem = st.session_state.precos_banco.get("Siemens - Serviço Custo AI/AO", 750.0)
            linhas_servicos.append({"Categoria": "Serviços de Lógica", "Item": "Serviços de lógica (Siemens): Pontos Analógicos", "Preço Unit.": pr_ai_siem, "Qtd": (total_ai_siemens + total_ao_siemens), "Custo Total": (total_ai_siemens + total_ao_siemens) * pr_ai_siem})
        if (total_di_siemens + total_do_siemens) > 0:
            pr_di_siem = st.session_state.precos_banco.get("Siemens - Serviço Custo DI/DO", 180.0)
            linhas_servicos.append({"Categoria": "Serviços de Lógica", "Item": "Serviços de lógica (Siemens): Pontos Digitais", "Preço Unit.": pr_di_siem, "Qtd": (total_di_siemens + total_do_siemens), "Custo Total": (total_di_siemens + total_do_siemens) * pr_di_siem})
        if total_io_mercato > 0:
            pr_serv_merc = st.session_state.precos_banco.get("Mercato - Serviço Parametrização por Ponto", 80.0)
            linhas_servicos.append({"Categoria": "Serviços de Lógica", "Item": "Parametrização de Pontos (Mercato)", "Preço Unit.": pr_serv_merc, "Qtd": total_io_mercato, "Custo Total": total_io_mercato * pr_serv_merc})

        mao_de_obra_extra = (custo_base_schneider * 0.25) + (custo_base_siemens * 0.35) + (custo_base_mercato * 0.10)
        if mao_de_obra_extra > 0:
            linhas_servicos.append({"Categoria": "Serviços de Lógica", "Item": "Demais programações e desenvolvimentos", "Preço Unit.": mao_de_obra_extra, "Qtd": 1, "Custo Total": mao_de_obra_extra})
            
        df_serv = pd.DataFrame(linhas_servicos)
        subtotal_serv = df_serv['Custo Total'].sum() if not df_serv.empty else 0
        total_geral = subtotal_inst + subtotal_hw + subtotal_sw + subtotal_serv
        
        if total_geral > 0:
            st.markdown("### Resumo Estruturado")
            c1, c2, c3 = st.columns(3)
            c1.info(f"**Subtotal Instrumentação:**\nR$ {subtotal_inst:,.2f}")
            c2.warning(f"**Subtotal Hardware:**\nR$ {subtotal_hw:,.2f}")
            c3.success(f"**CUSTO TOTAL ESTIMADO:**\nR$ {total_geral:,.2f}")

            df_export_list = []
            if not df_inst.empty:
                df_export_list.append(df_inst.groupby(['Categoria', 'Item'], as_index=False).agg({'Preço Unit.': 'first', 'Qtd': 'sum', 'Custo Total': 'sum'}))
                df_export_list.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "INSTRUMENTAÇÃO DE CAMPO", "Preço Unit.": "-", "Qtd": "-", "Custo Total": subtotal_inst}]))
            if not df_hw.empty:
                df_export_list.append(df_hw.groupby(['Categoria', 'Item'], as_index=False).agg({'Preço Unit.': 'first', 'Qtd': 'sum', 'Custo Total': 'sum'}))
                df_export_list.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "HARDWARE E PAINÉIS", "Preço Unit.": "-", "Qtd": "-", "Custo Total": subtotal_hw}]))
            if not df_sw.empty:
                df_export_list.append(df_sw.groupby(['Categoria', 'Item'], as_index=False).agg({'Preço Unit.': 'first', 'Qtd': 'sum', 'Custo Total': 'sum'}))
                df_export_list.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "SOFTWARE", "Preço Unit.": "-", "Qtd": "-", "Custo Total": subtotal_sw}]))
            if not df_serv.empty:
                df_export_list.append(df_serv)
                df_export_list.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "SERVIÇOS E LÓGICA", "Preço Unit.": "-", "Qtd": "-", "Custo Total": subtotal_serv}]))
            df_export_list.append(pd.DataFrame([{"Categoria": "TOTAL GERAL", "Item": "ORÇAMENTO COMPLETO", "Preço Unit.": "-", "Qtd": "-", "Custo Total": total_geral}]))
            df_exportacao = pd.concat(df_export_list, ignore_index=True)
            
            def format_currency(val):
                try: return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                except: return val
            df_display = df_exportacao.copy()
            df_display['Preço Unit.'] = df_display['Preço Unit.'].apply(format_currency)
            df_display['Custo Total'] = df_display['Custo Total'].apply(format_currency)
            st.dataframe(df_display, use_container_width=True)
            
            # --- EXPANDER PARA O DESCRITIVO COMERCIAL ---
            with st.expander("📄 Gerar Descritivo Detalhado para Proposta Comercial", expanded=False):
                st.markdown("<div style='background-color:#E3F2FD; padding:20px; border-radius:10px;'>", unsafe_allow_html=True)
                for t_com in descritivo_comercial_linhas:
                    st.markdown(t_com)
                    st.markdown("<hr style='border:1px solid #B0BEC5'>", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
            
            df_pontos = pd.DataFrame(linhas_pontos)
            if not df_pontos.empty:
                total_qtd = pd.to_numeric(df_pontos['Quantidade Total'], errors='coerce').fillna(0).sum()
                linha_total = pd.DataFrame([{"Painel": "TOTAL GERAL", "Grupo/Equipamento": "-", "Instrumento": "-", "Quantidade Total": total_qtd, "Entrada Digital (DI)": df_pontos['Entrada Digital (DI)'].sum(), "Saída Digital (DO)": df_pontos['Saída Digital (DO)'].sum(), "Entrada Analógica (AI)": df_pontos['Entrada Analógica (AI)'].sum(), "Saída Analógica (AO)": df_pontos['Saída Analógica (AO)'].sum()}])
                df_pontos = pd.concat([df_pontos, linha_total], ignore_index=True)

            buffer = io.BytesIO()
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "Detalhamento Financeiro"
            ws1.views.sheetView[0].showGridLines = True
            
            # --- NOVO CABEÇALHO DO EXCEL (Mais largo) ---
            ws1.row_dimensions[1].height = 35
            ws1.row_dimensions[2].height = 25
            ws1.row_dimensions[3].height = 25
            ws1.row_dimensions[4].height = 25
            
            nome_projeto_header = st.session_state.nome_projeto_orcamento if st.session_state.nome_projeto_orcamento else "PROJETO NÃO NOMEADO"
            
            ws1.merge_cells("C1:E1")
            ws1.cell(row=1, column=3, value="DESCRIÇÃO TÉCNICA E ORÇAMENTÁRIA DE SISTEMAS DE AUTOMAÇÃO").font = Font(name="Arial", size=12, bold=True, color="1C8590")
            ws1.cell(row=1, column=3).alignment = Alignment(horizontal="center", vertical="center")
            
            ws1.merge_cells("C2:E2"); ws1.cell(row=2, column=3, value=f"PROJETO: {nome_projeto_header.upper()}").font = Font(name="Arial", size=10, bold=True, color="333333")
            ws1.merge_cells("C3:E3"); ws1.cell(row=3, column=3, value=f"DATA/HORA EMISSÃO: {datetime.now(fuso_br).strftime('%d/%m/%Y %H:%M:%S')}").font = Font(name="Arial", size=10, color="555555")
            ws1.merge_cells("C4:E4"); ws1.cell(row=4, column=3, value=f"RESPONSÁVEL TÉCNICO: {st.session_state.nome_exibicao.upper()}").font = Font(name="Arial", size=10, color="555555")
            
            for r in range(2, 5): ws1.cell(row=r, column=3).alignment = Alignment(horizontal="center", vertical="center")
            
            start_row = 6
            if ARQUIVO_LOGO:
                try:
                    from openpyxl.drawing.image import Image as OpenpyxlImage
                    img = OpenpyxlImage(ARQUIVO_LOGO)
                    img.width = 180
                    img.height = 50
                    ws1.add_image(img, "A1")
                except: pass
                
            fill_header = PatternFill(start_color="1C8590", end_color="1C8590", fill_type="solid")
            font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
            border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
            
            for r_idx, row in enumerate(dataframe_to_rows(df_exportacao, index=False, header=True), start=start_row):
                for c_idx, value in enumerate(row, start=1):
                    cell = ws1.cell(row=r_idx, column=c_idx, value=value)
                    cell.border = border_thin
                    if r_idx == start_row:
                        cell.fill = fill_header; cell.font = font_header; cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        is_subtotal = "SUBTOTAL" in str(ws1.cell(row=r_idx, column=1).value) or "TOTAL GERAL" in str(ws1.cell(row=r_idx, column=1).value)
                        cell.font = Font(name="Arial", size=10, bold=is_subtotal)
                        if is_subtotal: cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                        if c_idx in [3, 5]: 
                            if str(value).strip() != "-": 
                                try: cell.value = float(value); cell.number_format = '"R$" #,##0.00'
                                except: pass
                            cell.alignment = Alignment(horizontal="right")
                        elif c_idx == 4: cell.alignment = Alignment(horizontal="center")
            
            # --- CAIXA DE TEXTO DESCRITIVO NO EXCEL ---
            end_row_table = start_row + len(df_exportacao) + 2
            num_linhas_texto = len(texto_descritivo_final.split('\n'))
            tamanho_caixa = max(10, num_linhas_texto + 2) 
            
            ws1.merge_cells(start_row=end_row_table, start_column=1, end_row=end_row_table+tamanho_caixa, end_column=5)
            cell_desc = ws1.cell(row=end_row_table, column=1, value=texto_descritivo_final)
            cell_desc.font = Font(name="Arial", size=10, italic=False, color="333333")
            cell_desc.alignment = Alignment(vertical="top", wrap_text=True)
            cell_desc.fill = PatternFill(start_color="F2F4F4", end_color="F2F4F4", fill_type="solid")
            
            for r in range(end_row_table, end_row_table+tamanho_caixa+1):
                for c in range(1, 6): ws1.cell(row=r, column=c).border = border_thin
            
            for col in ws1.columns:
                max_len = max(len(str(cell.value or '')) for cell in col if cell.row <= start_row + len(df_exportacao))
                ws1.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 12)

            if not df_pontos.empty:
                ws2 = wb.create_sheet(title="Matriz de Pontos (IO)")
                ws2.views.sheetView[0].showGridLines = True
                for r_idx, row in enumerate(dataframe_to_rows(df_pontos, index=False, header=True), start=1):
                    for c_idx, value in enumerate(row, start=1):
                        cell = ws2.cell(row=r_idx, column=c_idx, value=value)
                        cell.border = border_thin
                        if r_idx == 1: cell.fill = fill_header; cell.font = font_header; cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            cell.font = Font(name="Arial", size=10)
                            if c_idx >= 4: cell.alignment = Alignment(horizontal="center")
                for col in ws2.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    ws2.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 12)
            
            wb.save(buffer)
            buffer.seek(0)
            st.download_button(label="📥 Exportar Orçamento Final para Excel", data=buffer.getvalue(), file_name="orcamento_dimensionado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.markdown("---")
            if st.button("☁️ Salvar Orçamento Final e Gerar Revisão", type="primary", use_container_width=True):
                if not st.session_state.nome_projeto_orcamento: st.warning("⚠️ Atenção: Preencha o 'Nome do Orçamento / Projeto' antes de salvar.")
                else:
                    try:
                        sh = conectar_google_sheets()
                        try: ws_hist_orc = sh.worksheet("Historico_Orcamentos")
                        except:
                            ws_hist_orc = sh.add_worksheet(title="Historico_Orcamentos", rows="1000", cols="8")
                            ws_hist_orc.append_row(["Data/Hora", "Nome do Projeto", "Revisão", "Subtotal Hardware", "Serviços de Lógica", "Custo Total Estimado", "Configuracao_JSON", "Usuário"])
                        todas_linhas = ws_hist_orc.get_all_values()
                        contagem_revisoes = sum(1 for r in todas_linhas[1:] if r[1].strip().upper() == st.session_state.nome_projeto_orcamento.strip().upper())
                        revisao_atual = f"R-{contagem_revisoes:02d}"
                        nova_linha = [datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M:%S"), st.session_state.nome_projeto_orcamento, revisao_atual, f"R$ {subtotal_hw:.2f}".replace('.', ','), f"R$ {subtotal_serv:.2f}".replace('.', ','), f"R$ {total_geral:.2f}".replace('.', ','), json.dumps(st.session_state.paineis_auto), st.session_state.usuario_logado]
                        ws_hist_orc.append_row(nova_linha)
                        st.cache_data.clear() # Limpa o cache para que a tabela do banco recarregue no visual
                        st.success(f"✅ Sucesso! Orçamento para '{st.session_state.nome_projeto_orcamento}' salvo com a revisão {revisao_atual}!")
                    except Exception as e: st.error(f"Erro ao salvar: {e}")
