import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
from datetime import date, datetime, timezone, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import google.generativeai as genai
from PIL import Image

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="App SIARCON - Propostas e Custos", layout="wide", page_icon="📄")

# ==========================================
# FUNÇÃO CAÇA-LOGO (Evita erros de nome de arquivo)
# ==========================================
def buscar_logo():
    nomes_possiveis = ["SIARCON.png", "SIARCON .png", "siarcon.png", "Siarcon.png", "logo.png"]
    for nome in nomes_possiveis:
        if os.path.exists(nome):
            return nome
    return None

ARQUIVO_LOGO = buscar_logo()

# ==========================================
# 🟢 CONEXÃO COM O GOOGLE SHEETS E IA
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
        st.error(f"Erro na conexão com Google Sheets: {e}. Verifique o link da planilha.")
        st.stop()

# Configuração da Inteligência Artificial (Google Gemini)
try:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model_ia = genai.GenerativeModel('gemini-1.5-flash')
    ia_disponivel = True
except Exception as e:
    ia_disponivel = False
    erro_ia = e

# Define o fuso horário de Brasília (UTC-3)
fuso_br = timezone(timedelta(hours=-3))

# ==========================================
# 🔐 CONTROLE DE ACESSO E LOGIN POR PERFIL
# ==========================================
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None
if "nome_exibicao" not in st.session_state:
    st.session_state.nome_exibicao = ""
if "menu_selecionado" not in st.session_state:
    st.session_state.menu_selecionado = "🏠 Tela Inicial"

# Se não estiver logado, exibe a tela de login
if st.session_state.usuario_logado is None:
    st.markdown("""
        <style>
        /* Puxa tudo o máximo para cima possível */
        .block-container {
            padding-top: 0rem !important;
            margin-top: -2rem !important;
        }
        
        /* Oculta os cabeçalhos nativos e EVITA bloqueio de clique com display:none */
        header {display: none !important;}
        [data-testid="collapsedControl"] {display: none !important;}
        
        /* Gradiente de fundo */
        .stApp {
            background: linear-gradient(135deg, #1C8590 0%, #8FD3B5 100%) !important;
        }
        
        /* Estiliza o Form do Streamlit para virar a "Caixa Branca" */
        [data-testid="stForm"] {
            background-color: white;
            border-radius: 12px;
            padding: 30px;
            box-shadow: 0 8px 24px rgba(0,0,0,0.15);
            border: none;
            position: relative;
            z-index: 100;
        }
        
        /* Ajuste dos botões dentro do Form */
        [data-testid="stFormSubmitButton"] button {
            background-color: #2b7bc4 !important;
            color: white !important;
            font-weight: bold !important;
            border-radius: 6px !important;
            height: 45px !important;
            border: none !important;
            margin-top: 15px !important;
        }
        [data-testid="stFormSubmitButton"] button:hover {
            background-color: #1a5c96 !important;
        }
        
        /* Estiliza os inputs de texto */
        input {
            border-bottom: 2px solid #ccc !important;
            border-top: none !important;
            border-left: none !important;
            border-right: none !important;
            border-radius: 0 !important;
            background-color: transparent !important;
            box-shadow: none !important;
            padding-left: 0px !important;
        }
        input:focus {
            border-bottom: 2px solid #1C8590 !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # Layout de 3 colunas para manter a caixa centralizada
    col1, col2, col3 = st.columns([1, 1.2, 1])

    with col2:
        # Centraliza o logo fora do form
        st.write("")
        col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
        with col_img2:
            if ARQUIVO_LOGO:
                st.image(ARQUIVO_LOGO, use_container_width=True)
            else:
                st.markdown("<h2 style='text-align: center; color: white; margin-bottom:0;'>SIARCON</h2>", unsafe_allow_html=True)
        
        # Caixa branca interativa e segura (Nativa do Streamlit)
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
                    "giovanna.ribeiro": "1234",
                    "aline.ferraz": "1234",
                    "janaina.dias": "1234",
                    "victor.hugo": "1234",
                    "rodrigo.ribeiro": "1234",
                    "engenharia": "1234",
                    "suprimentos": "1234",
                    "obras": "1234"
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
                else:
                    st.error("❌ Usuário ou senha incorretos.")
                    
    st.stop() # Bloqueia o carregamento do resto do site enquanto não houver login

# Restaura o padding normal e mostra o cabeçalho para a aplicação principal
st.markdown("""
    <style>
    .block-container { padding-top: 3rem !important; }
    header {display: flex !important;}
    [data-testid="collapsedControl"] {display: flex !important;}
    </style>
""", unsafe_allow_html=True)

# === MENU LATERAL PRINCIPAL ===
if ARQUIVO_LOGO:
    st.sidebar.image(ARQUIVO_LOGO, use_container_width=True)
else:
    st.sidebar.markdown("### SIARCON")
    
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
menu_ui = st.sidebar.radio(
    "Módulos do Sistema",
    opcoes_menu,
    index=opcoes_menu.index(st.session_state.menu_selecionado)
)

st.sidebar.markdown("---")

# Atualiza a sessão de navegação de forma reativa
if menu_ui != st.session_state.menu_selecionado:
    st.session_state.menu_selecionado = menu_ui
    st.rerun()

# ==============================================================================
# TELA 0: HOME / TELA INICIAL (DASHBOARD)
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
    modo_preenchimento = st.sidebar.radio(
        "Como deseja preencher o Escopo Técnico?",
        ["📋 Preenchimento Manual", "📊 Automático (Excel)"]
    )

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

    try:
        df_escopos, df_exclusoes, df_resp, df_clientes, df_coberturas = carregar_dados_propostas()
    except Exception as e:
        st.error("Erro ao ler banco de dados. Verifique a planilha.")
        st.stop()

    def formatar_data_portugues(dt):
        meses = {1:'Janeiro', 2:'Fevereiro', 3:'Março', 4:'Abril', 5:'Maio', 6:'Junho',
                 7:'Julho', 8:'Agosto', 9:'Setembro', 10:'Outubro', 11:'Novembro', 12:'Dezembro'}
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

    with st.expander("➕ Cadastrar Responsabilidade"):
        with st.form("nova_resp"):
            nr_c, nr_l = st.text_input("Título Curto"), st.text_input("Texto Completo")
            if st.form_submit_button("💾 Salvar") and nr_c and nr_l:
                salvar_no_banco("Responsabilidades", [nr_c, nr_l])
                st.rerun()

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
        
        with st.expander("➕ Cadastrar NOVO Item de Escopo no Banco"):
            with st.form("novo_esc"):
                cats_existentes = sorted(df_escopos['Categoria'].unique().tolist()) if 'Categoria' in df_escopos.columns else []
                c_cat, c_tit, c_txt = st.columns([0.3, 0.3, 0.4])
                opcao_cat = c_cat.selectbox("Categoria", ["Nova Categoria..."] + cats_existentes)
                cat_final = c_cat.text_input("Nome da Categoria") if opcao_cat == "Nova Categoria..." else opcao_cat
                ne_tit, ne_txt = c_tit.text_input("Título Curto"), c_txt.text_input("Texto Completo")
                if st.form_submit_button("💾 Salvar Item") and cat_final and ne_tit and ne_txt:
                    salvar_no_banco("Escopos", [cat_final, ne_tit, ne_txt])
                    st.rerun()

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
                
                if df_orc.shape[1] < 14:
                    df_orc = df_orc.reindex(columns=list(range(14)))

                for idx, row in df_orc.iterrows():
                    texto_col_e = str(row[4]).upper().strip() if 4 < len(row) else ""
                    texto_col_c = str(row[2]).upper().strip() if 2 < len(row) else ""
                    
                    if "PREÇO VENDA" in texto_col_e or "PRECO VENDA" in texto_col_e or "PREÇO VENDA" in texto_col_c or "PRECO VENDA" in texto_col_c:
                        val_bruto = row[9] 
                        if pd.notna(val_bruto):
                            if isinstance(val_bruto, (int, float)):
                                valor_total_calculado = float(val_bruto)
                            else:
                                val_str = str(val_bruto).upper().replace("R$", "").strip()
                                val_str = val_str.replace(".", "").replace(",", ".")
                                try: valor_total_calculado = float(val_str)
                                except: pass
                        break

                categoria_atual_nome = "ESCOPO GERAL"
                categoria_atual_indice = ""
                itens_detalhados = []
                itens_eap = []
                contador_item = 1
                
                for index, row in df_orc.iterrows():
                    descricao = str(row[2]).strip() if 2 < len(row) else ""
                    if "CUSTO INDIRETO" in descricao.upper() or "CUSTOS INDIRETOS" in descricao.upper(): break 
                    
                    col_b = str(row[1]).strip() if 1 < len(row) else ""
                    unidade = str(row[3]).strip() if 3 < len(row) else ""
                    quantidade = row[4] if 4 < len(row) else pd.NA
                    
                    if pd.isna(descricao) or descricao == "" or descricao.upper() in ["NAN", "DESCRIÇÃO DOS MATERIAIS", "DESCRICAO DOS MATERIAIS", "DESCRIÇÃO", "ITEM"]:
                        continue
                        
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
                            if isinstance(quantidade, (int, float)):
                                qtd_fmt = int(quantidade) if float(quantidade).is_integer() else round(float(quantidade), 2)
                            else:
                                try: 
                                    q_str = str(quantidade).replace(",", ".")
                                    qtd_fmt = int(float(q_str)) if float(q_str).is_integer() else round(float(q_str), 2)
                                except: qtd_fmt = quantidade
                                    
                            uni_fmt = f" {unidade}" if unidade.lower() not in ["nan", "", "-"] else ""
                            nome_resumido = descricao.split('|')[0].strip()
                            if len(nome_resumido) > 80: nome_resumido = nome_resumido[:80] + "..."
                            
                            if "AMORTECEDOR" not in descricao.upper():
                                indice_item = f"{categoria_atual_indice}.{contador_item}" if categoria_atual_indice else str(contador_item)
                                itens_eap.append({'indice': indice_item, 'nome': nome_resumido})
                                contador_item += 1 
                            
                            texto_item = f"Fornecimento / Instalação de {qtd_fmt}{uni_fmt} - {descricao}."
                            itens_detalhados.append(texto_item)
                        else:
                            if len(itens_detalhados) > 0:
                                itens_detalhados[-1] += f"\n\n{descricao}"
                
                if categoria_atual_nome != "ESCOPO GERAL" and len(itens_detalhados) > 0:
                    escopo_estruturado.append({'nome': f"{categoria_atual_indice} - {categoria_atual_nome}".strip(' -'), 'itens': itens_detalhados})
                    eap_estruturada.append({'indice': categoria_atual_indice, 'categoria': categoria_atual_nome.upper(), 'itens': itens_eap})
                    
                if len(escopo_estruturado) > 0: st.success(f"✅ Planilha processada com sucesso.")
            except Exception as e:
                st.error(f"Erro ao processar a planilha: {e}")

    st.markdown("---")
    st.header("5. Exclusões")

    with st.expander("➕ Cadastrar Exclusão"):
        with st.form("nova_exc"):
            nex_c, nex_l = st.text_input("Título Curto"), st.text_input("Texto Completo")
            if st.form_submit_button("💾 Salvar") and nex_c and nex_l:
                salvar_no_banco("Exclusoes", [nex_c, nex_l])
                st.rerun()

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
        except Exception as e:
            st.error(f"Erro ao gerar o Word: {e}")


# ==============================================================================
# MÓDULO 2: SISTEMA DE AUTOMAÇÃO
# ==============================================================================
elif st.session_state.menu_selecionado == "🔌 Levantamento de Automação":
    
    st.title("🔌 Engenharia e Custos - Automação e Infra")
    st.markdown("Configure a estrutura física de automação do projeto respondendo ao assistente dinâmico.")
    
    # === INICIALIZAÇÃO SEGURA DE VARIÁVEIS ===
    if 'nome_projeto_orcamento' not in st.session_state: st.session_state.nome_projeto_orcamento = ""
    if 'projeto_para_abrir' not in st.session_state: st.session_state.projeto_para_abrir = None
    if 'dados_projeto_abrir' not in st.session_state: st.session_state.dados_projeto_abrir = {}
    if 'orcamento' not in st.session_state: st.session_state.orcamento = []
    if 'historico_precos' not in st.session_state: st.session_state.historico_precos = []
    if 'wizard_ativo' not in st.session_state: st.session_state.wizard_ativo = False
    if 'paineis_auto' not in st.session_state: st.session_state.paineis_auto = []
        
    nome_proj = st.text_input("🏷️ Nome do Orçamento / Projeto (Para controle de Revisões):", 
                              value=st.session_state.nome_projeto_orcamento,
                              placeholder="Ex: Instalação Farmacêutica Bloco B")
    st.session_state.nome_projeto_orcamento = nome_proj
    st.markdown("---")

    # ==========================================
    # DATABASE DE PREÇOS E REGRAS SIARCON
    # ==========================================
    REGRA_IO = {
        "Transmissor de pressão Dif. Para ar (Vazão de ar)": {"AI": 1, "AO": 1, "DI": 1, "DO": 1},
        "Transmissor de temperatura e umidade para duto": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Válvula de controle proporcional com atuador": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Relé de Corrente - Status Compressor": {"AI": 0, "AO": 0, "DI": 1, "DO": 2},
        "Termostato de segurança (Proteção da resistência)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Pressostato diferencial para ar (Proteção da resistência)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Transmissor de temperatura para duto": {"AI": 1, "AO": 1, "DI": 0, "DO": 0},
        "Resistência de aquecimento (Equipamento)": {"AI": 0, "AO": 1, "DI": 2, "DO": 1},
        "Resistência de aquecimento (Duto)": {"AI": 0, "AO": 1, "DI": 2, "DO": 1},
        "Válvula de controle de água gelada proporcional": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula de controle de água quente proporcional": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula de controle de vapor proporcional": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula motorizada Bypass Proporcional (até 2.1/2\")": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula motorizada Bypass Proporcional (3\" ou 4\")": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula motorizada Bypass Proporcional (5\")": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula motorizada Bypass Proporcional (6\")": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula motorizada Bypass Proporcional (8\")": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Transmissor de pressão para água": {"AI": 1, "AO": 2, "DI": 0, "DO": 0},
        "Tranmissor de vazão para água": {"AI": 1, "AO": 1, "DI": 0, "DO": 0},
        "Válvula bloqueio motorizada": {"AI": 0, "AO": 0, "DI": 0, "DO": 1},
        "Chave de fluxo": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Bombas (I/O para controlador)": {"AI": 0, "AO": 1, "DI": 1, "DO": 1},
        "Tanques (I/O para controlador)": {"AI": 1, "AO": 0, "DI": 1, "DO": 1},
        "Pressostato - Filtro G4": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Pressostato - Filtro F9": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Pressostato - Filtro H13/H14": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Status funcionamento ventilador ou exaustor (partida direta)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Transmissor de pressão diferencial - Filtro G4": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de pressão diferencial - Filtro F9": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Tranmissor de pressão diferencial - Filtro H13": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de pressão diferencial entre salas": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de pressão diferencial entre salas com display": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura Ambiente": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura ambiente com display": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura e umidade ambiente": {"AI": 2, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura e umidade ambiente com display": {"AI": 2, "AO": 0, "DI": 0, "DO": 0}
    }

    banco_padrao_precos = {
        "Transmissor de pressão Dif. Para ar (Vazão de ar)": 1490.00,
        "Transmissor de temperatura e umidade para duto": 2050.00,
        "Válvula de controle proporcional com atuador": 0.00,
        "Relé de Corrente - Status Compressor": 150.00,
        "Termostato de segurança (Proteção da resistência)": 250.00,
        "Pressostato diferencial para ar (Proteção da resistência)": 349.00,
        "Transmissor de temperatura para duto": 800.00,
        "Resistência de aquecimento (Equipamento)": 0.00,
        "Resistência de aquecimento (Duto)": 0.00,
        "Válvula de controle de água gelada proporcional": 0.00,
        "Válvula de controle de água quente proporcional": 0.00,
        "Válvula de controle de vapor proporcional": 0.00,
        "Válvula motorizada Bypass Proporcional (até 2.1/2\")": 2690.00,
        "Válvula motorizada Bypass Proporcional (3\" ou 4\")": 4950.00,
        "Válvula motorizada Bypass Proporcional (5\")": 6450.00,
        "Válvula motorizada Bypass Proporcional (6\")": 7900.00,
        "Válvula motorizada Bypass Proporcional (8\")": 9200.00,
        "Transmissor de pressão para água": 1359.00,
        "Tranmissor de vazão para água": 3550.00,
        "Válvula bloqueio motorizada": 0.00,
        "Chave de fluxo": 349.00,
        "Bombas (I/O para controlador)": 0.00,
        "Tanques (I/O para controlador)": 0.00,
        "Pressostato - Filtro G4": 349.00,
        "Pressostato - Filtro F9": 349.00,
        "Pressostato - Filtro H13/H14": 349.00,
        "Status funcionamento ventilador ou exaustor (partida direta)": 349.00,
        "Transmissor de pressão diferencial - Filtro G4": 1490.00,
        "Transmissor de pressão diferencial - Filtro F9": 1490.00,
        "Tranmissor de pressão diferencial - Filtro H13": 1490.00,
        "Transmissor de pressão diferencial entre salas": 1490.00,
        "Transmissor de pressão diferencial entre salas com display": 2110.00,
        "Transmissor de temperatura Ambiente": 2050.00,
        "Transmissor de temperatura ambiente com display": 2650.00,
        "Transmissor de temperatura e umidade ambiente": 2050.00,
        "Transmissor de temperatura e umidade ambiente com display": 2650.00,
        "MP-C-15A": 4649.49, "MP-C-18A": 5185.54, "MP-C-24A": 7290.75, "MP-C-36A": 9459.08,
        "Custo AI/AO": 565.00, "Custo DI/DO": 120.00,
        "Licença Supervisório - SEM CFR-21 (Base)": 23000.00,
        "Licença Supervisório - SEM CFR-21 (Por Ponto I/O)": 100.00,
        "Licença Supervisório - COM CFR-21 (Base)": 23000.00,
        "Licença Supervisório - COM CFR-21 (Por Ponto I/O)": 285.00,
        "Licença Supervisório - Schneider EBO (Base)": 13000.00,
        "Licença Supervisório - Schneider EBO (Por Ponto I/O)": 110.00
    }

    OPCOES_SUPERVISAO = {
        "Sem Supervisório": {"base": 0.0, "por_ponto": 0.0},
        "Sistema supervisório SEM certificação CFR-21": {"base": 23000.0, "por_ponto": 100.0},
        "Sistema supervisório COM certificação CFR-21": {"base": 23000.0, "por_ponto": 285.0},
        "Sistema de monitoramento Schneider EBO": {"base": 13000.0, "por_ponto": 110.0}
    }

    if 'banco_precos_carregado' not in st.session_state:
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
        st.session_state.banco_precos_carregado = True

    # Trava de atualização de preços caso novos itens não existam na nuvem
    for k_n, v_n in banco_padrao_precos.items():
        if k_n not in st.session_state.precos_banco: 
            st.session_state.precos_banco[k_n] = v_n

    GRUPOS_INSTRUMENTOS = {
        "🔹 Controle (HVAC e Máquinas)": [
            "Transmissor de pressão Dif. Para ar (Vazão de ar)",
            "Transmissor de temperatura e umidade para duto", "Transmissor de temperatura para duto",
            "Válvula de controle proporcional com atuador", "Válvula de controle de água gelada proporcional",
            "Válvula de controle de água quente proporcional", "Válvula de controle de vapor proporcional",
            "Relé de Corrente - Status Compressor", "Termostato de segurança (Proteção da resistência)",
            "Pressostato diferencial para ar (Proteção da resistência)", "Resistência de aquecimento (Equipamento)", "Resistência de aquecimento (Duto)"
        ],
        "💧 Controle (Central de Água Gelada - CAG)": [
            "Válvula motorizada Bypass Proporcional (até 2.1/2\")", "Válvula motorizada Bypass Proporcional (3\" ou 4\")",
            "Válvula motorizada Bypass Proporcional (5\")", "Válvula motorizada Bypass Proporcional (6\")",
            "Válvula motorizada Bypass Proporcional (8\")", "Transmissor de pressão para água",
            "Tranmissor de vazão para água", "Válvula bloqueio motorizada", "Chave de fluxo", "Bombas (I/O para controlador)", "Tanques (I/O para controlador)"
        ],
        "🔸 Monitoramento (Filtros e Status)": [
            "Pressostato - Filtro G4", "Pressostato - Filtro F9", "Pressostato - Filtro H13/H14",
            "Status funcionamento ventilador ou exaustor (partida direta)", "Transmissor de pressão diferencial - Filtro G4",
            "Transmissor de pressão diferencial - Filtro F9", "Tranmissor de pressão diferencial - Filtro H13"
        ],
        "🟢 Monitoramento de Ambiente (Salas)": [
            "Transmissor de pressão diferencial entre salas", "Transmissor de pressão diferencial entre salas com display",
            "Transmissor de temperatura Ambiente", "Transmissor de temperatura ambiente com display",
            "Transmissor de temperatura e umidade ambiente", "Transmissor de temperatura e umidade ambiente com display"
        ]
    }
    
    KITS_PADRAO = {
        "❄️ UTA Padrão - Água Gelada": {
            "Transmissor de pressão Dif. Para ar (Vazão de ar)": 1,
            "Transmissor de temperatura e umidade para duto": 1, "Válvula de controle de água gelada proporcional": 1,
            "Pressostato - Filtro G4": 1, "Pressostato - Filtro F9": 1, "Pressostato - Filtro H13/H14": 1
        },
        "🌬️ UTA Padrão - Expansão Direta": {
            "Transmissor de pressão Dif. Para ar (Vazão de ar)": 1,
            "Transmissor de temperatura e umidade para duto": 1, "Relé de Corrente - Status Compressor": 2,
            "Pressostato - Filtro G4": 1, "Pressostato - Filtro F9": 1, "Pressostato - Filtro H13/H14": 1
        },
        "🔥 UTA Padrão - Água Gelada + Resistência": {
            "Transmissor de pressão Dif. Para ar (Vazão de ar)": 1,
            "Transmissor de temperatura e umidade para duto": 1, "Válvula de controle de água gelada proporcional": 1,
            "Pressostato - Filtro G4": 1, "Pressostato - Filtro F9": 1, "Pressostato - Filtro H13/H14": 1,
            "Termostato de segurança (Proteção da resistência)": 1, "Pressostato diferencial para ar (Proteção da resistência)": 1
        },
        "♨️ UTA Padrão - Expansão Direta + Resistência": {
            "Transmissor de pressão Dif. Para ar (Vazão de ar)": 1,
            "Transmissor de temperatura e umidade para duto": 1, "Relé de Corrente - Status Compressor": 2,
            "Pressostato - Filtro G4": 1, "Pressostato - Filtro F9": 1, "Pressostato - Filtro H13/H14": 1,
            "Termostato de segurança (Proteção da resistência)": 1, "Pressostato diferencial para ar (Proteção da resistência)": 1
        },
        "💨 Adicional: Ventilador/Exaustor (Inversor)": { "Transmissor de pressão Dif. Para ar (Vazão de ar)": 1 },
        "⚙️ Adicional: Ventilador/Exaustor (Partida Direta)": { "Status funcionamento ventilador ou exaustor (partida direta)": 1 }
    }

    PRECOS_IHM = {"Sem Interface (Cego)": 0.0, "IHM Básica 4.3\"": 1700.00, "IHM Padrão 7\"": 3400.00, "IHM Premium 10\"": 8500.00}

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

    # ==========================================
    # INTERFACE DE ABAS
    # ==========================================
    aba_auto, aba_planilhas, aba_infra, aba_precos, aba_resumo = st.tabs([
        "🚀 Dimensionamento de Automação", "🛠️ Planilhas Antigas", "🔌 Infraestrutura", "💲 Base de Preços", "📊 Orçamento Final"
    ])

    with aba_auto:
        
        # ASSISTENTE EM ETAPAS
        if not st.session_state.wizard_ativo:
            if st.button("➕ Criar Novo Quadro de Automação", type="primary"):
                st.session_state.wizard_ativo = True
                st.rerun()

        if st.session_state.wizard_ativo:
            with st.container(border=True):
                st.markdown("### 🧙‍♂️ Assistente de Configuração de Quadro")
                tipo_q = st.radio("1. Selecione o Tipo do Quadro:", ["Controle (HVAC/Máquinas)", "CAG (Central de Água Gelada)"], horizontal=True)
                sup_opt = st.radio("2. Este quadro fará parte de um Sistema de Supervisório?", ["Não", "Sim"], horizontal=True)
                soft_sel = "Sem Supervisório"
                if sup_opt == "Sim":
                    soft_sel = st.selectbox("Selecione o Software de Supervisão:", [
                        "Sistema supervisório SEM certificação CFR-21",
                        "Sistema supervisório COM certificação CFR-21",
                        "Sistema de monitoramento Schneider EBO"
                    ])
                
                tag_q = st.text_input("3. Insira a TAG / Identificação do Quadro (Ex: QTA-01, QD-CAG):")
                config_opt = st.radio("4. Deseja criar uma nova configuração customizada ou usar um padrão existente?", 
                                      ["Usar Padrão Existente (Kits)", "Criar Nova Configuração Customizada (Em Branco)"], horizontal=True)
                
                kit_final_selecionado = None
                if config_opt == "Usar Padrão Existente (Kits)":
                    opcoes_kits_filtrados = list(KITS_PADRAO.keys())
                    if "CAG" in tipo_q:
                        opcoes_kits_filtrados = [k for k in KITS_PADRAO.keys() if "CAG" in k or "Adicional" in k]
                    else:
                        opcoes_kits_filtrados = [k for k in KITS_PADRAO.keys() if "CAG" not in k]
                    
                    kit_final_selecionado = st.selectbox("Selecione o Modelo Padrão SIARCON:", ["Selecione..."] + opcoes_kits_filtrados)
                
                c_conf, c_canc = st.columns(2)
                if c_conf.button("🚀 Confirmar e Montar Quadro", use_container_width=True):
                    if not tag_q:
                        st.warning("⚠️ Insira uma TAG válida para identificar o quadro.")
                    elif config_opt == "Usar Padrão Existente (Kits)" and kit_final_selecionado == "Selecione...":
                        st.warning("⚠️ Selecione um kit padrão ou mude para configuração customizada.")
                    else:
                        novos_instrumentos = {k: 0 for k in REGRA_IO.keys()}
                        grupos_equip = []
                        
                        if config_opt == "Usar Padrão Existente (Kits)":
                            for item_nome, qtd_padrao in KITS_PADRAO[kit_final_selecionado].items():
                                if item_nome in novos_instrumentos: novos_instrumentos[item_nome] = qtd_padrao
                            nome_limpo = kit_final_selecionado.split(" ", 1)[1] if " " in kit_final_selecionado else kit_final_selecionado
                            grupos_equip.append({
                                "nome_grupo": f"{nome_limpo}", "multiplicador": 1, "instrumentos": novos_instrumentos, "tags_lista": [""]
                            })
                        else:
                            grupos_equip.append({
                                "nome_grupo": "Equipamento Customizado", "multiplicador": 1, "instrumentos": novos_instrumentos, "tags_lista": [""]
                            })
                            
                        st.session_state.paineis_auto.append({
                            "id": len(st.session_state.paineis_auto),
                            "nome": tag_q, "tipo": tipo_q, "supervisorio": soft_sel,
                            "modo_config": config_opt, "ihm": "IHM Padrão 7\"", "grupos_equipamentos": grupos_equip
                        })
                        st.session_state.wizard_ativo = False
                        st.rerun()
                        
                if c_canc.button("❌ Cancelar", use_container_width=True):
                    st.session_state.wizard_ativo = False
                    st.rerun()

        st.write("")

        # EXIBIÇÃO DOS QUADROS CRIADOS (DASHBOARD)
        for p_idx, p_data in enumerate(st.session_state.paineis_auto):
            with st.container(border=True):
                c_icone, c_nome_painel, c_ihm_painel = st.columns([0.5, 4, 2])
                c_icone.markdown("## 🎛️")
                p_data['nome'] = c_nome_painel.text_input("Identificação do Quadro", value=p_data['nome'], key=f"n_p_{p_idx}", label_visibility="collapsed")
                
                opcoes_ihm = list(PRECOS_IHM.keys())
                ihm_salva = p_data.get('ihm', opcoes_ihm[0])
                idx_ihm = opcoes_ihm.index(ihm_salva) if ihm_salva in opcoes_ihm else 2 
                p_data['ihm'] = c_ihm_painel.selectbox("Interface", opcoes_ihm, index=idx_ihm, key=f"i_p_{p_idx}", label_visibility="collapsed")
                
                st.caption(f"**Tipo:** {p_data.get('tipo', 'Controle')} | **Supervisão:** {p_data.get('supervisorio', 'Sem Supervisório')}")
                
                with st.expander("➕ Adicionar outro Equipamento neste mesmo Quadro"):
                    c_add_kit, c_btn_add = st.columns([3, 1])
                    sub_kit = c_add_kit.selectbox("Escolha o Equipamento:", ["Selecione..."] + list(KITS_PADRAO.keys()), key=f"sub_kit_{p_idx}")
                    if c_btn_add.button("Adicionar", key=f"btn_sub_add_{p_idx}", use_container_width=True):
                        if sub_kit != "Selecione...":
                            novos_inst = {k: 0 for k in REGRA_IO.keys()}
                            for item_nome, qtd_padrao in KITS_PADRAO[sub_kit].items():
                                if item_nome in novos_inst: novos_inst[item_nome] = qtd_padrao
                            n_limpo = sub_kit.split(" ", 1)[1] if " " in sub_kit else sub_kit
                            p_data['grupos_equipamentos'].append({
                                "nome_grupo": f"{n_limpo}", "multiplicador": 1, "instrumentos": novos_inst, "tags_lista": [""]
                            })
                            st.rerun()

                total_ai_painel = total_ao_painel = total_di_painel = total_do_painel = 0

                for g_idx, g_data in enumerate(p_data['grupos_equipamentos']):
                    with st.expander(f"📦 {g_data['nome_grupo']}"):
                        
                        qtd_key = f"m_g_{p_idx}_{g_idx}"
                        qtd_atual = st.session_state.get(qtd_key, g_data.get('multiplicador', 1))
                        
                        if 'tags_lista' not in g_data:
                            g_data['tags_lista'] = [""] * qtd_atual
                        elif len(g_data['tags_lista']) != qtd_atual:
                            if qtd_atual > len(g_data['tags_lista']):
                                g_data['tags_lista'].extend([""] * (qtd_atual - len(g_data['tags_lista'])))
                            else:
                                g_data['tags_lista'] = g_data['tags_lista'][:qtd_atual]

                        render_qtd = min(qtd_atual, 5) 
                        col_ratios = [3] + [1.5] * render_qtd + [1]
                        cols = st.columns(col_ratios)
                        
                        g_data['nome_grupo'] = cols[0].text_input("Equipamento", value=g_data['nome_grupo'], key=f"n_g_{p_idx}_{g_idx}")
                        
                        for i in range(render_qtd):
                            g_data['tags_lista'][i] = cols[i+1].text_input(f"TAG {i+1}", value=g_data['tags_lista'][i], key=f"t_g_{p_idx}_{g_idx}_{i}")
                            
                        g_data['multiplicador'] = cols[-1].number_input("Qtd", min_value=1, value=qtd_atual, key=qtd_key)
                        
                        if qtd_atual > 5:
                            st.caption("⚠️ Para mais de 5 equipamentos, as TAGs extras podem ser inseridas como anotações no final do projeto.")

                        with st.expander("⚙️ Ajuste Fino de Instrumentos (Engenharia)"):
                            for grupo_nome, lista_itens in GRUPOS_INSTRUMENTOS.items():
                                open_p = True if "Controle" in grupo_nome else False
                                with st.expander(grupo_nome, expanded=open_p):
                                    if "CAG" in grupo_nome: st.caption("💡 *Dica de Engenharia: Dividir a vazão em válvulas menores reduz custos de Bypass.*")
                                    cols_inst = st.columns(2)
                                    for i, inst in enumerate(lista_itens):
                                        if inst not in g_data['instrumentos']: g_data['instrumentos'][inst] = 0
                                        with cols_inst[i % 2]:
                                            g_data['instrumentos'][inst] = st.number_input(inst, min_value=0, step=1, value=g_data['instrumentos'][inst], key=f"inst_{p_idx}_{g_idx}_{inst}")
                            
                            if st.button("🗑️ Remover Máquina", key=f"del_{p_idx}_{g_idx}"):
                                p_data['grupos_equipamentos'].pop(g_idx)
                                st.rerun()

                    total_ai_g = total_ao_g = total_di_g = total_do_g = 0
                    for inst, q in g_data['instrumentos'].items():
                        io_vals = REGRA_IO.get(inst, {"AI": 0, "AO": 0, "DI": 0, "DO": 0})
                        total_ai_g += q * io_vals["AI"]
                        total_ao_g += q * io_vals["AO"]
                        total_di_g += q * io_vals["DI"]
                        total_do_g += q * io_vals["DO"]
                    
                    m_mult = g_data['multiplicador']
                    total_ai_painel += total_ai_g * m_mult
                    total_ao_painel += total_ao_g * m_mult
                    total_di_painel += total_di_g * m_mult
                    total_do_painel += total_do_g * m_mult

                total_io_pontos = total_ai_painel + total_ao_painel + total_di_painel + total_do_painel
                c36, c24, c18, c15 = dimensionar_controladores(total_io_pontos)
                nome_caixa, preco_caixa = calcular_painel_fisico(c36 + c24 + c18 + c15)
                
                st.markdown("##### 🧠 Estrutura de I/O do Quadro")
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("Total I/O", str(total_io_pontos)) 
                m2.metric("AI", total_ai_painel)
                m3.metric("AO", total_ao_painel)
                m4.metric("DI", total_di_painel)
                m5.metric("DO", total_do_painel)
                
                if st.button("🗑️ Deletar Todo este Quadro", key=f"del_quadro_{p_idx}"):
                    st.session_state.paineis_auto.pop(p_idx)
                    st.rerun()
        
        # ---------------------------------------------------------
        # BOTÃO: SALVAR RASCUNHO E SAIR DA TELA
        # ---------------------------------------------------------
        if st.session_state.paineis_auto:
            st.markdown("---")
            if st.button("💾 Salvar Rascunho e Sair (Retomar depois)", type="secondary", use_container_width=True):
                if not st.session_state.nome_projeto_orcamento:
                    st.warning("⚠️ Atenção: Preencha o 'Nome do Orçamento / Projeto' no topo da página antes de salvar o rascunho.")
                else:
                    try:
                        sh = conectar_google_sheets()
                        try: ws_hist_orc = sh.worksheet("Historico_Orcamentos")
                        except:
                            ws_hist_orc = sh.add_worksheet(title="Historico_Orcamentos", rows="1000", cols="8")
                            ws_hist_orc.append_row(["Data/Hora", "Nome do Projeto", "Revisão", "Subtotal Hardware", "Serviços de Lógica", "Custo Total Estimado", "Configuracao_JSON", "Usuário"])
                        
                        todas_linhas_existentes = ws_hist_orc.get_all_values()
                        contagem_revisoes = 0
                        if len(todas_linhas_existentes) > 1:
                            for r_row in todas_linhas_existentes[1:]:
                                if r_row[1].strip().upper() == st.session_state.nome_projeto_orcamento.strip().upper():
                                    contagem_revisoes += 1
                        
                        revisao_atual = f"R-{contagem_revisoes:02d}"
                        agora = datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M:%S")
                        json_config = json.dumps(st.session_state.paineis_auto)
                        
                        nova_linha_banco = [
                            agora, st.session_state.nome_projeto_orcamento, revisao_atual,
                            "R$ 0,00", "R$ 0,00", "R$ 0,00 (Rascunho)", json_config, st.session_state.usuario_logado
                        ]
                        ws_hist_orc.append_row(nova_linha_banco)
                        
                        st.session_state.paineis_auto = []
                        st.session_state.nome_projeto_orcamento = ""
                        st.toast("📝 Rascunho salvo na nuvem com sucesso! Tela limpa.", icon="💾")
                        st.rerun()
                    except Exception as e: st.error(f"Erro ao salvar rascunho: {e}")

        # --- FILTRAGEM INTELIGENTE NO HISTÓRICO ---
        st.markdown("---")
        with st.expander(f"📂 Abrir Orçamento Existente (Histórico de {st.session_state.nome_exibicao})"):
            try:
                sh = conectar_google_sheets()
                todas_linhas = sh.worksheet("Historico_Orcamentos").get_all_values()
                if len(todas_linhas) > 1:
                    dados_historico = todas_linhas[1:]
                    
                    for idx_rev, linha in enumerate(dados_historico[::-1]):
                        idx_real = len(dados_historico) - 1 - idx_rev
                        
                        usuario_registro = linha[7] if len(linha) > 7 else "rodrigo.ribeiro"
                        if usuario_registro.strip().lower() != st.session_state.usuario_logado.strip().lower():
                            continue
                            
                        with st.container():
                            c1, c2, c3, c4 = st.columns([1.5, 3, 1.5, 1])
                            d_h = linha[0]
                            n_p = linha[1]
                            rev = linha[2] if len(linha) >= 7 else "R-00"
                            tot_val = linha[5] if len(linha) >= 7 else linha[4]
                            j_salvo = linha[6] if len(linha) >= 7 else linha[5]
                            
                            c1.write(f"📅 {d_h}")
                            c2.write(f"**{n_p}** `({rev})`")
                            c3.write(tot_val)
                            if c4.button("📂 Carregar", key=f"btn_abrir_{idx_real}"):
                                st.session_state.projeto_para_abrir = idx_real
                                st.session_state.dados_projeto_abrir = {'nome': n_p, 'json': j_salvo}
                            st.markdown("---")
                            
                    if st.session_state.get('projeto_para_abrir') is not None:
                        d_a = st.session_state.get('dados_projeto_abrir', {})
                        st.warning(f"⚠️ Atenção: Carregar os dados de '{d_a['nome']}' irá substituir as configurações atuais da sua tela.")
                        c_sim, c_nao = st.columns(2)
                        if c_sim.button("✔️ Sim, substituir tela", use_container_width=True):
                            st.session_state.paineis_auto = json.loads(d_a['json'])
                            st.session_state.nome_projeto_orcamento = d_a['nome']
                            st.session_state.projeto_para_abrir = None
                            st.rerun()
                        if c_nao.button("❌ Não, cancelar", use_container_width=True):
                            st.session_state.projeto_para_abrir = None
                            st.rerun()
                else: st.write("Nenhum levantamento salvo neste perfil ainda.")
            except Exception as e: st.write("A aba 'Historico_Orcamentos' está vazia ou aguardando dados.")

    with aba_precos:
        st.header("Gestão da Base de Preços")
        df_precos = pd.DataFrame(list(st.session_state.precos_banco.items()), columns=["Item / Equipamento", "Valor Atual (R$)"])
        edited_df = st.data_editor(df_precos, use_container_width=True, hide_index=True)
        
        if st.button("💾 Salvar Novos Preços no Banco de Dados", type="primary"):
            alterou_algo = False
            novos_historicos = []
            for idx, row in edited_df.iterrows():
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
                    
                    linhas_h = [[h["Data/Hora"], h["Item Alterado"], h["Valor Antigo"], h["Novo Valor"]] for h in novos_historicos]
                    if linhas_h: ws_hist.append_rows(linhas_h)
                    st.success("✅ Preços atualizados na nuvem!")
                except Exception as e: st.error(f"Erro ao salvar: {e}")

    with aba_planilhas:
        st.header("Leitura das Planilhas Antigas")
        def converter_valor_plan(val):
            try:
                v = str(val).upper().replace('R$', '').strip()
                if v in ['NAN', 'NONE', '', '-']: return 0.0
                if ',' in v: v = v.replace('.', '').replace(',', '.')
                return float(v)
            except: return 0.0

        @st.cache_data
        def carregar_dados_planilhas():
            diretorio_atual = os.getcwd()
            pasta_dados = os.path.join(diretorio_atual, "dados")
            arquivos_na_pasta = os.listdir(pasta_dados) if os.path.exists(pasta_dados) else []
            nome_cag = next((f for f in arquivos_na_pasta if "CAG" in f.upper()), "CAG.csv")
            nome_ahu = next((f for f in arquivos_na_pasta if "AHU" in f.upper()), "AHU01.csv")
            nome_infra = next((f for f in arquivos_na_pasta if "INFRA" in f.upper()), "Infra.csv")
            cag_path = os.path.join(pasta_dados, nome_cag)
            ahu_path = os.path.join(pasta_dados, nome_ahu)
            infra_path = os.path.join(pasta_dados, nome_infra)
            def ler_csv_blindado(caminho, palavra_chave):
                try:
                    df = pd.read_csv(caminho, sep=',', header=None, dtype=str)
                    if len(df.columns) <= 2: df = pd.read_csv(caminho, sep=';', header=None, dtype=str)
                except: df = pd.read_csv(caminho, sep=';', header=None, dtype=str)
                header_idx = 0
                for i, row in df.iterrows():
                    if palavra_chave in " ".join([str(x).upper() for x in row.values]):
                        header_idx = i
                        break
                raw_cols = df.iloc[header_idx].astype(str).tolist()
                clean_cols = []
                for j, col in enumerate(raw_cols):
                    c = col.strip().upper()
                    if c in ['NAN', '', 'NONE', 'UNNAMED']: clean_cols.append(f"VAZIA_{j}")
                    elif c in clean_cols: clean_cols.append(f"{c}_{j}")
                    else: clean_cols.append(c)
                df.columns = clean_cols
                df = df.iloc[header_idx+1:].reset_index(drop=True)
                return df
            try:
                return ler_csv_blindado(cag_path, "ITEM"), ler_csv_blindado(ahu_path, "ITEM"), ler_csv_blindado(infra_path, "INSTRUMENTAÇÃO")
            except: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        cag_df, ahu_df, infra_df = carregar_dados_planilhas()
        if not cag_df.empty and not ahu_df.empty:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("CAG (Planilha)")
                for index, row in cag_df.iterrows():
                    c1, c2 = st.columns([3, 1])
                    item_nome = str(row.get('ITEM', 'Item Desconhecido'))
                    valor_unit = converter_valor_plan(row.get('VALOR UNITÁRIO', 0))
                    c1.write(f"{item_nome} (R$ {valor_unit:.2f})")
                    qtd = c2.number_input(f"Qtd", min_value=0, value=0, key=f"cag_plan_{index}")
                    if qtd > 0: st.session_state.orcamento.append({"Categoria": "Manual - Planilha", "Item": item_nome, "Quantidade": qtd, "Custo_Total": qtd * valor_unit})
            with col2:
                st.subheader("AHU01 (Planilha)")
                for index, row in ahu_df.iterrows():
                    c1, c2 = st.columns([3, 1])
                    item_nome = str(row.get('ITEM', 'Item Desconhecido'))
                    valor_unit = converter_valor_plan(row.get('VALOR UNITÁRIO', 0))
                    c1.write(f"{item_nome} (R$ {valor_unit:.2f})")
                    qtd = c2.number_input(f"Qtd", min_value=0, value=0, key=f"ahu_plan_{index}")
                    if qtd > 0: st.session_state.orcamento.append({"Categoria": "Manual - Planilha", "Item": item_nome, "Quantidade": qtd, "Custo_Total": qtd * valor_unit})

    with aba_infra:
        st.header("Cálculo de Infraestrutura")
        if not infra_df.empty and 'INSTRUMENTAÇÃO' in infra_df.columns:
            for index, row in infra_df.iterrows():
                tipo_inst = str(row['INSTRUMENTAÇÃO']).strip()
                custo_cabo = converter_valor_plan(row.iloc[6]) if len(row) > 6 else 0.0
                custo_infra = converter_valor_plan(row.get('INFRA', 0))
                if tipo_inst in ["", "NAN", "NONE", "-", "EQUIPAMENTOS", "TOTAL EQUIPAMENTOS"]: continue
                st.write(f"**{tipo_inst}** (Cabo: R${custo_cabo:.2f}/m | Infra: R${custo_infra:.2f}/m)")
                c1, c2 = st.columns(2)
                qtd_inst = c1.number_input("Qtd. de Instrumentos", min_value=0, value=0, key=f"infra_qtd_{index}")
                dist_media = c2.number_input("Distância Média (m)", min_value=0.0, value=0.0, step=1.0, key=f"infra_dist_{index}")
                if qtd_inst > 0 and dist_media > 0:
                    metragem = qtd_inst * dist_media
                    st.session_state.orcamento.append({"Categoria": "Infraestrutura", "Item": f"Cabo ({tipo_inst})", "Quantidade": metragem, "Custo_Total": metragem * custo_cabo})
                    st.session_state.orcamento.append({"Categoria": "Infraestrutura", "Item": f"Infra ({tipo_inst})", "Quantidade": metragem, "Custo_Total": metragem * custo_infra})

    with aba_resumo:
        st.header("Consolidação Financeira do Orçamento")
        linhas_resumo = []
        linhas_pontos = []
        softwares_incluidos = {}

        for p in st.session_state.paineis_auto:
            total_ai_painel = total_ao_painel = total_di_painel = total_do_painel = 0
            for g in p['grupos_equipamentos']:
                mult = g.get('multiplicador', 1)
                
                lista_tags = [t for t in g.get('tags_lista', []) if t.strip() != ""]
                str_tags = f" (TAGs: {', '.join(lista_tags)})" if len(lista_tags) > 0 else ""
                nome_equip = f"{g['nome_grupo']}{str_tags}"
                
                for inst, qtd in g['instrumentos'].items():
                    if qtd > 0:
                        qtd_final = qtd * mult
                        preco_item = st.session_state.precos_banco.get(inst, 0.0)
                        
                        io_vals = REGRA_IO.get(inst, {"AI": 0, "AO": 0, "DI": 0, "DO": 0})
                        
                        total_ai_painel += qtd_final * io_vals["AI"]
                        total_ao_painel += qtd_final * io_vals["AO"]
                        total_di_painel += qtd_final * io_vals["DI"]
                        total_do_painel += qtd_final * io_vals["DO"]
                        linhas_resumo.append({"Categoria": f"{p['nome']} - Campo", "Item": f"{inst} ({nome_equip})", "Preço Unit.": preco_item, "Qtd": qtd_final, "Custo Total": qtd_final * preco_item})
                        linhas_pontos.append({"Painel": p['nome'], "Grupo/Equipamento": nome_equip, "Instrumento": inst, "Quantidade Total": qtd_final, "Entrada Digital (DI)": qtd_final * io_vals["DI"], "Saída Digital (DO)": qtd_final * io_vals["DO"], "Entrada Analógica (AI)": qtd_final * io_vals["AI"], "Saída Analógica (AO)": qtd_final * io_vals["AO"]})

            tot_io_painel = total_ai_painel + total_ao_painel + total_di_painel + total_do_painel
            if tot_io_painel > 0:
                custo_ana = (total_ai_painel + total_ao_painel) * st.session_state.precos_banco.get("Custo AI/AO", 565.0)
                custo_dig = (total_di_painel + total_do_painel) * st.session_state.precos_banco.get("Custo DI/DO", 120.0)
                linhas_resumo.append({"Categoria": f"{p['nome']} - I/Os", "Item": "Pontos Analógicos (AI/AO)", "Preço Unit.": st.session_state.precos_banco.get("Custo AI/AO", 565.0), "Qtd": (total_ai_painel + total_ao_painel), "Custo Total": custo_ana})
                linhas_resumo.append({"Categoria": f"{p['nome']} - I/Os", "Item": "Pontos Digitais (DI/DO)", "Preço Unit.": st.session_state.precos_banco.get("Custo DI/DO", 120.0), "Qtd": (total_di_painel + total_do_painel), "Custo Total": custo_dig})
                
                c36, c24, c18, c15 = dimensionar_controladores(tot_io_painel)
                if c36 > 0: linhas_resumo.append({"Categoria": f"{p['nome']} - MPC", "Item": "Controlador MP-C-36A", "Preço Unit.": st.session_state.precos_banco.get("MP-C-36A", 9459.0), "Qtd": c36, "Custo Total": c36 * st.session_state.precos_banco.get("MP-C-36A", 9459.0)})
                if c24 > 0: linhas_resumo.append({"Categoria": f"{p['nome']} - MPC", "Item": "Controlador MP-C-24A", "Preço Unit.": st.session_state.precos_banco.get("MP-C-24A", 7290.0), "Qtd": c24, "Custo Total": c24 * st.session_state.precos_banco.get("MP-C-24A", 7290.0)})
                if c18 > 0: linhas_resumo.append({"Categoria": f"{p['nome']} - MPC", "Item": "Controlador MP-C-18A", "Preço Unit.": st.session_state.precos_banco.get("MP-C-18A", 5185.0), "Qtd": c18, "Custo Total": c18 * st.session_state.precos_banco.get("MP-C-18A", 5185.0)})
                if c15 > 0: linhas_resumo.append({"Categoria": f"{p['nome']} - MPC", "Item": "Controlador MP-C-15A", "Preço Unit.": st.session_state.precos_banco.get("MP-C-15A", 4649.0), "Qtd": c15, "Custo Total": c15 * st.session_state.precos_banco.get("MP-C-15A", 4649.0)})
                
                nome_caixa, preco_caixa = calcular_painel_fisico(c36 + c24 + c18 + c15)
                linhas_resumo.append({"Categoria": f"{p['nome']} - Estrutura Fís.", "Item": nome_caixa, "Preço Unit.": preco_caixa, "Qtd": 1, "Custo Total": preco_caixa})
                if PRECOS_IHM[p['ihm']] > 0: linhas_resumo.append({"Categoria": f"{p['nome']} - Estrutura Fís.", "Item": p['ihm'], "Preço Unit.": PRECOS_IHM[p['ihm']], "Qtd": 1, "Custo Total": PRECOS_IHM[p['ihm']]})

                s_type = p.get('supervisorio', "Sem Supervisório")
                if s_type != "Sem Supervisório":
                    if s_type not in softwares_incluidos: softwares_incluidos[s_type] = 0
                    softwares_incluidos[s_type] += tot_io_painel

        for item in st.session_state.orcamento:
            linhas_resumo.append({"Categoria": item['Categoria'], "Item": item['Item'], "Preço Unit.": item['Custo_Total']/item['Quantidade'] if item['Quantidade'] > 0 else 0, "Qtd": item['Quantidade'], "Custo Total": item['Custo_Total']})

        for s_name, pts_total in softwares_incluidos.items():
            b_k, p_k = "", ""
            if "SEM certificação" in s_name:
                b_k, p_k = "Licença Supervisório - SEM CFR-21 (Base)", "Licença Supervisório - SEM CFR-21 (Por Ponto I/O)"
            elif "COM certificação" in s_name:
                b_k, p_k = "Licença Supervisório - COM CFR-21 (Base)", "Licença Supervisório - COM CFR-21 (Por Ponto I/O)"
            else:
                b_k, p_k = "Licença Supervisório - Schneider EBO (Base)", "Licença Supervisório - Schneider EBO (Por Ponto I/O)"
            
            p_base = st.session_state.precos_banco.get(b_k, 23000.0)
            p_pto = st.session_state.precos_banco.get(p_k, 100.0)
            
            linhas_resumo.append({"Categoria": "🖥️ Software de Supervisão", "Item": f"Licença Base: {s_name}", "Preço Unit.": p_base, "Qtd": 1, "Custo Total": p_base})
            if p_pto > 0 and pts_total > 0:
                linhas_resumo.append({"Categoria": "🖥️ Software de Supervisão", "Item": f"Pontos Licenciados no Software ({pts_total} canais)", "Preço Unit.": p_pto, "Qtd": pts_total, "Custo Total": pts_total * p_pto})

        if len(linhas_resumo) > 0:
            df_final = pd.DataFrame(linhas_resumo)
            df_agrupado = df_final.groupby(['Categoria', 'Item', 'Preço Unit.'], as_index=False).agg({'Qtd': 'sum', 'Custo Total': 'sum'})
            subtotal_materiais = df_agrupado['Custo Total'].sum()
            custo_servicos_logica = subtotal_materiais * 0.25  
            total_projeto = subtotal_materiais + custo_servicos_logica
            
            st.dataframe(df_agrupado.style.format({'Preço Unit.': 'R$ {:.2f}', 'Custo Total': 'R$ {:.2f}'}), use_container_width=True)
            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            c1.info(f"**Subtotal Materiais/Hardware:**\nR$ {subtotal_materiais:,.2f}")
            c2.warning(f"**Serviços de Lógica (25%):**\nR$ {custo_servicos_logica:,.2f}")
            c3.success(f"**CUSTO TOTAL ESTIMADO:**\nR$ {total_projeto:,.2f}")
            
            df_pontos = pd.DataFrame(linhas_pontos)
            if not df_pontos.empty:
                total_qtd = df_pontos['Quantidade Total'].sum()
                linha_total = pd.DataFrame([{"Painel": "TOTAL GERAL", "Grupo/Equipamento": "-", "Instrumento": "-", "Quantidade Total": total_qtd, "Entrada Digital (DI)": df_pontos['Entrada Digital (DI)'].sum(), "Saída Digital (DO)": df_pontos['Saída Digital (DO)'].sum(), "Entrada Analógica (AI)": df_pontos['Entrada Analógica (AI)'].sum(), "Saída Analógica (AO)": df_pontos['Saída Analógica (AO)'].sum()}])
                df_pontos = pd.concat([df_pontos, linha_total], ignore_index=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_exportacao = pd.concat([
                    df_agrupado, 
                    pd.DataFrame([{'Categoria': 'Serviços', 'Item': 'Mão de Obra Lógica', 'Preço Unit.': custo_servicos_logica, 'Qtd': 1, 'Custo Total': custo_servicos_logica}]), 
                    pd.DataFrame([{'Categoria': 'TOTAL', 'Item': 'Geral', 'Preço Unit.': '-', 'Qtd': '-', 'Custo Total': total_projeto}])
                ], ignore_index=True)
                df_exportacao.to_excel(writer, index=False, sheet_name='Detalhamento Financeiro')
                if not df_pontos.empty: df_pontos.to_excel(writer, index=False, sheet_name='Matriz de Pontos (IO)')
            
            st.download_button(label="📥 Exportar Orçamento Final para Excel", data=buffer.getvalue(), file_name="orcamento_dimensionado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.markdown("---")
            
            if st.button("☁️ Salvar Orçamento Final e Gerar Revisão", type="primary", use_container_width=True):
                if not st.session_state.nome_projeto_orcamento: 
                    st.warning("⚠️ Atenção: Preencha o 'Nome do Orçamento / Projeto' antes de salvar.")
                else:
                    try:
                        sh = conectar_google_sheets()
                        try: ws_hist_orc = sh.worksheet("Historico_Orcamentos")
                        except:
                            ws_hist_orc = sh.add_worksheet(title="Historico_Orcamentos", rows="1000", cols="8")
                            ws_hist_orc.append_row(["Data/Hora", "Nome do Projeto", "Revisão", "Subtotal Hardware", "Serviços de Lógica", "Custo Total Estimado", "Configuracao_JSON", "Usuário"])
                        
                        todas_linhas_existentes = ws_hist_orc.get_all_values()
                        contagem_revisoes = 0
                        if len(todas_linhas_existentes) > 1:
                            for r_row in todas_linhas_existentes[1:]:
                                if r_row[1].strip().upper() == st.session_state.nome_projeto_orcamento.strip().upper():
                                    contagem_revisoes += 1
                        
                        revisao_atual = f"R-{contagem_revisoes:02d}"
                        agora = datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M:%S")
                        json_config = json.dumps(st.session_state.paineis_auto)
                        
                        nova_linha_banco = [
                            agora, st.session_state.nome_projeto_orcamento, revisao_atual,
                            f"R$ {subtotal_materiais:.2f}".replace('.', ','), f"R$ {custo_servicos_logica:.2f}".replace('.', ','), f"R$ {total_projeto:.2f}".replace('.', ','), json_config, st.session_state.usuario_logado
                        ]
                        ws_hist_orc.append_row(nova_linha_banco)
                        st.success(f"✅ Sucesso! Orçamento para '{st.session_state.nome_projeto_orcamento}' salvo com a revisão {revisao_atual}!")
                    except Exception as e: st.error(f"Erro ao salvar: {e}")
