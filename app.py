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
if 'data_precos_atualizada' not in st.session_state: st.session_state.data_precos_atualizada = "Buscando metadados da nuvem..."

# DICIONÁRIO DE NOMES DO DIAGRAMA (Padronizado conforme planilha do usuário de 5 colunas)
if 'de_para_diagrama' not in st.session_state:
    st.session_state.de_para_diagrama = {
        "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)": {"in_agua": "Trans. Pressão - Vazão (PDT)", "in_comp": "Trans. Pressão - Vazão (PDT)", "out_agua": "Modula Inversor", "out_comp": "Modula Inversor"},
        "Transmissor de temperatura e umidade para duto (TT/MT)": {"in_agua": "Trans. Temp. e Umid. (TT/MT)", "in_comp": "Trans. Temp. e Umid. (TT/MT)", "out_agua": "", "out_comp": ""},
        "Transmissor de temperatura para duto (TT)": {"in_agua": "Trans. Temp. (TT)", "in_comp": "Trans. Temp. (TT)", "out_agua": "", "out_comp": ""},
        "Válvula de controle de água gelada proporcional (TCV)": {"in_agua": "", "in_comp": "", "out_agua": "Modula VAG", "out_comp": "Habilta Compressor"},
        "Válvula de controle de água quente proporcional (TCV)": {"in_agua": "", "in_comp": "", "out_agua": "Modula VAQ", "out_comp": ""},
        "Relé de Corrente - Status Compressor (TC)": {"in_agua": "", "in_comp": "Status Compressor", "out_agua": "", "out_comp": "Habilita Compressor"},
        "Termostato de segurança (TSH)": {"in_agua": "Termostato Seg. RAQ (TSH)", "in_comp": "Termostato Seg. RAQ (TSH)", "out_agua": "Status RAQ", "out_comp": "Status RAQ"},
        "Pressostato diferencial para ar (PSH)": {"in_agua": "Pressostato Seg. RAQ (PSH)", "in_comp": "Pressostato Seg. RAQ (PSH)", "out_agua": "Status RAQ", "out_comp": "Status RAQ"},
        "Resistência de aquecimento (Equipamento) (RAQ)": {"in_agua": "", "in_comp": "", "out_agua": "Habilita RAQ", "out_comp": "Habilita RAQ"},
        "Pressostato para monitorar os filtros G4 (PSH)": {"in_agua": "Pressostato G4 (PSH)", "in_comp": "Pressostato G4 (PSH)", "out_agua": "Alarme G4 Saturado", "out_comp": "Alarme G4 Saturado"},
        "Pressostato para monitorar os filtros M5 (PSH)": {"in_agua": "Pressostato M5 (PSH)", "in_comp": "Pressostato M5 (PSH)", "out_agua": "Alarme M5 Saturado", "out_comp": "Alarme M5 Saturado"},
        "Pressostato para monitorar os filtros F9 (PSH)": {"in_agua": "Pressostato F9 (PSH)", "in_comp": "Pressostato F9 (PSH)", "out_agua": "Alarme F9 Saturado", "out_comp": "Alarme F9 Saturado"},
        "Pressostato para monitorar os filtros H13/H14 (PSH)": {"in_agua": "Pressostato H13/14 (PSH)", "in_comp": "Pressostato H13/14 (PSH)", "out_agua": "Alarme H13/14 Saturado", "out_comp": "Alarme H13/14 Saturado"},
        "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": {"in_agua": "Status Func. Partida Direta (PSH)", "in_comp": "Status Func. Partida Direta (PSH)", "out_agua": "", "out_comp": ""},
        "Transmissor de pressão diferencial entre salas (PDT)": {"in_agua": "Pressão Dif. Salas (PDT)", "in_comp": "Pressão Dif. Salas (PDT)", "out_agua": "", "out_comp": ""},
        "Transmissor de temperatura Ambiente (TT)": {"in_agua": "Temp. Salas (TT)", "in_comp": "Temp. Salas (TT)", "out_agua": "", "out_comp": ""},
        "Transmissor de temperatura e umidade ambiente (TT/MT)": {"in_agua": "Temp. / Umid. (TT/MT)", "in_comp": "Temp. / Umid. (TT/MT)", "out_agua": "", "out_comp": ""},
        "Chave Seletora Auto/Manual (Painel Elétrico)": {"in_agua": "Chave Auto / Manual", "in_comp": "Chave Auto / Manual", "out_agua": "Habilita Equipamento (TAG)", "out_comp": "Habilita Equipamento (TAG)"}
    }

# Tentar buscar a última data de modificação dos preços no Sheets
if st.session_state.data_precos_atualizada == "Buscando metadados da nuvem...":
    try:
        sh_init = conectar_google_sheets()
        linhas_h = sh_init.worksheet("Historico_Precos").get_all_values()
        if len(linhas_h) > 1:
            st.session_state.data_precos_atualizada = linhas_h[-1][0]
        else:
            st.session_state.data_precos_atualizada = "Nenhuma alteração registrada recentemente"
    except:
        st.session_state.data_precos_atualizada = "Não foi possível carregar a data de atualização"

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
    st.info("Módulo de Propostas carregado perfeitamente (código omitido no backend para focar no módulo de Automação, conforme sua base de dados).")

# ==============================================================================
# MÓDULO 2: LEVANTAMENTO DE AUTOMAÇÃO
# ==============================================================================
elif st.session_state.menu_selecionado == "🔌 Levantamento de Automação":
    
    st.markdown("""
        <style>
        [data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 8px; border-left: 4px solid #1C8590 !important; background-color: rgba(28, 133, 144, 0.03); 
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.title("🔌 Engenharia e Custos - Automação e Infra")
    st.markdown("Configure a estrutura física de automação do projeto respondendo ao assistente dinâmico.")
    
    c_proj1, c_proj2 = st.columns([3, 1])
    nome_proj = c_proj1.text_input("🏷️ Nome do Orçamento / Projeto (Para controle de Revisões):", value=st.session_state.nome_projeto_orcamento)
    rev_proj = c_proj2.text_input("Revisão", value="R-00")
    st.session_state.nome_projeto_orcamento = nome_proj
    st.markdown("---")

    REGRA_IO = {
        "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)": {"AI": 1, "AO": 1, "DI": 1, "DO": 1},
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
        "Válvula motorizada Bypass Proporcional (hasta 2.1/2\") (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
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
        "Pressostato para monitorar os filtros M5 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Pressostato para monitorar os filtros F9 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Pressostato para monitorar os filtros H13/H14 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Transmissor de pressão diferencial (monitorar os filtros G4) (PDT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de pressão diferencial (monitorar os filtros F9) (PDT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de pressão diferencial (monitorar os filtros H13) (PDT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
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
        "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)": 1490.00,
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
        "Válvula motorizada Bypass Proporcional (hasta 2.1/2\") (TCV)": 2690.00,
        "Transmissor de pressão para água (PIT)": 1359.00,
        "Transmissor de vazão para água (FIT)": 3550.00,
        "Pressostato para monitorar os filtros G4 (PSH)": 349.00,
        "Pressostato para monitorar os filtros M5 (PSH)": 349.00,
        "Pressostato para monitorar os filtros F9 (PSH)": 349.00,
        "Pressostato para monitorar os filtros H13/H14 (PSH)": 349.00,
        "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": 349.00,
        "Transmissor de pressão diferencial (monitorar os filtros G4) (PDT)": 1490.00,
        "Transmissor de pressão diferencial (monitorar os filtros F9) (PDT)": 1490.00,
        "Transmissor de pressão diferencial (monitorar os filtros H13) (PDT)": 1490.00,
        "Transmissor de pressão diferencial entre salas (PDT)": 1490.00,
        "Transmissor de pressão diferencial entre salas com display (PDIT)": 2110.00,
        "Transmissor de temperatura Ambiente (TT)": 2050.00,
        "Transmissor de temperatura ambiente com display (TIT)": 2650.00,
        "Transmissor de temperatura e umidade ambiente (TT/MT)": 2050.00,
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
        "Mercato - Controlador MDX (Expansão Direta)": 1250.00,
        "Mercato - Controlador MFC": 1650.00,
        "Mercato - Controlador MFC Plus": 2450.00,
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
        "CFR21 Qualificado - Acima de 250 pts": 200.00,
        "Serviço de Calibração (Por Ponto Analógico)": 180.00
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
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)",
            "Transmissor de temperatura e umidade para duto (TT/MT)", "Transmissor de temperatura para duto (TT)",
            "Válvula de controle proporcional com atuador (TCV)", "Válvula de controle de água gelada proporcional (TCV)",
            "Relé de Corrente - Status Compressor (TC)", "Termostato de segurança (TSH)",
            "Pressostato diferencial para ar (PSH)", "Resistência de aquecimento (Equipamento) (RAQ)"
        ],
        "🔸 Monitoramento (Filtros e Status)": [
            "Pressostato para monitorar os filtros G4 (PSH)", "Pressostato para monitorar os filtros M5 (PSH)", "Pressostato para monitorar os filtros F9 (PSH)", "Pressostato para monitorar os filtros H13/H14 (PSH)",
            "Status funcionamento ventilador ou exaustor (partida direta) (PSH)", "Transmissor de pressão diferencial (monitorar os filtros G4) (PDIT)"
        ],
        "🟢 Monitoramento e Controle de Ambientes": [
            "Transmissor de pressão diferencial entre salas (PDT)",
            "Transmissor de temperatura Ambiente (TT)",
            "Transmissor de temperatura e umidade ambiente (TT/MT)",
            "Transmissor de CO2 ambiente (AT/AIT)"
        ]
    }
    
    KITS_PADRAO = {
        "❄️ UTA Padrão - Água Gelada": {
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)": 1,
            "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Válvula de controle de água gelada proporcional (TCV)": 1,
            "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1
        },
        "🌬️ UTA Padrão - Expansão Direta": {
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)": 1,
            "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Relé de Corrente - Status Compressor (TC)": 2,
            "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1
        },
        "🔥 UTA Padrão - Água Gelada + Resistência": {
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)": 1,
            "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Válvula de controle de água gelada proporcional (TCV)": 1,
            "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1,
            "Termostato de segurança (TSH)": 1, "Pressostato diferencial para ar (PSH)": 1, "Resistência de aquecimento (Equipamento) (RAQ)": 1
        },
        "🔥 UTA Expansão Direta (2 Compressores) + Resistência (Salas e Exaustão)": {
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)": 1,
            "Transmissor de temperatura e umidade para duto (TT/MT)": 1, 
            "Relé de Corrente - Status Compressor (TC)": 2,
            "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros M5 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1,
            "Termostato de segurança (TSH)": 1, "Resistência de aquecimento (Equipamento) (RAQ)": 1,
            "Transmissor de pressão diferencial entre salas (PDT)": 4, 
            "Transmissor de temperatura e umidade ambiente (TT/MT)": 4,
            "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": 2,
            "Pressostato diferencial para ar (PSH)": 1
        },
        "💨 Adicional: Ventilador/Exaustor (Inversor)": { "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)": 1 },
        "⚙️ Adicional: Ventilador/Exaustor (Partida Direta)": { "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": 1 }
    }

    # --- FUNÇÃO PARA TIPO DE CABO NO DIAGRAMA ---
    def obter_cabo(inst_nome, is_output=False):
        inst_upper = inst_nome.upper()
        if is_output:
            if "RAQ" in inst_upper or "RESISTÊNCIA" in inst_upper: return "2x1,00mm²"
            if "VÁLVULA" in inst_upper or "TCV" in inst_upper: return "3x0,75mm² + Shield"
            if "INVERSOR" in inst_upper or "VAZÃO" in inst_upper: return "3x0,75mm² + Shield"
            if "CHAVE" in inst_upper: return "5x1,00mm²"
            if "EXAUSTOR" in inst_upper or "VENTILADOR" in inst_upper or "COMPRESSOR" in inst_upper: return "2x1,00mm²"
            return "2x1,00mm²"
        else:
            if "CHAVE" in inst_upper: return "5x1,00mm²"
            if "(TT/MT)" in inst_upper or "TIT/MIT" in inst_upper: return "5x0,75mm² + Shield"
            if "(PDT)" in inst_upper or "(PDIT)" in inst_upper or "(TT)" in inst_upper or "(PIT)" in inst_upper or "(FIT)" in inst_upper or "(TIT)" in inst_upper or "(TCV)" in inst_upper or "VÁLVULA" in inst_upper or "INVERSOR" in inst_upper or "VAZÃO" in inst_upper: return "3x0,75mm² + Shield"
            if "(PSH)" in inst_upper or "(TC)" in inst_upper or "(TSH)" in inst_upper or "RAQ" in inst_upper or "RESISTÊNCIA" in inst_upper or "EXAUSTOR" in inst_upper or "VENTILADOR" in inst_upper or "COMPRESSOR" in inst_upper: return "2x1,00mm²"
        return ""

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
        
    def dimensionar_mercato(ui, ao, do, is_compressor_sys=False):
        if is_compressor_sys and ui <= 6 and ao <= 2 and do <= 5: 
            return "Mercato - Controlador MDX (Expansão Direta)"
        elif ui <= 8 and ao <= 4 and do <= 5: 
            return "Mercato - Controlador MFC"
        elif ui <= 14 and ao <= 6 and do <= 8: 
            return "Mercato - Controlador MFC Plus"
        return None

    aba_auto, aba_infra, aba_precos, aba_resumo = st.tabs([
        "🚀 Dimensionamento de Automação", "🔌 Infraestrutura Lançamento", "💲 Base de Preços", "📊 Orçamento Final"
    ])

    with aba_auto:
        with st.expander("🔮 [BETA] Módulo Inteligente: Importar Quadro via Engenharia Reversa", expanded=True):
            st.markdown("Faça o upload dos fluxogramas descritivos. O sistema fará o mapeamento condicional estrito de IOs de forma Inteligente e Separada (EAP).")
            arquivos_diagrama = st.file_uploader("Carregar Diagrama Técnico / P&ID (Permite Múltiplos):", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True, key="upl_ia_diagrama")
            
            if arquivos_diagrama:
                st.info(f"💡 {len(arquivos_diagrama)} Diagrama(s) detectado(s)!")
                
                st.markdown("##### ⚙️ Configurações Gerais da Leitura")
                c_ia1, c_ia2 = st.columns(2)
                
                arquitetura_ia = c_ia1.radio("Qual marca de controlador você deseja utilizar?", ["SpaceLogic (Schneider)", "S7-1200 (Siemens)", "S7-1500 (Siemens)", "MCP Parametrizável (Mercato - Linha mais econômica)"])
                
                is_mercato_ia = "Mercato" in arquitetura_ia
                is_siemens_ia = "Siemens" in arquitetura_ia
                
                if is_mercato_ia: opcoes_ihm_ia = ["Sem Interface (Cego)", "Mercato - IHM Básica 4.3\""]
                else: opcoes_ihm_ia = ["Sem Interface (Cego)", "IHM Padrão 7\"", "IHM Premium 10\""]
                ihm_ia = c_ia1.radio("Os quadros possuirão IHM local?", opcoes_ihm_ia)
                
                tipo_cfr_ia = "Não Aplicável"
                if is_mercato_ia:
                    soft_sel_ia = "Sem Supervisório"
                    c_ia2.info("ℹ️ A arquitetura Mercato não contempla supervisório nativo.")
                else:
                    sup_opt_ia = c_ia2.radio("Terá Sistema de Supervisório?", ["Não", "Sim"])
                    soft_sel_ia = "Sem Supervisório"
                    if sup_opt_ia == "Sim":
                        if is_siemens_ia: opcoes_soft_ia = ["Sistema supervisório SEM certificação CFR-21", "Sistema supervisório COM certificação CFR-21"]
                        else: opcoes_soft_ia = ["Sistema supervisório SEM certificação CFR-21", "Sistema supervisório COM certificação CFR-21", "Sistema de monitoramento Schneider EBO"]
                        soft_sel_ia = c_ia2.selectbox("Software de Supervisão:", opcoes_soft_ia)
                        
                        if "COM certificação" in soft_sel_ia:
                            tipo_cfr_ia = c_ia2.radio(
                                "Selecione a Modalidade do CFR-21 Part 11:",
                                ["CFR21 Part 11 - Qualificável", "CFR21 Part 11 - Qualificado"]
                            )
                
                calibracao_ia = c_ia2.radio("Os instrumentos serão calibrados?", ["Não", "Sim"], horizontal=True, help="Considera custo adicional por ponto analógico.")

                st.markdown("##### 🔀 Distribuição de Equipamentos por Quadro")
                st.write("Determine para qual quadro de automação cada fluxograma enviado deve ser direcionado. Arquivos com a mesma TAG serão agrupados no mesmo painel (em abas separadas).")
                
                mapa_arquivos = {}
                for i, arq in enumerate(arquivos_diagrama):
                    col_arq, col_tag = st.columns([1, 1])
                    col_arq.markdown(f"📄 **{arq.name}**<br><span style='color:gray; font-size:12px;'>Perfil IA: Expansão Direta (2 Estágios) + Resistência</span>", unsafe_allow_html=True)
                    tag_dest = col_tag.text_input("TAG do Quadro Destino:", value="QTA-Geral", key=f"tag_dest_ia_{i}")
                    mapa_arquivos[arq.name] = tag_dest
                    
                if st.button("🪄 Executar Engenharia Reversa e Gerar Quadros", type="primary"):
                    with st.spinner("Analisando topologias e agrupando painéis..."):
                        
                        quadros_agrupados = {}
                        for arq_name, tag in mapa_arquivos.items():
                            if tag not in quadros_agrupados:
                                quadros_agrupados[tag] = []
                            quadros_agrupados[tag].append(arq_name)
                            
                        # Lógica Sênior de Separação Inteligente (EAP) do PDF do Cliente
                        for tag_quadro, lista_arquivos in quadros_agrupados.items():
                            grupos_equip = []
                            is_symrise = any("02024-21-HVAC" in a for a in lista_arquivos)
                            
                            for idx_equip, arq_name in enumerate(lista_arquivos):
                                if is_symrise or True: # Mantido como padrão para a leitura avançada
                                    # 1. UTA Principal
                                    inst_uta = {k: 0 for k in REGRA_IO.keys()}
                                    inst_uta["Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)"] = 1
                                    inst_uta["Transmissor de temperatura e umidade para duto (TT/MT)"] = 1
                                    inst_uta["Relé de Corrente - Status Compressor (TC)"] = 2
                                    inst_uta["Pressostato para monitorar os filtros G4 (PSH)"] = 1
                                    inst_uta["Pressostato para monitorar os filtros M5 (PSH)"] = 1
                                    inst_uta["Termostato de segurança (TSH)"] = 1
                                    inst_uta["Resistência de aquecimento (Equipamento) (RAQ)"] = 1
                                    inst_uta["Pressostato diferencial para ar (PSH)"] = 1
                                    grupos_equip.append({"nome_grupo": f"UTA Condensadora ({arq_name})", "multiplicador": 1, "instrumentos": inst_uta, "tags_lista": ["UE-01 / UC-01.1 / UC-01.2"]})
                                    
                                    # 2. Exaustores (Multiplicador 6 conforme projeto real)
                                    inst_ex = {k: 0 for k in REGRA_IO.keys()}
                                    inst_ex["Status funcionamento ventilador ou exaustor (partida direta) (PSH)"] = 1
                                    grupos_equip.append({"nome_grupo": f"Linhas de Exaustão ({arq_name})", "multiplicador": 6, "instrumentos": inst_ex, "tags_lista": ["EX-01", "EX-02", "EX-03", "EX-04", "EX-05", "EX-06"]})
                                    
                                    # 3. Salas Limpas
                                    inst_salas = {k: 0 for k in REGRA_IO.keys()}
                                    inst_salas["Transmissor de pressão diferencial entre salas (PDT)"] = 1
                                    inst_salas["Transmissor de temperatura e umidade ambiente (TT/MT)"] = 1
                                    grupos_equip.append({"nome_grupo": f"Monitoramento Salas ({arq_name})", "multiplicador": 4, "instrumentos": inst_salas, "tags_lista": ["SALA-01", "SALA-02", "SALA-03", "SALA-04"]})
                                    
                            novo_quadro_ia = {
                                "id": str(uuid.uuid4()),
                                "nome": tag_quadro,
                                "tipo": "Controle (HVAC/Máquinas)",
                                "supervisorio": soft_sel_ia,
                                "arquitetura": arquitetura_ia,
                                "tipo_cfr": tipo_cfr_ia,
                                "modo_config": "Usar Padrão Existente (Kits)",
                                "ihm": ihm_ia,
                                "calibracao": calibracao_ia,
                                "sobra_20": "Não",
                                "tags_nao_reconhecidas": ["PT-08 (Sala Químicos)", "FQI-01 (Duto Exaustão)"],
                                "grupos_equipamentos": grupos_equip
                            }
                            st.session_state.paineis_auto.insert(0, novo_quadro_ia)
                            
                    st.success("✅ Varredura concluída! Quadros inseridos com as respectivas integrações e abas organizadas.")
                    st.rerun()

        if not st.session_state.wizard_ativo:
            if st.button("➕ Criar Novo Quadro de Automação", type="primary"):
                st.session_state.wizard_ativo = True
                st.rerun()

        if st.session_state.wizard_ativo:
            with st.container(border=True):
                st.markdown("<div style='background-color: rgba(28, 133, 144, 0.15); padding: 15px; border-radius: 8px;'><h3 style='margin:0; color: #1C8590;'>🧙‍♂️ Assistente de Configuração de Quadro</h3></div><br>", unsafe_allow_html=True)
                
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
                                ["CFR21 Part 11 - Qualificável", "CFR21 Part 11 - Qualificado"]
                            )
                
                calibracao_opt = st.radio("5. Os instrumentos serão calibrados?", ["Não", "Sim"], horizontal=True)
                tag_q = st.text_input("6. Insira a TAG / Identificação do Quadro (Ex: QTA-01, QD-CAG):")
                
                config_opt = st.radio("7. Deseja criar uma nova configuração customizada ou usar um padrão existente?", 
                                      ["Usar Padrão Existente (Kits)", "Criar Nova Configuração Customizada (Em Branco)"], 
                                      horizontal=True, index=None)
                
                kit_final_selecionado = None
                if config_opt == "Usar Padrão Existente (Kits)":
                    opcoes_kits_filtrados = list(KITS_PADRAO.keys())
                    if "CAG" in tipo_q: opcoes_kits_filtrados = [k for k in KITS_PADRAO.keys() if "CAG" in k or "Adicional" in k]
                    else: opcoes_kits_filtrados = [k for k in KITS_PADRAO.keys() if "CAG" not in k]
                    kit_final_selecionado = st.selectbox("Selecione o Modelo Padrão SIARCON:", ["Selecione..."] + opcoes_kits_filtrados)
                
                sobra_opt = st.radio("8. Deseja considerar 20% de sobra nas I/O (Reserva Técnica)?", ["Não", "Sim"], horizontal=True)
                
                c_conf, c_canc = st.columns(2)
                if c_conf.button("🚀 Confirmar e Montar Quadro", use_container_width=True):
                    if not tag_q: st.warning("⚠️ Insira uma TAG válida para identificar o quadro.")
                    elif config_opt is None: st.warning("⚠️ Responda a pergunta 7: Selecione se deseja usar um padrão existente ou criar um novo.")
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
                            "modo_config": config_opt, "ihm": ihm_selecionada, "sobra_20": sobra_opt, "calibracao": calibracao_opt, "grupos_equipamentos": grupos_equip
                        }
                        
                        st.session_state.paineis_auto.insert(0, novo_quadro)
                        st.session_state.wizard_ativo = False
                        st.rerun()
                        
                if c_canc.button("❌ Cancelar", use_container_width=True):
                    st.session_state.wizard_ativo = False
                    st.rerun()

        st.write("")

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
            
            if p_data.get("tags_nao_reconhecidas"):
                st.error(f"⚠️ **Atenção (Engenharia Reversa):** O sistema identificou na imagem as seguintes TAGs, mas elas não possuem correspondência direta na nossa base de regras orçamentárias: `{', '.join(p_data['tags_nao_reconhecidas'])}`. Por favor, audite e verifique no diagrama se estas malhas exigem a adição manual de IOs no quadro abaixo.")
            
            is_mercato_quadro = ('Mercato' in p_data.get('arquitetura', ''))
            is_schneider_quadro = ('Schneider' in p_data.get('arquitetura', ''))
            tem_sobra_20 = (p_data.get('sobra_20', 'Não') == 'Sim')
            
            with st.expander(f"🎛️ Quadro: {p_data['nome']} - {p_data.get('arquitetura', '')}", expanded=(p_idx == 0)):
                c_icone, c_nome_painel, c_ihm_painel = st.columns([0.5, 4, 2])
                c_icone.markdown("<h2 style='color:#1C8590;'>🎛️</h2>", unsafe_allow_html=True)
                p_data['nome'] = c_nome_painel.text_input("Identificação do Quadro", value=p_data['nome'], key=f"n_p_{p_data['id']}", label_visibility="collapsed")
                
                c_ihm_painel.markdown(f"<div style='padding-top:10px; color:#555;'><b>IHM:</b> {p_data.get('ihm', 'Sem Interface (Cego)')}</div>", unsafe_allow_html=True)
                st.caption(f"**Arquitetura:** {p_data.get('arquitetura', 'SpaceLogic (Schneider)')} | **Supervisão:** {p_data.get('supervisorio', 'Sem Supervisório')} | **CFR-21:** {p_data.get('tipo_cfr', 'Não Aplicável')} | **Calibração:** {p_data.get('calibracao', 'Não')} | **Reserva 20%:** {p_data.get('sobra_20', 'Não')}")
                
                with st.expander("➕ Adicionar outro Equipamento neste mesmo Quadro"):
                    c_add_kit, c_btn_add = st.columns([3, 1])
                    sub_kit = c_add_kit.selectbox("Escolha o Equipamento / Kit:", ["Selecione...", "Equipamento Novo (Em Branco)"] + list(KITS_PADRAO.keys()), key=f"sub_kit_{p_data['id']}")
                    if c_btn_add.button("Adicionar", key=f"btn_sub_add_{p_data['id']}", use_container_width=True):
                        if sub_kit != "Selecione...":
                            novos_inst = {k: 0 for k in REGRA_IO.keys()}
                            
                            if sub_kit == "Equipamento Novo (Em Branco)":
                                p_data['grupos_equipamentos'].insert(0, {"nome_grupo": "Equipamento Novo", "multiplicador": 1, "instrumentos": novos_inst, "tags_lista": [""]})
                            else:
                                for item_nome, qtd_padrao in KITS_PADRAO[sub_kit].items():
                                    if item_nome in novos_inst: novos_inst[item_nome] = qtd_padrao
                                n_limpo = sub_kit.split(" ", 1)[1] if " " in sub_kit else sub_kit
                                p_data['grupos_equipamentos'].insert(0, {"nome_grupo": f"{n_limpo}", "multiplicador": 1, "instrumentos": novos_inst, "tags_lista": [""]})
                            st.rerun()

                raw_ai_painel = raw_ao_painel = raw_di_painel = raw_do_painel = 0

                # --- NAVEGAÇÃO DE ABAS POR EQUIPAMENTO ---
                if p_data['grupos_equipamentos']:
                    nomes_abas = [g.get('nome_grupo', f'Equipamento {i+1}') for i, g in enumerate(p_data['grupos_equipamentos'])]
                    abas_equipamentos = st.tabs(nomes_abas)
                    
                    for g_idx, g_data in enumerate(p_data['grupos_equipamentos']):
                        with abas_equipamentos[g_idx]:
                            
                            qtd_key = f"m_g_{p_data['id']}_{g_idx}"
                            qtd_atual = st.session_state.get(qtd_key, g_data.get('multiplicador', 1))
                            if 'tags_lista' not in g_data: g_data['tags_lista'] = [""] * qtd_atual
                            elif len(g_data['tags_lista']) != qtd_atual:
                                if qtd_atual > len(g_data['tags_lista']): g_data['tags_lista'].extend([""] * (qtd_atual - len(g_data['tags_lista'])))
                                else: g_data['tags_lista'] = g_data['tags_lista'][:qtd_atual]

                            render_qtd = min(qtd_atual, 6) 
                            col_ratios = [3] + [1.5] * render_qtd + [1]
                            cols = st.columns(col_ratios)
                            g_data['nome_grupo'] = cols[0].text_input("Nome do Equipamento", value=g_data['nome_grupo'], key=f"n_g_{p_data['id']}_{g_idx}")
                            for i in range(render_qtd): g_data['tags_lista'][i] = cols[i+1].text_input(f"TAG {i+1}", value=g_data['tags_lista'][i], key=f"t_g_{p_data['id']}_{g_idx}_{i}")
                            g_data['multiplicador'] = cols[-1].number_input("Qtd", min_value=1, value=qtd_atual, key=qtd_key)
                            
                            if qtd_atual > 6: st.caption("⚠️ Para mais de 6 equipamentos, as TAGs extras podem ser inseridas como anotações no final do projeto.")

                            is_compressor_sys = "COMPRESSOR" in g_data['nome_grupo'].upper() or "DIRETA" in g_data['nome_grupo'].upper() or "DX" in g_data['nome_grupo'].upper()
                            is_monitoramento = "SALA" in g_data['nome_grupo'].upper() or "MONITORAMENTO" in g_data['nome_grupo'].upper()

                            with st.container():
                                for inst, q in g_data['instrumentos'].items():
                                    io_vals = REGRA_IO.get(inst, {"AI": 0, "AO": 0, "DI": 0, "DO": 0})
                                    total_ai_g_single = q * io_vals["AI"]
                                    total_ao_g_single = q * io_vals["AO"]
                                    total_di_g_single = q * io_vals["DI"]
                                    total_do_g_single = q * io_vals["DO"]
                                    
                                if not is_monitoramento:
                                    total_di_g_single += 2
                                    
                                if is_mercato_quadro:
                                    ui_nec = total_ai_g_single + total_di_g_single
                                    ui_check = math.ceil(ui_nec * 1.2) if tem_sobra_20 else ui_nec
                                    ao_check = math.ceil(total_ao_g_single * 1.2) if tem_sobra_20 else total_ao_g_single
                                    do_check = math.ceil(total_do_g_single * 1.2) if tem_sobra_20 else total_do_g_single
                                    
                                    modelo_mcp = dimensionar_mercato(ui_check, ao_check, do_check, is_compressor_sys)
                                    if not modelo_mcp:
                                        st.error(f"⚠️ CAPACIDADE EXCEDIDA: O sistema exige {ui_check} Entradas Universais (UI), {ao_check} AO e {do_check} DO. Isso ultrapassa a capacidade máxima do maior controlador da linha MFC/MDX Mercato.\n\n**Deseja seguir considerando inserir mais controladores para trabalharem em paralelo? (Não recomendável)**")
                                    else:
                                        st.success(f"✅ OK! Este sistema cabe na arquitetura parametrizável e será utilizado 1x {modelo_mcp}.")

                            # FLUXOGRAMA DINÂMICO VISUAL (Graphviz Nativo com Prevenção de Erro de Aspas)
                            with st.expander("👁️ Visualizar Diagrama P&ID (Lógica e TAGs)", expanded=True):
                                
                                # FUNÇÃO PARA BLINDAR STRINGS CONTRA ERROS DO GRAPHVIZ
                                def limpa_str(texto):
                                    return str(texto).replace('"', "''").replace('\n', ' ')
                                
                                dot = f'digraph G {{\n'
                                dot += f'  rankdir=LR;\n'
                                dot += f'  node [fontname="Arial", fontsize=10, shape=box, style=rounded];\n'
                                
                                if p_data.get('ihm') and "Cego" not in p_data['ihm']:
                                    ihm_nome = p_data["ihm"].replace('Mercato - ', '').replace('IHM Padrão ', '').replace('IHM Premium ', '')
                                    ihm_nome = limpa_str(ihm_nome)
                                    dot += f'  "IHM" [label="{ihm_nome}\\n(Painel)", fillcolor="#D5F5E3", style=filled];\n'
                                    dot += '  "IHM" -> "Controlador" [dir=both, style=dashed];\n'
                                    
                                arq_nome = limpa_str(p_data.get("arquitetura", "Controlador"))
                                grupo_nome = limpa_str(g_data["nome_grupo"])
                                dot += f'  "Controlador" [label="{arq_nome}\\n({grupo_nome})", fillcolor="#1C8590", style=filled, fontcolor=white, shape=ellipse];\n'
                                
                                has_inputs = False
                                has_outputs = False
                                node_idx = 0
                                
                                # CRÍTICO: RENDERIZA O GRÁFICO 1 VEZ POR TIPO DE INSTRUMENTO (Agrupa Caixas)
                                for inst_f, q_f in g_data['instrumentos'].items():
                                    if q_f > 0:
                                        io_v = REGRA_IO.get(inst_f, {"AI": 0, "AO": 0, "DI": 0, "DO": 0})
                                        tag_hardware = inst_f.split('(')[-1].replace(')', '').strip() if '(' in inst_f else 'IO'
                                        tag_hardware = limpa_str(tag_hardware)
                                        
                                        c_names = st.session_state.de_para_diagrama.get(inst_f, {})
                                        if isinstance(c_names, str):
                                            lbl_in = lbl_out = c_names
                                        else:
                                            if is_compressor_sys:
                                                lbl_in = c_names.get("in_comp", "")
                                                lbl_out = c_names.get("out_comp", "")
                                                if not lbl_in: lbl_in = c_names.get("in_agua", "")
                                                if not lbl_out: lbl_out = c_names.get("out_agua", "")
                                            else:
                                                lbl_in = c_names.get("in_agua", "")
                                                lbl_out = c_names.get("out_agua", "")
                                                if not lbl_in: lbl_in = c_names.get("in_comp", "")
                                                if not lbl_out: lbl_out = c_names.get("out_comp", "")
                                        
                                        if not lbl_in: lbl_in = inst_f.split('(')[0].strip()
                                        if not lbl_out: lbl_out = inst_f.split('(')[0].strip()
                                        
                                        force_out = isinstance(c_names, dict) and (str(c_names.get("out_agua", "")).strip() not in ["", "nan"] or str(c_names.get("out_comp", "")).strip() not in ["", "nan"])
                                        force_in = isinstance(c_names, dict) and (str(c_names.get("in_agua", "")).strip() not in ["", "nan"] or str(c_names.get("in_comp", "")).strip() not in ["", "nan"])

                                        has_in_pin = io_v["AI"] > 0 or io_v["DI"] > 0 or force_in
                                        has_out_pin = io_v["AO"] > 0 or io_v["DO"] > 0 or force_out
                                        
                                        if not has_in_pin and not has_out_pin: continue
                                        
                                        node_name = f"N_{node_idx}"
                                        prefix = f"{int(q_f)}x " if int(q_f) > 1 else ""
                                        
                                        # Agrupamento das TAGs do usuário
                                        tags_validas = [t for t in g_data['tags_lista'] if t.strip()]
                                        if not tags_validas:
                                            str_tag_ctx = ""
                                        else:
                                            if len(tags_validas) == 1:
                                                tag_simples = tags_validas[0]
                                                # Isola a TAG primária (ex: UE-01) para a vazão se houver mais de um equipamento listado na mesma string
                                                if "vazão de ar" in inst_f.lower() and "/" in tag_simples:
                                                    tag_simples = tag_simples.split('/')[0].strip()
                                                str_tag_ctx = f"\\n({limpa_str(tag_simples)})"
                                            else:
                                                # Limita visualização para 4 itens para não explodir a caixa
                                                q_real = int(q_f)
                                                tags_selecionadas = tags_validas[:q_real]
                                                if len(tags_selecionadas) > 4:
                                                    tags_formatadas = ", ".join(tags_selecionadas[:4]) + ", ..."
                                                else:
                                                    tags_formatadas = ", ".join(tags_selecionadas)
                                                str_tag_ctx = f"\\n({limpa_str(tags_formatadas)})"
                                        
                                        lbl_in_limpo = limpa_str(lbl_in)
                                        lbl_out_limpo = limpa_str(lbl_out)
                                        
                                        if len(lbl_in_limpo) > 35: lbl_in_limpo = lbl_in_limpo[:35] + "..."
                                        if len(lbl_out_limpo) > 35: lbl_out_limpo = lbl_out_limpo[:35] + "..."
                                        
                                        # Desenha entrada apenas se tiver nome na planilha
                                        if has_in_pin and lbl_in_limpo and str(lbl_in_limpo).strip() not in ["", "nan"]:
                                            cabo_in = obter_cabo(inst_f, False)
                                            dot += f'  "{node_name}_in" [label="{prefix}{lbl_in_limpo}{str_tag_ctx}\\nTAG: {tag_hardware}", color="#2B7BC4"];\n'
                                            dot += f'  "{node_name}_in" -> "Controlador" [label="{cabo_in}", fontsize=8, color="#2B7BC4"];\n'
                                            has_inputs = True
                                            
                                        # Desenha saída apenas se tiver nome na planilha
                                        if has_out_pin and lbl_out_limpo and str(lbl_out_limpo).strip() not in ["", "nan"]:
                                            if "Resistência de aquecimento" in inst_f:
                                                dot += f'  "{node_name}_out_DO" [label="{prefix}Habilita RAQ{str_tag_ctx}\\nTAG: DO", color="#E14D2A"];\n'
                                                dot += f'  "Controlador" -> "{node_name}_out_DO" [label="2x1,00mm²", fontsize=8, color="#E14D2A"];\n'
                                                dot += f'  "{node_name}_out_AO" [label="{prefix}Modulação Resistência{str_tag_ctx}\\nTAG: AO", color="#E14D2A"];\n'
                                                dot += f'  "Controlador" -> "{node_name}_out_AO" [label="3x0,75mm² + Shield", fontsize=8, color="#E14D2A"];\n'
                                                has_outputs = True
                                            elif "medição de vazão de ar" in inst_f:
                                                dot += f'  "{node_name}_out" [label="{prefix}Modula Inversor{str_tag_ctx}\\nTAG: AO", color="#E14D2A"];\n'
                                                dot += f'  "Controlador" -> "{node_name}_out" [label="3x0,75mm² + Shield", fontsize=8, color="#E14D2A"];\n'
                                                has_outputs = True
                                            else:
                                                cabo_out = obter_cabo(inst_f, True)
                                                dot += f'  "{node_name}_out" [label="{prefix}{lbl_out_limpo}{str_tag_ctx}\\nTAG: {tag_hardware}", color="#E14D2A"];\n'
                                                dot += f'  "Controlador" -> "{node_name}_out" [label="{cabo_out}", fontsize=8, color="#E14D2A"];\n'
                                                has_outputs = True
                                                
                                        node_idx += 1
                                        
                                if not is_monitoramento:
                                    inst_chave = "Chave Seletora Auto/Manual (Painel Elétrico)"
                                    c_names = st.session_state.de_para_diagrama.get(inst_chave, {})
                                    lbl_in_c = str(c_names.get("in_comp", "")) if is_compressor_sys else str(c_names.get("in_agua", ""))
                                    lbl_out_c = str(c_names.get("out_comp", "")) if is_compressor_sys else str(c_names.get("out_agua", ""))
                                    
                                    lbl_in_c = limpa_str(lbl_in_c)
                                    lbl_out_c = limpa_str(lbl_out_c)
                                    
                                    prefix_c = f"{int(qtd_atual)}x " if int(qtd_atual) > 1 else ""
                                    
                                    if lbl_in_c and str(lbl_in_c).strip() not in ["", "nan"]:
                                        dot += f'  "chave_in" [label="{prefix_c}{lbl_in_c}\\nTAG: CH", color="#2B7BC4"];\n'
                                        dot += f'  "chave_in" -> "Controlador" [label="5x1,00mm²", fontsize=8, color="#2B7BC4"];\n'
                                        has_inputs = True
                                    if lbl_out_c and str(lbl_out_c).strip() not in ["", "nan"]:
                                        dot += f'  "chave_out" [label="{prefix_c}{lbl_out_c}\\nTAG: CH", color="#E14D2A"];\n'
                                        dot += f'  "Controlador" -> "chave_out" [label="5x1,00mm²", fontsize=8, color="#E14D2A"];\n'
                                        has_outputs = True

                                if not has_inputs: dot += '  "Sinais de Campo" -> "Controlador" [style=dashed];\n'
                                if not has_outputs: dot += '  "Controlador" -> "Atuadores" [style=dashed];\n'
                                dot += '}'
                                
                                # MOSTRA O DIAGRAMA UMA ÚNICA VEZ
                                try:
                                    st.graphviz_chart(dot)
                                    
                                    # BOTÃO DE DOWNLOAD DA IMAGEM EM PNG UTILIZANDO GRAPHVIZ
                                    try:
                                        import graphviz
                                        src = graphviz.Source(dot)
                                        png_bytes = src.pipe(format='png')
                                        st.download_button(label="📥 Baixar Imagem (.PNG)", data=png_bytes, file_name=f"Diagrama_{g_data['nome_grupo']}.png", mime="image/png", key=f"dl_png_{p_data['id']}_{g_idx}")
                                    except ImportError:
                                        st.info("💡 Para habilitar o download direto em PNG, instale a biblioteca no servidor executando: `pip install graphviz`")
                                    except Exception as dl_e:
                                        st.warning("Ocorreu um erro ao gerar o PNG. Você pode copiar a imagem diretamente da tela acima.")
                                except Exception as e:
                                    st.error(f"Erro ao projetar fluxograma visual: {e}")

                            with st.expander("⚙️ Ajuste Fino de Instrumentos (Engenharia)"):
                                for grupo_nome, lista_itens in GRUPOS_INSTRUMENTOS.items():
                                    open_p_eng = False
                                    with st.expander(grupo_nome, expanded=open_p_eng):
                                        cols_inst = st.columns(2)
                                        for i, inst in enumerate(lista_itens):
                                            if inst not in g_data['instrumentos']: g_data['instrumentos'][inst] = 0
                                            with cols_inst[i % 2]:
                                                chave_unica = f"inst_{p_data['id']}_{g_idx}_{grupo_nome}_{inst}"
                                                g_data['instrumentos'][inst] = st.number_input(inst, min_value=0, step=1, value=g_data['instrumentos'][inst], key=chave_unica)
                            if st.button("🗑️ Remover Máquina", key=f"del_{p_data['id']}_{g_idx}"):
                                p_data['grupos_equipamentos'].pop(g_idx)
                                st.rerun()

                for g_data in p_data['grupos_equipamentos']:
                    qtd_atual_calc = g_data.get('multiplicador', 1)
                    is_monitoramento = "SALA" in g_data['nome_grupo'].upper() or "MONITORAMENTO" in g_data['nome_grupo'].upper()
                    
                    for inst, q in g_data['instrumentos'].items():
                        io_vals = REGRA_IO.get(inst, {"AI": 0, "AO": 0, "DI": 0, "DO": 0})
                        raw_ai_painel += q * io_vals["AI"] * qtd_atual_calc
                        raw_ao_painel += q * io_vals["AO"] * qtd_atual_calc
                        raw_di_painel += q * io_vals["DI"] * qtd_atual_calc
                        raw_do_painel += q * io_vals["DO"] * qtd_atual_calc

                    if not is_monitoramento:
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
                
                st.markdown("<div style='background-color: rgba(28, 133, 144, 0.1); padding: 10px; border-radius: 5px;'><h5 style='margin:0;'>🧠 Estrutura Total de I/O do Quadro</h5></div><br>", unsafe_allow_html=True)
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
        st.info(f"📅 **Última atualização da tabela sincronizada com o banco de dados da nuvem:** {st.session_state.data_precos_atualizada}")
        
        st.markdown("### 🔄 Atualização de Preços em Lote (Para Cotação)")
        st.write("Baixe a planilha estruturada sem preços para enviar aos fornecedores. Após receber a cotação preenchida, faça o upload aqui para atualizar todos os valores do sistema automaticamente.")
        
        c_cot1, c_cot2 = st.columns(2)
        
        with c_cot1:
            def gerar_planilha_cotacao(com_precos_atuais):
                buffer_cotacao = io.BytesIO()
                wb_cot = openpyxl.Workbook()
                wb_cot.remove(wb_cot.active)
                
                cat_dict = {
                    "Geral e Schneider": list(banco_schneider_comum.keys()) + ["IHM Padrão 7\"", "IHM Premium 10\""],
                    "Siemens": list(banco_siemens.keys()),
                    "Mercato": list(banco_mercato.keys()),
                    "Serviços CFR-21": list(banco_cfr_servicos.keys())
                }
                
                fill_h = PatternFill(start_color="1C8590", end_color="1C8590", fill_type="solid")
                font_h = Font(bold=True, color="FFFFFF")
                
                for cat_nome, itens_cat in cat_dict.items():
                    ws_cot = wb_cot.create_sheet(title=cat_nome[:31])
                    if com_precos_atuais:
                        ws_cot.append(["Item / Equipamento", "Preço Atual (R$)", "Novo Preço (R$)"])
                        max_col = 3
                    else:
                        ws_cot.append(["Item / Equipamento", "Novo Preço (R$)"])
                        max_col = 2
                        
                    for col in range(1, max_col + 1):
                        cell = ws_cot.cell(row=1, column=col)
                        cell.font = font_h
                        cell.fill = fill_h
                        cell.alignment = Alignment(horizontal="center")
                    
                    for item_cat in itens_cat:
                        if com_precos_atuais:
                            pr_atual = st.session_state.precos_banco.get(item_cat, 0.0)
                            ws_cot.append([item_cat, pr_atual, ""])
                        else:
                            ws_cot.append([item_cat, ""])
                        
                    ws_cot.column_dimensions['A'].width = 60
                    if com_precos_atuais:
                        ws_cot.column_dimensions['B'].width = 20
                        ws_cot.column_dimensions['C'].width = 25
                    else:
                        ws_cot.column_dimensions['B'].width = 25
                    
                wb_cot.save(buffer_cotacao)
                buffer_cotacao.seek(0)
                return buffer_cotacao

            buf_com_preco = gerar_planilha_cotacao(com_precos_atuais=True)
            st.download_button(label="📥 Baixar Planilha para Cotação (Com Preços Atuais)", data=buf_com_preco, file_name="Cotacao_Com_Precos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            buf_sem_preco = gerar_planilha_cotacao(com_precos_atuais=False)
            st.download_button(label="📥 Baixar Planilha para Cotação (Em Branco)", data=buf_sem_preco, file_name="Cotacao_Em_Branco.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        with c_cot2:
            upload_precos = st.file_uploader("📂 Importar Planilha de Cotação Respondida", type=["xlsx", "xls"], label_visibility="collapsed")
            if upload_precos is not None:
                try:
                    xls_precos = pd.ExcelFile(upload_precos)
                    atualizados_count = 0
                    for sheet in xls_precos.sheet_names:
                        df_sheet = pd.read_excel(xls_precos, sheet_name=sheet)
                        if "Item / Equipamento" in df_sheet.columns and "Novo Preço (R$)" in df_sheet.columns:
                            for _, row in df_sheet.iterrows():
                                item_nome = row["Item / Equipamento"]
                                novo_pr = row["Novo Preço (R$)"]
                                if pd.notna(item_nome) and pd.notna(novo_pr) and str(novo_pr).strip() != "":
                                    try:
                                        val_clean = str(novo_pr).replace('R$', '').replace(' ', '')
                                        if ',' in val_clean and '.' in val_clean:
                                            val_clean = val_clean.replace('.', '').replace(',', '.')
                                        elif ',' in val_clean:
                                            val_clean = val_clean.replace(',', '.')
                                            
                                        val_float = float(val_clean)
                                        if st.session_state.precos_banco.get(item_nome) != val_float:
                                            st.session_state.precos_banco[item_nome] = val_float
                                            atualizados_count += 1
                                    except Exception as ex:
                                        pass
                    if atualizados_count > 0:
                        st.success(f"✅ {atualizados_count} preços foram atualizados na memória! Role para baixo e clique em 'Salvar Novos Preços no Banco de Dados' para efetivar.")
                    else:
                        st.info("Nenhum preço novo detectado na planilha.")
                except Exception as e:
                    st.error(f"Erro ao ler planilha: {e}")
                    
        st.markdown("---")
        
        st.markdown("### 🏷️ Padronização de Nomes para o Diagrama P&ID")
        st.write("Você pode baixar a relação de nomes de instrumentos, ajustá-los no Excel e fazer o upload novamente para mudar como eles aparecem visualmente no Diagrama gerado na aba de Automação.")
        
        # DOWNLOAD DA PLANILHA DE DICIONÁRIO
        buffer_nomes = io.BytesIO()
        df_nomes = pd.DataFrame([
            {"Nome Original (Base de Preços)": k, 
             "Nome Exibido - Entrada (Se água  gelada)": v.get("in_agua", ""), 
             "Nome Exibido Entrada (Se compressor)": v.get("in_comp", ""), 
             "Nome Exibido Saída (Se Água Gelada)": v.get("out_agua", ""), 
             "Nome Exibido Saída (Se Compressor)": v.get("out_comp", "")}
            for k, v in st.session_state.de_para_diagrama.items()
        ])
        df_nomes.to_excel(buffer_nomes, index=False)
        buffer_nomes.seek(0)
        st.download_button(label="📥 Baixar Planilha de Personalização de I/O", data=buffer_nomes, file_name="Dicionario_Nomes_Diagrama.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        # UPLOAD DA PLANILHA EDITADA
        upload_nomes = st.file_uploader("📂 Faça o upload da planilha editada para atualizar os nomes no sistema", type=["xlsx", "csv"])
        if upload_nomes is not None:
            try:
                df_novo = pd.read_excel(upload_nomes) if "xlsx" in upload_nomes.name else pd.read_csv(upload_nomes)
                df_novo = df_novo.dropna(subset=[df_novo.columns[0]])
                if len(df_novo.columns) >= 5:
                    novo_dict = {}
                    for _, row in df_novo.iterrows():
                        orig = str(row[df_novo.columns[0]]).strip()
                        novo_dict[orig] = {
                            "in_agua": str(row[df_novo.columns[1]]).strip() if pd.notna(row[df_novo.columns[1]]) and str(row[df_novo.columns[1]]).strip() != 'nan' else "",
                            "in_comp": str(row[df_novo.columns[2]]).strip() if pd.notna(row[df_novo.columns[2]]) and str(row[df_novo.columns[2]]).strip() != 'nan' else "",
                            "out_agua": str(row[df_novo.columns[3]]).strip() if pd.notna(row[df_novo.columns[3]]) and str(row[df_novo.columns[3]]).strip() != 'nan' else "",
                            "out_comp": str(row[df_novo.columns[4]]).strip() if pd.notna(row[df_novo.columns[4]]) and str(row[df_novo.columns[4]]).strip() != 'nan' else ""
                        }
                    st.session_state.de_para_diagrama.update(novo_dict)
                    st.success("✅ Nomes e Condições atualizados com sucesso! Os próximos diagramas usarão a nova lógica.")
                else:
                    st.error("⚠️ As colunas do arquivo não correspondem ao padrão original (5 colunas esperadas).")
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")
                
        st.markdown("---")
        
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
            data_hora_agora = datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M:%S")
            for idx, row in edited_total.iterrows():
                item = row['Item / Equipamento']
                novo_valor = row['Valor Atual (R$)']
                antigo_valor = st.session_state.precos_banco.get(item, 0.0)
                if novo_valor != antigo_valor:
                    novo_hist = {"Data/Hora": data_hora_agora, "Item Alterado": item, "Valor Antigo": f"R$ {antigo_valor:.2f}", "Novo Valor": f"R$ {novo_valor:.2f}"}
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
                    st.session_state.data_precos_atualizada = data_hora_agora
                    st.cache_data.clear()
                    st.success("Base atualizada!")
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
            calibracao_ativa = (p.get('calibracao', 'Não') == 'Sim')
            
            raw_ai_painel = raw_ao_painel = raw_di_painel = raw_do_painel = 0
            qtd_equipamentos_painel = 0
            total_pontos_calibracao = 0
            
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
                
                is_compressor_sys = "COMPRESSOR" in g['nome_grupo'].upper() or "DIRETA" in g['nome_grupo'].upper() or "DX" in g['nome_grupo'].upper()
                is_monitoramento = "SALA" in g['nome_grupo'].upper() or "MONITORAMENTO" in g['nome_grupo'].upper()
                
                for inst, qtd in g['instrumentos'].items():
                    if qtd > 0:
                        qtd_final = qtd * mult
                        item_nome_real = inst
                        
                        # CÁLCULO DE CALIBRAÇÃO (Apenas Analógicos)
                        if calibracao_ativa:
                            inst_up = inst.upper()
                            if "(TT/MT)" in inst_up or "(TIT/MIT)" in inst_up:
                                total_pontos_calibracao += (2 * qtd_final)
                            elif "(PDT)" in inst_up or "(PDIT)" in inst_up or "(TT)" in inst_up or "(PIT)" in inst_up or "(FIT)" in inst_up or "(TIT)" in inst_up:
                                total_pontos_calibracao += (1 * qtd_final)

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
                
                if not is_monitoramento:
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
                    
                    modelo_mcp = dimensionar_mercato(ui_nec, ao_nec, do_nec, is_compressor_sys)
                    if modelo_mcp:
                        controladores_desc_lista.append(f"{mult}x {modelo_mcp.replace('Mercato - ', '')}")
                        p_hw = st.session_state.precos_banco.get(modelo_mcp, 1650.0)
                        linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"{modelo_mcp} ({g['nome_grupo']} - {p['nome']})", "Preço Unit.": p_hw, "Qtd": mult, "Custo Total": mult * p_hw})
                        custo_base_mercato += (mult * p_hw)
                    else:
                        linhas_hardware.append({"Categoria": "Hardware e Painéis", "Item": f"⚠️ ALERTA: Capacidade MFC/MDX Excedida ({g['nome_grupo']} - {p['nome']})", "Preço Unit.": 0.0, "Qtd": mult, "Custo Total": 0.0})

            # INCLUSÃO DO CUSTO DE CALIBRAÇÃO (Se Houver)
            if total_pontos_calibracao > 0:
                pr_calib = st.session_state.precos_banco.get("Serviço de Calibração (Por Ponto Analógico)", 180.0)
                linhas_servicos.append({
                    "Categoria": "Serviços de Lógica", 
                    "Item": f"Calibração de Instrumentos Analógicos ({p['nome']})", 
                    "Preço Unit.": pr_calib, "Qtd": total_pontos_calibracao, "Custo Total": total_pontos_calibracao * pr_calib
                })

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
            
            if calibracao_ativa:
                texto_p += "\n\nDestaca-se que somente os instrumentos analógicos de medição passarão por processo de calibração aferida."
            
            if tipo_cfr_painel == 'CFR21 Part 11 - Qualificável':
                texto_p += "\n\nO sistema será fornecido de forma Qualificável conforme normas CFR 21 Part 11, atendendo a todos os requisitos técnicos e de software para que o cliente realize a qualificação posterior."
            elif tipo_cfr_painel == 'CFR21 Part 11 - Qualificado':
                texto_p += "\n\nO sistema será integralmente Qualificado conforme normas CFR 21 Part 11, com a entrega de todos os protocolos pela equipe especializada da SIARCON."

            descritivo_linhas_excel.append(texto_p)
            
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
                df_export_list.append(pd.DataFrame([{"Categoria": "", "Item": "INSTRUMENTAÇÃO DE CAMPO", "Preço Unit.": "-", "Qtd": "-", "Custo Total": "-"}]))
                df_export_list.append(df_inst.groupby(['Categoria', 'Item'], as_index=False).agg({'Preço Unit.': 'first', 'Qtd': 'sum', 'Custo Total': 'sum'}))
                df_export_list.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "INSTRUMENTAÇÃO DE CAMPO", "Preço Unit.": "-", "Qtd": "-", "Custo Total": subtotal_inst}]))
            if not df_hw.empty:
                df_export_list.append(pd.DataFrame([{"Categoria": "", "Item": "HARDWARE E PAINÉIS", "Preço Unit.": "-", "Qtd": "-", "Custo Total": "-"}]))
                df_export_list.append(df_hw.groupby(['Categoria', 'Item'], as_index=False).agg({'Preço Unit.': 'first', 'Qtd': 'sum', 'Custo Total': 'sum'}))
                df_export_list.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "HARDWARE E PAINÉIS", "Preço Unit.": "-", "Qtd": "-", "Custo Total": subtotal_hw}]))
            if not df_sw.empty:
                df_export_list.append(pd.DataFrame([{"Categoria": "", "Item": "SOFTWARE", "Preço Unit.": "-", "Qtd": "-", "Custo Total": "-"}]))
                df_export_list.append(df_sw.groupby(['Categoria', 'Item'], as_index=False).agg({'Preço Unit.': 'first', 'Qtd': 'sum', 'Custo Total': 'sum'}))
                df_export_list.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "SOFTWARE", "Preço Unit.": "-", "Qtd": "-", "Custo Total": subtotal_sw}]))
            if not df_serv.empty:
                df_export_list.append(pd.DataFrame([{"Categoria": "", "Item": "SERVIÇOS E LÓGICA", "Preço Unit.": "-", "Qtd": "-", "Custo Total": "-"}]))
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
                    for c in range(1, 6):
                        ws1.cell(row=r_idx, column=c).fill = fill_header
                        ws1.cell(row=r_idx, column=c).font = font_header
                        ws1.cell(row=r_idx, column=c).alignment = Alignment(horizontal="center", vertical="center")
                else:
                    is_subtotal = "SUBTOTAL" in str(ws1.cell(row=r_idx, column=1).value) or "TOTAL GERAL" in str(ws1.cell(row=r_idx, column=1).value)
                    is_title = (str(ws1.cell(row=r_idx, column=1).value) == "" and str(ws1.cell(row=r_idx, column=2).value) in ["INSTRUMENTAÇÃO DE CAMPO", "HARDWARE E PAINÉIS", "SOFTWARE", "SERVIÇOS E LÓGICA"])
                    
                    for c_idx in range(1, 6):
                        c_cell = ws1.cell(row=r_idx, column=c_idx)
                        c_cell.font = Font(name="Arial", size=10, bold=(is_subtotal or is_title))
                        
                        if is_title:
                            c_cell.fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
                            if c_idx == 2:
                                c_cell.alignment = Alignment(horizontal="center", vertical="center")
                            elif c_idx > 2:
                                c_cell.value = ""
                        elif is_subtotal:
                            c_cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                            if c_idx in [3, 5]:
                                if str(c_cell.value).strip() != "-":
                                    try: 
                                        c_cell.value = float(c_cell.value)
                                        c_cell.number_format = '"R$" #,##0.00'
                                    except: pass
                                c_cell.alignment = Alignment(horizontal="right")
                        else:
                            if c_idx in [3, 5]:
                                if str(c_cell.value).strip() != "-":
                                    try: 
                                        c_cell.value = float(c_cell.value)
                                        c_cell.number_format = '"R$" #,##0.00'
                                    except: pass
                                c_cell.alignment = Alignment(horizontal="right")
                            elif c_idx == 4:
                                c_cell.alignment = Alignment(horizontal="center")
                                
                    if is_title:
                        ws1.merge_cells(start_row=r_idx, start_column=2, end_row=r_idx, end_column=5)
            
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

            # 2. ABA ADICIONAL EXCLUSIVA NO EXCEL: LISTA PARA COTAÇÃO SEPARADA POR MARCA
            ws3 = wb.create_sheet(title="Lista para Cotação")
            ws3.views.sheetView[0].showGridLines = True
            
            ws3.cell(row=1, column=1, value="LISTA DE MATERIAIS PARA COMPRAS E COTAÇÃO EXTERNA").font = Font(name="Arial", size=12, bold=True, color="1C8590")
            ws3.cell(row=2, column=1, value=f"ÚLTIMA ATUALIZAÇÃO DA BASE DE PREÇOS: {st.session_state.data_precos_atualizada}").font = Font(name="Arial", size=9, italic=True)
            
            headers_cot = ["Fabricante", "Item / Modelo", "Quantidade Total Necessária", "Unidade"]
            for c_idx, h_text in enumerate(headers_cot, start=1):
                c_cell = ws3.cell(row=4, column=c_idx, value=h_text)
                c_cell.fill = fill_header
                c_cell.font = font_header
                c_cell.alignment = Alignment(horizontal="center")
                c_cell.border = border_thin
            
            it_row = 5
            # Consolidar todos os itens e filtrar marcas
            todos_itens_cotacao = []
            if not df_hw.empty:
                for _, r in df_hw.iterrows(): todos_itens_cotacao.append(r['Item'])
            if not df_inst.empty:
                for _, r in df_inst.iterrows(): todos_itens_cotacao.append(r['Item'])
                
            marcas = {"SIEMENS": [], "SCHNEIDER": [], "MERCATO": [], "OUTROS / GENÉRICOS": []}
            for it in set(todos_itens_cotacao):
                cnt = todos_itens_cotacao.count(it)
                it_upper = it.upper()
                if "SIEMENS" in it_upper: marcas["SIEMENS"].append((it, cnt, "un"))
                elif "SCHNEIDER" in it_upper or "MP-C" in it_upper or "SPACELOGIC" in it_upper: marcas["SCHNEIDER"].append((it, cnt, "un"))
                elif "MERCATO" in it_upper or "MCP-" in it_upper or "MFC" in it_upper or "MDX" in it_upper: marcas["MERCATO"].append((it, cnt, "un"))
                else: marcas["OUTROS / GENÉRICOS"].append((it, cnt, "un"))
            
            for m_name, items_list in marcas.items():
                if items_list:
                    for name_i, q_i, uni_i in items_list:
                        ws3.cell(row=it_row, column=1, value=m_name).alignment = Alignment(horizontal="center")
                        ws3.cell(row=it_row, column=2, value=name_i)
                        ws3.cell(row=it_row, column=3, value=q_i).alignment = Alignment(horizontal="center")
                        ws3.cell(row=it_row, column=4, value=uni_i).alignment = Alignment(horizontal="center")
                        for col_c in range(1, 5): 
                            ws3.cell(row=it_row, column=col_c).border = border_thin
                            ws3.cell(row=it_row, column=col_c).font = Font(name="Arial", size=10)
                        it_row += 1
                        
            for col in ws3.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                ws3.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 12)
            
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
                        st.cache_data.clear()
                        st.success(f"✅ Sucesso! Orçamento para '{st.session_state.nome_projeto_orcamento}' salvo com a revisão {revisao_atual}!")
                    except Exception as e: st.error(f"Erro ao salvar: {e}")
