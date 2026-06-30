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
        "Chave Seletora Auto/Manual (Painel Elétrico)": {"in_agua": "Chave Auto / Manual", "in_comp": "Chave Auto / Manual", "out_agua": "Habilita Equipamento (TAG)", "out_comp": "Habilita Equipamento (TAG)"},
        "Chave de fluxo para água (FS/CF)": {"in_agua": "Status Fluxo de Água", "in_comp": "", "out_agua": "", "out_comp": ""}
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
                    "suprimentos": "1234", "obras": "1234", "ricardo.pires": "1234"
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
        "Chave de fluxo para água (FS/CF)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
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
        "Chave de fluxo para água (FS/CF)": 450.00,
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
            "Transmissor de temperatura e umidade para duto (TT/MT)", 
            "Transmissor de temperatura para duto (TT)",
            "Válvula de controle proporcional com atuador (TCV)", 
            "Válvula de controle de água gelada proporcional (TCV)",
            "Válvula de controle de água quente proporcional (TCV)",
            "Válvula de controle de vapor proporcional (TCV)",
            "Relé de Corrente - Status Compressor (TC)", 
            "Termostato de segurança (TSH)",
            "Pressostato diferencial para ar (PSH)", 
            "Resistência de aquecimento (Equipamento) (RAQ)"
        ],
        "🔸 Monitoramento (Filtros e Status)": [
            "Pressostato para monitorar os filtros G4 (PSH)", 
            "Pressostato para monitorar os filtros M5 (PSH)", 
            "Pressostato para monitorar os filtros F9 (PSH)", 
            "Pressostato para monitorar os filtros H13/H14 (PSH)",
            "Status funcionamento ventilador ou exaustor (partida direta) (PSH)", 
            "Transmissor de pressão diferencial (monitorar os filtros G4) (PDT)", 
            "Transmissor de pressão diferencial (monitorar os filtros F9) (PDT)",
            "Transmissor de pressão diferencial (monitorar os filtros H13) (PDT)",
            "Chave de fluxo para água (FS/CF)"
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
            if "(PSH)" in inst_upper or "(TC)" in inst_upper or "(TSH)" in inst_upper or "RAQ" in inst_upper or "RESISTÊNCIA" in inst_upper or "EXAUSTOR" in inst_upper or "VENTILADOR" in inst_upper or "COMPRESSOR" in inst_upper or "FLUXO" in inst_upper or "FS" in inst_upper or "CF" in inst_upper: return "2x1,00mm²"
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
        
    def get_specific_tags(inst_nome, tags_lista, is_compressor_sys):
        tags_validas = [t for t in tags_lista if t.strip()]
        if not tags_validas: return ""
        if is_compressor_sys:
            inst_up = inst_nome.upper()
            if "COMPRESSOR" in inst_up or "TC" in inst_up or "CONDENSADOR" in inst_up:
                subset = [t for t in tags_validas if "UC" in t.upper() or "COND" in t.upper() or "COMP" in t.upper()]
                if subset: return "/".join(subset)
            else:
                subset = [t for t in tags_validas if "UE" in t.upper() or "EVAP" in t.upper() or "UTA" in t.upper()]
                if subset: return "/".join(subset)
        return "/".join(tags_validas)

    aba_auto, aba_infra, aba_precos, aba_resumo = st.tabs([
        "🚀 Dimensionamento de Automação", "🔌 Infraestrutura Lançamento", "💲 Base de Preços", "📊 Orçamento Final"
    ])

    help_cfr = "Qualificável: O sistema é entregue pronto e aberto para qualificação por terceiros.\nQualificado: A SIARCON executa e entrega a documentação de validação (Protocolos - CFR-21 Part 11)."

    with aba_auto:
        with st.expander("🔮 Módulo de Importação em Lote: Mapeamento Sênior", expanded=True):
            st.markdown("Para garantir 100% de precisão sem depender de conexões instáveis com IA para leitura de PDF na nuvem, utilize o mapeamento rápido. O sistema estruturará os quadros e as lógicas de forma exata.")
            
            arquivos_diagrama = st.file_uploader("Carregar Diagramas / P&ID (Permite Múltiplos):", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)
            
            if arquivos_diagrama:
                st.info(f"💡 {len(arquivos_diagrama)} arquivo(s) carregado(s). Verifique as configurações antes de gerar.")
                
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
                                ["CFR21 Part 11 - Qualificável", "CFR21 Part 11 - Qualificado"], help=help_cfr
                            )
                
                calibracao_ia = c_ia2.radio("Os instrumentos serão calibrados?", ["Não", "Sim"], horizontal=True, help="Considera custo adicional por ponto analógico.")
                sobra_ia = c_ia2.radio("Considerar 20% de sobra nas I/O (Reserva Técnica)?", ["Não", "Sim"], horizontal=True)

                st.markdown("##### 🔀 Validação e Mapeamento de Arquivos")
                st.write("Valide rapidamente a classificação de cada documento importado para a injeção da engenharia:")
                
                mapa_arquivos = []
                for i, arq in enumerate(arquivos_diagrama):
                    with st.container(border=True):
                        c_n, c_t, c_f, c_q = st.columns([2.5, 2, 3, 1.5])
                        c_n.markdown(f"📄 **{arq.name}**")
                        
                        n_up = arq.name.upper()
                        def_tipo = "UTA / Máquina Principal"
                        if "EXAUST" in n_up or "VENT" in n_up: def_tipo = "Exaustão / Ventilação"
                        elif "SALA" in n_up or "MONIT" in n_up: def_tipo = "Monitoramento de Salas"
                        
                        tipo = c_t.selectbox("Classe:", ["UTA / Máquina Principal", "Exaustão / Ventilação", "Monitoramento de Salas"], index=["UTA / Máquina Principal", "Exaustão / Ventilação", "Monitoramento de Salas"].index(def_tipo), key=f"tipo_{i}")
                        
                        filtros = []
                        if tipo != "Monitoramento de Salas":
                            def_filt = ["G4 (PSH)"] if tipo == "UTA / Máquina Principal" else []
                            filtros = c_f.multiselect("Filtros / Adicionais:", ["G4 (PSH)", "M5 (PSH)", "F9 (PDT)", "H13/H14 (PDT)", "Resistência Elétrica"], default=def_filt, key=f"filt_{i}")
                        
                        tag_d = c_q.text_input("Quadro Destino:", value="QTA-Geral", key=f"qta_{i}")
                        mapa_arquivos.append({"nome": arq.name, "tipo": tipo, "filtros": filtros, "quadro": tag_d})
                    
                if st.button("🪄 Gerar Quadros", type="primary"):
                    with st.spinner("Estruturando engenharia e agrupando painéis..."):
                        quadros_agrupados = {}
                        for config in mapa_arquivos:
                            t_q = config["quadro"]
                            if t_q not in quadros_agrupados: quadros_agrupados[t_q] = []
                            quadros_agrupados[t_q].append(config)
                            
                        for tag_quadro, listas_configs in quadros_agrupados.items():
                            grupos_equip = []
                            for config in listas_configs:
                                inst_dict = {k: 0 for k in REGRA_IO.keys()}
                                nome_base = config["nome"].rsplit('.', 1)[0]
                                
                                if config["tipo"] == "UTA / Máquina Principal":
                                    inst_dict["Transmissor de pressão dif. para ar (medição de vazão de ar) (PDT)"] = 1
                                    inst_dict["Transmissor de temperatura e umidade para duto (TT/MT)"] = 1
                                    inst_dict["Válvula de controle de água gelada proporcional (TCV)"] = 1
                                    inst_dict["Relé de Corrente - Status Compressor (TC)"] = 2
                                    if "Resistência Elétrica" in config["filtros"]:
                                        inst_dict["Resistência de aquecimento (Equipamento) (RAQ)"] = 1
                                        inst_dict["Termostato de segurança (TSH)"] = 1
                                        inst_dict["Pressostato diferencial para ar (PSH)"] = 1
                                elif config["tipo"] == "Exaustão / Ventilação":
                                    inst_dict["Status funcionamento ventilador ou exaustor (partida direta) (PSH)"] = 1
                                elif config["tipo"] == "Monitoramento de Salas":
                                    inst_dict["Transmissor de pressão diferencial entre salas (PDT)"] = 1
                                    inst_dict["Transmissor de temperatura e umidade ambiente (TT/MT)"] = 1
                                
                                if "G4 (PSH)" in config["filtros"]: inst_dict["Pressostato para monitorar os filtros G4 (PSH)"] = 1
                                if "M5 (PSH)" in config["filtros"]: inst_dict["Pressostato para monitorar os filtros M5 (PSH)"] = 1
                                if "F9 (PDT)" in config["filtros"]: inst_dict["Transmissor de pressão diferencial (monitorar os filtros F9) (PDT)"] = 1
                                if "H13/H14 (PDT)" in config["filtros"]: inst_dict["Transmissor de pressão diferencial (monitorar os filtros H13) (PDT)"] = 1

                                grupos_equip.append({"nome_grupo": f"{config['tipo']} ({nome_base})", "multiplicador": 1, "instrumentos": inst_dict.copy(), "tags_lista": [""]})
                                
                            novo_quadro_ia = {
                                "id": str(uuid.uuid4()), "nome": tag_quadro, "tipo": "Controle (HVAC/Máquinas)",
                                "supervisorio": soft_sel_ia, "arquitetura": arquitetura_ia, "tipo_cfr": tipo_cfr_ia,
                                "modo_config": "Importado", "ihm": ihm_ia, "calibracao": calibracao_ia,
                                "sobra_20": sobra_ia, "tags_nao_reconhecidas": [], "grupos_equipamentos": grupos_equip
                            }
                            st.session_state.paineis_auto.insert(0, novo_quadro_ia)
                    st.success("✅ Varredura e organização concluída com sucesso!")
                    st.rerun()

        if not st.session_state.wizard_ativo:
            if st.button("➕ Criar Novo Quadro Manualmente", type="primary"):
                st.session_state.wizard_ativo = True
                st.rerun()

        if st.session_state.wizard_ativo:
            with st.container(border=True):
                st.markdown("<div style='background-color: rgba(28, 133, 144, 0.15); padding: 15px; border-radius: 8px;'><h3 style='margin:0; color: #1C8590;'>🧙‍♂️ Assistente de Configuração</h3></div><br>", unsafe_allow_html=True)
                
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
                                ["CFR21 Part 11 - Qualificável", "CFR21 Part 11 - Qualificado"], help=help_cfr
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
                    elif config_opt is None: st.warning("⚠️ Responda a pergunta 7.")
                    elif config_opt == "Usar Padrão Existente (Kits)" and kit_final_selecionado == "Selecione...": st.warning("⚠️ Selecione um kit padrão.")
                    else:
                        novos_instrumentos = {k: 0 for k in REGRA_IO.keys()}
                        grupos_equip = []
                        if config_opt == "Usar Padrão Existente (Kits)":
                            for item_nome, qtd_padrao in KITS_PADRAO[kit_final_selecionado].items():
                                if item_nome in novos_instrumentos: novos_instrumentos[item_nome] = qtd_padrao
                            nome_limpo = kit_final_selecionado.split(" ", 1)[1] if " " in kit_final_selecionado else kit_final_selecionado
                            grupos_equip.append({"nome_grupo": f"{nome_limpo}", "multiplicador": 1, "instrumentos": novos_instrumentos.copy(), "tags_lista": [""]})
                        else:
                            grupos_equip.append({"nome_grupo": "Equipamento Novo", "multiplicador": 1, "instrumentos": novos_instrumentos.copy(), "tags_lista": [""]})

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
                                p_data['grupos_equipamentos'].insert(0, {"nome_grupo": "Equipamento Novo", "multiplicador": 1, "instrumentos": novos_inst.copy(), "tags_lista": [""]})
                            else:
                                for item_nome, qtd_padrao in KITS_PADRAO[sub_kit].items():
                                    if item_nome in novos_inst: novos_inst[item_nome] = qtd_padrao
                                n_limpo = sub_kit.split(" ", 1)[1] if " " in sub_kit else sub_kit
                                p_data['grupos_equipamentos'].insert(0, {"nome_grupo": f"{n_limpo}", "multiplicador": 1, "instrumentos": novos_inst.copy(), "tags_lista": [""]})
                            st.rerun()

                raw_ai_painel = raw_ao_painel = raw_di_painel = raw_do_painel = 0

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
                            
                            is_compressor_sys = "COMPRESSOR" in g_data['nome_grupo'].upper() or "DIRETA" in g_data['nome_grupo'].upper() or "DX" in g_data['nome_grupo'].upper()
                            
                            auto_mon = "SALA" in g_data['nome_grupo'].upper() or "MONITORAMENTO" in g_data['nome_grupo'].upper()
                            is_monitoramento = st.checkbox("📍 Exclusivamente para Monitoramento de Salas (Desabilita chaves de intertravamento)", value=auto_mon, key=f"chk_mon_{p_data['id']}_{g_idx}")

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
                                        st.error(f"⚠️ CAPACIDADE EXCEDIDA para Mercato (UI: {ui_check}, AO: {ao_check}, DO: {do_check}).")

                            with st.expander("👁️ Visualizar Diagrama P&ID (Lógica e TAGs)", expanded=False):
                                def limpa_str(texto): return str(texto).replace('"', "''").replace('\n', ' ')
                                
                                dot = f'digraph G {{\n  rankdir=LR;\n  node [fontname="Arial", fontsize=10, shape=box, style=rounded];\n'
                                if p_data.get('ihm') and "Cego" not in p_data['ihm']:
                                    ihm_nome = p_data["ihm"].replace('Mercato - ', '').replace('IHM Padrão ', '').replace('IHM Premium ', '')
                                    dot += f'  "IHM" [label="{limpa_str(ihm_nome)}\\n(Painel)", fillcolor="#D5F5E3", style=filled];\n  "IHM" -> "Controlador" [dir=both, style=dashed];\n'
                                    
                                arq_nome = limpa_str(p_data.get("arquitetura", "Controlador"))
                                grupo_nome = limpa_str(g_data["nome_grupo"])
                                dot += f'  "Controlador" [label="{arq_nome}\\n({grupo_nome})", fillcolor="#1C8590", style=filled, fontcolor=white, shape=ellipse];\n'
                                
                                has_inputs = has_outputs = False
                                node_idx = 0
                                group_boxes = is_monitoramento or ("EXAUST" in grupo_nome.upper())
                                
                                for inst_f, q_f in g_data['instrumentos'].items():
                                    if q_f > 0:
                                        q_int = int(q_f)
                                        io_v = REGRA_IO.get(inst_f, {"AI": 0, "AO": 0, "DI": 0, "DO": 0})
                                        tag_hardware = inst_f.split('(')[-1].replace(')', '').strip() if '(' in inst_f else 'IO'
                                        tag_hardware = limpa_str(tag_hardware)
                                        
                                        c_names = st.session_state.de_para_diagrama.get(inst_f, {})
                                        if isinstance(c_names, str): lbl_in = lbl_out = c_names
                                        else:
                                            if is_compressor_sys:
                                                lbl_in = c_names.get("in_comp", c_names.get("in_agua", ""))
                                                lbl_out = c_names.get("out_comp", c_names.get("out_agua", ""))
                                            else:
                                                lbl_in = c_names.get("in_agua", c_names.get("in_comp", ""))
                                                lbl_out = c_names.get("out_agua", c_names.get("out_comp", ""))
                                        
                                        if not lbl_in: lbl_in = inst_f.split('(')[0].strip()
                                        if not lbl_out: lbl_out = inst_f.split('(')[0].strip()
                                        
                                        force_out = isinstance(c_names, dict) and (str(c_names.get("out_agua", "")).strip() not in ["", "nan"] or str(c_names.get("out_comp", "")).strip() not in ["", "nan"])
                                        force_in = isinstance(c_names, dict) and (str(c_names.get("in_agua", "")).strip() not in ["", "nan"] or str(c_names.get("in_comp", "")).strip() not in ["", "nan"])

                                        has_in_pin = io_v["AI"] > 0 or io_v["DI"] > 0 or force_in
                                        has_out_pin = io_v["AO"] > 0 or io_v["DO"] > 0 or force_out
                                        
                                        if not has_in_pin and not has_out_pin: continue
                                        
                                        tags_inst = [t for t in g_data['tags_lista'] if t.strip()]
                                        if is_compressor_sys and ("UTA" in grupo_nome.upper() or "SISTEMA" in grupo_nome.upper()):
                                            if "COMPRESSOR" in inst_f.upper() or "TC" in inst_f.upper() or "CONDENSADOR" in inst_f.upper():
                                                tags_inst = [t for t in tags_inst if "UC" in t.upper() or "COND" in t.upper() or "COMP" in t.upper()]
                                            else:
                                                tags_inst = [t for t in tags_inst if "UE" in t.upper() or "EVAP" in t.upper() or "UTA" in t.upper()]
                                            if not tags_inst: tags_inst = [t for t in g_data['tags_lista'] if t.strip()]
                                        
                                        lbl_in_limpo = limpa_str(lbl_in)
                                        lbl_out_limpo = limpa_str(lbl_out)
                                        if len(lbl_in_limpo) > 35: lbl_in_limpo = lbl_in_limpo[:35] + "..."
                                        if len(lbl_out_limpo) > 35: lbl_out_limpo = lbl_out_limpo[:35] + "..."
                                        
                                        if group_boxes:
                                            node_name = f"N_{node_idx}_grp"
                                            prefix = f"{q_int}x "
                                            str_tags = ", ".join(tags_inst)
                                            str_tag_ctx = f"\\n({limpa_str(str_tags)})" if str_tags else ""
                                            
                                            if has_in_pin and lbl_in_limpo and str(lbl_in_limpo).strip() not in ["", "nan"]:
                                                dot += f'  "{node_name}_in" [label="{prefix}{lbl_in_limpo}{str_tag_ctx}\\nTAG: {tag_hardware}", color="#2B7BC4"];\n'
                                                dot += f'  "{node_name}_in" -> "Controlador" [label="{obter_cabo(inst_f, False)}", fontsize=8, color="#2B7BC4"];\n'
                                                has_inputs = True
                                                
                                            if not is_monitoramento and has_out_pin and lbl_out_limpo and str(lbl_out_limpo).strip() not in ["", "nan"]:
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
                                                    dot += f'  "{node_name}_out" [label="{prefix}{lbl_out_limpo}{str_tag_ctx}\\nTAG: {tag_hardware}", color="#E14D2A"];\n'
                                                    dot += f'  "Controlador" -> "{node_name}_out" [label="{obter_cabo(inst_f, True)}", fontsize=8, color="#E14D2A"];\n'
                                                    has_outputs = True
                                        else:
                                            for idx_q in range(q_int):
                                                node_name = f"N_{node_idx}_{idx_q}"
                                                lbl_suf = f" {idx_q+1}" if q_int > 1 else ""
                                                tag_contexto = tags_inst[idx_q % len(tags_inst)] if tags_inst else ""
                                                if "vazão de ar" in inst_f.lower() and "/" in tag_contexto:
                                                    tag_contexto = tag_contexto.split('/')[0].strip()
                                                str_tag_ctx = f"\\n({limpa_str(tag_contexto)})" if tag_contexto else ""
                                                
                                                if has_in_pin and lbl_in_limpo and str(lbl_in_limpo).strip() not in ["", "nan"]:
                                                    dot += f'  "{node_name}_in" [label="{lbl_in_limpo}{lbl_suf}{str_tag_ctx}\\nTAG: {tag_hardware}", color="#2B7BC4"];\n'
                                                    dot += f'  "{node_name}_in" -> "Controlador" [label="{obter_cabo(inst_f, False)}", fontsize=8, color="#2B7BC4"];\n'
                                                    has_inputs = True
                                                    
                                                if not is_monitoramento and has_out_pin and lbl_out_limpo and str(lbl_out_limpo).strip() not in ["", "nan"]:
                                                    if "Resistência de aquecimento" in inst_f:
                                                        dot += f'  "{node_name}_out_DO" [label="Habilita RAQ{lbl_suf}{str_tag_ctx}\\nTAG: DO", color="#E14D2A"];\n'
                                                        dot += f'  "Controlador" -> "{node_name}_out_DO" [label="2x1,00mm²", fontsize=8, color="#E14D2A"];\n'
                                                        dot += f'  "{node_name}_out_AO" [label="Modulação Resistência{lbl_suf}{str_tag_ctx}\\nTAG: AO", color="#E14D2A"];\n'
                                                        dot += f'  "Controlador" -> "{node_name}_out_AO" [label="3x0,75mm² + Shield", fontsize=8, color="#E14D2A"];\n'
                                                        has_outputs = True
                                                    elif "medição de vazão de ar" in inst_f:
                                                        dot += f'  "{node_name}_out" [label="Modula Inversor{lbl_suf}{str_tag_ctx}\\nTAG: AO", color="#E14D2A"];\n'
                                                        dot += f'  "Controlador" -> "{node_name}_out" [label="3x0,75mm² + Shield", fontsize=8, color="#E14D2A"];\n'
                                                        has_outputs = True
                                                    else:
                                                        dot += f'  "{node_name}_out" [label="{lbl_out_limpo}{lbl_suf}{str_tag_ctx}\\nTAG: {tag_hardware}", color="#E14D2A"];\n'
                                                        dot += f'  "Controlador" -> "{node_name}_out" [label="{obter_cabo(inst_f, True)}", fontsize=8, color="#E14D2A"];\n'
                                                        has_outputs = True
                                        node_idx += 1
                                        
                                if not is_monitoramento:
                                    inst_chave = "Chave Seletora Auto/Manual (Painel Elétrico)"
                                    c_names = st.session_state.de_para_diagrama.get(inst_chave, {})
                                    lbl_in_c = limpa_str(str(c_names.get("in_comp", "")) if is_compressor_sys else str(c_names.get("in_agua", "")))
                                    lbl_out_c = limpa_str(str(c_names.get("out_comp", "")) if is_compressor_sys else str(c_names.get("out_agua", "")))
                                    prefix_c = f"{int(qtd_atual)}x " if int(qtd_atual) > 1 and group_boxes else ""
                                    
                                    if lbl_in_c.strip() not in ["", "nan"]:
                                        dot += f'  "chave_in" [label="{prefix_c}{lbl_in_c}\\nTAG: CH", color="#2B7BC4"];\n'
                                        dot += f'  "chave_in" -> "Controlador" [label="5x1,00mm²", fontsize=8, color="#2B7BC4"];\n'
                                        has_inputs = True
                                    if lbl_out_c.strip() not in ["", "nan"]:
                                        dot += f'  "chave_out" [label="{prefix_c}{lbl_out_c}\\nTAG: CH", color="#E14D2A"];\n'
                                        dot += f'  "Controlador" -> "chave_out" [label="5x1,00mm²", fontsize=8, color="#E14D2A"];\n'
                                        has_outputs = True

                                if not has_inputs: dot += '  "Sinais de Campo" -> "Controlador" [style=dashed];\n'
                                if not has_outputs and not is_monitoramento: dot += '  "Controlador" -> "Atuadores" [style=dashed];\n'
                                dot += '}'
                                
                                try:
                                    st.graphviz_chart(dot)
                                    try:
                                        import graphviz
                                        src = graphviz.Source(dot)
                                        png_bytes = src.pipe(format='png')
                                        st.download_button(label="📥 Baixar Imagem (.PNG)", data=png_bytes, file_name=f"Diagrama_{g_data['nome_grupo']}.png", mime="image/png", key=f"dl_png_{p_data['id']}_{g_idx}")
                                    except: pass
                                except: pass

                            with st.expander("⚙️ Ajuste Fino de Instrumentos (Engenharia)"):
                                for grupo_nome, lista_itens in GRUPOS_INSTRUMENTOS.items():
                                    with st.expander(grupo_nome, expanded=False):
                                        cols_inst = st.columns(2)
                                        for i, inst in enumerate(lista_itens):
                                            if inst not in g_data['instrumentos']: g_data['instrumentos'][inst] = 0
                                            g_data['instrumentos'][inst] = cols_inst[i % 2].number_input(inst, min_value=0, step=1, value=g_data['instrumentos'][inst], key=f"inst_{p_data['id']}_{g_idx}_{grupo_nome}_{inst}")
                            st.write("")
                            if st.button("🗑️ Remover Máquina", key=f"del_{p_data['id']}_{g_idx}"):
                                p_data['grupos_equipamentos'].pop(g_idx)
                                st.rerun()

                for g_data in p_data['grupos_equipamentos']:
                    qtd_atual_calc = g_data.get('multiplicador', 1)
                    is_mon = st.session_state.get(f"chk_mon_{p_data['id']}_{p_data['grupos_equipamentos'].index(g_data)}", ("SALA" in g_data['nome_grupo'].upper() or "MONITORAMENTO" in g_data['nome_grupo'].upper()))
                    for inst, q in g_data['instrumentos'].items():
                        io_vals = REGRA_IO.get(inst, {"AI": 0, "AO": 0, "DI": 0, "DO": 0})
                        raw_ai_painel += q * io_vals["AI"] * qtd_atual_calc
                        raw_ao_painel += q * io_vals["AO"] * qtd_atual_calc
                        raw_di_painel += q * io_vals["DI"] * qtd_atual_calc
                        raw_do_painel += q * io_vals["DO"] * qtd_atual_calc

                    if not is_mon:
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
            
            if c_bot_salvar.button("💾 Salvar Rascunho e Sair", type="primary", use_container_width=True):
                if not st.session_state.nome_projeto_orcamento: st.warning("⚠️ Atenção: Preencha o 'Nome do Orçamento / Projeto'.")
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
                        st.toast("📝 Rascunho salvo na nuvem!", icon="💾")
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
                                    st.toast("🗑️ Orçamento excluído!", icon="✅")
                                    st.rerun()
                                except Exception as e: st.error(f"Erro ao excluir: {e}")
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
            except: pass

    with aba_infra:
        st.header("Cálculo de Infraestrutura Lançamento Manual")
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
        st.info(f"📅 **Última atualização:** {st.session_state.data_precos_atualizada}")
        
        c_cot1, c_cot2 = st.columns(2)
        with c_cot1:
            def gerar_planilha_cotacao(com_precos_atuais):
                buffer_cotacao = io.BytesIO()
                wb_cot = openpyxl.Workbook()
                wb_cot.remove(wb_cot.active)
                cat_dict = {
                    "Geral e Schneider": list(banco_schneider_comum.keys()) + ["IHM Padrão 7\"", "IHM Premium 10\""],
                    "Siemens": list(banco_siemens.keys()), "Mercato": list(banco_mercato.keys()), "Serviços CFR-21": list(banco_cfr_servicos.keys())
                }
                fill_h = PatternFill(start_color="1C8590", end_color="1C8590", fill_type="solid")
                font_h = Font(bold=True, color="FFFFFF")
                for cat_nome, itens_cat in cat_dict.items():
                    ws_cot = wb_cot.create_sheet(title=cat_nome[:31])
                    if com_precos_atuais: ws_cot.append(["Item / Equipamento", "Preço Atual (R$)", "Novo Preço (R$)"]); max_col = 3
                    else: ws_cot.append(["Item / Equipamento", "Novo Preço (R$)"]); max_col = 2
                    for col in range(1, max_col + 1):
                        cell = ws_cot.cell(row=1, column=col); cell.font = font_h; cell.fill = fill_h; cell.alignment = Alignment(horizontal="center")
                    for item_cat in itens_cat:
                        if com_precos_atuais: ws_cot.append([item_cat, st.session_state.precos_banco.get(item_cat, 0.0), ""])
                        else: ws_cot.append([item_cat, ""])
                    ws_cot.column_dimensions['A'].width = 60
                    if com_precos_atuais: ws_cot.column_dimensions['B'].width = 20; ws_cot.column_dimensions['C'].width = 25
                    else: ws_cot.column_dimensions['B'].width = 25
                wb_cot.save(buffer_cotacao); buffer_cotacao.seek(0)
                return buffer_cotacao

            st.download_button("📥 Planilha Cotação (Com Preços Atuais)", data=gerar_planilha_cotacao(True), file_name="Cotacao_Com_Precos.xlsx", use_container_width=True)
            st.download_button("📥 Planilha Cotação (Em Branco)", data=gerar_planilha_cotacao(False), file_name="Cotacao_Em_Branco.xlsx", use_container_width=True)

        with c_cot2:
            upload_precos = st.file_uploader("📂 Importar Planilha de Cotação Respondida", type=["xlsx", "xls"], label_visibility="collapsed")
            if upload_precos is not None:
                try:
                    xls_precos = pd.ExcelFile(upload_precos); atualizados_count = 0
                    for sheet in xls_precos.sheet_names:
                        df_sheet = pd.read_excel(xls_precos, sheet_name=sheet)
                        if "Item / Equipamento" in df_sheet.columns and "Novo Preço (R$)" in df_sheet.columns:
                            for _, row in df_sheet.iterrows():
                                item_nome = row["Item / Equipamento"]
                                novo_pr = row["Novo Preço (R$)"]
                                if pd.notna(item_nome) and pd.notna(novo_pr) and str(novo_pr).strip() != "":
                                    try:
                                        val_clean = str(novo_pr).replace('R$', '').replace(' ', '')
                                        if ',' in val_clean and '.' in val_clean: val_clean = val_clean.replace('.', '').replace(',', '.')
                                        elif ',' in val_clean: val_clean = val_clean.replace(',', '.')
                                        val_float = float(val_clean)
                                        if st.session_state.precos_banco.get(item_nome) != val_float:
                                            st.session_state.precos_banco[item_nome] = val_float; atualizados_count += 1
                                    except: pass
                    if atualizados_count > 0: st.success(f"✅ {atualizados_count} preços atualizados! Salve no banco de dados abaixo.")
                except: st.error("Erro ao ler planilha.")
                    
        st.markdown("---")
        st.subheader("Bases de Preços (Edição Manual)")
        edited_geral = st.data_editor(pd.DataFrame([{"Item / Equipamento": k, "Valor Atual (R$)": st.session_state.precos_banco.get(k, 0.0)} for k in list(banco_schneider_comum.keys()) + ["IHM Padrão 7\"", "IHM Premium 10\""]]), use_container_width=True, hide_index=True)
        st.subheader("Base Siemens")
        edited_siemens = st.data_editor(pd.DataFrame([{"Item / Equipamento": k, "Valor Atual (R$)": st.session_state.precos_banco.get(k, 0.0)} for k in banco_siemens.keys()]), use_container_width=True, hide_index=True)
        st.subheader("Base Mercato")
        edited_mercato = st.data_editor(pd.DataFrame([{"Item / Equipamento": k, "Valor Atual (R$)": st.session_state.precos_banco.get(k, 0.0)} for k in banco_mercato.keys()]), use_container_width=True, hide_index=True)
        st.subheader("Serviços CFR-21")
        edited_cfr = st.data_editor(pd.DataFrame([{"Item / Equipamento": k, "Valor Atual (R$)": st.session_state.precos_banco.get(k, 0.0)} for k in banco_cfr_servicos.keys()]), use_container_width=True, hide_index=True)
        
        if st.button("💾 Salvar Novos Preços no Banco de Dados", type="primary"):
            alt = False; nh = []
            edt_tot = pd.concat([edited_geral, edited_siemens, edited_mercato, edited_cfr], ignore_index=True)
            dh_agora = datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M:%S")
            
            for _, r in edt_tot.iterrows():
                it = r['Item / Equipamento']; n_val = r['Valor Atual (R$)']; a_val = st.session_state.precos_banco.get(it, 0.0)
                if n_val != a_val:
                    nh.append({"Data/Hora": dh_agora, "Item Alterado": it, "Valor Antigo": f"R$ {a_val:.2f}", "Novo Valor": f"R$ {n_val:.2f}"})
                    st.session_state.precos_banco[it] = n_val; alt = True
                    
            if alt:
                try:
                    sh = conectar_google_sheets()
                    try: ws_p = sh.worksheet("Precos")
                    except: ws_p = sh.add_worksheet("Precos", 100, 2); ws_p.append_row(["Item", "Valor"])
                    ws_p.clear(); ws_p.append_rows([["Item", "Valor"]] + [[k, v] for k, v in st.session_state.precos_banco.items()])
                    try: ws_h = sh.worksheet("Historico_Precos")
                    except: ws_h = sh.add_worksheet("Historico_Precos", 1000, 4); ws_h.append_row(["Data/Hora", "Item Alterado", "Valor Antigo", "Novo Valor"])
                    if nh: ws_h.append_rows([[h["Data/Hora"], h["Item Alterado"], h["Valor Antigo"], h["Novo Valor"]] for h in nh])
                    
                    st.session_state.data_precos_atualizada = dh_agora
                    st.cache_data.clear()
                    st.toast("✅ Base atualizada!", icon="💾"); st.rerun()
                except Exception as e: st.error(f"Erro: {e}")
            else: st.info("Sem alterações.")

    with aba_resumo:
        st.header("Consolidação Financeira do Orçamento")
        linhas_inst, linhas_hw, linhas_sw, linhas_pt, linhas_srv = [], [], [], [], []
        sw_inc = {}; t_ai_sch = t_ao_sch = t_di_sch = t_do_sch = 0; t_ai_siem = t_ao_siem = t_di_siem = t_do_siem = 0; t_io_merc = 0
        c_sch = c_siem = c_merc = 0.0
        desc_linhas = ["DESCRIÇÃO TÉCNICA DO ESCOPO CONTEMPLADO:\n"]; desc_comercial = []

        for p in st.session_state.paineis_auto:
            arq_at = p.get('arquitetura', 'SpaceLogic (Schneider)')
            is_siem = 'Siemens' in arq_at; is_merc = 'Mercato' in arq_at; is_sch = 'Schneider' in arq_at
            tem_s20 = p.get('sobra_20', 'Não') == 'Sim'; cal_ativa = p.get('calibracao', 'Não') == 'Sim'
            t_cfr_p = p.get('tipo_cfr', 'Não Aplicável')
            
            r_ai = r_ao = r_di = r_do = q_eq = t_calib = 0
            l_eq_nm = []; tem_res = tem_f_pdt = tem_f_psh = False; l_inst_nm = set(); l_inst_det = []; ctrl_desc = []
            
            for g_idx, g in enumerate(p['grupos_equipamentos']):
                mult = g.get('multiplicador', 1); q_eq += mult
                l_tags = [t for t in g.get('tags_lista', []) if t.strip()]
                str_tags = f" [TAGs: {', '.join(l_tags)}]" if l_tags else ""
                n_limpo = g['nome_grupo'].replace("Equipamento Novo", "").strip() or "Equipamento"
                l_eq_nm.append(f"{mult}x {n_limpo}{str_tags}")
                
                is_mon = st.session_state.get(f"chk_mon_{p['id']}_{g_idx}", ("SALA" in g['nome_grupo'].upper() or "MONITORAMENTO" in g['nome_grupo'].upper()))
                is_comp = "COMPRESSOR" in g['nome_grupo'].upper() or "DIRETA" in g['nome_grupo'].upper()
                
                ai_g = ao_g = di_g = do_g = 0
                for inst, qtd in g['instrumentos'].items():
                    if qtd > 0:
                        q_f = qtd * mult; i_real = inst
                        if cal_ativa:
                            i_up = inst.upper()
                            if "(TT/MT)" in i_up or "(TIT/MIT)" in i_up: t_calib += (2 * q_f)
                            elif "(PDT)" in i_up or "(PDIT)" in i_up or "(TT)" in i_up or "(PIT)" in i_up or "(FIT)" in i_up or "(TIT)" in i_up: t_calib += q_f
                        
                        if is_merc:
                            if "temperatura para duto" in inst: i_real = "Mercato - Sensor de Temperatura NTC (Duto)"
                            elif "temperatura Ambiente" in inst: i_real = "Mercato - Sensor de Temperatura NTC (Ambiente)"
                            elif "ambiente com display" in inst: i_real = "Mercato - Sensor de Temperatura NTC com Display (Ambiente)"
                        elif is_sch:
                            if "temperatura para duto" in inst: i_real = "Schneider - Sensor de Temperatura NTC (Duto)"
                            elif "temperatura Ambiente" in inst: i_real = "Schneider - Sensor de Temperatura NTC (Ambiente)"
                        
                        nm_curto = re.sub(r'\s*\([A-Z/]+\)$', '', i_real.replace("Mercato - ", "").replace("Schneider - ", "").replace("Siemens - ", ""))
                        l_inst_nm.add(nm_curto)
                        
                        if "resistência" in inst.lower() or "raq" in inst.lower() or "tsh" in inst.lower(): tem_res = True
                        if "filtro" in inst.lower():
                            if "pdit" in inst.lower() or "pdt" in inst.lower(): tem_f_pdt = True
                            if "psh" in inst.lower() or "pressostato" in inst.lower(): tem_f_psh = True

                        f_i = "Medição Genérica"
                        if "vazão de ar" in inst.lower(): f_i = "Medição da Vazão de Ar"
                        elif "temperatura e umidade" in inst.lower(): f_i = "Medição de Temperatura e Umidade"
                        elif "temperatura" in inst.lower(): f_i = "Medição de Temperatura"
                        elif "proporcional" in inst.lower(): f_i = "Controle Proporcional Térmico"
                        elif "compressor" in inst.lower(): f_i = "Status de Operação do Compressor"
                        elif "termostato" in inst.lower(): f_i = "Segurança de Sobreaquecimento"
                        elif "resistência" in inst.lower() and "pressostato" not in inst.lower(): f_i = "Aquecimento"
                        elif "pressostato diferencial para ar" in inst.lower() and "resistência" in inst.lower(): f_i = "Segurança da Resistência por Fluxo de Ar"
                        elif "pressostato para monitorar" in inst.lower(): f_i = "Alarme de Saturação de Filtro"
                        elif "pressão diferencial (monitorar" in inst.lower(): f_i = "Monitoramento da Saturação do Filtro"
                        elif "ventilador" in inst.lower() or "exaustor" in inst.lower(): f_i = "Status de Operação do Ventilador/Exaustor"
                        
                        l_inst_det.append((nm_curto, q_f, f_i))
                        
                        pr_i = st.session_state.precos_banco.get(i_real, st.session_state.precos_banco.get(inst, 0.0))
                        io_v = REGRA_IO.get(inst, {"AI":0, "AO":0, "DI":0, "DO":0})
                        
                        r_ai += q_f * io_v["AI"]; r_ao += q_f * io_v["AO"]; r_di += q_f * io_v["DI"]; r_do += q_f * io_v["DO"]
                        ai_g += qtd * io_v["AI"]; ao_g += qtd * io_v["AO"]; di_g += qtd * io_v["DI"]; do_g += qtd * io_v["DO"]
                        
                        c_tot = q_f * pr_i
                        linhas_inst.append({"Categoria": "Instrumentação", "Item": f"{i_real} ({n_limpo} - {p['nome']})", "Preço Unit.": pr_i, "Qtd": q_f, "Custo Total": c_tot})
                        linhas_pt.append({"Painel": p['nome'], "Grupo/Equipamento": n_limpo, "Instrumento": i_real, "Quantidade Total": q_f, "DI": q_f*io_v["DI"], "DO": q_f*io_v["DO"], "AI": q_f*io_v["AI"], "AO": q_f*io_v["AO"]})
                        
                        if is_siem: c_siem += c_tot
                        elif is_merc: c_merc += c_tot
                        else: c_sch += c_tot
                
                if not is_mon:
                    r_di += (2 * mult); di_g += 2
                    linhas_pt.append({"Painel": p['nome'], "Grupo/Equipamento": n_limpo, "Instrumento": "Chave Auto/Manual", "Quantidade Total": mult, "DI": 2*mult, "DO": 0, "AI": 0, "AO": 0})

                if is_merc:
                    u_n = ai_g + di_g + (math.ceil((ai_g+di_g)*0.2) if tem_s20 else 0)
                    a_n = ao_g + (math.ceil(ao_g*0.2) if tem_s20 else 0)
                    d_n = do_g + (math.ceil(do_g*0.2) if tem_s20 else 0)
                    m_mcp = dimensionar_mercato(u_n, a_n, d_n, is_comp)
                    if m_mcp:
                        ctrl_desc.append(f"{mult}x {m_mcp.replace('Mercato - ', '')}")
                        p_hw = st.session_state.precos_banco.get(m_mcp, 1650.0)
                        linhas_hw.append({"Categoria": "Hardware", "Item": f"{m_mcp} ({n_limpo} - {p['nome']})", "Preço Unit.": p_hw, "Qtd": mult, "Custo Total": mult * p_hw})
                        c_merc += (mult * p_hw)

            if t_calib > 0:
                pr_cal = st.session_state.precos_banco.get("Serviço de Calibração (Por Ponto Analógico)", 180.0)
                linhas_srv.append({"Categoria": "Serviços", "Item": f"Calibração Analógicos ({p['nome']})", "Preço Unit.": pr_cal, "Qtd": t_calib, "Custo Total": t_calib * pr_cal})

            rs_ai = math.ceil(r_ai*0.2) if tem_s20 else 0; rs_ao = math.ceil(r_ao*0.2) if tem_s20 else 0
            rs_di = math.ceil(r_di*0.2) if tem_s20 else 0; rs_do = math.ceil(r_do*0.2) if tem_s20 else 0
            t_hw = r_ai+rs_ai + r_ao+rs_ao + r_di+rs_di + r_do+rs_do
            
            if rs_ai>0 or rs_ao>0 or rs_di>0 or rs_do>0:
                linhas_pt.append({"Painel": p['nome'], "Grupo/Equipamento": "Reserva 20%", "Instrumento": "Sobras", "Quantidade Total": "-", "DI": rs_di, "DO": rs_do, "AI": rs_ai, "AO": rs_ao})

            if is_siem: t_ai_siem += r_ai; t_ao_siem += r_ao; t_di_siem += r_di; t_do_siem += r_do
            elif is_merc: t_io_merc += (r_ai+r_ao+r_di+r_do)
            else: t_ai_sch += r_ai; t_ao_sch += r_ao; t_di_sch += r_di; t_do_sch += r_do
            
            if t_hw > 0:
                if is_merc:
                    n_cx, p_cx = calcular_painel_fisico(q_eq)
                    linhas_hw.append({"Categoria": "Hardware", "Item": f"Painel: {n_cx} ({p['nome']})", "Preço Unit.": p_cx, "Qtd": 1, "Custo Total": p_cx}); c_merc += p_cx
                else:
                    n_cx, p_cx = calcular_painel_fisico(t_hw/15)
                    linhas_hw.append({"Categoria": "Hardware", "Item": f"Painel: {n_cx} ({p['nome']})", "Preço Unit.": p_cx, "Qtd": 1, "Custo Total": p_cx})
                    if is_siem:
                        c_siem += p_cx
                        hw_s = dimensionar_siemens_1200(r_ai+rs_ai, r_ao+rs_ao, r_di+rs_di, r_do+rs_do) if '1200' in arq_at else dimensionar_siemens_1500(r_ai+rs_ai, r_ao+rs_ao, r_di+rs_di, r_do+rs_do)
                        for i_hw, q_hw in hw_s.items():
                            if q_hw > 0:
                                ctrl_desc.append(f"{q_hw}x {i_hw.replace('Siemens - ', '')}")
                                pr = st.session_state.precos_banco.get(i_hw, 0.0)
                                linhas_hw.append({"Categoria": "Hardware", "Item": f"{i_hw} ({p['nome']})", "Preço Unit.": pr, "Qtd": q_hw, "Custo Total": q_hw * pr}); c_siem += (q_hw * pr)
                    else:
                        c_sch += p_cx
                        c36, c24, c18, c15 = dimensionar_controladores(t_hw)
                        if c36>0: ctrl_desc.append(f"{c36}x MP-C-36A"); linhas_hw.append({"Categoria": "Hardware", "Item": f"MP-C-36A ({p['nome']})", "Preço Unit.": st.session_state.precos_banco.get("MP-C-36A", 9459.0), "Qtd": c36, "Custo Total": c36*st.session_state.precos_banco.get("MP-C-36A", 9459.0)}); c_sch += c36*st.session_state.precos_banco.get("MP-C-36A", 9459.0)
                        if c24>0: ctrl_desc.append(f"{c24}x MP-C-24A"); linhas_hw.append({"Categoria": "Hardware", "Item": f"MP-C-24A ({p['nome']})", "Preço Unit.": st.session_state.precos_banco.get("MP-C-24A", 7290.0), "Qtd": c24, "Custo Total": c24*st.session_state.precos_banco.get("MP-C-24A", 7290.0)}); c_sch += c24*st.session_state.precos_banco.get("MP-C-24A", 7290.0)
                        if c18>0: ctrl_desc.append(f"{c18}x MP-C-18A"); linhas_hw.append({"Categoria": "Hardware", "Item": f"MP-C-18A ({p['nome']})", "Preço Unit.": st.session_state.precos_banco.get("MP-C-18A", 5185.0), "Qtd": c18, "Custo Total": c18*st.session_state.precos_banco.get("MP-C-18A", 5185.0)}); c_sch += c18*st.session_state.precos_banco.get("MP-C-18A", 5185.0)
                        if c15>0: ctrl_desc.append(f"{c15}x MP-C-15A"); linhas_hw.append({"Categoria": "Hardware", "Item": f"MP-C-15A ({p['nome']})", "Preço Unit.": st.session_state.precos_banco.get("MP-C-15A", 4649.0), "Qtd": c15, "Custo Total": c15*st.session_state.precos_banco.get("MP-C-15A", 4649.0)}); c_sch += c15*st.session_state.precos_banco.get("MP-C-15A", 4649.0)
                
                if p.get('ihm') and "Cego" not in p['ihm']:
                    pr_ihm = st.session_state.precos_banco.get(p['ihm'], 0.0)
                    if pr_ihm > 0: 
                        linhas_hw.append({"Categoria": "Hardware", "Item": f"IHM: {p['ihm']} ({p['nome']})", "Preço Unit.": pr_ihm, "Qtd": 1, "Custo Total": pr_ihm})
                        if is_merc: c_merc += pr_ihm
                        elif is_siem: c_siem += pr_ihm
                        else: c_sch += pr_ihm

                s_tp = p.get('supervisorio', "Sem Supervisório")
                if s_tp != "Sem Supervisório":
                    if is_sch:
                        pr_as = st.session_state.precos_banco.get("Schneider - Servidor de Automação (SpaceLogic AS-P/AS-B)", 9500.0)
                        ctrl_desc.append("1x AS-P/AS-B")
                        linhas_hw.append({"Categoria": "Hardware", "Item": f"AS-P/AS-B ({p['nome']})", "Preço Unit.": pr_as, "Qtd": 1, "Custo Total": pr_as}); c_sch += pr_as
                    k_s = (s_tp, t_cfr_p)
                    if k_s not in sw_inc: sw_inc[k_s] = 0
                    sw_inc[k_s] += (r_ai+r_ao+r_di+r_do)

            i_desc = f"com IHM {p['ihm'].replace('Mercato - ', '')}" if "Cego" not in p['ihm'] else "sem interface IHM"
            s_desc = "Stand-alone" if "Sem" in p['supervisorio'] else "integrado ao supervisório EBO" if "EBO" in p['supervisorio'] else "integrado ao supervisório"
            t_filtro = "monitoramento da saturação dos filtros" if tem_f_pdt else "alarme de saturação de filtros" if tem_f_psh else ""
            t_res = "controle da resistência elétrica" if tem_res else ""
            t_extra = ", incluindo " + " e ".join([t for t in [t_filtro, t_res] if t]) if t_filtro or t_res else ""
            
            txt_p = f"Sistema para {', '.join(l_eq_nm)}{t_extra}.\nQuadro [TAG: {p['nome']}] {i_desc}, em {arq_at.replace(' - Linha mais econômica', '')} ({', '.join(ctrl_desc)}), {s_desc}. Visualização e controle de:\n• Status de operação.\n"
            if tem_f_pdt or tem_f_psh: txt_p += "• Monitoramento de saturação de filtros.\n"
            if tem_res: txt_p += "• Acionamento da resistência.\n"
            txt_p += f"• Leitura de instrumentos ({', '.join(list(l_inst_nm))}).\n"
            if cal_ativa: txt_p += "\nInstrumentos analógicos serão calibrados aferidos."
            if t_cfr_p == 'CFR21 Part 11 - Qualificável': txt_p += "\nFornecimento Qualificável (CFR 21 Part 11)."
            elif t_cfr_p == 'CFR21 Part 11 - Qualificado': txt_p += "\nFornecimento 100% Qualificado (CFR 21 Part 11) via SIARCON."
            
            desc_linhas.append(txt_p)
            desc_comercial.append(txt_p.replace("\n", "\n\n"))

        desc_final = "\n\n----------------------------------------------------\n\n".join(desc_linhas)

        for (sn, tc), pts in sw_inc.items():
            b_k = "Licença Supervisório - SEM CFR-21 (Base)" if "SEM" in sn else "Licença Supervisório - COM CFR-21 (Base)" if "COM" in sn else "Licença Supervisório - Schneider EBO (Base)"
            p_k = "Licença Supervisório - SEM CFR-21 (Por Ponto I/O)" if "SEM" in sn else "Licença Supervisório - COM CFR-21 (Por Ponto I/O)" if "COM" in sn else "Licença Supervisório - Schneider EBO (Por Ponto I/O)"
            pb = st.session_state.precos_banco.get(b_k, 23000.0); pp = st.session_state.precos_banco.get(p_k, 100.0)
            linhas_sw.append({"Categoria": "Software", "Item": f"Base: {sn}", "Preço Unit.": pb, "Qtd": 1, "Custo Total": pb})
            if pp>0 and pts>0: linhas_sw.append({"Categoria": "Software", "Item": f"Pontos Licenciados", "Preço Unit.": pp, "Qtd": pts, "Custo Total": pts*pp})
            
            if "COM" in sn:
                cu = 0.0
                if tc == "CFR21 Part 11 - Qualificável": cu = st.session_state.precos_banco.get("CFR21 Qualificável - Até 100 pts", 70.0) if pts<=100 else st.session_state.precos_banco.get("CFR21 Qualificável - 101 a 250 pts", 50.0) if pts<=250 else st.session_state.precos_banco.get("CFR21 Qualificável - Acima de 250 pts", 30.0)
                else: cu = st.session_state.precos_banco.get("CFR21 Qualificado - Até 30 pts", 400.0) if pts<=30 else st.session_state.precos_banco.get("CFR21 Qualificado - 31 a 60 pts", 350.0) if pts<=60 else st.session_state.precos_banco.get("CFR21 Qualificado - Acima de 250 pts", 200.0)
                linhas_srv.append({"Categoria": "Serviços", "Item": f"CFR21 ({tc}) - Por Ponto", "Preço Unit.": cu, "Qtd": pts, "Custo Total": pts*cu})

        for av in st.session_state.orcamento:
            linhas_inst.append({"Categoria": "Instrumentação", "Item": av['Item'], "Preço Unit.": av['Custo_Total']/av['Quantidade'] if av['Quantidade']>0 else 0, "Qtd": av['Quantidade'], "Custo Total": av['Custo_Total']})

        df_i, df_h, df_s = pd.DataFrame(linhas_inst), pd.DataFrame(linhas_hw), pd.DataFrame(linhas_sw)
        s_i = df_i['Custo Total'].sum() if not df_i.empty else 0
        s_h = df_h['Custo Total'].sum() if not df_h.empty else 0
        s_s = df_s['Custo Total'].sum() if not df_s.empty else 0
        
        if (t_ai_sch+t_ao_sch)>0: linhas_srv.append({"Categoria": "Serviços", "Item": "Lógica (Sch) - AI/AO", "Preço Unit.": st.session_state.precos_banco.get("Custo AI/AO", 565.0), "Qtd": t_ai_sch+t_ao_sch, "Custo Total": (t_ai_sch+t_ao_sch)*st.session_state.precos_banco.get("Custo AI/AO", 565.0)})
        if (t_di_sch+t_do_sch)>0: linhas_srv.append({"Categoria": "Serviços", "Item": "Lógica (Sch) - DI/DO", "Preço Unit.": st.session_state.precos_banco.get("Custo DI/DO", 120.0), "Qtd": t_di_sch+t_do_sch, "Custo Total": (t_di_sch+t_do_sch)*st.session_state.precos_banco.get("Custo DI/DO", 120.0)})
        if (t_ai_siem+t_ao_siem)>0: linhas_srv.append({"Categoria": "Serviços", "Item": "Lógica (Siem) - AI/AO", "Preço Unit.": st.session_state.precos_banco.get("Siemens - Serviço Custo AI/AO", 750.0), "Qtd": t_ai_siem+t_ao_siem, "Custo Total": (t_ai_siem+t_ao_siem)*st.session_state.precos_banco.get("Siemens - Serviço Custo AI/AO", 750.0)})
        if (t_di_siem+t_do_siem)>0: linhas_srv.append({"Categoria": "Serviços", "Item": "Lógica (Siem) - DI/DO", "Preço Unit.": st.session_state.precos_banco.get("Siemens - Serviço Custo DI/DO", 180.0), "Qtd": t_di_siem+t_do_siem, "Custo Total": (t_di_siem+t_do_siem)*st.session_state.precos_banco.get("Siemens - Serviço Custo DI/DO", 180.0)})
        if t_io_merc>0: linhas_srv.append({"Categoria": "Serviços", "Item": "Parametrização (Mercato)", "Preço Unit.": st.session_state.precos_banco.get("Mercato - Serviço Parametrização por Ponto", 80.0), "Qtd": t_io_merc, "Custo Total": t_io_merc*st.session_state.precos_banco.get("Mercato - Serviço Parametrização por Ponto", 80.0)})

        mo_ex = (c_sch*0.25) + (c_siem*0.35) + (c_merc*0.10)
        if mo_ex>0: linhas_srv.append({"Categoria": "Serviços", "Item": "Programações e Desenv. (25 a 35% HW)", "Preço Unit.": mo_ex, "Qtd": 1, "Custo Total": mo_ex})
        
        df_sr = pd.DataFrame(linhas_srv)
        s_sr = df_sr['Custo Total'].sum() if not df_sr.empty else 0
        t_geral = s_i + s_h + s_s + s_sr
        
        if t_geral > 0:
            c1, c2, c3 = st.columns(3)
            c1.info(f"**Instrumentação:**\nR$ {s_i:,.2f}"); c2.warning(f"**Hardware:**\nR$ {s_h:,.2f}"); c3.success(f"**TOTAL:**\nR$ {t_geral:,.2f}")

            expl = []
            if not df_i.empty: expl.append(pd.DataFrame([{"Categoria": "", "Item": "INSTRUMENTAÇÃO", "Preço Unit.": "-", "Qtd": "-", "Custo Total": "-"}])), expl.append(df_i.groupby(['Categoria', 'Item'], as_index=False).agg({'Preço Unit.': 'first', 'Qtd': 'sum', 'Custo Total': 'sum'})), expl.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "INSTRUMENTAÇÃO", "Preço Unit.": "-", "Qtd": "-", "Custo Total": s_i}]))
            if not df_h.empty: expl.append(pd.DataFrame([{"Categoria": "", "Item": "HARDWARE", "Preço Unit.": "-", "Qtd": "-", "Custo Total": "-"}])), expl.append(df_h.groupby(['Categoria', 'Item'], as_index=False).agg({'Preço Unit.': 'first', 'Qtd': 'sum', 'Custo Total': 'sum'})), expl.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "HARDWARE", "Preço Unit.": "-", "Qtd": "-", "Custo Total": s_h}]))
            if not df_s.empty: expl.append(pd.DataFrame([{"Categoria": "", "Item": "SOFTWARE", "Preço Unit.": "-", "Qtd": "-", "Custo Total": "-"}])), expl.append(df_s.groupby(['Categoria', 'Item'], as_index=False).agg({'Preço Unit.': 'first', 'Qtd': 'sum', 'Custo Total': 'sum'})), expl.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "SOFTWARE", "Preço Unit.": "-", "Qtd": "-", "Custo Total": s_s}]))
            if not df_sr.empty: expl.append(pd.DataFrame([{"Categoria": "", "Item": "SERVIÇOS", "Preço Unit.": "-", "Qtd": "-", "Custo Total": "-"}])), expl.append(df_sr), expl.append(pd.DataFrame([{"Categoria": "SUBTOTAL", "Item": "SERVIÇOS", "Preço Unit.": "-", "Qtd": "-", "Custo Total": s_sr}]))
            expl.append(pd.DataFrame([{"Categoria": "TOTAL GERAL", "Item": "ORÇAMENTO COMPLETO", "Preço Unit.": "-", "Qtd": "-", "Custo Total": t_geral}]))
            df_exp = pd.concat(expl, ignore_index=True)
            
            df_disp = df_exp.copy()
            df_disp['Preço Unit.'] = df_disp['Preço Unit.'].apply(lambda x: f"R$ {float(x):,.2f}".replace(",","X").replace(".",",").replace("X",".") if x!="-" else "-")
            df_disp['Custo Total'] = df_disp['Custo Total'].apply(lambda x: f"R$ {float(x):,.2f}".replace(",","X").replace(".",",").replace("X",".") if x!="-" else "-")
            st.dataframe(df_disp, use_container_width=True)
            
            with st.expander("📄 Gerar Descritivo Detalhado"):
                st.markdown("<div style='background-color:#E3F2FD; padding:20px; border-radius:10px;'>", unsafe_allow_html=True)
                for t_c in desc_comercial: st.markdown(t_c); st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
            
            df_p = pd.DataFrame(linhas_pt)
            if not df_p.empty: df_p = pd.concat([df_p, pd.DataFrame([{"Painel": "TOTAL", "Grupo/Equipamento": "-", "Instrumento": "-", "Quantidade Total": pd.to_numeric(df_p['Quantidade Total'], errors='coerce').fillna(0).sum(), "DI": df_p['DI'].sum(), "DO": df_p['DO'].sum(), "AI": df_p['AI'].sum(), "AO": df_p['AO'].sum()}])], ignore_index=True)

            wb = openpyxl.Workbook(); ws1 = wb.active; ws1.title = "Financeiro"
            ws1.row_dimensions[1].height = 35; ws1.row_dimensions[2].height = ws1.row_dimensions[3].height = ws1.row_dimensions[4].height = 25
            np_hdr = st.session_state.nome_projeto_orcamento or "PROJETO NÃO NOMEADO"
            ws1.merge_cells("C1:E1"); ws1.cell(1, 3, "DESCRIÇÃO DE SISTEMAS").font = Font(name="Arial", size=12, bold=True, color="1C8590"); ws1.cell(1,3).alignment = Alignment(horizontal="center", vertical="center")
            ws1.merge_cells("C2:E2"); ws1.cell(2, 3, f"PROJETO: {np_hdr.upper()}").font = Font(name="Arial", size=10, bold=True)
            ws1.merge_cells("C3:E3"); ws1.cell(3, 3, f"DATA: {datetime.now(fuso_br).strftime('%d/%m/%Y %H:%M')}").font = Font(name="Arial", size=10)
            ws1.merge_cells("C4:E4"); ws1.cell(4, 3, f"RESPONSÁVEL: {st.session_state.nome_exibicao.upper()}").font = Font(name="Arial", size=10)
            for r in range(2, 5): ws1.cell(r, 3).alignment = Alignment(horizontal="center", vertical="center")
            
            if ARQUIVO_LOGO:
                try:
                    from openpyxl.drawing.image import Image as OpxImg
                    img = OpxImg(ARQUIVO_LOGO); img.width = 180; img.height = 50; ws1.add_image(img, "A1")
                except: pass
                
            fl_h = PatternFill(start_color="1C8590", end_color="1C8590", fill_type="solid"); ft_h = Font(name="Arial", size=11, bold=True, color="FFFFFF")
            bd_t = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
            
            sr = 6
            for ri, row in enumerate(dataframe_to_rows(df_exp, index=False, header=True), sr):
                for ci, val in enumerate(row, 1):
                    c = ws1.cell(ri, ci, val); c.border = bd_t
                if ri == sr:
                    for c_idx in range(1, 6): ws1.cell(ri, c_idx).fill = fl_h; ws1.cell(ri, c_idx).font = ft_h; ws1.cell(ri, c_idx).alignment = Alignment(horizontal="center", vertical="center")
                else:
                    is_sub = "SUBTOTAL" in str(ws1.cell(ri, 1).value) or "TOTAL" in str(ws1.cell(ri, 1).value)
                    is_tit = str(ws1.cell(ri, 1).value) == ""
                    for c_idx in range(1, 6):
                        cc = ws1.cell(ri, c_idx); cc.font = Font(name="Arial", size=10, bold=(is_sub or is_tit))
                        if is_tit:
                            cc.fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
                            if c_idx == 2: cc.alignment = Alignment(horizontal="center", vertical="center")
                            elif c_idx > 2: cc.value = ""
                        elif is_sub:
                            cc.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                            if c_idx in [3, 5] and str(cc.value).strip() != "-":
                                try: cc.value = float(cc.value); cc.number_format = '"R$" #,##0.00'
                                except: pass
                            if c_idx in [3, 5]: cc.alignment = Alignment(horizontal="right")
                        else:
                            if c_idx in [3, 5] and str(cc.value).strip() != "-":
                                try: cc.value = float(cc.value); cc.number_format = '"R$" #,##0.00'
                                except: pass
                            if c_idx in [3, 5]: cc.alignment = Alignment(horizontal="right")
                            elif c_idx == 4: cc.alignment = Alignment(horizontal="center")
                    if is_tit: ws1.merge_cells(start_row=ri, start_column=2, end_row=ri, end_column=5)
            
            er = sr + len(df_exp) + 2
            tcx = max(10, len(desc_final.split('\n')) + 2) 
            ws1.merge_cells(start_row=er, start_column=1, end_row=er+tcx, end_column=5)
            cd = ws1.cell(er, 1, desc_final); cd.font = Font(name="Arial", size=10); cd.alignment = Alignment(vertical="top", wrap_text=True); cd.fill = PatternFill(start_color="F2F4F4", end_color="F2F4F4", fill_type="solid")
            for r in range(er, er+tcx+1):
                for c in range(1, 6): ws1.cell(r, c).border = bd_t
            for col in ws1.columns: ws1.column_dimensions[get_column_letter(col[0].column)].width = max(max(len(str(c.value or '')) for c in col if c.row <= sr + len(df_exp)) + 4, 12)

            if not df_p.empty:
                ws2 = wb.create_sheet("Matriz de IO")
                for ri, row in enumerate(dataframe_to_rows(df_p, index=False, header=True), 1):
                    for ci, val in enumerate(row, 1):
                        c = ws2.cell(ri, ci, val); c.border = bd_t
                        if ri == 1: c.fill = fl_h; c.font = ft_h; c.alignment = Alignment(horizontal="center", vertical="center")
                        else: c.font = Font(name="Arial", size=10); c.alignment = Alignment(horizontal="center") if ci >= 4 else Alignment(horizontal="left")
                for col in ws2.columns: ws2.column_dimensions[get_column_letter(col[0].column)].width = max(max(len(str(c.value or '')) for c in col) + 4, 12)

            ws3 = wb.create_sheet("Lista para Cotação")
            ws3.cell(1, 1, "LISTA DE MATERIAIS PARA COMPRAS").font = Font(name="Arial", size=12, bold=True, color="1C8590")
            for ci, ht in enumerate(["Fabricante", "Item / Modelo", "Qtd", "Unidade"], 1):
                c = ws3.cell(4, ci, ht); c.fill = fl_h; c.font = ft_h; c.alignment = Alignment(horizontal="center"); c.border = bd_t
            
            ir = 5; todos_it = [r['Item'] for _, r in df_h.iterrows()] + [r['Item'] for _, r in df_i.iterrows()]
            marcas = {"SIEMENS": [], "SCHNEIDER": [], "MERCATO": [], "OUTROS": []}
            for it in set(todos_it):
                cnt = todos_it.count(it); iu = it.upper()
                if "SIEMENS" in iu: marcas["SIEMENS"].append((it, cnt, "un"))
                elif "SCHNEIDER" in iu or "MP-C" in iu or "SPACELOGIC" in iu: marcas["SCHNEIDER"].append((it, cnt, "un"))
                elif "MERCATO" in iu or "MCP-" in iu or "MFC" in iu or "MDX" in iu: marcas["MERCATO"].append((it, cnt, "un"))
                else: marcas["OUTROS"].append((it, cnt, "un"))
            
            for mn, il in marcas.items():
                if il:
                    for ni, qi, ui in il:
                        ws3.cell(ir, 1, mn).alignment = Alignment(horizontal="center"); ws3.cell(ir, 2, ni); ws3.cell(ir, 3, qi).alignment = Alignment(horizontal="center"); ws3.cell(ir, 4, ui).alignment = Alignment(horizontal="center")
                        for cc in range(1, 5): ws3.cell(ir, cc).border = bd_t; ws3.cell(ir, cc).font = Font(name="Arial", size=10)
                        ir += 1
            for col in ws3.columns: ws3.column_dimensions[get_column_letter(col[0].column)].width = max(max(len(str(c.value or '')) for c in col) + 4, 12)
            
            buf = io.BytesIO(); wb.save(buf); buf.seek(0)
            st.download_button("📥 Exportar Excel", buf.getvalue(), "orcamento.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.markdown("---")
            if st.button("☁️ Salvar e Gerar Revisão", type="primary", use_container_width=True):
                if not st.session_state.nome_projeto_orcamento: st.warning("Preencha o Nome do Projeto.")
                else:
                    try:
                        sh = conectar_google_sheets()
                        try: ws_h = sh.worksheet("Historico_Orcamentos")
                        except: ws_h = sh.add_worksheet("Historico_Orcamentos", 1000, 8); ws_h.append_row(["Data/Hora", "Nome do Projeto", "Revisão", "Subtotal Hardware", "Serviços de Lógica", "Custo Total Estimado", "Configuracao_JSON", "Usuário"])
                        rev = f"R-{sum(1 for r in ws_h.get_all_values()[1:] if r[1].strip().upper() == st.session_state.nome_projeto_orcamento.strip().upper()):02d}"
                        ws_h.append_row([datetime.now(fuso_br).strftime("%d/%m/%Y %H:%M:%S"), st.session_state.nome_projeto_orcamento, rev, f"R$ {s_h:.2f}".replace('.', ','), f"R$ {s_sr:.2f}".replace('.', ','), f"R$ {t_geral:.2f}".replace('.', ','), json.dumps(st.session_state.paineis_auto), st.session_state.usuario_logado])
                        st.cache_data.clear(); st.success(f"✅ Salvo como {rev}!")
                    except Exception as e: st.error(f"Erro: {e}")
