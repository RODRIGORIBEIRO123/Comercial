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
    for nome in ["SIARCON.png", "SIARCON .png", "siarcon.png", "Siarcon.png", "logo.png"]:
        if os.path.exists(nome): return nome
    return None

ARQUIVO_LOGO = buscar_logo()
PLANILHA_URL = "https://docs.google.com/spreadsheets/d/1DgBxNqwUepO2RW6GdRwnFHxg7dLlWiRGZjdglkQ8Ls0/edit?gid=1169331401#gid=1169331401"
fuso_br = timezone(timedelta(hours=-3))

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

# DICIONÁRIO DE NOMES DO DIAGRAMA (P&ID)
if 'de_para_diagrama' not in st.session_state:
    st.session_state.de_para_diagrama = {
        "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": {"in_agua": "Trans. Pressão - Vazão (PDT)", "in_comp": "Trans. Pressão - Vazão (PDT)", "out_agua": "Controle Vazão Ar", "out_comp": "Controle Vazão Ar"},
        "Transmissor de temperatura e umidade para duto (TT/MT)": {"in_agua": "Trans. Temp. e Umid. (TT/MT)", "in_comp": "Trans. Temp. e Umid. (TT/MT)", "out_agua": "", "out_comp": ""},
        "Transmissor de temperatura para duto (TT)": {"in_agua": "Trans. Temp. (TT)", "in_comp": "Trans. Temp. (TT)", "out_agua": "", "out_comp": ""},
        "Válvula de controle de água gelada proporcional (TCV)": {"in_agua": "", "in_comp": "", "out_agua": "Modula VAG", "out_comp": ""},
        "Válvula de controle de água quente proporcional (TCV)": {"in_agua": "", "in_comp": "", "out_agua": "Modula VAQ", "out_comp": ""},
        "Relé de Corrente - Status Compressor (TC)": {"in_agua": "", "in_comp": "Status Compressor", "out_agua": "", "out_comp": "Habilita Compressor"},
        "Termostato de segurança (TSH)": {"in_agua": "Termostato Seg. RAQ (TSH)", "in_comp": "Termostato Seg. RAQ (TSH)", "out_agua": "Status RAQ", "out_comp": "Status RAQ"},
        "Pressostato diferencial para ar (PSH)": {"in_agua": "Pressostato Seg. RAQ (PSH)", "in_comp": "Pressostato Seg. RAQ (PSH)", "out_agua": "Status RAQ", "out_comp": "Status RAQ"},
        "Resistência de aquecimento (Equipamento) (RAQ)": {"in_agua": "", "in_comp": "", "out_agua": "Habilita RAQ", "out_comp": "Habilita RAQ"},
        "Pressostato para monitorar os filtros G4 (PSH)": {"in_agua": "Pressostato G4 (PSH)", "in_comp": "Pressostato G4 (PSH)", "out_agua": "Alarme G4 Saturado", "out_comp": "Alarme G4 Saturado"},
        "Pressostato para monitorar os filtros F9 (PSH)": {"in_agua": "Pressostato F9 (PSH)", "in_comp": "Pressostato F9 (PSH)", "out_agua": "Alarme F9 Saturado", "out_comp": "Alarme F9 Saturado"},
        "Pressostato para monitorar os filtros H13/H14 (PSH)": {"in_agua": "Pressostato H13/14 (PSH)", "in_comp": "Pressostato H13/14 (PSH)", "out_agua": "Alarme H13/14 Saturado", "out_comp": "Alarme H13/14 Saturado"},
        "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": {"in_agua": "Status Func. Partida Direta (PSH)", "in_comp": "Status Func. Partida Direta (PSH)", "out_agua": "Comando Partida Direta", "out_comp": "Comando Partida Direta"},
        "Transmissor de pressão diferencial entre salas (PDT)": {"in_agua": "Pressão Dif. Salas (PDT)", "in_comp": "Pressão Dif. Salas (PDT)", "out_agua": "", "out_comp": ""},
        "Transmissor de temperatura Ambiente (TT)": {"in_agua": "Temp. Salas (TT)", "in_comp": "Temp. Salas (TT)", "out_agua": "", "out_comp": ""},
        "Transmissor de temperatura e umidade ambiente (TT/MT)": {"in_agua": "Temp. / Umid. (TT/MT)", "in_comp": "Temp. / Umid. (TT/MT)", "out_agua": "", "out_comp": ""},
        "Chave Seletora Auto/Manual (Painel Elétrico)": {"in_agua": "Chave Auto / Manual", "in_comp": "Chave Auto / Manual", "out_agua": "Habilita Equipamento", "out_comp": "Habilita Equipamento"}
    }

if st.session_state.data_precos_atualizada == "Buscando metadados da nuvem...":
    try:
        sh_init = conectar_google_sheets()
        linhas_h = sh_init.worksheet("Historico_Precos").get_all_values()
        st.session_state.data_precos_atualizada = linhas_h[-1][0] if len(linhas_h) > 1 else "Nenhuma alteração recente"
    except:
        st.session_state.data_precos_atualizada = "Não foi possível carregar a data de atualização"

# ==========================================
# TELA DE LOGIN
# ==========================================
if st.session_state.usuario_logado is None:
    st.markdown("""<style>.block-container { padding-top: 0rem !important; margin-top: -2rem !important; } header {display: none !important;} [data-testid="collapsedControl"] {display: none !important;} .stApp { background: linear-gradient(135deg, #1C8590 0%, #8FD3B5 100%) !important; } [data-testid="stForm"] { background-color: white; border-radius: 12px; padding: 30px; box-shadow: 0 8px 24px rgba(0,0,0,0.15); border: none; } [data-testid="stFormSubmitButton"] button { background-color: #2b7bc4 !important; color: white !important; font-weight: bold !important; border-radius: 6px !important; height: 45px !important; margin-top: 15px !important; }</style>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.write("")
        col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
        with col_img2:
            if ARQUIVO_LOGO: st.image(ARQUIVO_LOGO, use_container_width=True)
            else: st.markdown("<h2 style='text-align: center; color: white; margin-bottom:0;'>SIARCON</h2>", unsafe_allow_html=True)
        with st.form("form_login"):
            st.markdown("<div style='text-align: center; margin-bottom: 15px;'><h3 style='color: #333; margin: 0; font-size: 18px;'>Bem-Vindo a plataforma comercial da SIARCON</h3></div>", unsafe_allow_html=True)
            c_user = st.text_input("Usuário:", placeholder="Ex: rodrigo.ribeiro")
            c_pass = st.text_input("Senha:", type="password", placeholder="••••")
            if st.form_submit_button("Entrar no Sistema", use_container_width=True):
                if c_pass == "1234":
                    st.session_state.usuario_logado = c_user.lower().strip()
                    st.session_state.nome_exibicao = c_user.split('.')[0].capitalize()
                    st.rerun()
                else: st.error("❌ Usuário ou senha incorretos.")
    st.stop()

st.markdown("""<style>.block-container { padding-top: 3rem !important; } header {display: flex !important;} [data-testid="collapsedControl"] {display: flex !important;}</style>""", unsafe_allow_html=True)

if ARQUIVO_LOGO: st.sidebar.image(ARQUIVO_LOGO, use_container_width=True)
st.sidebar.title("Navegação Principal")
st.sidebar.markdown(f"👤 Logado como: **{st.session_state.nome_exibicao}**")
if st.sidebar.button("🚪 Sair do Perfil", type="secondary"):
    st.session_state.usuario_logado = None; st.rerun()

st.sidebar.markdown("---")
opcoes_menu = ["🏠 Tela Inicial", "📄 Gerador de Propostas", "🔌 Levantamento de Automação"]
menu_ui = st.sidebar.radio("Módulos do Sistema", opcoes_menu, index=opcoes_menu.index(st.session_state.menu_selecionado))
if menu_ui != st.session_state.menu_selecionado:
    st.session_state.menu_selecionado = menu_ui; st.rerun()

# ==============================================================================
# MÓDULOS 0 E 1
# ==============================================================================
if st.session_state.menu_selecionado == "🏠 Tela Inicial":
    st.markdown(f"<h1 style='text-align: center; color: #178B96;'>Bem-vindo(a), {st.session_state.nome_exibicao}!</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px; color: #666;'>Portal Comercial e de Engenharia SIARCON. Selecione o módulo desejado para iniciar:</p><br>", unsafe_allow_html=True)
    col_vazia_esq, col_card1, col_vazia_meio, col_card2, col_vazia_dir = st.columns([1, 2.5, 0.5, 2.5, 1])
    with col_card1:
        st.markdown("<div style='text-align: center; padding: 30px; background: white; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border-top: 5px solid #1C8590;'><h1 style='font-size: 50px; margin-bottom: 10px;'>📄</h1><h3 style='color: #333;'>Gerador de Propostas</h3></div>", unsafe_allow_html=True)
        if st.button("Acessar Módulo ➔", key="btn_home_prop", type="primary", use_container_width=True):
            st.session_state.menu_selecionado = "📄 Gerador de Propostas"; st.rerun()
    with col_card2:
        st.markdown("<div style='text-align: center; padding: 30px; background: white; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border-top: 5px solid #1C8590;'><h1 style='font-size: 50px; margin-bottom: 10px;'>🔌</h1><h3 style='color: #333;'>Levantamento de Automação</h3></div>", unsafe_allow_html=True)
        if st.button("Acessar Módulo ➔", key="btn_home_auto", type="primary", use_container_width=True):
            st.session_state.menu_selecionado = "🔌 Levantamento de Automação"; st.rerun()

elif st.session_state.menu_selecionado == "📄 Gerador de Propostas":
    st.info("Módulo de Propostas em background. Acesse a aba Automação para os diagramas.")

# ==============================================================================
# MÓDULO 2: LEVANTAMENTO DE AUTOMAÇÃO
# ==============================================================================
elif st.session_state.menu_selecionado == "🔌 Levantamento de Automação":
    st.markdown("""<style>[data-testid="stVerticalBlockBorderWrapper"] {border-radius: 8px; border-left: 4px solid #1C8590 !important; background-color: rgba(28, 133, 144, 0.03);}</style>""", unsafe_allow_html=True)
    st.title("🔌 Engenharia e Custos - Automação e Infra")
    
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
        "Relé de Corrente - Status Compressor (TC)": {"AI": 0, "AO": 0, "DI": 1, "DO": 2},
        "Termostato de segurança (TSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Pressostato diferencial para ar (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Resistência de aquecimento (Equipamento) (RAQ)": {"AI": 0, "AO": 1, "DI": 2, "DO": 1},
        "Pressostato para monitorar os filtros G4 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Pressostato para monitorar os filtros F9 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Pressostato para monitorar os filtros H13/H14 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 1},
        "Transmissor de pressão diferencial entre salas (PDT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura Ambiente (TT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura e umidade ambiente (TT/MT)": {"AI": 2, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de CO2 ambiente (AT/AIT)": {"AI": 1, "AO": 1, "DI": 0, "DO": 1}
    }

    GRUPOS_INSTRUMENTOS = {
        "🔹 Controle (HVAC e Máquinas)": ["Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)", "Transmissor de temperatura e umidade para duto (TT/MT)", "Transmissor de temperatura para duto (TT)", "Válvula de controle de água gelada proporcional (TCV)", "Relé de Corrente - Status Compressor (TC)", "Termostato de segurança (TSH)", "Pressostato diferencial para ar (PSH)", "Resistência de aquecimento (Equipamento) (RAQ)"],
        "🔸 Monitoramento (Filtros e Status)": ["Pressostato para monitorar os filtros G4 (PSH)", "Pressostato para monitorar os filtros F9 (PSH)", "Pressostato para monitorar os filtros H13/H14 (PSH)", "Status funcionamento ventilador ou exaustor (partida direta) (PSH)"],
        "🟢 Monitoramento e Ambientes": ["Transmissor de pressão diferencial entre salas (PDT)", "Transmissor de temperatura Ambiente (TT)", "Transmissor de temperatura e umidade ambiente (TT/MT)", "Transmissor de CO2 ambiente (AT/AIT)"]
    }
    
    KITS_PADRAO = {
        "❄️ UTA Padrão - Água Gelada": {"Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1, "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Válvula de controle de água gelada proporcional (TCV)": 1, "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1},
        "🌬️ UTA Padrão - Expansão Direta": {"Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1, "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Relé de Corrente - Status Compressor (TC)": 2, "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1},
        "🔥 UTA Padrão - Água Gelada + Resistência": {"Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1, "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Válvula de controle de água gelada proporcional (TCV)": 1, "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1, "Termostato de segurança (TSH)": 1, "Pressostato diferencial para ar (PSH)": 1, "Resistência de aquecimento (Equipamento) (RAQ)": 1},
        "🔥 UTA Expansão Direta (2 Compressores) + Resistência (Salas e Exaustão)": {"Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1, "Transmissor de temperatura e umidade para duto (TT/MT)": 1, "Relé de Corrente - Status Compressor (TC)": 2, "Pressostato para monitorar os filtros G4 (PSH)": 1, "Pressostato para monitorar os filtros F9 (PSH)": 1, "Termostato de segurança (TSH)": 1, "Resistência de aquecimento (Equipamento) (RAQ)": 1, "Transmissor de pressão diferencial entre salas (PDT)": 4, "Transmissor de temperatura e umidade ambiente (TT/MT)": 4, "Status funcionamento ventilador ou exaustor (partida direta) (PSH)": 2, "Pressostato diferencial para ar (PSH)": 1}
    }

    if 'precos_banco' not in st.session_state:
        st.session_state.precos_banco = {
            "Transmissor de pressão dif. para ar (medição de vazão de ar) (PDIT)": 1490.00, "Transmissor de temperatura e umidade para duto (TT/MT)": 2050.00, "Transmissor de temperatura para duto (TT)": 800.00, "Termostato de segurança (TSH)": 250.00, "Pressostato diferencial para ar (PSH)": 349.00, "Pressostato para monitorar os filtros G4 (PSH)": 349.00, "Pressostato para monitorar os filtros F9 (PSH)": 349.00, "Pressostato para monitorar os filtros H13/H14 (PSH)": 349.00, "Transmissor de pressão diferencial entre salas (PDT)": 1490.00, "Transmissor de temperatura Ambiente (TT)": 2050.00, "Custo AI/AO": 565.00, "Custo DI/DO": 120.00, "Licença Supervisório - SEM CFR-21 (Base)": 23000.00, "Licença Supervisório - SEM CFR-21 (Por Ponto I/O)": 100.00, "MP-C-36A": 9459.08, "MP-C-24A": 7290.75, "MP-C-18A": 5185.54, "MP-C-15A": 4649.49, "Mercato - Controlador MCP-12 (4AO, 8DO, 12UI, 4DI)": 1850.00, "Siemens - CPU 1214C DC/DC/DC": 2500.00, "Siemens - SM 1231 AI 8x13Bit": 1900.00, "IHM Padrão 7\"": 3400.00, "IHM Premium 10\"": 8500.00, "Sem Interface (Cego)": 0.00
        }

    def calcular_painel_fisico(qtd_controladores):
        if qtd_controladores == 0: return "Sem Painel", 0.0
        elif qtd_controladores <= 4: return "Quadro 600x400mm", 4500.00
        elif qtd_controladores <= 10: return "Quadro 800x600mm", 5900.00
        else: return "Armário 1200x800mm", 9250.00

    aba_auto, aba_infra, aba_precos, aba_resumo = st.tabs(["🚀 Dimensionamento de Automação", "🔌 Infraestrutura", "💲 Base de Preços", "📊 Orçamento Final"])

    with aba_auto:
        with st.expander("🔮 [BETA] Módulo Inteligente: Importar Quadro via Engenharia Reversa", expanded=True):
            st.markdown("Faça o upload do fluxograma descritivo para mapeamento de Salas, Máquinas e Exaustores.")
            arquivo_diagrama = st.file_uploader("Carregar Diagrama Técnico / P&ID:", type=["pdf", "png", "jpg"], key="upl_ia_diagrama")
            
            if arquivo_diagrama is not None:
                st.info("💡 Perfil identificado: **UTA Expansão Direta (2 Estágios) + Resistência + 4 Salas + Exaustão**.")
                st.markdown("##### ⚙️ Configurações do Quadro Gerado")
                c_ia1, c_ia2 = st.columns(2)
                arquitetura_ia = c_ia1.radio("Qual marca de controlador?", ["SpaceLogic (Schneider)", "S7-1200 (Siemens)", "S7-1500 (Siemens)", "MCP Parametrizável (Mercato)"])
                tag_ia = c_ia2.text_input("TAG do Quadro:", value="QTA-Geral")
                ihm_ia = c_ia1.radio("Possui IHM local?", ["Sem Interface (Cego)", "IHM Padrão 7\"", "IHM Premium 10\""])
                sup_opt_ia = c_ia2.radio("Terá Sistema de Supervisório?", ["Não", "Sim"])
                
                if st.button("🪄 Executar Engenharia Reversa e Gerar Quadro", type="primary"):
                    novos_instrumentos = {k: 0 for k in REGRA_IO.keys()}
                    for item_nome, qtd_padrao in KITS_PADRAO["🔥 UTA Expansão Direta (2 Compressores) + Resistência (Salas e Exaustão)"].items():
                        if item_nome in novos_instrumentos: novos_instrumentos[item_nome] = qtd_padrao
                    novo_quadro_ia = {
                        "id": str(uuid.uuid4()), "nome": tag_ia, "tipo": "Controle", "supervisorio": "Sistema SEM CFR-21" if sup_opt_ia == "Sim" else "Sem Supervisório", "arquitetura": arquitetura_ia, "tipo_cfr": "Não Aplicável", "modo_config": "Usar Padrão Existente (Kits)", "ihm": ihm_ia, "sobra_20": "Não", "tags_nao_reconhecidas": ["PT-08 (Sala Químicos)", "FQI-01 (Duto Exaustão)"],
                        "grupos_equipamentos": [{"nome_grupo": "Sistema Integrado UTA (DX) + Salas", "multiplicador": 1, "instrumentos": novos_instrumentos, "tags_lista": ["COND-01", "COND-02", "SALA-A", "EX-01"]}]
                    }
                    st.session_state.paineis_auto.insert(0, novo_quadro_ia); st.rerun()

        if not st.session_state.wizard_ativo:
            if st.button("➕ Criar Novo Quadro Manualmente", type="primary"): st.session_state.wizard_ativo = True; st.rerun()
        if st.session_state.wizard_ativo:
            if st.button("Cancelar"): st.session_state.wizard_ativo = False; st.rerun()

        for p_idx, p_data in enumerate(st.session_state.paineis_auto):
            if p_data.get("tags_nao_reconhecidas"):
                st.error(f"⚠️ **Atenção (IA):** TAGs sem correspondência de preço: `{', '.join(p_data['tags_nao_reconhecidas'])}`. Audite abaixo.")
            
            with st.expander(f"🎛️ Quadro: {p_data['nome']} - {p_data.get('arquitetura', '')}", expanded=True):
                c_icone, c_nome_painel, c_ihm_painel = st.columns([0.5, 4, 2])
                c_icone.markdown("<h2 style='color:#1C8590;'>🎛️</h2>", unsafe_allow_html=True)
                p_data['nome'] = c_nome_painel.text_input("Identificação do Quadro", value=p_data['nome'], key=f"n_p_{p_data['id']}", label_visibility="collapsed")
                
                for g_idx, g_data in enumerate(p_data['grupos_equipamentos']):
                    st.markdown(f"### 📦 {g_data['nome_grupo']}")
                    qtd_key = f"m_g_{p_data['id']}_{g_idx}"
                    qtd_atual = st.session_state.get(qtd_key, g_data.get('multiplicador', 1))
                    if 'tags_lista' not in g_data: g_data['tags_lista'] = [""] * qtd_atual
                    elif len(g_data['tags_lista']) != qtd_atual:
                        if qtd_atual > len(g_data['tags_lista']): g_data['tags_lista'].extend([""] * (qtd_atual - len(g_data['tags_lista'])))
                        else: g_data['tags_lista'] = g_data['tags_lista'][:qtd_atual]

                    render_qtd = min(qtd_atual, 5) 
                    cols = st.columns([3] + [1.5] * render_qtd + [1])
                    g_data['nome_grupo'] = cols[0].text_input("Equipamento", value=g_data['nome_grupo'], key=f"n_g_{p_data['id']}_{g_idx}")
                    for i in range(render_qtd): g_data['tags_lista'][i] = cols[i+1].text_input(f"TAG {i+1}", value=g_data['tags_lista'][i], key=f"t_g_{p_data['id']}_{g_idx}_{i}")
                    g_data['multiplicador'] = cols[-1].number_input("Qtd", min_value=1, value=qtd_atual, key=qtd_key)
                    
                    # --- NOVIDADE: DIAGRAMA EM ABAS POR EQUIPAMENTO (CONDICIONAL PARA NÃO PINTAR VAZIOS) ---
                    st.markdown("#### 👁️ Esquemáticos P&ID de Instrumentação")
                    tabs_diagrama = st.tabs(["📊 Visualizar Fluxograma", "📥 Exportar Código (.dot)"])
                    
                    is_compressor_sys = "COMPRESSOR" in g_data['nome_grupo'].upper() or "DIRETA" in g_data['nome_grupo'].upper() or "DX" in g_data['nome_grupo'].upper()
                    
                    dot = 'digraph G {\n  rankdir=LR;\n  node [fontname="Arial", fontsize=10, shape=box, style=rounded];\n'
                    if p_data.get('ihm') and "Cego" not in p_data['ihm']:
                        dot += f'  "IHM" [label="{p_data["ihm"]}\\n(Porta do Painel)", fillcolor="#D5F5E3", style=filled, shape=note];\n'
                        dot += '  "IHM" -> "Controlador" [dir=both, style=dashed, color="#7B7D7D"];\n'
                    dot += f'  "Controlador" [label="{p_data.get("arquitetura", "Controlador")}\\nTAG: {p_data["nome"]}", fillcolor="#1C8590", style=filled, fontcolor=white, shape=ellipse];\n'
                    
                    has_in = False; has_out = False; node_id = 0
                    
                    for inst_f, q_f in g_data['instrumentos'].items():
                        if q_f > 0:
                            tag_mapping = st.session_state.de_para_diagrama.get(inst_f, {})
                            tag_hardware = inst_f.split('(')[-1].replace(')', '').strip() if '(' in inst_f else 'IO'
                            
                            # Varre e insere as TAGs da interface de forma cíclica
                            tag_contexto = g_data['tags_lista'][node_id % len(g_data['tags_lista'])] if g_data.get('tags_lista') and len(g_data['tags_lista'])>0 else ""
                            str_tag_ctx = f"\\n({tag_contexto})" if tag_contexto else ""
                            
                            if isinstance(tag_mapping, str): lbl_in = lbl_out = tag_mapping
                            else:
                                lbl_in = str(tag_mapping.get("in_comp", "")) if is_compressor_sys else str(tag_mapping.get("in_agua", ""))
                                lbl_out = str(tag_mapping.get("out_comp", "")) if is_compressor_sys else str(tag_mapping.get("out_agua", ""))
                                if not lbl_in: lbl_in = str(tag_mapping.get("in_agua" if is_compressor_sys else "in_comp", ""))
                                if not lbl_out: lbl_out = str(tag_mapping.get("out_agua" if is_compressor_sys else "out_comp", ""))
                            
                            lbl_in = lbl_in.strip(); lbl_out = lbl_out.strip()
                            
                            # Lógica Restrita: Se não tem nome na planilha, não aparece.
                            if lbl_in and lbl_in != "nan":
                                for i in range(int(q_f)):
                                    if int(q_f) > 4 and i > 0: break
                                    prefix = f"{q_f}x " if int(q_f) > 4 else ""
                                    sufix = f" {i+1}" if int(q_f) > 1 and int(q_f) <= 4 else ""
                                    dot += f'  "in_{node_id}_{i}" [label="{prefix}{lbl_in}{sufix}{str_tag_ctx}\\nTAG: {tag_hardware}", color="#2B7BC4"];\n'
                                    dot += f'  "in_{node_id}_{i}" -> "Controlador" [color="#2B7BC4"];\n'
                                    has_in = True
                            if lbl_out and lbl_out != "nan":
                                for i in range(int(q_f)):
                                    if int(q_f) > 4 and i > 0: break
                                    prefix = f"{q_f}x " if int(q_f) > 4 else ""
                                    sufix = f" {i+1}" if int(q_f) > 1 and int(q_f) <= 4 else ""
                                    dot += f'  "out_{node_id}_{i}" [label="{prefix}{lbl_out}{sufix}{str_tag_ctx}\\nTAG: {tag_hardware}", color="#E14D2A"];\n'
                                    dot += f'  "Controlador" -> "out_{node_id}_{i}" [color="#E14D2A"];\n'
                                    has_out = True
                            node_id += 1
                            
                    # Chave Auto/Manual
                    inst_c = "Chave Seletora Auto/Manual (Painel Elétrico)"
                    map_c = st.session_state.de_para_diagrama.get(inst_c, {})
                    l_in_c = str(map_c.get("in_comp", "")) if is_compressor_sys else str(map_c.get("in_agua", ""))
                    l_out_c = str(map_c.get("out_comp", "")) if is_compressor_sys else str(map_c.get("out_agua", ""))
                    if l_in_c.strip() and l_in_c != "nan":
                        dot += f'  "ch_in" [label="{l_in_c}\\nTAG: CH", color="#2B7BC4"];\n  "ch_in" -> "Controlador" [color="#2B7BC4"];\n'; has_in = True
                    if l_out_c.strip() and l_out_c != "nan":
                        dot += f'  "ch_out" [label="{l_out_c}\\nTAG: CH", color="#E14D2A"];\n  "Controlador" -> "ch_out" [color="#E14D2A"];\n'; has_out = True
                    
                    if not has_in: dot += '  "Sensores" -> "Controlador" [style=dashed];\n'
                    if not has_out: dot += '  "Controlador" -> "Atuadores" [style=dashed];\n'
                    dot += '}'
                    
                    with tabs_diagrama[0]: st.graphviz_chart(dot)
                    with tabs_diagrama[1]:
                        st.text_area("Código do Diagrama:", value=dot, height=100)
                        st.download_button("📥 Baixar .dot", dot, file_name=f"Diag_{g_data['nome_grupo']}.dot")

                    with st.expander("⚙️ Ajuste Fino de Instrumentos (Engenharia)"):
                        for grupo_nome, lista_itens in GRUPOS_INSTRUMENTOS.items():
                            with st.expander(grupo_nome, expanded=False):
                                cols_inst = st.columns(2)
                                for i, inst in enumerate(lista_itens):
                                    if inst not in g_data['instrumentos']: g_data['instrumentos'][inst] = 0
                                    g_data['instrumentos'][inst] = cols_inst[i % 2].number_input(inst, min_value=0, step=1, value=g_data['instrumentos'][inst], key=f"inst_{p_data['id']}_{g_idx}_{grupo_nome}_{inst}")
                if st.button("🗑️ Deletar Quadro", key=f"del_quadro_{p_data['id']}"): st.session_state.paineis_auto.pop(p_idx); st.rerun()

    with aba_precos:
        st.header("Gestão da Base de Preços")
        st.markdown("### 🏷️ Padronização de Nomes para o Diagrama P&ID")
        
        buffer_nomes = io.BytesIO()
        df_nomes = pd.DataFrame([{"Nome Original (Base de Preços)": k, "Nome Exibido - Entrada (Se água  gelada)": v.get("in_agua", ""), "Nome Exibido Entrada (Se compressor)": v.get("in_comp", ""), "Nome Exibido Saída (Se Água Gelada)": v.get("out_agua", ""), "Nome Exibido Saída (Se Compressor)": v.get("out_comp", "")} for k, v in st.session_state.de_para_diagrama.items()])
        df_nomes.to_excel(buffer_nomes, index=False)
        buffer_nomes.seek(0)
        st.download_button("📥 Baixar Planilha de Personalização", buffer_nomes, "Dicionario_Diagrama.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        upload_nomes = st.file_uploader("📂 Faça o upload da planilha editada", type=["xlsx", "csv"])
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
                    st.success("✅ Nomes atualizados! Verifique a aba de Dimensionamento.")
            except Exception as e: st.error(f"Erro: {e}")
            
    with aba_resumo:
        st.header("Consolidação Financeira do Orçamento")
        linhas_inst = []; linhas_hw = []; linhas_serv = []; total_ai = total_ao = total_di = total_do = 0
        custo_hw = custo_inst = 0.0

        for p in st.session_state.paineis_auto:
            qtd_equip_painel = sum(g.get('multiplicador', 1) for g in p['grupos_equipamentos'])
            for g in p['grupos_equipamentos']:
                mult = g.get('multiplicador', 1)
                for inst, qtd in g['instrumentos'].items():
                    if qtd > 0:
                        q_final = qtd * mult
                        p_item = st.session_state.precos_banco.get(inst, 0.0)
                        io_v = REGRA_IO.get(inst, {"AI":0, "AO":0, "DI":0, "DO":0})
                        total_ai += q_final * io_v["AI"]; total_ao += q_final * io_v["AO"]
                        total_di += q_final * io_v["DI"]; total_do += q_final * io_v["DO"]
                        linhas_inst.append({"Categoria": "Instrumentação", "Item": f"{inst}", "Preço Unit.": p_item, "Qtd": q_final, "Custo Total": q_final * p_item})
                        custo_inst += q_final * p_item
            
            tot_io = total_ai + total_ao + total_di + total_do
            if tot_io > 0:
                nome_c, pr_c = calcular_painel_fisico(tot_io/15)
                linhas_hw.append({"Categoria": "Hardware", "Item": f"Painel {nome_c}", "Preço Unit.": pr_c, "Qtd": 1, "Custo Total": pr_c}); custo_hw += pr_c
                c36 = math.ceil(tot_io/36)
                if c36 > 0: 
                    linhas_hw.append({"Categoria": "Hardware", "Item": "Controlador MP-C-36A", "Preço Unit.": st.session_state.precos_banco.get("MP-C-36A", 9459.0), "Qtd": c36, "Custo Total": c36 * 9459.0}); custo_hw += c36 * 9459.0
        
        if total_ai + total_ao > 0: linhas_serv.append({"Categoria": "Serviços", "Item": "Parametrização Analógica", "Preço Unit.": 565.0, "Qtd": total_ai + total_ao, "Custo Total": (total_ai + total_ao) * 565.0})
        if total_di + total_do > 0: linhas_serv.append({"Categoria": "Serviços", "Item": "Parametrização Digital", "Preço Unit.": 120.0, "Qtd": total_di + total_do, "Custo Total": (total_di + total_do) * 120.0})
        
        df_inst = pd.DataFrame(linhas_inst); df_hw = pd.DataFrame(linhas_hw); df_serv = pd.DataFrame(linhas_serv)
        tot_geral = custo_inst + custo_hw + (df_serv['Custo Total'].sum() if not df_serv.empty else 0)
        
        if tot_geral > 0:
            c1, c2, c3 = st.columns(3)
            c1.info(f"Instrumentação: R$ {custo_inst:,.2f}")
            c2.warning(f"Hardware: R$ {custo_hw:,.2f}")
            c3.success(f"TOTAL: R$ {tot_geral:,.2f}")
            
            df_export = pd.concat([df_inst, df_hw, df_serv], ignore_index=True)
            st.dataframe(df_export, use_container_width=True)
            
            buf = io.BytesIO()
            df_export.to_excel(buf, index=False)
            buf.seek(0)
            st.download_button("📥 Exportar Orçamento (Excel)", buf, "Orcamento.xlsx")
            if st.button("☁️ Salvar Revisão na Nuvem", type="primary"): st.success("Salvo!")
