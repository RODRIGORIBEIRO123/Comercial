import streamlit as st
import pandas as pd
import io
import json
import math
import re
import uuid
from datetime import datetime, timezone, timedelta
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

# ... (código anterior de Login e Menu ignorado para brevidade) ...
# [Inserir aqui o código de Login e Menu que você já tem]

# === MÓDULO 2: LEVANTAMENTO DE AUTOMAÇÃO (Aba de Preços) ===
# ... (código das abas anteriores) ...

# === ABA PREÇOS (Onde está o botão de salvar) ===
    with aba_precos:
        st.header("Gestão da Base de Preços")
        st.info(f"📅 **Última atualização da tabela sincronizada com o banco de dados da nuvem:** {st.session_state.data_precos_atualizada}")
        
        # ... (Upload e Download de planilhas) ...

        if st.button("💾 Salvar Novos Preços no Banco de Dados", type="primary"):
            alterou_algo = False
            novos_historicos = []
            
            # (assumindo que 'edited_geral', etc., estão definidos acima conforme seu código)
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
                    
                    # --- A CORREÇÃO ESTÁ AQUI ---
                    st.session_state.data_precos_atualizada = data_hora_agora
                    st.cache_data.clear()
                    st.toast("✅ Base de preços atualizada com sucesso!", icon="💾")
                    st.rerun() # Isso força a página a atualizar e exibir a nova data
                except Exception as e: 
                    st.error(f"Erro ao salvar: {e}")
            else:
                st.info("Nenhuma alteração detectada para salvar.")
