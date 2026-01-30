import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date

# ==============================================================================
# CONFIGURAÇÃO DE LINKS (SUBSTITUA AQUI PELOS SEUS LINKS DO GOOGLE SHEETS)
# ==============================================================================
# Cole o link CSV da aba "Escopos" dentro das aspas abaixo:
URL_ESCOPOS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9Dlv9q_qBgpCIwY6cQAfWTYY6JXO9ILRMN_NT_QNjFiWAy2N5W9QqjP51U2fAnE2mi-RCEtj5l2wG/pub?gid=221408068&single=true&output=csv"

# Cole o link CSV da aba "Exclusoes" dentro das aspas abaixo:
URL_EXCLUSOES = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9Dlv9q_qBgpCIwY6cQAfWTYY6JXO9ILRMN_NT_QNjFiWAy2N5W9QqjP51U2fAnE2mi-RCEtj5l2wG/pub?gid=1129521636&single=true&output=csv"
# ==============================================================================

# Configuração da página do site
st.set_page_config(page_title="Gerador de Propostas SIARCON", layout="wide", page_icon="📄")

st.title("📄 Gerador de Propostas - SIARCON")
st.markdown("---")

# --- FUNÇÃO PARA CARREGAR DADOS DO GOOGLE SHEETS ---
@st.cache_data(ttl=60) # Atualiza os dados a cada 60 segundos para não ficar lento
def carregar_dados():
    try:
        # Lê os dados direto dos links CSV
        df_esc = pd.read_csv(URL_ESCOPOS)
        df_exc = pd.read_csv(URL_EXCLUSOES)
        return df_esc, df_exc
    except Exception as e:
        return None, e

# Carrega os dados
df_escopos, df_exclusoes = carregar_dados()

# Verifica se deu erro ao carregar
if df_escopos is None:
    st.error(f"❌ Erro ao conectar com o Google Sheets!")
    st.warning("Verifique se os links 'URL_ESCOPOS' e 'URL_EXCLUSOES' no código estão corretos e publicados como CSV.")
    st.stop()

# --- FORMULÁRIO DE ENTRADA ---
st.info("Preencha os dados abaixo para gerar o documento Word automaticamente.")

col1, col2, col3 = st.columns(3)
with col1:
    cliente = st.text_input("Nome do Cliente", placeholder="Ex: Farmacêutica XYZ")
with col2:
    numero_prop = st.text_input("Número da Proposta", placeholder="Ex: P-2026-050")
with col3:
    prazo = st.text_input("Prazo de Execução", placeholder="Ex: 45 dias")

# --- SELEÇÃO DE ESCOPOS TÉCNICOS ---
st.markdown("### 🛠️ Escopo Técnico")

# Pega todas as categorias únicas (ex: Dutos, Água Gelada, VRF)
categorias = df_escopos['Categoria'].unique()
itens_selecionados_texto = []

# Cria uma caixinha expansível para cada categoria
for cat in categorias:
    with st.expander(f"Categoria: {cat}"):
        # Filtra os itens daquela categoria
        itens_da_categoria = df_escopos[df_escopos['Categoria'] == cat]
        
        # Cria o menu de seleção
        selecionados = st.multiselect(
            f"Selecione os itens de {cat}:",
            options=itens_da_categoria['Titulo'].tolist(),
            key=cat
        )
        
        # Pega o Texto Completo dos itens que foram selecionados
        textos = itens_da_categoria[itens_da_categoria['Titulo'].isin(selecionados)]['Texto_Completo'].tolist()
        itens_selecionados_texto.extend(textos)

# --- SELEÇÃO DE EXCLUSÕES ---
st.markdown("### 🚫 Exclusões")
# Por padrão, já deixamos todas as exclusões marcadas para não esquecer nada
todas_exclusoes = df_exclusoes['Titulo'].tolist()
exclusoes_selecionadas = st.multiselect(
    "Itens não inclusos no fornecimento:",
    options=todas_exclusoes,
    default=todas_exclusoes
)

# --- BOTÃO DE GERAR ---
st.markdown("---")
if st.button("🚀 Gerar Proposta (.docx)", type="primary"):
    
    # 1. Validação básica
    if not cliente or not numero_prop:
        st.warning("⚠️ Por favor, preencha o Nome do Cliente e o Número da Proposta.")
    else:
        try:
            # 2. Prepara os dados para o Template
            # Recupera os textos completos das exclusões selecionadas
            textos_exclusao_final = df_exclusoes[df_exclusoes['Titulo'].isin(exclusoes_selecionadas)]['Texto_Completo'].tolist()
            
            contexto = {
                'nome_cliente': cliente,
                'numero_proposta': numero_prop,
                'data_hoje': date.today().strftime("%d/%m/%Y"),
                'prazo_entrega': prazo,
                'lista_escopos': itens_selecionados_texto,
                'lista_exclusoes': textos_exclusao_final,
                'revisao': "R-00"
            }

            # 3. Abre o Template Word e preenche
            # O arquivo Template_Siarcon.docx DEVE estar junto no GitHub
            doc = DocxTemplate("Template_Siarcon.docx")
            doc.render(contexto)

            # 4. Salva o arquivo na memória (Buffer) para baixar
            buffer_arquivo = io.BytesIO()
            doc.save(buffer_arquivo)
            buffer_arquivo.seek(0)

            # 5. Mostra mensagem de sucesso e botão de download
            st.success(f"✅ Proposta para {cliente} gerada com sucesso!")
            
            st.download_button(
                label="📥 Clique para Baixar o Word",
                data=buffer_arquivo,
                file_name=f"Proposta_{numero_prop}_{cliente.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except FileNotFoundError:
            st.error("❌ Erro: O arquivo 'Template_Siarcon.docx' não foi encontrado no GitHub.")
            st.info("Certifique-se de que você fez o upload do arquivo Word para o repositório.")
        except Exception as e:
            st.error(f"❌ Ocorreu um erro inesperado: {e}")

