import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date

# ==============================================================================
# 🔗 CONFIGURAÇÃO DOS LINKS DO GOOGLE SHEETS
# ==============================================================================
# 1. Cole o link CSV da aba "Escopos" AQUI EMBAIXO:
URL_ESCOPOS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9Dlv9q_qBgpCIwY6cQAfWTYY6JXO9ILRMN_NT_QNjFiWAy2N5W9QqjP51U2fAnE2mi-RCEtj5l2wG/pub?gid=221408068&single=true&output=csv"

# 2. Cole o link CSV da aba "Exclusoes" AQUI EMBAIXO:
URL_EXCLUSOES = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9Dlv9q_qBgpCIwY6cQAfWTYY6JXO9ILRMN_NT_QNjFiWAy2N5W9QqjP51U2fAnE2mi-RCEtj5l2wG/pub?gid=1129521636&single=true&output=csv"
# ==============================================================================

st.set_page_config(page_title="Gerador Propostas SIARCON", layout="wide", page_icon="📄")
st.title("📄 Gerador de Propostas - SIARCON")
st.markdown("---")

# --- 1. CARREGAMENTO SEGURO DOS DADOS ---
@st.cache_data(ttl=60)
def carregar_dados():
    try:
        # Tenta ler as planilhas
        df_esc = pd.read_csv(URL_ESCOPOS)
        df_exc = pd.read_csv(URL_EXCLUSOES)
        
        # Limpa nomes das colunas (remove espaços extras)
        df_esc.columns = df_esc.columns.str.strip()
        df_exc.columns = df_exc.columns.str.strip()
        
        return df_esc, df_exc
        
    except Exception as e:
        # Se der erro, retorna None e o texto do erro
        return None, str(e)

# Executa o carregamento
resultado = carregar_dados()

# Verifica se o carregamento funcionou
if resultado[0] is None:
    # Se falhou, para tudo e mostra o erro na tela
    st.error("❌ ERRO CRÍTICO: Não foi possível carregar o Banco de Dados.")
    st.warning("Verifique se os LINKS do Google Sheets no código estão corretos (formato CSV e publicados na web).")
    st.code(f"Detalhe do erro: {resultado[1]}") # Mostra o erro técnico
    st.stop() # PARA A EXECUÇÃO AQUI para não dar NameError lá embaixo

# Se chegou aqui, é seguro desempacotar as variáveis
df_escopos, df_exclusoes = resultado

# --- 2. INTERFACE ---
col1, col2, col3 = st.columns(3)
with col1:
    cliente = st.text_input("Nome do Cliente", placeholder="Ex: Farmacêutica XYZ")
with col2:
    numero_prop = st.text_input("Número da Proposta", placeholder="Ex: P-2026-050")
with col3:
    prazo = st.text_input("Prazo de Execução", placeholder="Ex: 45 dias")

st.markdown("### 🛠️ Escopo Técnico")

# --- 3. LÓGICA DE SELEÇÃO E AGRUPAMENTO ---
if 'Categoria' not in df_escopos.columns:
    st.error("A planilha de Escopos não tem a coluna 'Categoria'. Verifique o arquivo.")
    st.stop()

categorias_ordenadas = df_escopos['Categoria'].unique()
itens_selecionados_titulos = []

for cat in categorias_ordenadas:
    with st.expander(f"Categoria: {cat}"):
        df_cat = df_escopos[df_escopos['Categoria'] == cat]
        selecao = st.multiselect(
            f"Itens de {cat}:",
            options=df_cat['Titulo'].tolist(),
            key=cat
        )
        itens_selecionados_titulos.extend(selecao)

st.markdown("### 🚫 Exclusões")
todas_exclusoes = df_exclusoes['Titulo'].tolist()
exclusoes_selecionadas = st.multiselect(
    "Itens não inclusos:",
    options=todas_exclusoes,
    default=todas_exclusoes
)

# --- 4. GERAÇÃO DO ARQUIVO ---
st.markdown("---")
if st.button("🚀 Gerar Proposta (.docx)", type="primary"):
    if not cliente or not numero_prop:
        st.warning("Preencha Cliente e Número da Proposta.")
    else:
        try:
            escopo_estruturado = []
            contador_categoria = 1
            
            for cat in categorias_ordenadas:
                # Filtra itens selecionados desta categoria
                df_itens = df_escopos[
                    (df_escopos['Categoria'] == cat) & 
                    (df_escopos['Titulo'].isin(itens_selecionados_titulos))
                ]
                
                if not df_itens.empty:
                    grupo = {
                        'indice': f"1.{contador_categoria}",
                        'nome_categoria': cat.upper(),
                        'lista_itens': df_itens['Texto_Completo'].tolist()
                    }
                    escopo_estruturado.append(grupo)
                    contador_categoria += 1
            
            # Prepara exclusões
            lista_exclusoes = df_exclusoes[
                df_exclusoes['Titulo'].isin(exclusoes_selecionadas)
            ]['Texto_Completo'].tolist()

            contexto = {
                'nome_cliente': cliente,
                'numero_proposta': numero_prop,
                'data_hoje': date.today().strftime("%d/%m/%Y"),
                'prazo_entrega': prazo,
                'escopo_estruturado': escopo_estruturado,
                'lista_exclusoes': lista_exclusoes,
                'revisao': "R-00"
            }

            doc = DocxTemplate("Template_Siarcon.docx")
            doc.render(contexto)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("Proposta gerada com sucesso!")
            st.download_button(
                label="📥 Baixar Word",
                data=buffer,
                file_name=f"Proposta_{numero_prop}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Erro na geração: {e}")
