import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date

# ==============================================================================
# 🔗 CONFIGURAÇÃO DOS LINKS DO GOOGLE SHEETS
# ==============================================================================
# Passo importante: Substitua os links abaixo pelos seus links CSV "Publicados na Web"
URL_ESCOPOS = "https://docs.google.com/spreadsheets/d/e/SUBSTITUA_PELO_SEU_LINK_ESCOPOS/pub?gid=0&single=true&output=csv"
URL_EXCLUSOES = "https://docs.google.com/spreadsheets/d/e/SUBSTITUA_PELO_SEU_LINK_EXCLUSOES/pub?gid=0&single=true&output=csv"
# ==============================================================================

# Configuração da página
st.set_page_config(page_title="Gerador Propostas SIARCON", layout="wide", page_icon="📄")

st.title("📄 Gerador de Propostas - SIARCON")
st.markdown("---")

# --- 1. CARREGAR DADOS (CACHE) ---
@st.cache_data(ttl=60)
def carregar_dados():
    try:
        # Lê direto do Google Sheets
        df_esc = pd.read_csv(URL_ESCOPOS)
        df_exc = pd.read_csv(URL_EXCLUSOES)
        
        # Limpeza básica nas colunas
        df_esc.columns = df_esc.columns.str.strip()
        df_exc.columns = df_exc.columns.str.strip()
        
        return df_esc, df_exc
    except Exception as e:
        # A CORREÇÃO ESTÁ AQUI: Retornamos o TEXTO do erro (str(e)), não o objeto erro
        return None, str(e)

# --- 2. INTERFACE DE DADOS DO CLIENTE ---
st.subheader("📝 Dados do Projeto")
col1, col2, col3 = st.columns(3)
with col1:
    cliente = st.text_input("Nome do Cliente", placeholder="Ex: Farmacêutica XYZ")
with col2:
    numero_prop = st.text_input("Número da Proposta", placeholder="Ex: P-2026-050")
with col3:
    prazo = st.text_input("Prazo de Execução", placeholder="Ex: 45 dias")

# --- 3. SELEÇÃO DE ESCOPO TÉCNICO ---
st.markdown("---")
st.subheader("🛠️ Escopo Técnico")

# Lista para guardar TÍTULOS de tudo que foi selecionado
# Usamos session_state ou apenas uma lista acumuladora simples aqui
itens_selecionados_titulos = []

# Pega as categorias na ordem original do Excel (para manter Água Gelada antes de Elétrica, etc)
categorias_ordenadas = df_escopos['Categoria'].unique()

for cat in categorias_ordenadas:
    with st.expander(f"Categoria: {cat}"):
        # Filtra itens dessa categoria
        df_cat = df_escopos[df_escopos['Categoria'] == cat]
        
        # Cria o checkbox
        selecao = st.multiselect(
            f"Selecione os itens de {cat}:",
            options=df_cat['Titulo'].tolist(),
            key=cat
        )
        itens_selecionados_titulos.extend(selecao)

# --- 4. SELEÇÃO DE EXCLUSÕES ---
st.markdown("---")
st.subheader("🚫 Exclusões")
todas_exclusoes = df_exclusoes['Titulo'].tolist()
exclusoes_selecionadas = st.multiselect(
    "Itens não inclusos no fornecimento:",
    options=todas_exclusoes,
    default=todas_exclusoes
)

# --- 5. BOTÃO E LÓGICA DE GERAÇÃO ---
st.markdown("---")
if st.button("🚀 Gerar Proposta (.docx)", type="primary"):
    
    if not cliente or not numero_prop:
        st.warning("⚠️ Preencha o Cliente e o Número da Proposta.")
    else:
        try:
            # === A MÁGICA DO AGRUPAMENTO ===
            # O objetivo é criar a estrutura: 1.1 Categoria -> Lista de Itens
            
            escopo_estruturado = []
            contador_categoria = 1
            
            # Varre as categorias na ordem do Excel novamente
            for cat in categorias_ordenadas:
                # Descobre quais itens DESSA categoria o usuário marcou
                # Filtra o Excel: Categoria É IGUAL a atual E o Título ESTÁ na lista de selecionados
                df_itens_selecionados = df_escopos[
                    (df_escopos['Categoria'] == cat) & 
                    (df_escopos['Titulo'].isin(itens_selecionados_titulos))
                ]
                
                # Se tiver pelo menos um item selecionado, cria o grupo
                if not df_itens_selecionados.empty:
                    grupo = {
                        'indice': f"1.{contador_categoria}",  # Gera 1.1, 1.2, 1.3...
                        'nome_categoria': cat.upper(),        # Ex: SISTEMA DE ÁGUA GELADA
                        'lista_itens': df_itens_selecionados['Texto_Completo'].tolist()
                    }
                    escopo_estruturado.append(grupo)
                    contador_categoria += 1
            
            # Prepara as exclusões (apenas lista de textos)
            lista_exclusoes_texto = df_exclusoes[
                df_exclusoes['Titulo'].isin(exclusoes_selecionadas)
            ]['Texto_Completo'].tolist()

            # Dicionário final para o Word
            contexto = {
                'nome_cliente': cliente,
                'numero_proposta': numero_prop,
                'data_hoje': date.today().strftime("%d/%m/%Y"),
                'prazo_entrega': prazo,
                'escopo_estruturado': escopo_estruturado, # <--- Lista de Grupos
                'lista_exclusoes': lista_exclusoes_texto,
                'revisao': "R-00"
            }

            # Carrega Template e Renderiza
            doc = DocxTemplate("Template_Siarcon.docx")
            doc.render(contexto)

            # Salva na memória
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            st.success(f"✅ Proposta gerada! Baixe abaixo:")
            
            st.download_button(
                label="📥 Baixar Documento Word",
                data=buffer,
                file_name=f"Proposta_{numero_prop}_{cliente.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except FileNotFoundError:
            st.error("Erro: O arquivo 'Template_Siarcon.docx' não foi encontrado no GitHub.")
        except Exception as e:
            st.error(f"Erro inesperado: {e}")

