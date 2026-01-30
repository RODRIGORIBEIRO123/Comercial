import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date
import locale

# Tenta configurar data em português (pode variar dependendo do servidor, mas tentamos)
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except:
    pass

# ==============================================================================
# 🔗 CONFIGURAÇÃO DOS LINKS (ATUALIZE COM SEUS LINKS CSV)
# ==============================================================================
URL_ESCOPOS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9Dlv9q_qBgpCIwY6cQAfWTYY6JXO9ILRMN_NT_QNjFiWAy2N5W9QqjP51U2fAnE2mi-RCEtj5l2wG/pub?gid=221408068&single=true&output=csv"
URL_EXCLUSOES = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9Dlv9q_qBgpCIwY6cQAfWTYY6JXO9ILRMN_NT_QNjFiWAy2N5W9QqjP51U2fAnE2mi-RCEtj5l2wG/pub?gid=1129521636&single=true&output=csv"
URL_RESPONSABILIDADES = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR9Dlv9q_qBgpCIwY6cQAfWTYY6JXO9ILRMN_NT_QNjFiWAy2N5W9QqjP51U2fAnE2mi-RCEtj5l2wG/pub?gid=1382076169&single=true&output=csv"
# ==============================================================================

st.set_page_config(page_title="Gerador Propostas SIARCON", layout="wide", page_icon="📄")
st.title("📄 Gerador de Propostas - SIARCON")
st.markdown("---")

# --- 1. CARREGAMENTO DE DADOS ---
@st.cache_data(ttl=60)
def carregar_dados():
    try:
        df_esc = pd.read_csv(URL_ESCOPOS)
        df_exc = pd.read_csv(URL_EXCLUSOES)
        df_resp = pd.read_csv(URL_RESPONSABILIDADES) # Nova tabela
        
        # Limpeza
        for df in [df_esc, df_exc, df_resp]:
            df.columns = df.columns.str.strip()
            
        return df_esc, df_exc, df_resp
    except Exception as e:
        return None, None, None, str(e)

res = carregar_dados()
if res[0] is None:
    st.error("❌ Erro ao carregar dados. Verifique os links CSV.")
    st.stop()

df_escopos, df_exclusoes, df_responsabilidades = res

# --- 2. DADOS DO PROJETO (CABEÇALHO) ---
st.header("1. Dados do Projeto")
col_a, col_b = st.columns(2)
with col_a:
    cidade_data = st.text_input("Local e Data", value=f"Limeira, {date.today().strftime('%d de %B de %Y')}")
    nome_contato = st.text_input("Nome do Contato")
    email_contato = st.text_input("Email")
    fone_contato = st.text_input("Telefone")
with col_b:
    nome_empresa = st.text_input("Nome da Empresa (Cliente)")
    nome_projeto = st.text_input("Nome do Projeto")
    local_projeto = st.text_input("Cidade/Estado da Obra")
    numero_prop = st.text_input("Número da Proposta", value=f"P-{date.today().year}-XXX")

# --- 3. COBERTURA DO FORNECIMENTO ---
st.markdown("---")
st.header("2. Cobertura e Documentos")

texto_padrao_cobertura = """Os custos aqui apresentados compreendem: instalação com fornecimento de equipamentos, materiais e mão-de-obra, despesas de viagem, alimentação, com todos os impostos inclusos, excluindo os itens expressamente informados na seção exclusão."""

texto_cobertura = st.text_area("Texto de Cobertura do Fornecimento", value=texto_padrao_cobertura, height=100)
lista_docs = st.text_area("Projetos/Documentos de Referência", placeholder="Liste aqui os projetos recebidos (ex: Planta baixa rev.02, Diagrama unifilar...)", height=100)

# --- 4. RESPONSABILIDADES DO CLIENTE ---
st.markdown("---")
st.header("3. Responsabilidades do Cliente")
with st.expander("Selecione as obrigações do cliente (Checklist)", expanded=False):
    todas_resp = df_responsabilidades['Texto_Completo'].tolist()
    # Por padrão, marca todas
    resp_selecionadas = st.multiselect(
        "Itens de Responsabilidade do Cliente:",
        options=todas_resp,
        default=todas_resp
    )

# --- 5. ESCOPO TÉCNICO (SIARCON) ---
st.markdown("---")
st.header("4. Obrigações SIARCON (Escopo Técnico)")

intro_servico = st.text_area("Introdução / Descrição Geral do Serviço", placeholder="Ex: Trata-se de fornecimento e instalação de rede de dutos para o novo sistema...", height=100)

st.subheader("Seleção de Itens por Disciplina")
escopo_estruturado = []
contador_cat = 1

# Garante a ordem correta das categorias
categorias = df_escopos['Categoria'].unique()

for cat in categorias:
    # Checkbox para "Ligar" a categoria
    col_check, col_exp = st.columns([0.05, 0.95])
    with col_check:
        ativo = st.checkbox("", key=f"chk_{cat}")
    with col_exp:
        # Se ativado, mostra o expander com os itens
        if ativo:
            with st.expander(f"**{cat}** (Selecionado)", expanded=True):
                df_cat = df_escopos[df_escopos['Categoria'] == cat]
                # Checkbox para selecionar itens específicos
                itens_sel = st.multiselect(
                    f"Itens de {cat}:",
                    options=df_cat['Titulo'].tolist(),
                    default=df_cat['Titulo'].tolist(), # Já vem tudo marcado pra facilitar
                    key=f"multi_{cat}"
                )
                
                if itens_sel:
                    # Busca os textos completos
                    textos = df_cat[df_cat['Titulo'].isin(itens_sel)]['Texto_Completo'].tolist()
                    escopo_estruturado.append({
                        'indice': f"1.{contador_cat}", # Ajuste conforme a numeração desejada no doc (ex: 6.1, 6.2)
                        'nome': cat.upper(),
                        'itens': textos
                    })
                    contador_cat += 1
        else:
            st.write(f"{cat}") # Mostra apenas o nome cinza se não estiver ativo

# --- 6. EXCLUSÕES ---
st.markdown("---")
st.header("5. Exclusões")
with st.expander("Itens Exclusos (Clique para alterar)", expanded=False):
    todas_exc = df_exclusoes['Texto_Completo'].tolist() # Usando texto completo direto
    exc_selecionadas = st.multiselect(
        "Selecione as exclusões:",
        options=todas_exc,
        default=todas_exc
    )

# --- 7. CONDIÇÕES COMERCIAIS ---
st.markdown("---")
st.header("6. Condições Comerciais")
col_c1, col_c2 = st.columns(2)
with col_c1:
    mes_base = st.text_input("Mês/Ano Base", value=date.today().strftime("%B/%Y"))
with col_c2:
    valor_total = st.text_input("Valor Total (R$)", placeholder="Ex: R$ 150.000,00")

# --- GERAÇÃO ---
st.markdown("---")
if st.button("🚀 GERAR PROPOSTA COMPLETA", type="primary"):
    
    # Monta o contexto para o Word
    contexto = {
        # Cabeçalho
        'cidade_data': cidade_data,
        'nome_empresa': nome_empresa,
        'nome_contato': nome_contato,
        'fone': fone_contato,
        'email': email_contato,
        'nome_projeto': nome_projeto,
        'local_projeto': local_projeto,
        'numero_proposta': numero_prop,
        
        # Cobertura
        'texto_cobertura': texto_cobertura,
        'lista_docs': lista_docs,
        
        # Listas
        'lista_resp_cliente': resp_selecionadas,
        
        # Escopo
        'intro_servico': intro_servico,
        'escopo_estruturado': escopo_estruturado,
        
        # Exclusão
        'lista_exclusoes': exc_selecionadas,
        
        # Comercial
        'mes_base': mes_base,
        'valor_total': valor_total,
        'revisao': "R-00"
    }
    
    try:
        doc = DocxTemplate("Template_Siarcon.docx")
        doc.render(contexto)
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success("✅ Proposta Gerada!")
        st.download_button(
            label="📥 Baixar Arquivo Word",
            data=buffer,
            file_name=f"Proposta_{numero_prop}_{nome_empresa}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"Erro ao gerar documento: {e}")

