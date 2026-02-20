import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador Propostas SIARCON", layout="wide", page_icon="📄")
st.title("📄 Gerador de Propostas Automático - SIARCON")

# === NOME EXATO DA PLANILHA DO BANCO DE DADOS ===
PLANILHA_NOME = "DB_Propostas_Siarcon" 

# --- CONEXÃO SEGURA ---
@st.cache_resource
def conectar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open(PLANILHA_NOME)
    except Exception as e:
        st.error(f"Erro na conexão com Google Sheets: {e}")
        st.stop()

# --- CARREGAR DADOS ---
def carregar_dados():
    sh = conectar_google_sheets()
    def ler_aba(nome):
        try: return pd.DataFrame(sh.worksheet(nome).get_all_records())
        except: return pd.DataFrame()

    return ler_aba("Exclusoes"), ler_aba("Responsabilidades"), ler_aba("Clientes"), ler_aba("Coberturas")

def salvar_no_banco(aba, dados_lista):
    try:
        sh = conectar_google_sheets()
        sh.worksheet(aba).append_row(dados_lista)
        st.cache_data.clear()
        st.toast(f"✅ Salvo em {aba} com sucesso!", icon="💾")
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")

# Carrega Banco de Dados
try:
    df_exclusoes, df_resp, df_clientes, df_coberturas = carregar_dados()
except Exception as e:
    st.error("Erro ao ler banco de dados. Verifique a planilha.")
    st.stop()

def formatar_data_portugues(dt):
    meses = {1:'Janeiro', 2:'Fevereiro', 3:'Março', 4:'Abril', 5:'Maio', 6:'Junho',
             7:'Julho', 8:'Agosto', 9:'Setembro', 10:'Outubro', 11:'Novembro', 12:'Dezembro'}
    return f"Limeira, {dt.day} de {meses[dt.month]} de {dt.year}"

# ==============================================================================
# 1. DADOS DO PROJETO E CLIENTE
# ==============================================================================
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

# ==============================================================================
# 2. COBERTURA
# ==============================================================================
st.markdown("---")
st.header("2. Cobertura")

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

# ==============================================================================
# 3. RESPONSABILIDADES DO CLIENTE
# ==============================================================================
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

# ==============================================================================
# 4. ESCOPO TÉCNICO (AUTOMATIZADO VIA EXCEL)
# ==============================================================================
st.markdown("---")
st.header("4. Escopo Técnico (Upload da Planilha)")

intro = st.text_area("Introdução do Escopo", value="Trata-se do fornecimento de materiais e mão de obra conforme itens abaixo:")

arquivo_excel = st.file_uploader("📂 Faça o upload da Planilha Orçamentária (.xlsx)", type=["xlsx"])

escopo_estruturado = []
valor_total_calculado = 0.0

if arquivo_excel is not None:
    try:
        # Lê a planilha inteira sem forçar colunas específicas para não dar erro "out-of-bounds"
        df_orc = pd.read_excel(arquivo_excel, header=None)
        
        categoria_atual = "ESCOPO GERAL"
        itens_da_categoria = []
        contador_cat = 1
        
        # Garante que a planilha tem pelo menos até a coluna C (índice 2)
        if df_orc.shape[1] > 2:
            for index, row in df_orc.iterrows():
                
                # Acesso seguro às colunas: só puxa o dado se a coluna existir na planilha enviada
                descricao = str(row[2]).strip() if 2 in df_orc.columns else ""
                unidade = str(row[3]).strip() if 3 in df_orc.columns else ""
                quantidade = row[4] if 4 in df_orc.columns else pd.NA
                v_mat = row[12] if 12 in df_orc.columns else 0.0
                v_mao = row[13] if 13 in df_orc.columns else 0.0
                
                # Pula linhas vazias ou cabeçalhos
                if pd.isna(descricao) or descricao == "" or descricao.lower() in ["nan", "descrição dos materiais"]:
                    continue
                    
                # LÓGICA DE DETECÇÃO: Se a Quantidade estiver vazia, é um Título/Categoria
                if pd.isna(quantidade) or str(quantidade).strip() in ["", "nan", "-"]:
                    # Salva a categoria anterior antes de criar a nova
                    if len(itens_da_categoria) > 0:
                        escopo_estruturado.append({
                            'indice': f"1.{contador_cat}",
                            'nome': categoria_atual.upper(),
                            'itens': itens_da_categoria
                        })
                        contador_cat += 1
                        itens_da_categoria = [] 
                    
                    categoria_atual = descricao
                    
                # Se a Quantidade NÃO está vazia, é um Item de Escopo
                else:
                    try: qtd_fmt = int(float(quantidade)) if float(quantidade).is_integer() else float(quantidade)
                    except: qtd_fmt = quantidade
                    
                    uni_fmt = f" {unidade}" if unidade.lower() not in ["nan", "", "-"] else ""
                    
                    texto_item = f"Fornecimento / Instalação de {qtd_fmt}{uni_fmt} - {descricao}."
                    itens_da_categoria.append(texto_item)
                    
                    # Soma os valores para o total
                    try: valor_total_calculado += float(v_mat)
                    except: pass
                    try: valor_total_calculado += float(v_mao)
                    except: pass
            
            # Adiciona a última categoria que ficou no loop
            if len(itens_da_categoria) > 0:
                escopo_estruturado.append({
                    'indice': f"1.{contador_cat}",
                    'nome': categoria_atual.upper(),
                    'itens': itens_da_categoria
                })
                
            st.success(f"✅ Planilha lida com sucesso! Encontrados {len(escopo_estruturado)} grupos de itens.")
        else:
            st.error("⚠️ A planilha enviada parece estar vazia ou não tem colunas suficientes.")
            
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")

# ==============================================================================
# 5. EXCLUSÕES
# ==============================================================================
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

# ==============================================================================
# 6. COMERCIAL
# ==============================================================================
st.markdown("---")
st.header("6. Comercial")

valor_formatado_sugerido = f"R$ {valor_total_calculado:_.2f}".replace('.', ',').replace('_', '.')

c_v, c_m = st.columns(2)
valor = c_v.text_input("Valor Total (R$) - Somado do Excel:", value=valor_formatado_sugerido if valor_total_calculado > 0 else "")
mes = c_m.text_input("Mês/Ano Base", value=f"{hoje.month}/{hoje.year}")

# ==============================================================================
# BOTÃO GERAR
# ==============================================================================
st.markdown("---")
if st.button("🚀 GERAR PROPOSTA (.DOCX)", type="primary"):
    
    if len(escopo_estruturado) == 0:
        st.warning("⚠️ Você não fez o upload da planilha de Orçamento. O escopo sairá vazio.")
    
    contexto = {
        'data_formatada': data_txt,
        'nome_contato': nome_contato, 'fone': fone, 'email': email,
        'nome_cliente': nome_cliente, 'nome_projeto': nome_projeto, 'cidade_estado': cidade_estado,
        'numero_proposta': num_prop,
        'texto_cobertura': texto_cob_final,
        'tem_docs': tem_docs, 'docs_referencia': lista_docs,
        'lista_resp_cliente': resp_final,
        'escopo_estruturado': escopo_estruturado, 
        'lista_exclusoes': exc_final,
        'intro_servico': intro,
        'mes_base': mes, 'valor_total': valor,
        'revisao': "R-00"
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
        st.error(f"Erro ao gerar: {e}")
