import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador Propostas SIARCON", layout="wide", page_icon="📄")

# === MENU LATERAL ===
st.sidebar.image("https://via.placeholder.com/150x50.png?text=SIARCON", use_container_width=True)
st.sidebar.title("Navegação")
modo_preenchimento = st.sidebar.radio(
    "Como deseja preencher o Escopo Técnico?",
    ["📋 Preenchimento Manual", "📊 Automático (Excel)"]
)

st.title("📄 Gerador de Propostas - SIARCON")

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
    df_escopos, df_exclusoes, df_resp, df_clientes, df_coberturas = carregar_dados()
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
# 4. ESCOPO TÉCNICO
# ==============================================================================
st.markdown("---")
intro = st.text_area("Introdução do Escopo", value="Trata-se do fornecimento de materiais e mão de obra conforme itens abaixo:")

escopo_estruturado = [] 
eap_estruturada = []    
valor_total_calculado = 0.0

# ---------------------------------------------------------
# MODO 1: PREENCHIMENTO MANUAL
# ---------------------------------------------------------
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
                        
                        lista_eap.append({'indice': f"{contador_cat}.{contador_item}", 'nome': item_curto})
                        
                        texto_final = texto_base
                        if qtd > 1: texto_final += f" — Qtd: {qtd}."
                        lista_detalhada.append(texto_final)
                        
                        contador_item += 1
                    
                    eap_estruturada.append({'indice': str(contador_cat), 'categoria': cat.upper(), 'itens': lista_eap})
                    escopo_estruturado.append({'nome': f"{contador_cat}. {cat.upper()}", 'itens': lista_detalhada})
                    contador_cat += 1

# ---------------------------------------------------------
# MODO 2: PREENCHIMENTO AUTOMÁTICO (EXCEL)
# ---------------------------------------------------------
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

            # --- PRÉ-VARREDURA: Achar Valor Total na Coluna J e parar sem bugar ---
            for idx, row in df_orc.iterrows():
                texto_col_e = str(row[4]).upper().strip() if 4 < len(row) else ""
                texto_col_c = str(row[2]).upper().strip() if 2 < len(row) else ""
                
                if "PREÇO VENDA" in texto_col_e or "PRECO VENDA" in texto_col_e or "PREÇO VENDA" in texto_col_c or "PRECO VENDA" in texto_col_c:
                    val_bruto = row[9] # Coluna J
                    if pd.notna(val_bruto):
                        # BLINDAGEM: Se o Excel mandar como float, a gente usa direto sem manipular string!
                        if isinstance(val_bruto, (int, float)):
                            valor_total_calculado = float(val_bruto)
                        else:
                            val_str = str(val_bruto).upper().replace("R$", "").strip()
                            val_str = val_str.replace(".", "").replace(",", ".")
                            try: valor_total_calculado = float(val_str)
                            except: pass
                    break # Para de procurar o preço

            # --- LEITURA PRINCIPAL DO ESCOPO ---
            categoria_atual_nome = "ESCOPO GERAL"
            categoria_atual_indice = ""
            itens_detalhados = []
            itens_eap = []
            contador_item = 1
            
            for index, row in df_orc.iterrows():
                
                descricao = str(row[2]).strip() if 2 < len(row) else ""
                
                # PARADA ABSOLUTA: Se achar CUSTO INDIRETO na descrição, quebra o loop inteiro.
                if "CUSTO INDIRETO" in descricao.upper() or "CUSTOS INDIRETOS" in descricao.upper():
                    break 
                
                col_b = str(row[1]).strip() if 1 < len(row) else ""
                unidade = str(row[3]).strip() if 3 < len(row) else ""
                quantidade = row[4] if 4 < len(row) else pd.NA
                
                if pd.isna(descricao) or descricao == "" or descricao.upper() in ["NAN", "DESCRIÇÃO DOS MATERIAIS", "DESCRICAO DOS MATERIAIS", "DESCRIÇÃO", "ITEM"]:
                    continue
                    
                # É CATEGORIA / TÍTULO DA EAP? (Apenas se tiver número na Coluna ITEM)
                is_header = False
                if col_b and col_b.lower() != 'nan' and col_b != '-':
                    if col_b[0].isdigit():
                        is_header = True
                        
                if is_header:
                    # Salva a categoria anterior
                    if categoria_atual_nome != "ESCOPO GERAL" and len(itens_detalhados) > 0:
                        escopo_estruturado.append({'nome': f"{categoria_atual_indice} - {categoria_atual_nome}".strip(' -'), 'itens': itens_detalhados})
                        eap_estruturada.append({'indice': categoria_atual_indice, 'categoria': categoria_atual_nome.upper(), 'itens': itens_eap})
                    
                    categoria_atual_indice = col_b
                    categoria_atual_nome = descricao
                    itens_detalhados = []
                    itens_eap = []
                    contador_item = 1
                    
                # É ITEM OU COMPLEMENTO DE ITEM
                else:
                    has_qty = not pd.isna(quantidade) and str(quantidade).strip() not in ["", "nan", "-"]
                    
                    # Tem quantidade = Item Novo
                    if has_qty:
                        # BLINDAGEM DE QUANTIDADE: Para não sair "205.894" como quantidade (quando pega sujeira do cabeçalho)
                        if isinstance(quantidade, (int, float)):
                            qtd_fmt = int(quantidade) if float(quantidade).is_integer() else round(float(quantidade), 2)
                        else:
                            try: 
                                q_str = str(quantidade).replace(",", ".")
                                qtd_fmt = int(float(q_str)) if float(q_str).is_integer() else round(float(q_str), 2)
                            except: 
                                qtd_fmt = quantidade
                                
                        uni_fmt = f" {unidade}" if unidade.lower() not in ["nan", "", "-"] else ""
                        
                        # Resumo Inteligente (Corta no "|")
                        nome_resumido = descricao.split('|')[0].strip()
                        if len(nome_resumido) > 80: nome_resumido = nome_resumido[:80] + "..."
                        
                        indice_item = f"{categoria_atual_indice}.{contador_item}" if categoria_atual_indice else str(contador_item)
                        itens_eap.append({'indice': indice_item, 'nome': nome_resumido})
                        
                        # Item Detalhado
                        texto_item = f"Fornecimento / Instalação de {qtd_fmt}{uni_fmt} - {descricao}."
                        itens_detalhados.append(texto_item)
                        
                        contador_item += 1
                        
                    # Não tem quantidade = Complemento do item de cima (ex: Pintura, Automação...)
                    else:
                        if len(itens_detalhados) > 0:
                            itens_detalhados[-1] += f" {descricao}"
            
            # Adiciona o último bloco de categorias lido
            if categoria_atual_nome != "ESCOPO GERAL" and len(itens_detalhados) > 0:
                escopo_estruturado.append({'nome': f"{categoria_atual_indice} - {categoria_atual_nome}".strip(' -'), 'itens': itens_detalhados})
                eap_estruturada.append({'indice': categoria_atual_indice, 'categoria': categoria_atual_nome.upper(), 'itens': itens_eap})
                
            if len(escopo_estruturado) > 0:
                st.success(f"✅ Planilha processada com sucesso! Estrutura EAP e valor total importados.")
            else:
                st.warning(f"⚠️ A aba '{nome_aba}' foi lida, mas não encontrei itens válidos na Coluna B.")
                
        except Exception as e:
            st.error(f"Erro ao processar a planilha: {e}")

# ==============================================================================
# 5. EXCLUSÕES
# ==============================================================================
st.markdown("
