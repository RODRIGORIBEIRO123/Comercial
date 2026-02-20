import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador Propostas SIARCON", layout="wide", page_icon="📄")

# === MENU LATERAL (COMO SE FOSSEM 2 PÁGINAS) ===
st.sidebar.image("https://via.placeholder.com/150x50.png?text=SIARCON", use_container_width=True) # Pode trocar pela URL do seu logo depois
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

# Carrega Banco de Dados
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
# 4. ESCOPO TÉCNICO (MUTA DE ACORDO COM O MENU LATERAL)
# ==============================================================================
st.markdown("---")
intro = st.text_area("Introdução do Escopo", value="Trata-se do fornecimento de materiais e mão de obra conforme itens abaixo:")

escopo_estruturado = [] # Lista com os textos detalhados
eap_estruturada = []    # Lista com os textos curtos (Índice/EAP)
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
                        
                        # Monta EAP (Texto Curto)
                        lista_eap.append({'indice': f"{contador_cat}.{contador_item}", 'nome': item_curto})
                        
                        # Monta Escopo Detalhado (Texto Longo)
                        texto_final = texto_base
                        if qtd > 1: texto_final += f" — Qtd: {qtd}."
                        lista_detalhada.append(texto_final)
                        
                        contador_item += 1
                    
                    # Salva nas listas principais
                    eap_estruturada.append({'indice': str(contador_cat), 'categoria': cat.upper(), 'itens': lista_eap})
                    escopo_estruturado.append({'nome': cat.upper(), 'itens': lista_detalhada})
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
            
            categoria_atual = "ESCOPO GERAL"
            itens_detalhados = []
            itens_eap = []
            contador_cat = 1
            contador_item = 1
            
            for index, row in df_orc.iterrows():
                descricao = str(row[2]).strip()
                unidade = str(row[3]).strip()
                quantidade = row[4]
                v_mat, v_mao = row[12], row[13]
                
                if pd.isna(descricao) or descricao == "" or descricao.lower() in ["nan", "descrição dos materiais", "descrição"]:
                    continue
                    
                # É Categoria (Sem Quantidade)
                if pd.isna(quantidade) or str(quantidade).strip() in ["", "nan", "-"]:
                    if len(itens_detalhados) > 0:
                        escopo_estruturado.append({'nome': categoria_atual.upper(), 'itens': itens_detalhados})
                        eap_estruturada.append({'indice': str(contador_cat), 'categoria': categoria_atual.upper(), 'itens': itens_eap})
                        contador_cat += 1
                        itens_detalhados, itens_eap = [], []
                        contador_item = 1
                    
                    categoria_atual = descricao
                    
                # É Item de Escopo (Com Quantidade)
                else:
                    try: qtd_fmt = int(float(quantidade)) if float(quantidade).is_integer() else float(quantidade)
                    except: qtd_fmt = quantidade
                    
                    uni_fmt = f" {unidade}" if unidade.lower() not in ["nan", "", "-"] else ""
                    
                    # Para a EAP (Nome Curto)
                    itens_eap.append({'indice': f"{contador_cat}.{contador_item}", 'nome': descricao})
                    
                    # Para o Escopo Detalhado
                    texto_item = f"Fornecimento / Instalação de {qtd_fmt}{uni_fmt} - {descricao}."
                    itens_detalhados.append(texto_item)
                    
                    try: valor_total_calculado += float(v_mat) if not pd.isna(v_mat) else 0.0
                    except: pass
                    try: valor_total_calculado += float(v_mao) if not pd.isna(v_mao) else 0.0
                    except: pass
                    
                    contador_item += 1
            
            # Adiciona o último bloco
            if len(itens_detalhados) > 0:
                escopo_estruturado.append({'nome': categoria_atual.upper(), 'itens': itens_detalhados})
                eap_estruturada.append({'indice': str(contador_cat), 'categoria': categoria_atual.upper(), 'itens': itens_eap})
                
            if len(escopo_estruturado) > 0:
                st.success(f"✅ Planilha processada com sucesso! Encontrados {len(escopo_estruturado)} grupos de itens.")
            else:
                st.warning(f"⚠️ A aba '{nome_aba}' foi lida, mas não encontrei itens válidos.")
                
        except Exception as e:
            st.error(f"Erro ao processar a planilha: {e}")

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
# Se preencheu manual, vem vazio. Se foi Excel, vem preenchido.
valor = c_v.text_input("Valor Total (R$):", value=valor_formatado_sugerido if valor_total_calculado > 0 else "")
mes = c_m.text_input("Mês/Ano Base", value=f"{hoje.month}/{hoje.year}")

# ==============================================================================
# BOTÃO GERAR
# ==============================================================================
st.markdown("---")
if st.button("🚀 GERAR PROPOSTA (.DOCX)", type="primary"):
    
    if len(escopo_estruturado) == 0:
        st.warning("⚠️ O escopo técnico está vazio.")
    
    contexto = {
        'data_formatada': data_txt,
        'nome_contato': nome_contato, 'fone': fone, 'email': email,
        'nome_cliente': nome_cliente, 'nome_projeto': nome_projeto, 'cidade_estado': cidade_estado,
        'numero_proposta': num_prop,
        'texto_cobertura': texto_cob_final,
        'tem_docs': tem_docs, 'docs_referencia': lista_docs,
        'lista_resp_cliente': resp_final,
        
        # NOVAS VARIÁVEIS PARA O WORD
        'eap_estruturada': eap_estruturada, 
        'escopo_estruturado': escopo_estruturado, 
        
        'lista_exclusoes': exc_final,
        'intro_servico': intro,
        'mes_base': mes, 'valor_total': valor,
        'revisao': "R-00"
    }

    try:
        # ATENÇÃO: Confirme se o nome do seu template atual é este abaixo
        doc = DocxTemplate("Template_Siarcon_Auto.docx") 
        doc.render(contexto)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        st.success("✅ Proposta Gerada!")
        st.download_button("📥 Baixar Arquivo Word", bio, f"Proposta_{num_prop}.docx")
    except Exception as e:
        st.error(f"Erro ao gerar o Word: {e}")
