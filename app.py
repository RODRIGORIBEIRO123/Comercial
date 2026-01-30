import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador Propostas SIARCON", layout="wide", page_icon="📄")
st.title("📄 Gerador de Propostas - SIARCON")

# ==============================================================================
# ⚙️ CONFIGURAÇÃO DA PLANILHA
# ==============================================================================
# Substitua pelo nome EXATO da sua planilha no Google (o que aparece na aba do navegador)
PLANILHA_NOME = "DB_Propostas_Siarcon" 

# --- CONEXÃO SEGURA COM GOOGLE SHEETS ---
@st.cache_resource
def conectar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open(PLANILHA_NOME)
    except Exception as e:
        st.error(f"❌ Erro ao conectar com o Google Sheets: {e}")
        st.info("Verifique se o nome da planilha está correto e se o arquivo 'secrets' está configurado no Streamlit.")
        st.stop()

# --- FUNÇÃO ROBUSTA PARA CARREGAR DADOS ---
def carregar_dados():
    sh = conectar_google_sheets()
    
    def ler_aba_segura(nome_aba, colunas_obrigatorias):
        try:
            ws = sh.worksheet(nome_aba)
            dados = ws.get_all_records()
            df = pd.DataFrame(dados)
            
            # SE A PLANILHA ESTIVER VAZIA
            if df.empty:
                return pd.DataFrame(columns=colunas_obrigatorias)

            # LIMPEZA DE COLUNAS (Remove espaços extras que causam erro)
            df.columns = df.columns.str.strip()
            
            # VERIFICA SE AS COLUNAS EXISTEM
            faltando = [col for col in colunas_obrigatorias if col not in df.columns]
            if faltando:
                st.error(f"⚠️ Erro na aba '{nome_aba}': As colunas {faltando} não foram encontradas.")
                st.write(f"Colunas lidas: {list(df.columns)}")
                st.info(f"Por favor, renomeie a linha 1 da sua planilha para: {colunas_obrigatorias}")
                st.stop()
                
            return df
            
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"❌ A aba '{nome_aba}' não existe na planilha '{PLANILHA_NOME}'. Crie-a no Google Sheets.")
            st.stop()

    # Lê e valida cada aba
    df_esc = ler_aba_segura("Escopos", ["Categoria", "Titulo_Curto", "Texto_Completo"])
    df_exc = ler_aba_segura("Exclusoes", ["Titulo_Curto", "Texto_Completo"])
    df_resp = ler_aba_segura("Responsabilidades", ["Titulo_Curto", "Texto_Completo"])
    
    return df_esc, df_exc, df_resp

# --- FUNÇÃO PARA SALVAR NO BANCO ---
def salvar_no_banco(aba, dados_lista):
    try:
        sh = conectar_google_sheets()
        ws = sh.worksheet(aba)
        ws.append_row(dados_lista)
        st.cache_data.clear() # Limpa o cache
        st.toast(f"✅ Item salvo com sucesso em {aba}!", icon="💾")
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")

# Tenta carregar os dados
df_escopos, df_exclusoes, df_resp = carregar_dados()

# --- FUNÇÃO DE DATA EM PORTUGUÊS ---
def formatar_data_portugues(dt):
    meses = {1:'Janeiro', 2:'Fevereiro', 3:'Março', 4:'Abril', 5:'Maio', 6:'Junho',
             7:'Julho', 8:'Agosto', 9:'Setembro', 10:'Outubro', 11:'Novembro', 12:'Dezembro'}
    return f"Limeira, {dt.day} de {meses[dt.month]} de {dt.year}"

# ==============================================================================
# 🖥️ INTERFACE DO USUÁRIO
# ==============================================================================

# 1. DADOS DO PROJETO
st.header("1. Dados do Projeto")
c1, c2 = st.columns(2)
hoje = date.today()
data_txt = formatar_data_portugues(hoje)

with c1:
    st.info(f"📅 {data_txt}")
    nome_contato = st.text_input("Nome do Contato")
    fone = st.text_input("Telefone")
    email = st.text_input("Email")

with c2:
    cliente = st.text_input("Cliente (Empresa)")
    projeto = st.text_input("Nome do Projeto")
    local = st.text_input("Cidade/Estado", placeholder="Ex: Limeira - SP")
    num_prop = st.text_input("Nº Proposta", value=f"P-{hoje.year}-XXX")

# 2. COBERTURA
st.markdown("---")
st.header("2. Cobertura")

opcoes_cob = [
    "Os custos aqui apresentados compreendem: instalação com fornecimento de equipamentos, materiais e mão-de-obra, despesas de viagem, alimentação, com todos os impostos inclusos, excluindo os itens expressamente informados na seção exclusão.",
    "Fornecimento apenas de mão de obra.",
    "Outro (Personalizado)"
]
sel_cob = st.selectbox("Selecione o Texto Padrão:", opcoes_cob)
texto_cob_final = st.text_area("Texto Final (Editável):", value=sel_cob, height=80)

tem_docs = st.checkbox("Incluir Documentos de Referência?", value=True)
lista_docs = st.text_area("Lista de Documentos:") if tem_docs else ""

# 3. RESPONSABILIDADES DO CLIENTE
st.markdown("---")
st.header("3. Responsabilidades do Cliente")

# Cadastro
with st.expander("➕ Cadastrar NOVA Responsabilidade (Salvar no Banco)"):
    with st.form("nova_resp"):
        nr_curto = st.text_input("Título Curto (Menu)")
        nr_longo = st.text_input("Texto Completo (Proposta)")
        if st.form_submit_button("💾 Salvar"):
            if nr_curto and nr_longo:
                salvar_no_banco("Responsabilidades", [nr_curto, nr_longo])
                st.rerun()
            else:
                st.warning("Preencha todos os campos.")

# Seleção
dict_resp = dict(zip(df_resp['Titulo_Curto'], df_resp['Texto_Completo']))
sel_resp = st.multiselect("Selecione as obrigações:", list(dict_resp.keys()), default=list(dict_resp.keys()))
# Recupera o texto longo
resp_final = [dict_resp[k] for k in sel_resp if k in dict_resp]

# 4. ESCOPO TÉCNICO
st.markdown("---")
st.header("4. Escopo Técnico")
intro = st.text_area("Introdução do Escopo")

# Cadastro
with st.expander("➕ Cadastrar NOVO Item de Escopo (Salvar no Banco)"):
    with st.form("novo_esc"):
        # Pega categorias existentes ou deixa criar nova
        cats_existentes = df_escopos['Categoria'].unique().tolist() if not df_escopos.empty else []
        c_cat, c_tit, c_txt = st.columns([0.3, 0.3, 0.4])
        
        cat_selecionada = c_cat.selectbox("Categoria Existente", ["Nova Categoria..."] + cats_existentes)
        if cat_selecionada == "Nova Categoria...":
            cat_final = c_cat.text_input("Digite o nome da Nova Categoria")
        else:
            cat_final = cat_selecionada
            
        ne_tit = c_tit.text_input("Nome Equipamento (Menu)")
        ne_txt = c_txt.text_input("Descrição Técnica")
        
        if st.form_submit_button("💾 Salvar"):
            if cat_final and ne_tit and ne_txt:
                salvar_no_banco("Escopos", [cat_final, ne_tit, ne_txt])
                st.rerun()
            else:
                st.warning("Preencha todos os campos.")

escopo_final = []
contador = 1
st.subheader("Seleção por Disciplina")

# Garante que temos categorias
if not df_escopos.empty:
    categorias = df_escopos['Categoria'].unique()
    for cat in categorias:
        if st.checkbox(f"📁 {cat}", key=cat):
            df_c = df_escopos[df_escopos['Categoria'] == cat]
            
            # Menu com nomes curtos
            itens = st.multiselect(f"Equipamentos de {cat}", df_c['Titulo_Curto'].tolist(), key=f"sel_{cat}")
            
            lista_textos = []
            if itens:
                st.caption(f"Detalhes - {cat}:")
                for item in itens:
                    col_q, col_c = st.columns([0.2, 0.8])
                    qtd = col_q.number_input(f"Qtd ({item})", min_value=1, value=1, key=f"q_{cat}_{item}")
                    comp = col_c.text_input(f"Complemento ({item})", placeholder="Marca, Modelo...", key=f"c_{cat}_{item}")
                    
                    # Monta a frase: Fornecimento de X Item, Y.
                    texto = f"Fornecimento de {int(qtd)} {item}"
                    if comp: texto += f", {comp}"
                    texto += "."
                    lista_textos.append(texto)
                
                escopo_final.append({'indice': f"1.{contador}", 'nome': cat.upper(), 'itens': lista_textos})
                contador += 1
                st.markdown("---")
else:
    st.info("Nenhum escopo cadastrado. Use o formulário acima para adicionar.")

# 5. EXCLUSÕES
st.markdown("---")
st.header("5. Exclusões")

# Cadastro
with st.expander("➕ Cadastrar NOVA Exclusão (Salvar no Banco)"):
    with st.form("nova_exc"):
        nex_c = st.text_input("Título Curto")
        nex_l = st.text_input("Texto Completo")
        if st.form_submit_button("💾 Salvar"):
            if nex_c and nex_l:
                salvar_no_banco("Exclusoes", [nex_c, nex_l])
                st.rerun()
            else:
                st.warning("Preencha todos os campos")

dict_exc = dict(zip(df_exclusoes['Titulo_Curto'], df_exclusoes['Texto_Completo']))
sel_exc = st.multiselect("Exclusões:", list(dict_exc.keys()), default=list(dict_exc.keys()))
exc_final = [dict_exc[k] for k in sel_exc if k in dict_exc]

# 6. COMERCIAL
st.markdown("---")
st.header("6. Comercial")
c_v, c_m = st.columns(2)
valor = c_v.text_input("Valor Total (R$)")
mes = c_m.text_input("Mês/Ano Base", value=f"{hoje.month}/{hoje.year}")

# BOTÃO GERAR
st.markdown("---")
if st.button("🚀 GERAR PROPOSTA (.DOCX)", type="primary"):
    
    contexto = {
        'data_formatada': data_txt,
        'nome_contato': nome_contato, 'fone': fone, 'email': email,
        'nome_cliente': cliente, 'nome_projeto': projeto, 'cidade_estado': local,
        'numero_proposta': num_prop,
        
        'texto_cobertura': texto_cob_final,
        'tem_docs': tem_docs, 'docs_referencia': lista_docs,
        
        'lista_resp_cliente': resp_final,
        'escopo_estruturado': escopo_final,
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
        
        st.success("✅ Proposta Gerada com Sucesso!")
        st.download_button(
            label="📥 Baixar Arquivo Word",
            data=bio,
            file_name=f"Proposta_{num_prop}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"Erro ao gerar documento: {e}")
