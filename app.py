import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date, datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="App SIARCON - Propostas e Custos", layout="wide", page_icon="📄")

# === MENU LATERAL PRINCIPAL ===
st.sidebar.image("https://via.placeholder.com/150x50.png?text=SIARCON", use_container_width=True)
st.sidebar.title("Navegação Principal")

menu_selecionado = st.sidebar.radio(
    "Selecione o módulo:",
    ["📄 Gerador de Propostas", "💰 Estimativa de Custos"]
)

st.sidebar.markdown("---")

# ==============================================================================
# MÓDULO 1: GERADOR DE PROPOSTAS
# ==============================================================================
if menu_selecionado == "📄 Gerador de Propostas":
    
    st.sidebar.subheader("Opções da Proposta")
    modo_preenchimento = st.sidebar.radio(
        "Como deseja preencher o Escopo Técnico?",
        ["📋 Preenchimento Manual", "📊 Automático (Excel)"]
    )

    st.title("📄 Gerador de Propostas - SIARCON")

    PLANILHA_NOME = "DB_Propostas_Siarcon" 

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

    def carregar_dados_propostas():
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
        df_escopos, df_exclusoes, df_resp, df_clientes, df_coberturas = carregar_dados_propostas()
    except Exception as e:
        st.error("Erro ao ler banco de dados. Verifique a planilha.")
        st.stop()

    def formatar_data_portugues(dt):
        meses = {1:'Janeiro', 2:'Fevereiro', 3:'Março', 4:'Abril', 5:'Maio', 6:'Junho',
                 7:'Julho', 8:'Agosto', 9:'Setembro', 10:'Outubro', 11:'Novembro', 12:'Dezembro'}
        return f"Limeira, {dt.day} de {meses[dt.month]} de {dt.year}"

    # --- 1. DADOS DO PROJETO E CLIENTE ---
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

    # --- 2. COBERTURA ---
    st.markdown("---")
    st.header("2. COBERTURA")

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

    # --- 3. RESPONSABILIDADES DO CLIENTE ---
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

    # --- 4. ESCOPO TÉCNICO ---
    st.markdown("---")
    intro = st.text_area("Introdução do Escopo", value="Trata-se do fornecimento de materiais e mão de obra conforme itens abaixo:")

    escopo_estruturado = [] 
    eap_estruturada = []    
    valor_total_calculado = 0.0

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
                            
                            if "AMORTECEDOR" not in item_curto.upper():
                                lista_eap.append({'indice': f"{contador_cat}.{contador_item}", 'nome': item_curto})
                                contador_item += 1 
                            
                            texto_final = texto_base
                            if qtd > 1: texto_final += f" — Qtd: {qtd}."
                            lista_detalhada.append(texto_final)
                            
                        eap_estruturada.append({'indice': str(contador_cat), 'categoria': cat.upper(), 'itens': lista_eap})
                        escopo_estruturado.append({'nome': f"{contador_cat}. {cat.upper()}", 'itens': lista_detalhada})
                        contador_cat += 1

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

                for idx, row in df_orc.iterrows():
                    texto_col_e = str(row[4]).upper().strip() if 4 < len(row) else ""
                    texto_col_c = str(row[2]).upper().strip() if 2 < len(row) else ""
                    
                    if "PREÇO VENDA" in texto_col_e or "PRECO VENDA" in texto_col_e or "PREÇO VENDA" in texto_col_c or "PRECO VENDA" in texto_col_c:
                        val_bruto = row[9] 
                        if pd.notna(val_bruto):
                            if isinstance(val_bruto, (int, float)):
                                valor_total_calculado = float(val_bruto)
                            else:
                                val_str = str(val_bruto).upper().replace("R$", "").strip()
                                val_str = val_str.replace(".", "").replace(",", ".")
                                try: valor_total_calculado = float(val_str)
                                except: pass
                        break

                categoria_atual_nome = "ESCOPO GERAL"
                categoria_atual_indice = ""
                itens_detalhados = []
                itens_eap = []
                contador_item = 1
                
                for index, row in df_orc.iterrows():
                    descricao = str(row[2]).strip() if 2 < len(row) else ""
                    
                    if "CUSTO INDIRETO" in descricao.upper() or "CUSTOS INDIRETOS" in descricao.upper():
                        break 
                    
                    col_b = str(row[1]).strip() if 1 < len(row) else ""
                    unidade = str(row[3]).strip() if 3 < len(row) else ""
                    quantidade = row[4] if 4 < len(row) else pd.NA
                    
                    if pd.isna(descricao) or descricao == "" or descricao.upper() in ["NAN", "DESCRIÇÃO DOS MATERIAIS", "DESCRICAO DOS MATERIAIS", "DESCRIÇÃO", "ITEM"]:
                        continue
                        
                    is_header = False
                    if col_b and col_b.lower() != 'nan' and col_b != '-':
                        if col_b[0].isdigit():
                            is_header = True
                            
                    if is_header:
                        if categoria_atual_nome != "ESCOPO GERAL" and len(itens_detalhados) > 0:
                            escopo_estruturado.append({'nome': f"{categoria_atual_indice} - {categoria_atual_nome}".strip(' -'), 'itens': itens_detalhados})
                            eap_estruturada.append({'indice': categoria_atual_indice, 'categoria': categoria_atual_nome.upper(), 'itens': itens_eap})
                        
                        categoria_atual_indice = col_b
                        categoria_atual_nome = descricao
                        itens_detalhados = []
                        itens_eap = []
                        contador_item = 1
                        
                    else:
                        has_qty = not pd.isna(quantidade) and str(quantidade).strip() not in ["", "nan", "-"]
                        
                        if has_qty:
                            if isinstance(quantidade, (int, float)):
                                qtd_fmt = int(quantidade) if float(quantidade).is_integer() else round(float(quantidade), 2)
                            else:
                                try: 
                                    q_str = str(quantidade).replace(",", ".")
                                    qtd_fmt = int(float(q_str)) if float(q_str).is_integer() else round(float(q_str), 2)
                                except: 
                                    qtd_fmt = quantidade
                                    
                            uni_fmt = f" {unidade}" if unidade.lower() not in ["nan", "", "-"] else ""
                            nome_resumido = descricao.split('|')[0].strip()
                            if len(nome_resumido) > 80: nome_resumido = nome_resumido[:80] + "..."
                            
                            if "AMORTECEDOR" not in descricao.upper():
                                indice_item = f"{categoria_atual_indice}.{contador_item}" if categoria_atual_indice else str(contador_item)
                                itens_eap.append({'indice': indice_item, 'nome': nome_resumido})
                                contador_item += 1 
                            
                            texto_item = f"Fornecimento / Instalação de {qtd_fmt}{uni_fmt} - {descricao}."
                            itens_detalhados.append(texto_item)
                            
                        else:
                            if len(itens_detalhados) > 0:
                                itens_detalhados[-1] += f"\n\n{descricao}"
                
                if categoria_atual_nome != "ESCOPO GERAL" and len(itens_detalhados) > 0:
                    escopo_estruturado.append({'nome': f"{categoria_atual_indice} - {categoria_atual_nome}".strip(' -'), 'itens': itens_detalhados})
                    eap_estruturada.append({'indice': categoria_atual_indice, 'categoria': categoria_atual_nome.upper(), 'itens': itens_eap})
                    
                if len(escopo_estruturado) > 0:
                    st.success(f"✅ Planilha processada! EAP sem amortecedores e linhas espaçadas geradas com sucesso.")
                else:
                    st.warning(f"⚠️ A aba '{nome_aba}' foi lida, mas não encontrei itens válidos na Coluna B.")
                    
            except Exception as e:
                st.error(f"Erro ao processar a planilha: {e}")

    # --- 5. EXCLUSÕES ---
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

    # --- 6. COMERCIAL ---
    st.markdown("---")
    st.header("6. Comercial")

    valor_formatado_sugerido = f"R$ {valor_total_calculado:_.2f}".replace('.', ',').replace('_', '.')

    c_v, c_m = st.columns(2)
    valor = c_v.text_input("Valor Total (R$):", value=valor_formatado_sugerido if valor_total_calculado > 0 else "")
    mes = c_m.text_input("Mês/Ano Base", value=f"{hoje.month}/{hoje.year}")

    # --- BOTÃO GERAR PROPOSTA ---
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
            'eap_estruturada': eap_estruturada, 
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
            st.error(f"Erro ao gerar o Word: {e}")


# ==============================================================================
# MÓDULO 2: ESTIMATIVA DE CUSTOS (NOVO & INTELIGENTE)
# ==============================================================================
elif menu_selecionado == "💰 Estimativa de Custos":
    
    st.title("💰 Engenharia e Custos - Automação e Infra")
    
    # NOVO CAMPO PARA O NOME DO PROJETO / HISTÓRICO
    nome_projeto_orcamento = st.text_input("🏷️ Nome do Projeto / Cliente (Para salvar no Histórico):", placeholder="Ex: Reforma UTA Siemens - Prédio 2")
    st.markdown("---")

    PLANILHA_NOME = "DB_Propostas_Siarcon" 

    @st.cache_resource
    def conectar_google_sheets():
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open(PLANILHA_NOME)

    # ==========================================
    # 1. INICIALIZAÇÃO E BASE DE DADOS (COM INTEGRAÇÃO GSHEETS)
    # ==========================================
    banco_padrao_precos = {
        "Transmissor de pressão Dif. Para ar (Vazão de ar) (PDIT)": 1490.00,
        "Transmissor de temperatura (TT) (Controle)": 800.00,
        "Transmissor de temperatura e umidade (TT/MT ou TMT) (Controle)": 2050.00,
        "Resistência aquecimento (RAQ) (Equipamento)": 0.0,
        "Resistência de aquecimento (RAQ) (Duto)": 0.0,
        "Válvula de água gelada (TCV)": 2650.00,
        "Válvula de água quente (TCV)": 3210.00,
        "Válvula de vapor (TCV)": 0.0,
        "Transmissor pressão para filtro G4 (PDIT)": 1490.00,
        "Transmissor pressão Filtro F9 (PDIT)": 1490.00,
        "Transmissor pressão filtro H13 (PDIT)": 1490.00,
        "Pressostato para filtro G4 (PSH)": 349.00,
        "Pressostato Filtro F9 (PSH)": 349.00,
        "Pressostato filtro H13 (PSH)": 349.00,
        "Transmissor de pressão diferencial (Pressão entre salas) (PDT)": 1490.00,
        "Transmissor de pressão diferencial com display (Pressão entre salas) (PDIT)": 2110.00,
        "Transmissor de temperatura (TT) (Ambiente)": 2050.00,
        "Transmissor de temperatura com display (TIT) (Ambiente)": 2650.00,
        "Transmissor de temperatura e umidade (TT/MT ou TMT) (Ambiente)": 2050.00,
        "Transmissor de temperatura e umidade com display (TT/MT ou TMT) (Ambiente)": 2650.00,
        "MP-C-15A": 4649.49,
        "MP-C-18A": 5185.54,
        "MP-C-24A": 7290.75,
        "MP-C-36A": 9459.08,
        "Custo AI/AO": 565.00,
        "Custo DI/DO": 120.00
    }

    # Carrega os preços do banco de dados online na inicialização
    if 'banco_precos_carregado' not in st.session_state:
        st.session_state.precos_banco = banco_padrao_precos.copy()
        st.session_state.historico_precos = []
        
        try:
            sh = conectar_google_sheets()
            # Puxa os preços
            try:
                df_p = pd.DataFrame(sh.worksheet("Precos").get_all_records())
                if not df_p.empty:
                    precos_bd = dict(zip(df_p['Item'], df_p['Valor']))
                    st.session_state.precos_banco.update(precos_bd)
            except: pass
            
            # Puxa o histórico
            try:
                df_h = pd.DataFrame(sh.worksheet("Historico_Precos").get_all_records())
                if not df_h.empty:
                    st.session_state.historico_precos = df_h.to_dict('records')
            except: pass
            
        except Exception as e:
            st.toast("Aviso: Iniciando com os preços padrão (Google Sheets não sincronizou)", icon="⚠️")
            
        st.session_state.banco_precos_carregado = True

    # Validação de Estrutura Antiga na memória (Reset se necessário)
    if 'paineis_auto' not in st.session_state or (len(st.session_state.paineis_auto) > 0 and 'grupos_equipamentos' not in st.session_state.paineis_auto[0]):
        st.session_state.paineis_auto = []
    
    if len(st.session_state.paineis_auto) > 0:
        if "Transmissor de pressão Dif. Para ar (Vazão de ar) (PDIT)" not in st.session_state.paineis_auto[0]['grupos_equipamentos'][0]['instrumentos']:
             st.session_state.paineis_auto = []

    if 'orcamento' not in st.session_state: st.session_state.orcamento = []

    # ==========================================
    # 2. DICIONÁRIO DE REGRAS DE ENGENHARIA
    # ==========================================
    GRUPOS_INSTRUMENTOS = {
        "🔹 Equipamento (Controle)": [
            "Transmissor de pressão Dif. Para ar (Vazão de ar) (PDIT)",
            "Transmissor de temperatura (TT) (Controle)",
            "Transmissor de temperatura e umidade (TT/MT ou TMT) (Controle)",
            "Resistência aquecimento (RAQ) (Equipamento)",
            "Resistência de aquecimento (RAQ) (Duto)",
            "Válvula de água gelada (TCV)",
            "Válvula de água quente (TCV)",
            "Válvula de vapor (TCV)"
        ],
        "🔸 Monitoramento (Equipamento)": [
            "Transmissor pressão para filtro G4 (PDIT)",
            "Transmissor pressão Filtro F9 (PDIT)",
            "Transmissor pressão filtro H13 (PDIT)",
            "Pressostato para filtro G4 (PSH)",
            "Pressostato Filtro F9 (PSH)",
            "Pressostato filtro H13 (PSH)"
        ],
        "🟢 Monitoramento (Ambiente)": [
            "Transmissor de pressão diferencial (Pressão entre salas) (PDT)",
            "Transmissor de pressão diferencial com display (Pressão entre salas) (PDIT)",
            "Transmissor de temperatura (TT) (Ambiente)",
            "Transmissor de temperatura com display (TIT) (Ambiente)",
            "Transmissor de temperatura e umidade (TT/MT ou TMT) (Ambiente)",
            "Transmissor de temperatura e umidade com display (TT/MT ou TMT) (Ambiente)"
        ]
    }

    REGRA_IO = {
        # Controle
        "Transmissor de pressão Dif. Para ar (Vazão de ar) (PDIT)": {"AI": 1, "AO": 1, "DI": 1, "DO": 1},
        "Transmissor de temperatura (TT) (Controle)": {"AI": 1, "AO": 1, "DI": 0, "DO": 0},
        "Transmissor de temperatura e umidade (TT/MT ou TMT) (Controle)": {"AI": 2, "AO": 2, "DI": 0, "DO": 0},
        "Resistência aquecimento (RAQ) (Equipamento)": {"AI": 0, "AO": 1, "DI": 2, "DO": 1},
        "Resistência de aquecimento (RAQ) (Duto)": {"AI": 0, "AO": 1, "DI": 2, "DO": 1},
        "Válvula de água gelada (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula de água quente (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},
        "Válvula de vapor (TCV)": {"AI": 0, "AO": 1, "DI": 0, "DO": 0},

        # Monitoramento (Equipamento)
        "Transmissor pressão para filtro G4 (PDIT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor pressão Filtro F9 (PDIT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor pressão filtro H13 (PDIT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Pressostato para filtro G4 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Pressostato Filtro F9 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},
        "Pressostato filtro H13 (PSH)": {"AI": 0, "AO": 0, "DI": 1, "DO": 0},

        # Monitoramento (Ambiente)
        "Transmissor de pressão diferencial (Pressão entre salas) (PDT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de pressão diferencial com display (Pressão entre salas) (PDIT)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura (TT) (Ambiente)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura com display (TIT) (Ambiente)": {"AI": 1, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura e umidade (TT/MT ou TMT) (Ambiente)": {"AI": 2, "AO": 0, "DI": 0, "DO": 0},
        "Transmissor de temperatura e umidade com display (TT/MT ou TMT) (Ambiente)": {"AI": 2, "AO": 0, "DI": 0, "DO": 0},
    }
    
    PRECOS_IHM = {
        "Sem IHM": 0.0,
        "IHM 4,3 polegadas": 1700.00,
        "IHM 7 polegadas": 3400.00,
        "IHM 10 polegadas": 8500.00
    }

    def calcular_painel_fisico(qtd_controladores):
        if qtd_controladores == 0: return "Sem Painel Físico", 0.0
        elif qtd_controladores <= 2: return "Painel 600x400mm", 4500.00
        elif qtd_controladores <= 5: return "Painel 800x600mm", 5900.00
        else: return "Painel 1200x600mm", 9250.00

    def dimensionar_controladores(total_io):
        c36 = c24 = c18 = c15 = 0
        rem = total_io
        while rem > 0:
            if rem > 24: c36 += 1; rem -= 36
            elif rem > 18: c24 += 1; rem -= 24
            elif rem > 15: c18 += 1; rem -= 18
            else: c15 += 1; rem -= 15
        return c36, c24, c18, c15

    # ==========================================
    # 3. INTERFACE DE ABAS
    # ==========================================
    aba_auto, aba_planilhas, aba_infra, aba_precos, aba_resumo = st.tabs([
        "🚀 Dimensionamento Automático", 
        "🛠️ Modo Planilhas Antigas", 
        "🔌 Infraestrutura", 
        "💲 Base de Preços",
        "📊 Resumo Final"
    ])

    with aba_auto:
        st.header("Motor de Dimensionamento SIARCON")
        st.markdown("Crie um painel principal e adicione **múltiplos grupos de equipamentos** dentro dele. O sistema calculará a otimização dos controladores em um único quadro.")
        
        if st.button("➕ Adicionar Novo Painel Físico"):
            st.session_state.paineis_auto.append({
                "id": len(st.session_state.paineis_auto),
                "nome": f"Quadro Automação {len(st.session_state.paineis_auto) + 1}",
                "ihm": "Sem IHM",
                "grupos_equipamentos": [
                    {
                        "nome_grupo": "Equipamento 1 (ex: UTA 01)",
                        "multiplicador": 1,
                        "instrumentos": {k: 0 for k in REGRA_IO.keys()}
                    }
                ]
            })

        # Iteração de cada PAINEL
        for p_idx, p_data in enumerate(st.session_state.paineis_auto):
            with st.expander(f"📦 {p_data['nome']}", expanded=True):
                
                # Configurações do Painel
                c_nome_painel, c_ihm_painel = st.columns([2, 1])
                p_data['nome'] = c_nome_painel.text_input("Identificação do Quadro Físico", value=p_data['nome'], key=f"n_p_{p_idx}")
                p_data['ihm'] = c_ihm_painel.selectbox("IHM Geral do Quadro", list(PRECOS_IHM.keys()), index=list(PRECOS_IHM.keys()).index(p_data['ihm']), key=f"i_p_{p_idx}")
                
                st.markdown("---")
                
                # Variáveis de Somatório do Painel Inteiro
                total_ai_painel = total_ao_painel = total_di_painel = total_do_painel = 0

                # Iteração de cada GRUPO DENTRO DO PAINEL
                for g_idx, g_data in enumerate(p_data['grupos_equipamentos']):
                    st.markdown(f"#### ⚙️ Grupo: {g_data['nome_grupo']}")
                    
                    cg_nome, cg_mult = st.columns([3, 1])
                    g_data['nome_grupo'] = cg_nome.text_input("Nome do Equipamento/Ambiente", value=g_data['nome_grupo'], key=f"n_g_{p_idx}_{g_idx}")
                    g_data['multiplicador'] = cg_mult.number_input("Qtd. de Equipamentos Iguais", min_value=1, value=g_data.get('multiplicador', 1), key=f"m_g_{p_idx}_{g_idx}")
                    
                    st.markdown("*(Selecione os instrumentos para APENAS 1 unidade deste grupo)*")
                    
                    # Variáveis do Grupo Específico
                    total_ai_grupo = total_ao_grupo = total_di_grupo = total_do_grupo = 0
                    
                    # Layout 3 colunas
                    for grupo_nome, lista_itens in GRUPOS_INSTRUMENTOS.items():
                        st.markdown(f"**{grupo_nome}**")
                        cols = st.columns(3)
                        
                        for i, inst in enumerate(lista_itens):
                            if inst not in g_data['instrumentos']:
                                g_data['instrumentos'][inst] = 0
                                
                            with cols[i % 3]:
                                qtd = st.number_input(inst, min_value=0, step=1, value=g_data['instrumentos'][inst], key=f"inst_{p_idx}_{g_idx}_{inst}")
                                g_data['instrumentos'][inst] = qtd
                                
                                # Soma I/O do grupo
                                total_ai_grupo += qtd * REGRA_IO[inst]["AI"]
                                total_ao_grupo += qtd * REGRA_IO[inst]["AO"]
                                total_di_grupo += qtd * REGRA_IO[inst]["DI"]
                                total_do_grupo += qtd * REGRA_IO[inst]["DO"]
                        st.write("") # Espaçamento
                    
                    st.divider() # Linha divisória após o grupo
                    
                    # Adiciona os totais do Grupo (multiplicados) ao somatório do Painel
                    mult = g_data['multiplicador']
                    total_ai_painel += total_ai_grupo * mult
                    total_ao_painel += total_ao_grupo * mult
                    total_di_painel += total_di_grupo * mult
                    total_do_painel += total_do_grupo * mult
                
                # Botão para adicionar mais equipamentos NO MESMO PAINEL
                if st.button(f"➕ Adicionar Equipamento Diferente neste Quadro", key=f"add_grp_{p_idx}"):
                    p_data['grupos_equipamentos'].append({
                        "nome_grupo": f"Equipamento {len(p_data['grupos_equipamentos']) + 1}",
                        "multiplicador": 1,
                        "instrumentos": {k: 0 for k in REGRA_IO.keys()}
                    })
                    st.rerun()

                # ---- FECHAMENTO E CÁLCULO GERAL DO PAINEL ----
                total_io_pontos = total_ai_painel + total_ao_painel + total_di_painel + total_do_painel
                c36, c24, c18, c15 = dimensionar_controladores(total_io_pontos)
                total_controladores = c36 + c24 + c18 + c15
                nome_caixa, preco_caixa = calcular_painel_fisico(total_controladores)

                st.markdown(f"### 📊 Resumo do {p_data['nome']}")
                res_c1, res_c2, res_c3 = st.columns(3)
                
                res_c1.info(f"**Pontos Físicos Totais ( {total_io_pontos} )**\n\nAI: {total_ai_painel} | AO: {total_ao_painel}\nDI: {total_di_painel} | DO: {total_do_painel}")
                
                txt_controladores = ""
                if c36 > 0: txt_controladores += f"• {c36}x MP-C-36A\n"
                if c24 > 0: txt_controladores += f"• {c24}x MP-C-24A\n"
                if c18 > 0: txt_controladores += f"• {c18}x MP-C-18A\n"
                if c15 > 0: txt_controladores += f"• {c15}x MP-C-15A\n"
                if txt_controladores == "": txt_controladores = "Nenhum I/O configurado."
                
                res_c2.success(f"**Controladores Otimizados**\n\n{txt_controladores}")
                res_c3.warning(f"**Estrutura Centralizada**\n\n• 1x {nome_caixa}\n• 1x {p_data['ihm']}")

    with aba_precos:
        st.header("Gestão da Base de Preços")
        st.markdown("Atualize os valores nesta tabela. O sistema usará esses preços para todos os cálculos de orçamentos e salvará o histórico diretamente no **Google Sheets**.")
        
        df_precos = pd.DataFrame(list(st.session_state.precos_banco.items()), columns=["Item / Equipamento", "Valor Atual (R$)"])
        edited_df = st.data_editor(df_precos, use_container_width=True, hide_index=True)
        
        if st.button("💾 Salvar Novos Preços no Banco de Dados", type="primary"):
            alterou_algo = False
            novos_historicos = []
            
            for idx, row in edited_df.iterrows():
                item = row['Item / Equipamento']
                novo_valor = row['Valor Atual (R$)']
                antigo_valor = st.session_state.precos_banco.get(item, 0.0)
                
                if novo_valor != antigo_valor:
                    novo_hist = {
                        "Data/Hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                        "Item Alterado": item,
                        "Valor Antigo": f"R$ {antigo_valor:.2f}",
                        "Novo Valor": f"R$ {novo_valor:.2f}"
                    }
                    st.session_state.historico_precos.append(novo_hist)
                    novos_historicos.append(novo_hist)
                    st.session_state.precos_banco[item] = novo_valor
                    alterou_algo = True
            
            if alterou_algo:
                # Rotina para Salvar no Banco de Dados Real com Auto-Criação de Abas
                try:
                    sh = conectar_google_sheets()
                    
                    # 1. Atualiza ou Cria a aba Precos
                    try:
                        ws_precos = sh.worksheet("Precos")
                    except:
                        ws_precos = sh.add_worksheet(title="Precos", rows="100", cols="2")
                        ws_precos.append_row(["Item", "Valor"])
                        
                    ws_precos.clear()
                    dados_precos_matriz = [["Item", "Valor"]] + [[k, v] for k, v in st.session_state.precos_banco.items()]
                    ws_precos.append_rows(dados_precos_matriz)
                    
                    # 2. Atualiza ou Cria a aba Historico_Precos
                    try:
                        ws_hist = sh.worksheet("Historico_Precos")
                    except:
                        ws_hist = sh.add_worksheet(title="Historico_Precos", rows="1000", cols="4")
                        ws_hist.append_row(["Data/Hora", "Item Alterado", "Valor Antigo", "Novo Valor"])
                        
                    linhas_h = []
                    for h in novos_historicos:
                        linhas_h.append([h["Data/Hora"], h["Item Alterado"], h["Valor Antigo"], h["Novo Valor"]])
                    if linhas_h:
                        ws_hist.append_rows(linhas_h)
                        
                    st.success("✅ Preços atualizados e gravados na nuvem com sucesso!")
                except Exception as e:
                    st.error(f"⚠️ Os preços foram atualizados nesta tela, mas houve um erro crítico ao acessar o banco. Erro: {e}.")
            else:
                st.info("Nenhuma alteração foi feita na tabela.")

        st.markdown("---")
        st.subheader("Histórico Geral de Atualizações de Preços")
        if st.session_state.historico_precos:
            st.dataframe(pd.DataFrame(st.session_state.historico_precos)[::-1], use_container_width=True, hide_index=True)
        else:
            st.write("Sem registros de alterações no banco.")

    with aba_planilhas:
        st.header("Leitura das Planilhas Antigas")
        st.markdown("Módulo mantido para compatibilidade com arquivos de csv antigos.")
        
        def converter_valor_plan(val):
            try:
                v = str(val).upper().replace('R$', '').strip()
                if v in ['NAN', 'NONE', '', '-']: return 0.0
                if ',' in v: v = v.replace('.', '').replace(',', '.')
                return float(v)
            except:
                return 0.0

        @st.cache_data
        def carregar_dados_planilhas():
            diretorio_atual = os.getcwd()
            pasta_dados = os.path.join(diretorio_atual, "dados")
            arquivos_na_pasta = os.listdir(pasta_dados) if os.path.exists(pasta_dados) else []
            
            nome_cag = next((f for f in arquivos_na_pasta if "CAG" in f.upper()), "CAG.csv")
            nome_ahu = next((f for f in arquivos_na_pasta if "AHU" in f.upper()), "AHU01.csv")
            nome_infra = next((f for f in arquivos_na_pasta if "INFRA" in f.upper()), "Infra.csv")
            
            cag_path = os.path.join(pasta_dados, nome_cag)
            ahu_path = os.path.join(pasta_dados, nome_ahu)
            infra_path = os.path.join(pasta_dados, nome_infra)

            def ler_csv_blindado(caminho, palavra_chave):
                try:
                    df = pd.read_csv(caminho, sep=',', header=None, dtype=str)
                    if len(df.columns) <= 2: df = pd.read_csv(caminho, sep=';', header=None, dtype=str)
                except Exception:
                    df = pd.read_csv(caminho, sep=';', header=None, dtype=str)
                    
                header_idx = 0
                for i, row in df.iterrows():
                    linha_texto = " ".join([str(x).upper() for x in row.values])
                    if palavra_chave in linha_texto:
                        header_idx = i
                        break
                        
                raw_cols = df.iloc[header_idx].astype(str).tolist()
                clean_cols = []
                for j, col in enumerate(raw_cols):
                    c = col.strip().upper()
                    if c in ['NAN', '', 'NONE', 'UNNAMED']: clean_cols.append(f"VAZIA_{j}")
                    elif c in clean_cols: clean_cols.append(f"{c}_{j}")
                    else: clean_cols.append(c)
                        
                df.columns = clean_cols
                df = df.iloc[header_idx+1:].reset_index(drop=True)
                return df

            try:
                cag_df = ler_csv_blindado(cag_path, "ITEM")
                ahu_df = ler_csv_blindado(ahu_path, "ITEM")
                infra_df = ler_csv_blindado(infra_path, "INSTRUMENTAÇÃO")
                
                if 'ITEM' in cag_df.columns:
                    cag_df = cag_df.dropna(subset=['ITEM'])
                    cag_df = cag_df[cag_df['ITEM'].astype(str).str.strip() != 'NAN']
                    
                if 'ITEM' in ahu_df.columns:
                    ahu_df = ahu_df.dropna(subset=['ITEM'])
                    ahu_df = ahu_df[ahu_df['ITEM'].astype(str).str.strip() != 'NAN']
                    
                if 'INSTRUMENTAÇÃO' in infra_df.columns:
                    infra_df = infra_df.dropna(subset=['INSTRUMENTAÇÃO'])
                    infra_df = infra_df[infra_df['INSTRUMENTAÇÃO'].astype(str).str.strip() != 'NAN']
                
                return cag_df, ahu_df, infra_df
            except Exception:
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        cag_df, ahu_df, infra_df = carregar_dados_planilhas()
        
        if not cag_df.empty and not ahu_df.empty:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("CAG (Planilha)")
                for index, row in cag_df.iterrows():
                    c1, c2 = st.columns([3, 1])
                    item_nome = str(row.get('ITEM', 'Item Desconhecido'))
                    valor_unit = converter_valor_plan(row.get('VALOR UNITÁRIO', 0))
                    c1.write(f"{item_nome} (R$ {valor_unit:.2f})")
                    qtd = c2.number_input(f"Qtd", min_value=0, value=0, key=f"cag_plan_{index}")
                    if qtd > 0:
                        st.session_state.orcamento.append({"Categoria": "Manual - Planilha", "Item": item_nome, "Quantidade": qtd, "Custo_Total": qtd * valor_unit})
                        
            with col2:
                st.subheader("AHU01 (Planilha)")
                for index, row in ahu_df.iterrows():
                    c1, c2 = st.columns([3, 1])
                    item_nome = str(row.get('ITEM', 'Item Desconhecido'))
                    valor_unit = converter_valor_plan(row.get('VALOR UNITÁRIO', 0))
                    c1.write(f"{item_nome} (R$ {valor_unit:.2f})")
                    qtd = c2.number_input(f"Qtd", min_value=0, value=0, key=f"ahu_plan_{index}")
                    if qtd > 0:
                        st.session_state.orcamento.append({"Categoria": "Manual - Planilha", "Item": item_nome, "Quantidade": qtd, "Custo_Total": qtd * valor_unit})

    with aba_infra:
        st.header("Cálculo de Infraestrutura")
        st.markdown("Insira a distância média para calcular cabos e eletrocalhas/tubulações.")
        
        if not infra_df.empty and 'INSTRUMENTAÇÃO' in infra_df.columns:
            for index, row in infra_df.iterrows():
                tipo_inst = str(row['INSTRUMENTAÇÃO']).strip()
                custo_cabo = converter_valor_plan(row.iloc[6]) if len(row) > 6 else 0.0
                custo_infra = converter_valor_plan(row.get('INFRA', 0))
                
                if tipo_inst in ["", "NAN", "NONE", "-", "EQUIPAMENTOS", "TOTAL EQUIPAMENTOS"]: continue

                st.write(f"**{tipo_inst}** (Cabo: R${custo_cabo:.2f}/m | Infra: R${custo_infra:.2f}/m)")
                c1, c2 = st.columns(2)
                qtd_inst = c1.number_input("Qtd. de Instrumentos", min_value=0, value=0, key=f"infra_qtd_{index}")
                dist_media = c2.number_input("Distância Média (m)", min_value=0.0, value=0.0, step=1.0, key=f"infra_dist_{index}")
                
                if qtd_inst > 0 and dist_media > 0:
                    metragem = qtd_inst * dist_media
                    st.session_state.orcamento.append({"Categoria": "Infraestrutura", "Item": f"Cabo ({tipo_inst})", "Quantidade": metragem, "Custo_Total": metragem * custo_cabo})
                    st.session_state.orcamento.append({"Categoria": "Infraestrutura", "Item": f"Infra ({tipo_inst})", "Quantidade": metragem, "Custo_Total": metragem * custo_infra})

    with aba_resumo:
        st.header("Consolidação Financeira do Orçamento")
        
        linhas_resumo = []
        linhas_pontos = []

        # 1. Processando Custos do Motor Automático (Hierarquia: Painel > Grupos)
        for p in st.session_state.paineis_auto:
            
            total_ai_painel = total_ao_painel = total_di_painel = total_do_painel = 0
            
            for g in p['grupos_equipamentos']:
                mult = g.get('multiplicador', 1)
                
                for inst, qtd in g['instrumentos'].items():
                    if qtd > 0:
                        qtd_final = qtd * mult
                        preco_item = st.session_state.precos_banco.get(inst, 0.0)
                        
                        total_ai_painel += qtd_final * REGRA_IO[inst]["AI"]
                        total_ao_painel += qtd_final * REGRA_IO[inst]["AO"]
                        total_di_painel += qtd_final * REGRA_IO[inst]["DI"]
                        total_do_painel += qtd_final * REGRA_IO[inst]["DO"]
                        
                        linhas_resumo.append({
                            "Categoria": f"{p['nome']} - Campo", 
                            "Item": f"{inst} ({g['nome_grupo']})", 
                            "Qtd": qtd_final, 
                            "Custo_Total": qtd_final * preco_item
                        })
                        
                        linhas_pontos.append({
                            "Painel": p['nome'],
                            "Grupo/Equipamento": g['nome_grupo'],
                            "Instrumento": inst,
                            "Quantidade Total": qtd_final,
                            "Entrada Digital (DI)": qtd_final * REGRA_IO[inst]["DI"],
                            "Saída Digital (DO)": qtd_final * REGRA_IO[inst]["DO"],
                            "Entrada Analógica (AI)": qtd_final * REGRA_IO[inst]["AI"],
                            "Saída Analógica (AO)": qtd_final * REGRA_IO[inst]["AO"]
                        })

            tot_io_painel = total_ai_painel + total_ao_painel + total_di_painel + total_do_painel

            if tot_io_painel > 0:
                custo_ana = (total_ai_painel + total_ao_painel) * st.session_state.precos_banco["Custo AI/AO"]
                custo_dig = (total_di_painel + total_do_painel) * st.session_state.precos_banco["Custo DI/DO"]
                linhas_resumo.append({"Categoria": f"{p['nome']} - I/Os", "Item": "Pontos Analógicos (AI/AO)", "Qtd": (total_ai_painel + total_ao_painel), "Custo_Total": custo_ana})
                linhas_resumo.append({"Categoria": f"{p['nome']} - I/Os", "Item": "Pontos Digitais (DI/DO)", "Qtd": (total_di_painel + total_do_painel), "Custo_Total": custo_dig})
                
                c36, c24, c18, c15 = dimensionar_controladores(tot_io_painel)
                if c36 > 0: linhas_resumo.append({"Categoria": f"{p['nome']} - MPC", "Item": "Controlador MP-C-36A", "Qtd": c36, "Custo_Total": c36 * st.session_state.precos_banco["MP-C-36A"]})
                if c24 > 0: linhas_resumo.append({"Categoria": f"{p['nome']} - MPC", "Item": "Controlador MP-C-24A", "Qtd": c24, "Custo_Total": c24 * st.session_state.precos_banco["MP-C-24A"]})
                if c18 > 0: linhas_resumo.append({"Categoria": f"{p['nome']} - MPC", "Item": "Controlador MP-C-18A", "Qtd": c18, "Custo_Total": c18 * st.session_state.precos_banco["MP-C-18A"]})
                if c15 > 0: linhas_resumo.append({"Categoria": f"{p['nome']} - MPC", "Item": "Controlador MP-C-15A", "Qtd": c15, "Custo_Total": c15 * st.session_state.precos_banco["MP-C-15A"]})
                
                nome_caixa, preco_caixa = calcular_painel_fisico(c36 + c24 + c18 + c15)
                linhas_resumo.append({"Categoria": f"{p['nome']} - Estrutura Fís.", "Item": nome_caixa, "Qtd": 1, "Custo_Total": preco_caixa})
                
                if PRECOS_IHM[p['ihm']] > 0:
                    linhas_resumo.append({"Categoria": f"{p['nome']} - Estrutura Fís.", "Item": p['ihm'], "Qtd": 1, "Custo_Total": PRECOS_IHM[p['ihm']]})

        for item in st.session_state.orcamento:
            linhas_resumo.append({"Categoria": item['Categoria'], "Item": item['Item'], "Qtd": item['Quantidade'], "Custo_Total": item['Custo_Total']})

        # Renderizando a Tabela Final e Gerando Arquivo Excel
        if len(linhas_resumo) > 0:
            df_final = pd.DataFrame(linhas_resumo)
            df_agrupado = df_final.groupby(['Categoria', 'Item'], as_index=False).agg({'Qtd': 'sum', 'Custo_Total': 'sum'})
            
            subtotal_materiais = df_agrupado['Custo_Total'].sum()
            custo_servicos_logica = subtotal_materiais * 0.25  
            total_projeto = subtotal_materiais + custo_servicos_logica
            
            df_servicos = pd.DataFrame([{'Categoria': 'Serviços / Mão de Obra', 'Item': 'Serviços de Lógica (25%)', 'Qtd': 1, 'Custo_Total': custo_servicos_logica}])
            df_total = pd.DataFrame([{'Categoria': 'TOTAL GERAL', 'Item': 'Custo Total Estimado', 'Qtd': '-', 'Custo_Total': total_projeto}])
            df_exportacao = pd.concat([df_agrupado, df_servicos, df_total], ignore_index=True)
            
            st.dataframe(df_agrupado, use_container_width=True)
            
            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            c1.info(f"**Subtotal Materiais/Hardware:**\nR$ {subtotal_materiais:,.2f}")
            c2.warning(f"**Serviços de Lógica (25%):**\nR$ {custo_servicos_logica:,.2f}")
            c3.success(f"**CUSTO TOTAL ESTIMADO:**\nR$ {total_projeto:,.2f}")
            
            # --- Gerar a Tabela de Pontos de I/O ---
            df_pontos = pd.DataFrame(linhas_pontos)
            if not df_pontos.empty:
                total_di = df_pontos['Entrada Digital (DI)'].sum()
                total_do = df_pontos['Saída Digital (DO)'].sum()
                total_ai = df_pontos['Entrada Analógica (AI)'].sum()
                total_ao = df_pontos['Saída Analógica (AO)'].sum()
                total_qtd = df_pontos['Quantidade Total'].sum()
                
                linha_total = pd.DataFrame([{
                    "Painel": "TOTAL GERAL",
                    "Grupo/Equipamento": "-",
                    "Instrumento": "-",
                    "Quantidade Total": total_qtd,
                    "Entrada Digital (DI)": total_di,
                    "Saída Digital (DO)": total_do,
                    "Entrada Analógica (AI)": total_ai,
                    "Saída Analógica (AO)": total_ao
                }])
                df_pontos = pd.concat([df_pontos, linha_total], ignore_index=True)

            # --- Criação do Arquivo Excel Final ---
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_exportacao.to_excel(writer, index=False, sheet_name='Detalhamento Financeiro')
                if not df_pontos.empty:
                    df_pontos.to_excel(writer, index=False, sheet_name='Matriz de Pontos (IO)')
            
            st.download_button(label="📥 Exportar Orçamento Final para Excel", data=buffer.getvalue(), file_name="orcamento_dimensionado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            st.markdown("---")
            
            # --- NOVO BOTÃO DE SALVAR COM AUTO-CRIAÇÃO DE ABA ---
            if st.button("☁️ Salvar Levantamento no Banco de Dados", type="secondary", use_container_width=True):
                if not nome_projeto_orcamento:
                    st.warning("⚠️ Atenção: Preencha o 'Nome do Projeto / Cliente' lá no topo da página antes de salvar.")
                else:
                    try:
                        sh = conectar_google_sheets()
                        
                        # Tenta achar a aba, se não achar, cria na hora
                        try:
                            ws_hist_orc = sh.worksheet("Historico_Orcamentos")
                        except:
                            ws_hist_orc = sh.add_worksheet(title="Historico_Orcamentos", rows="1000", cols="5")
                            ws_hist_orc.append_row(["Data/Hora", "Nome do Projeto", "Subtotal Hardware", "Serviços de Lógica", "Custo Total Estimado"])
                        
                        agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        nova_linha_banco = [
                            agora,
                            nome_projeto_orcamento,
                            f"R$ {subtotal_materiais:.2f}".replace('.', ','),
                            f"R$ {custo_servicos_logica:.2f}".replace('.', ','),
                            f"R$ {total_projeto:.2f}".replace('.', ',')
                        ]
                        ws_hist_orc.append_row(nova_linha_banco)
                        st.success(f"✅ Orçamento para '{nome_projeto_orcamento}' salvo com sucesso no Google Sheets!")
                    except Exception as e:
                        st.error(f"Erro ao salvar no banco. Detalhe técnico: {e}")

            # --- VISUALIZAR HISTÓRICO DENTRO DO APP ---
            with st.expander("📂 Ver Histórico de Levantamentos Salvos no Banco"):
                try:
                    sh = conectar_google_sheets()
                    df_hist_orc = pd.DataFrame(sh.worksheet("Historico_Orcamentos").get_all_records())
                    if not df_hist_orc.empty:
                        st.dataframe(df_hist_orc.iloc[::-1], use_container_width=True, hide_index=True)
                    else:
                        st.write("Nenhum levantamento salvo ainda.")
                except:
                    st.write("Nenhum levantamento foi salvo ou a aba ainda não existe.")

        else:
            st.info("Adicione painéis na aba 'Dimensionamento Automático' ou itens de Infraestrutura para visualizar o orçamento final.")

        # Limpa o orçamento infra/manual a cada ciclo
        st.session_state.orcamento = []
