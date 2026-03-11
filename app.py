import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
from datetime import date
import gspread
from oauth2client.service_account import ServiceAccountCredentials

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
    
    # Submenu específico da proposta na barra lateral
    st.sidebar.subheader("Opções da Proposta")
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
                            
                            if "AMORTECEDOR" not in item_curto.upper():
                                lista_eap.append({'indice': f"{contador_cat}.{contador_item}", 'nome': item_curto})
                                contador_item += 1 
                            
                            texto_final = texto_base
                            if qtd > 1: texto_final += f" — Qtd: {qtd}."
                            lista_detalhada.append(texto_final)
                            
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
    valor = c_v.text_input("Valor Total (R$):", value=valor_formatado_sugerido if valor_total_calculado > 0 else "")
    mes = c_m.text_input("Mês/Ano Base", value=f"{hoje.month}/{hoje.year}")

    # ==============================================================================
    # BOTÃO GERAR PROPOSTA
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
# MÓDULO 2: ESTIMATIVA DE CUSTOS (NOVO)
# ==============================================================================
elif menu_selecionado == "💰 Estimativa de Custos":
    
    st.title("💰 Estimativa de Custos - Automação e Infra")
    st.markdown("Selecione os equipamentos e infraestrutura para gerar a estimativa final.")

    # Inicializa o orçamento no estado da sessão (se não existir)
    if 'orcamento' not in st.session_state:
        st.session_state.orcamento = []

    # Regras de Custos Manuais (Sidebar)
    st.sidebar.header("⚙️ Regras e Taxas")
    taxa_mo = st.sidebar.number_input("Custo Mão de Obra (MO)", value=74.02, step=1.0)
    taxa_cabo_infra = st.sidebar.number_input("Custo Cabo + Infra", value=164.50, step=1.0)

    # Carregando as bases de custo do seu GitHub
    @st.cache_data
    def carregar_dados_custos():
        # Lembre-se de certificar que os arquivos estão na pasta 'dados' no GitHub com esses nomes
        cag_df = pd.read_csv("dados/CAG.csv", skiprows=3) 
        ahu_df = pd.read_csv("dados/AHU01.csv", skiprows=3)
        infra_df = pd.read_csv("dados/Infra.csv", skiprows=4)
        
        # Limpa linhas vazias baseadas na coluna ITEM para os equipamentos
        cag_df = cag_df.dropna(subset=['ITEM'])
        ahu_df = ahu_df.dropna(subset=['ITEM'])
        
        # Limpa linhas vazias baseadas na coluna de Instrumentação para a Infraestrutura
        # O .iloc[:, 0] pega a primeira coluna independentemente do nome (que deve ser a INSTRUMENTAÇÃO)
        infra_df = infra_df.dropna(subset=[infra_df.columns[0]])
        
        return cag_df, ahu_df, infra_df

    try:
        cag_df, ahu_df, infra_df = carregar_dados_custos()
        
        # Interface Principal (Abas)
        aba_equip, aba_infra, aba_resumo = st.tabs(["🛠️ Equipamentos", "🔌 Infraestrutura", "📊 Resumo e Exportação"])

        with aba_equip:
            st.header("Seleção de Instrumentos por Equipamento")
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("CAG")
                for index, row in cag_df.iterrows():
                    c1, c2 = st.columns([3, 1])
                    item_nome = row['ITEM']
                    valor_unit = float(row['VALOR UNITÁRIO']) if pd.notna(row['VALOR UNITÁRIO']) else 0.0
                    
                    c1.write(f"{item_nome} (R$ {valor_unit:.2f})")
                    qtd = c2.number_input(f"Qtd", min_value=0, value=0, key=f"cag_{index}")
                    
                    if qtd > 0:
                        st.session_state.orcamento.append({
                            "Categoria": "CAG",
                            "Item": item_nome,
                            "Quantidade": qtd,
                            "Valor_Unitario": valor_unit,
                            "Custo_Total": qtd * valor_unit
                        })
                        
            with col2:
                st.subheader("AHU01")
                for index, row in ahu_df.iterrows():
                    c1, c2 = st.columns([3, 1])
                    item_nome = row['ITEM']
                    valor_unit = float(row['VALOR UNITÁRIO']) if pd.notna(row['VALOR UNITÁRIO']) else 0.0
                    
                    c1.write(f"{item_nome} (R$ {valor_unit:.2f})")
                    qtd = c2.number_input(f"Qtd", min_value=0, value=0, key=f"ahu_{index}")
                    
                    if qtd > 0:
                        st.session_state.orcamento.append({
                            "Categoria": "AHU01",
                            "Item": item_nome,
                            "Quantidade": qtd,
                            "Valor_Unitario": valor_unit,
                            "Custo_Total": qtd * valor_unit
                        })

        with aba_infra:
            st.header("Cálculo de Infraestrutura")
            st.markdown("Insira a distância média para calcular cabos e eletrocalhas/tubulações.")
            
            for index, row in infra_df.iterrows():
                # Lendo usando o índice da coluna para evitar problemas de nomes duplicados na planilha
                tipo_inst = str(row.iloc[0]) # Coluna 0: INSTRUMENTAÇÃO
                custo_cabo = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0.0 # Coluna 5: CABO (R$)
                custo_infra = float(row.iloc[6]) if pd.notna(row.iloc[6]) else 0.0 # Coluna 6: INFRA (R$)
                
                # Pula linhas de cabeçalho perdidas na planilha
                if tipo_inst.strip() in ["", "nan", "-", "Equipamentos", "Total Equipamentos"]:
                    continue

                st.write(f"**{tipo_inst}** (Cabo: R${custo_cabo:.2f}/m | Infra: R${custo_infra:.2f}/m)")
                c1, c2 = st.columns(2)
                qtd_inst = c1.number_input("Qtd. de Instrumentos", min_value=0, value=0, key=f"infra_qtd_{index}")
                dist_media = c2.number_input("Distância Média (m)", min_value=0.0, value=0.0, step=1.0, key=f"infra_dist_{index}")
                
                if qtd_inst > 0 and dist_media > 0:
                    metragem = qtd_inst * dist_media
                    st.session_state.orcamento.append({
                        "Categoria": "Infraestrutura",
                        "Item": f"Cabo para {tipo_inst} ({metragem}m)",
                        "Quantidade": metragem,
                        "Valor_Unitario": custo_cabo,
                        "Custo_Total": metragem * custo_cabo
                    })
                    st.session_state.orcamento.append({
                        "Categoria": "Infraestrutura",
                        "Item": f"Infra Física para {tipo_inst} ({metragem}m)",
                        "Quantidade": metragem,
                        "Valor_Unitario": custo_infra,
                        "Custo_Total": metragem * custo_infra
                    })

        with aba_resumo:
            st.header("Resumo do Orçamento")
            
            if st.session_state.orcamento:
                df_resumo = pd.DataFrame(st.session_state.orcamento)
                
                # Agrupa e soma os custos caso a tela recarregue (evita duplicatas no session_state)
                df_resumo = df_resumo.groupby(['Categoria', 'Item', 'Valor_Unitario'], as_index=False).agg({'Quantidade': 'sum', 'Custo_Total': 'sum'})
                
                st.dataframe(df_resumo, use_container_width=True)
                
                total_projeto = df_resumo['Custo_Total'].sum()
                st.subheader(f"Custo Total de Materiais: R$ {total_projeto:,.2f}")
                
                # Opção para exportar para Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_resumo.to_excel(writer, index=False, sheet_name='Orcamento')
                
                st.download_button(
                    label="📥 Exportar Orçamento para Excel",
                    data=buffer.getvalue(),
                    file_name="orcamento_automacao.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("Selecione itens nas abas de Equipamentos e Infraestrutura para ver o resumo.")

        # Limpa o orcamento atual para evitar duplicação no próximo clique/recarregamento
        st.session_state.orcamento = []
        
    except FileNotFoundError:
        st.error("⚠️ Arquivos CSV não encontrados! Certifique-se de que os arquivos 'CAG.csv', 'AHU01.csv' e 'Infra.csv' estão dentro da pasta 'dados' no seu GitHub.")
    except Exception as e:
        st.error(f"Erro ao processar as planilhas de custo: {e}")
