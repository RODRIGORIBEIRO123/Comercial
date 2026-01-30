# Função para Ler Dados e Tratar Erros de Coluna
def carregar_dados():
    sh = conectar_google_sheets()
    
    # Função interna para ler aba e garantir colunas
    def ler_aba(nome_aba, colunas_esperadas):
        try:
            ws = sh.worksheet(nome_aba)
            dados = ws.get_all_records()
            
            if not dados:
                # Se estiver vazia, cria um DataFrame vazio com as colunas certas para não dar erro
                return pd.DataFrame(columns=colunas_esperadas)
                
            df = pd.DataFrame(dados)
            
            # Limpeza de colunas (tira espaços extras e joga para minúsculo para comparar)
            # Mas mantemos o nome original se estiver tudo certo
            # Aqui vamos forçar a padronização: remover espaços das colunas
            df.columns = df.columns.str.strip()
            
            # Verificação de Segurança
            for col in colunas_esperadas:
                if col not in df.columns:
                    st.error(f"⚠️ Erro na aba '{nome_aba}': Não encontrei a coluna '{col}'.")
                    st.write(f"Colunas encontradas: {list(df.columns)}")
                    st.stop()
            return df
            
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"❌ A aba '{nome_aba}' não existe na planilha!")
            st.stop()

    # Lê as 3 abas validando as colunas
    df_esc = ler_aba("Escopos", ["Categoria", "Titulo_Curto", "Texto_Completo"])
    df_exc = ler_aba("Exclusoes", ["Titulo_Curto", "Texto_Completo"])
    df_resp = ler_aba("Responsabilidades", ["Titulo_Curto", "Texto_Completo"])
    
    return df_esc, df_exc, df_resp
