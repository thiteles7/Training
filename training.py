import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import matplotlib.pyplot as plt
import io
import re
import unicodedata
from rapidfuzz import fuzz

# ========================
# Configuração do Banco de Dados
# ========================

DB_PATH = "report_history.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS report_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT,
        report_type TEXT,
        file_name TEXT,
        filter_options TEXT,
        user TEXT
    )
    """)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password TEXT
    )
    """)
    # Insere usuários padrão
    cursor.execute("INSERT OR IGNORE INTO users (username, password) VALUES (?, ?)", ("admin", "1234"))
    cursor.execute("INSERT OR IGNORE INTO users (username, password) VALUES (?, ?)", ("thiago", "fpsonery"))
    conn.commit()
    conn.close()

def check_login(username, password):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password))
    user = cursor.fetchone()
    conn.close()
    return user

def log_report(report_type, file_name, filter_options="", user="Desconhecido"):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        INSERT INTO report_history (timestamp, report_type, file_name, filter_options, user)
        VALUES (?, ?, ?, ?, ?)
    """, (timestamp, report_type, file_name, filter_options, user))
    conn.commit()
    conn.close()

# ========================
# Funções Utilitárias
# ========================

def safe_float(value):
    try:
        return float(str(value).strip())
    except Exception:
        return None

def extract_revision(rev_str):
    if isinstance(rev_str, str):
        digits = re.sub("[^0-9]", "", rev_str)
        return int(digits) if digits else None
    try:
        return int(rev_str)
    except Exception:
        return None

def normalize_text(text):
    try:
        text = str(text)
        text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8')
        return text.lower().strip()
    except Exception:
        return str(text).lower().strip()

# ========================
# Processamento dos Dados
# ========================

def process_data(team_file, train_file, control_file, training_type_file=None, unisea_file=None, fuzzy_threshold=80):
    try:
        # Lê o arquivo Team e separa as colunas de cargo
        df_team = pd.read_excel(team_file)
        if "Position in Matrix" not in df_team.columns:
            st.error("A coluna 'Position in Matrix' não foi encontrada no arquivo Team.xlsx.")
            return None
        df_team[['cargo_en_team', 'cargo_pt_team']] = df_team["Position in Matrix"].str.split("\n", n=1, expand=True)
        df_team['cargo_en_team'] = df_team['cargo_en_team'].str.strip()
        df_team['cargo_pt_team'] = df_team['cargo_pt_team'].str.strip()

        # Lê o arquivo de Treinamentos
        df_train = pd.read_excel(train_file).iloc[:, :6]
        df_train.columns = ['cargo_en_train', 'cargo_pt_train', 'procedimento_nome',
                            'procedimento_num_en', 'procedimento_num_pt', 'requisito']
        df_merged = pd.merge(df_team, df_train, left_on='cargo_pt_team', right_on='cargo_pt_train', how='left')
        df_merged['procedimento_num_assigned'] = df_merged.apply(
            lambda row: row['procedimento_num_pt'] if str(row.get('Nationality', '')).upper() == 'BR' 
                        else row['procedimento_num_en'], axis=1)
        df_merged['procedimento_num_alternative'] = df_merged.apply(
            lambda row: row['procedimento_num_en'] if str(row.get('Nationality', '')).upper() == 'BR'
                        else row['procedimento_num_pt'], axis=1)
        df_result = df_merged[['Unisea E-learning User', 'cargo_pt_team', 'cargo_en_team',
                               'procedimento_nome', 'procedimento_num_assigned',
                               'procedimento_num_alternative', 'requisito']]

        # Lê o arquivo de Controle
        df_control = pd.read_excel(control_file)
        df_control['nome_padrao'] = df_control.iloc[:, 0].astype(str).str.upper().str.strip()
        df_control['procedimento_num_controle'] = df_control.iloc[:, 4].astype(str).str.strip()
        df_control['procedimento_nome_controle'] = df_control.iloc[:, 5].astype(str).str.upper().str.strip()
        df_control['rev'] = df_control['procedimento_nome_controle'].str[-7:]
        df_control['status'] = df_control.iloc[:, 8]
        df_control['control_data_completo'] = pd.to_datetime(df_control.iloc[:, 9], errors='coerce')
        df_result['nome_padrao'] = df_result['Unisea E-learning User'].astype(str).str.upper().str.strip()

        # Função para casar informações de Controle
        def match_control(row, threshold=fuzzy_threshold):
            codigo_atribuido = str(row['procedimento_num_assigned']).strip()
            codigo_alternativo = (str(row['procedimento_num_alternative']).strip() if pd.notnull(row.get('procedimento_num_alternative')) else '')
            codigos = list({codigo_atribuido, codigo_alternativo}) if codigo_alternativo else [codigo_atribuido]
            nome_usuario = row['nome_padrao']
            candidatos = df_control[df_control['procedimento_num_controle'].isin(codigos)]
            if candidatos.empty:
                return pd.Series([None, None, None, None, 0])
            correspondencias_exatas = candidatos[candidatos['nome_padrao'] == nome_usuario]
            if not correspondencias_exatas.empty:
                if not correspondencias_exatas['control_data_completo'].dropna().empty:
                    melhor = correspondencias_exatas.loc[correspondencias_exatas['control_data_completo'].idxmax()]
                else:
                    melhor = correspondencias_exatas.iloc[0]
                return pd.Series([melhor['status'], melhor['control_data_completo'], 
                                  melhor['nome_padrao'], melhor['rev'], 100])
            melhor_score, melhor_candidato = 0, None
            for _, cand in candidatos.iterrows():
                score = fuzz.ratio(normalize_text(nome_usuario), normalize_text(cand['nome_padrao']))
                if score > melhor_score:
                    melhor_score, melhor_candidato = score, cand
            if melhor_score >= threshold and melhor_candidato is not None:
                return pd.Series([melhor_candidato['status'], melhor_candidato['control_data_completo'],
                                  melhor_candidato['nome_padrao'], melhor_candidato['rev'], melhor_score])
            return pd.Series([None, None, None, None, melhor_score])

        df_result[['control_status', 'control_data_completo', 'control_nome', 'control_rev', 'match_score']] = df_result.apply(match_control, axis=1)
        df_result['inconsistencia'] = df_result['control_status'].isnull() | (df_result['match_score'] < 100)

        # Se fornecido, processa o arquivo de Tipo de Treinamento
        if training_type_file is not None:
            df_type = pd.read_excel(training_type_file).iloc[:, :3]
            df_type.columns = ['procedimento_num_en_type', 'procedimento_num_pt_type', 'categoria']
            def get_categoria(procedimento):
                procedimento = str(procedimento).strip()
                match = df_type[(df_type['procedimento_num_en_type'].astype(str).str.strip() == procedimento) |
                                (df_type['procedimento_num_pt_type'].astype(str).str.strip() == procedimento)]
                if not match.empty:
                    return match.iloc[0]['categoria']
                return None
            df_result['categoria'] = df_result['procedimento_num_assigned'].apply(get_categoria)
        else:
            df_result['categoria'] = None

        # Se fornecido, processa o arquivo Unisea
        if unisea_file is not None:
            df_unisea = pd.read_excel(unisea_file)
            df_unisea = df_unisea.rename(columns={df_unisea.columns[0]: 'procedimento_num_unisea',
                                                  df_unisea.columns[9]: 'rev_unisea'})
            df_unisea['procedimento_num_unisea'] = df_unisea['procedimento_num_unisea'].astype(str).str.strip()
            df_result['procedimento_num_assigned'] = df_result['procedimento_num_assigned'].astype(str).str.strip()
            df_result = df_result.merge(df_unisea[['procedimento_num_unisea', 'rev_unisea']],
                                         left_on='procedimento_num_assigned',
                                         right_on='procedimento_num_unisea', how='left')
            df_result.drop(columns=['procedimento_num_unisea'], inplace=True)
            def compare_revs(row):
                if normalize_text(row.get('control_status')) != "completed":
                    return "Not started"
                rev_control_extracted = extract_revision(row['control_rev'])
                rev_unisea_extracted = extract_revision(row['rev_unisea'])
                if rev_control_extracted is None or rev_unisea_extracted is None:
                    return "OK"
                if rev_control_extracted == rev_unisea_extracted:
                    return "OK"
                else:
                    return "Retreinamento"
            df_result['status_final'] = df_result.apply(compare_revs, axis=1)
        else:
            df_result['status_final'] = df_result['control_status']

        colunas_final = ['Unisea E-learning User', 'cargo_pt_team', 'cargo_en_team', 'procedimento_nome',
                         'procedimento_num_assigned', 'procedimento_num_alternative', 'requisito',
                         'categoria', 'control_status', 'control_nome', 'control_rev', 'rev_unisea',
                         'status_final', 'control_data_completo', 'match_score', 'inconsistencia']
        df_final = df_result[colunas_final]
        return df_final

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar os dados: {e}")
        return None

# ========================
# Inicializa o Banco e Sistema de Login
# ========================

init_db()

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("Login")
    username = st.text_input("Usuário")
    password = st.text_input("Senha", type="password")
    if st.button("Entrar"):
        if check_login(username, password):
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success("Login realizado com sucesso!")
#            st.experimental_rerun()#
        else:
            st.error("Credenciais inválidas!")

# ========================
# Aplicação Principal
# ========================

if st.session_state.get('logged_in'):
    st.title(f"Relatório de Treinamento - FPSO | Logado como: {st.session_state.username}")
    
    # Cria navegação via Sidebar
    page = st.sidebar.radio("Selecione a página", ["Relatório", "Filtros", "Visualização", "Histórico"])
    
    # Guarda o dataframe processado em session_state para uso em outras páginas
    if 'df_final' not in st.session_state:
        st.session_state.df_final = None

    # ----- Página Relatório -----
    if page == "Relatório":
        st.header("Upload dos Arquivos")
        team_file = st.file_uploader("Team.xlsx", type=["xlsx"])
        train_file = st.file_uploader("Treinamentos.xlsx", type=["xlsx"])
        control_file = st.file_uploader("Controle.xlsx", type=["xlsx"])
        training_type_file = st.file_uploader("Listagem Tipo Treinamento (opcional)", type=["xlsx"])
        unisea_file = st.file_uploader("Planilha Unisea (opcional)", type=["xlsx"])
        fuzzy_threshold = st.number_input("Threshold Fuzzy:", min_value=0, max_value=100, value=80)
        
        if st.button("Processar Dados"):
            if not (team_file and train_file and control_file):
                st.error("É necessário enviar os arquivos Team, Treinamentos e Controle.")
            else:
                df_final = process_data(team_file, train_file, control_file, training_type_file, unisea_file, fuzzy_threshold)
                if df_final is not None:
                    st.session_state.df_final = df_final
                    st.success("Relatório processado com sucesso!")
                    st.write("Exibindo os 5 primeiros registros:")
                    st.dataframe(df_final.head())
                    # Exportação para Excel
                    buffer = io.BytesIO()
                    df_final.to_excel(buffer, index=False)
                    st.download_button(label="Baixar Excel",
                                       data=buffer,
                                       file_name=f"Status_Treinamento_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # ----- Página Filtros -----
    elif page == "Filtros":
        st.header("Filtros Avançados")
        if st.session_state.df_final is None:
            st.error("Nenhum dado processado para filtrar. Vá na página 'Relatório' e processe os dados.")
        else:
            df_final = st.session_state.df_final
            cargos = sorted(df_final['cargo_pt_team'].dropna().unique())
            cargo_selected = st.selectbox("Cargo", options=["Todos"] + cargos)
            status_selected = st.selectbox("Status", options=["Todos", "OK", "Retreinamento", "Not started"])
            data_inicial = st.date_input("Data Inicial")
            data_final = st.date_input("Data Final")
            
            filtered_df = df_final.copy()
            if cargo_selected != "Todos":
                filtered_df = filtered_df[filtered_df['cargo_pt_team'] == cargo_selected]
            if status_selected != "Todos":
                filtered_df = filtered_df[filtered_df['status_final'] == status_selected]
            if 'control_data_completo' in filtered_df.columns:
                filtered_df['control_data_completo'] = pd.to_datetime(filtered_df['control_data_completo'], errors='coerce')
                filtered_df = filtered_df[(filtered_df['control_data_completo'] >= pd.to_datetime(data_inicial)) & 
                                          (filtered_df['control_data_completo'] <= pd.to_datetime(data_final))]
            
            if filtered_df.empty:
                st.info("Nenhum registro encontrado com os filtros aplicados.")
            else:
                st.dataframe(filtered_df)
    
    # ----- Página Visualização -----
    elif page == "Visualização":
        st.header("Dashboard de Visualização")
        if st.session_state.df_final is None:
            st.error("Nenhum dado processado para visualizar. Vá na página 'Relatório'.")
        else:
            df_final = st.session_state.df_final
            # Gráfico de Pizza para Status Geral
            status_counts = df_final['status_final'].value_counts()
            labels = ['OK', 'Retreinamento', 'Not started']
            data = [status_counts.get(l, 0) for l in labels]
            fig1, ax1 = plt.subplots()
            ax1.pie(data, labels=labels, autopct='%1.1f%%', startangle=90)
            ax1.axis('equal')
            st.pyplot(fig1)
            
            # Gráfico de Barras para Status por Cargo
            if 'cargo_pt_team' in df_final.columns and 'status_final' in df_final.columns:
                group = df_final.groupby(['cargo_pt_team', 'status_final']).size().unstack(fill_value=0)
                fig2, ax2 = plt.subplots(figsize=(8, 4))
                group.plot(kind='bar', ax=ax2)
                ax2.set_title("Status por Cargo")
                st.pyplot(fig2)
    
    # ----- Página Histórico -----
    elif page == "Histórico":
        st.header("Histórico de Relatórios")
        try:
            conn = sqlite3.connect(DB_PATH)
            df_history = pd.read_sql_query("SELECT * FROM report_history ORDER BY id DESC", conn)
            conn.close()
            if df_history.empty:
                st.info("Nenhum registro de relatório.")
            else:
                st.dataframe(df_history)
        except Exception as e:
            st.error(f"Erro ao carregar o histórico: {e}")
