import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import matplotlib.pyplot as plt
import io
import re
import unicodedata
import os
from rapidfuzz import fuzz

# Módulos para envio de e-mail
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# ========================
# Configurações de e-mail (ajuste conforme necessário)
# ========================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USERNAME = "seuemail@gmail.com"      # Altere para seu e-mail
SMTP_PASSWORD = "suasenha"                  # Altere para sua senha (ou app password)
EMAIL_RECIPIENT = "destinatario@exemplo.com"  # E-mail que receberá a notificação

def send_email(subject, body, to_email, attachment_path=None):
    msg = MIMEMultipart()
    msg['From'] = SMTP_USERNAME
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    # Se houver arquivo para anexar, anexa-o
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
        msg.attach(part)
    
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        server.send_message(msg)
        server.quit()
        st.success("E-mail enviado com sucesso!")
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {e}")

# ========================
# Exibe a logo da empresa
# ========================
st.image("logoYP.png", width=200, caption="Yinson Production")
st.sidebar.image("logoYP.png", width=200, caption="Yinson Production")
# ========================
# Configuração do Banco de Dados
# ========================
DB_PATH = "report_history.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    # Cria a tabela de histórico, se não existir
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
    # Cria a tabela de usuários, se não existir (sem a coluna last_access por padrão)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password TEXT
    )
    """)
    # Verifica se a coluna last_access já existe na tabela users
    cursor.execute("PRAGMA table_info(users)")
    columns = [row[1] for row in cursor.fetchall()]  # row[1] contém o nome da coluna
    if "last_access" not in columns:
        cursor.execute("ALTER TABLE users ADD COLUMN last_access TEXT")
    
    # Insere usuários padrão (usando INSERT OR IGNORE para evitar duplicatas)
    cursor.execute("INSERT OR IGNORE INTO users (username, password, last_access) VALUES (?, ?, ?)", ("admin", "1234", None))
    cursor.execute("INSERT OR IGNORE INTO users (username, password, last_access) VALUES (?, ?, ?)", ("thiago", "fpsonery", None))
    conn.commit()
    conn.close()

def update_last_access(username):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("UPDATE users SET last_access = ? WHERE username = ?", (timestamp, username))
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

# Funções para gerenciamento de usuários pelo admin
def add_user(username, password):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO users (username, password, last_access) VALUES (?, ?, ?)", (username, password, None))
        conn.commit()
    except Exception as e:
        st.error(f"Erro ao cadastrar usuário: {e}")
    finally:
        conn.close()

def delete_user(username):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM users WHERE username = ?", (username,))
        conn.commit()
    except Exception as e:
        st.error(f"Erro ao deletar usuário: {e}")
    finally:
        conn.close()

def get_all_users():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT username, password, last_access FROM users", conn)
    conn.close()
    return df

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

        # Processa o arquivo de Tipo de Treinamento (opcional)
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

        # Processa o arquivo Unisea (opcional)
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
    st.image("logoYP.png", width=200, caption="Yinson Production")
    username = st.text_input("Usuário")
    password = st.text_input("Senha", type="password")
    if st.button("Entrar"):
        if check_login(username, password):
            st.session_state.logged_in = True
            st.session_state.username = username
            update_last_access(username)
            st.success("Login realizado com sucesso!")
        else:
            st.error("Credenciais inválidas!")

# ========================
# Aplicação Principal
# ========================
if st.session_state.get('logged_in'):
    st.title(f"Relatório de Treinamento - FPSO | Logado como: {st.session_state.username}")
    
    # Menu de navegação
    pages = ["Relatório", "Filtros", "Visualização", "Tabela Completa", "Uploads Salvos", "Histórico"]
    st.sidebar.write("Usuário logado:", st.session_state.username)
    if st.session_state.username.lower() == "admin":
        pages.append("Admin")
    
    page = st.sidebar.radio("Selecione a página", pages)
    
    if 'df_final' not in st.session_state:
        st.session_state.df_final = None

    # ----- Página Relatório (Upload + exportação + envio de e-mail) -----
    if page == "Relatório":
        st.header("Upload dos Arquivos")
        upload_option = st.radio("Selecione a opção", ["Novo Upload", "Usar Último Upload"])
        
        if upload_option == "Novo Upload":
            team_file = st.file_uploader("Team.xlsx", type=["xlsx"], key="team")
            train_file = st.file_uploader("Treinamentos.xlsx", type=["xlsx"], key="train")
            control_file = st.file_uploader("Controle.xlsx", type=["xlsx"], key="control")
            training_type_file = st.file_uploader("Listagem Tipo Treinamento (opcional)", type=["xlsx"], key="training_type")
            unisea_file = st.file_uploader("Planilha Unisea (opcional)", type=["xlsx"], key="unisea")
            fuzzy_threshold = st.number_input("Threshold Fuzzy:", min_value=0, max_value=100, value=80)
            
            if st.button("Processar Dados"):
                if not (team_file and train_file and control_file):
                    st.error("É necessário enviar os arquivos Team, Treinamentos e Controle.")
                else:
                    with st.spinner("Processando dados..."):
                        upload_dir = "uploaded_files"
                        if not os.path.exists(upload_dir):
                            os.makedirs(upload_dir)
                        timestamp_folder = datetime.now().strftime("%Y%m%d%H%M%S")
                        session_folder = os.path.join(upload_dir, timestamp_folder)
                        os.makedirs(session_folder)
                        
                        team_path = os.path.join(session_folder, "Team.xlsx")
                        with open(team_path, "wb") as f:
                            f.write(team_file.getbuffer())
                        
                        train_path = os.path.join(session_folder, "Treinamentos.xlsx")
                        with open(train_path, "wb") as f:
                            f.write(train_file.getbuffer())
                        
                        control_path = os.path.join(session_folder, "Controle.xlsx")
                        with open(control_path, "wb") as f:
                            f.write(control_file.getbuffer())
                        
                        training_type_path = None
                        if training_type_file:
                            training_type_path = os.path.join(session_folder, "Listagem_Tipo_Treinamento.xlsx")
                            with open(training_type_path, "wb") as f:
                                f.write(training_type_file.getbuffer())
                        
                        unisea_path = None
                        if unisea_file:
                            unisea_path = os.path.join(session_folder, "Planilha_Unisea.xlsx")
                            with open(unisea_path, "wb") as f:
                                f.write(unisea_file.getbuffer())
                        
                        df_final = process_data(team_path, train_path, control_path, training_type_path, unisea_path, fuzzy_threshold)
                        
                        # Salva o DataFrame final para uso futuro
                        final_data_path = os.path.join(session_folder, "final.xlsx")
                        if df_final is not None:
                            df_final.to_excel(final_data_path, index=False)
                    
                    if df_final is not None:
                        st.session_state.df_final = df_final
                        st.success("Relatório processado com sucesso!")
                        st.write("Exibindo os 5 primeiros registros:")
                        st.dataframe(df_final.head())
                        
                        buffer = io.BytesIO()
                        df_final.to_excel(buffer, index=False)
                        st.download_button(label="Baixar Tabela Completa",
                                           data=buffer,
                                           file_name=f"Status_Treinamento_Completo_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        
                        email_subject = "Relatório de Treinamento Finalizado"
                        email_body = "O relatório foi processado com sucesso. Em anexo, o arquivo final."
                        send_email(email_subject, email_body, EMAIL_RECIPIENT, attachment_path=final_data_path)
        
        else:  # Usar Último Upload com substituição individual
            upload_dir = "uploaded_files"
            if not os.path.exists(upload_dir):
                st.error("Nenhum upload encontrado. Por favor, faça um novo upload.")
            else:
                sessions = [os.path.join(upload_dir, d) for d in os.listdir(upload_dir) if os.path.isdir(os.path.join(upload_dir, d))]
                if not sessions:
                    st.error("Nenhum upload encontrado. Por favor, faça um novo upload.")
                else:
                    last_session = sorted(sessions)[-1]
                    last_session_name = os.path.basename(last_session)
                    try:
                        last_upload_date = datetime.strptime(last_session_name, "%Y%m%d%H%M%S")
                        last_upload_date_str = last_upload_date.strftime("%Y-%m-%d %H:%M:%S")
                    except Exception:
                        last_upload_date_str = "Data desconhecida"
                    
                    st.info(f"Último upload realizado em: {last_upload_date_str}")
                    st.write("Arquivos disponíveis no último upload:")
                    files_available = os.listdir(last_session)
                    st.write(files_available)
                    
                    st.markdown("### Substituir arquivos (opcional)")
                    team_file_new = st.file_uploader("Substituir Team.xlsx", type=["xlsx"], key="team_replace")
                    train_file_new = st.file_uploader("Substituir Treinamentos.xlsx", type=["xlsx"], key="train_replace")
                    control_file_new = st.file_uploader("Substituir Controle.xlsx", type=["xlsx"], key="control_replace")
                    training_type_file_new = st.file_uploader("Substituir Listagem Tipo Treinamento (opcional)", type=["xlsx"], key="training_type_replace")
                    unisea_file_new = st.file_uploader("Substituir Planilha Unisea (opcional)", type=["xlsx"], key="unisea_replace")
                    
                    fuzzy_threshold = st.number_input("Threshold Fuzzy:", min_value=0, max_value=100, value=80)
                    if st.button("Processar Dados do Último Upload"):
                        with st.spinner("Processando dados..."):
                            team_path = os.path.join(last_session, "Team.xlsx")
                            if team_file_new is not None:
                                with open(team_path, "wb") as f:
                                    f.write(team_file_new.getbuffer())
                            train_path = os.path.join(last_session, "Treinamentos.xlsx")
                            if train_file_new is not None:
                                with open(train_path, "wb") as f:
                                    f.write(train_file_new.getbuffer())
                            control_path = os.path.join(last_session, "Controle.xlsx")
                            if control_file_new is not None:
                                with open(control_path, "wb") as f:
                                    f.write(control_file_new.getbuffer())
                            training_type_path = os.path.join(last_session, "Listagem_Tipo_Treinamento.xlsx")
                            if training_type_file_new is not None:
                                with open(training_type_path, "wb") as f:
                                    f.write(training_type_file_new.getbuffer())
                            else:
                                if not os.path.exists(training_type_path):
                                    training_type_path = None
                            unisea_path = os.path.join(last_session, "Planilha_Unisea.xlsx")
                            if unisea_file_new is not None:
                                with open(unisea_path, "wb") as f:
                                    f.write(unisea_file_new.getbuffer())
                            else:
                                if not os.path.exists(unisea_path):
                                    unisea_path = None
                            
                            df_final = process_data(team_path, train_path, control_path, training_type_path, unisea_path, fuzzy_threshold)
                            
                            final_data_path = os.path.join(last_session, "final.xlsx")
                            if df_final is not None:
                                df_final.to_excel(final_data_path, index=False)
                        
                        if df_final is not None:
                            st.session_state.df_final = df_final
                            st.success("Relatório processado com sucesso!")
                            st.write("Exibindo os 5 primeiros registros:")
                            st.dataframe(df_final.head())
                            
                            buffer = io.BytesIO()
                            df_final.to_excel(buffer, index=False)
                            st.download_button(label="Baixar Tabela Completa",
                                               data=buffer,
                                               file_name=f"Status_Treinamento_Completo_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            
                            email_subject = "Relatório de Treinamento Finalizado"
                            email_body = "O relatório foi processado com sucesso. Em anexo, o arquivo final."
                            send_email(email_subject, email_body, EMAIL_RECIPIENT, attachment_path=final_data_path)
    
    # ----- Página Filtros (exportação personalizada) -----
    elif page == "Filtros":
        st.header("Filtros Avançados")
        if st.session_state.df_final is None:
            st.error("Nenhum dado processado para filtrar. Vá na página 'Relatório' e processe os dados.")
        else:
            df_final = st.session_state.df_final.copy()
            cargos = sorted(df_final['cargo_pt_team'].dropna().unique())
            cargo_selected = st.selectbox("Cargo", options=["Todos"] + cargos)
            status_selected = st.selectbox("Status", options=["Todos", "OK", "Retreinamento", "Not started"])
            data_inicial = st.date_input("Data Inicial")
            data_final = st.date_input("Data Final")
            
            if cargo_selected != "Todos":
                df_final = df_final[df_final['cargo_pt_team'] == cargo_selected]
            if status_selected != "Todos":
                df_final = df_final[df_final['status_final'] == status_selected]
            if 'control_data_completo' in df_final.columns:
                df_final['control_data_completo'] = pd.to_datetime(df_final['control_data_completo'], errors='coerce')
                df_final = df_final[(df_final['control_data_completo'] >= pd.to_datetime(data_inicial)) & 
                                    (df_final['control_data_completo'] <= pd.to_datetime(data_final))]
            
            if df_final.empty:
                st.info("Nenhum registro encontrado com os filtros aplicados.")
            else:
                st.dataframe(df_final)
                buffer = io.BytesIO()
                df_final.to_excel(buffer, index=False)
                st.download_button(label="Exportar Dados Filtrados",
                                   data=buffer,
                                   file_name=f"Status_Treinamento_Filtrado_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # ----- Página Visualização (Dashboard com Gráficos) -----
    elif page == "Visualização":
        st.header("Dashboard de Visualização")
        if st.session_state.df_final is None:
            st.error("Nenhum dado processado para visualizar. Vá na página 'Relatório'.")
        else:
            df_final = st.session_state.df_final
            # Gráfico de Pizza – Status Geral
            status_counts = df_final['status_final'].value_counts()
            labels = ['OK', 'Retreinamento', 'Not started']
            data = [status_counts.get(l, 0) for l in labels]
            fig1, ax1 = plt.subplots()
            ax1.pie(data, labels=labels, autopct='%1.1f%%', startangle=90)
            ax1.axis('equal')
            st.pyplot(fig1)
            
            # Gráfico de Barras – Status por Cargo
            if 'cargo_pt_team' in df_final.columns and 'status_final' in df_final.columns:
                group = df_final.groupby(['cargo_pt_team', 'status_final']).size().unstack(fill_value=0)
                fig2, ax2 = plt.subplots(figsize=(8, 4))
                group.plot(kind='bar', ax=ax2)
                ax2.set_title("Status por Cargo")
                st.pyplot(fig2)
    
    # ----- Página Tabela Completa (Simulação do Excel) -----
    elif page == "Tabela Completa":
        st.header("Tabela Completa")
        if st.session_state.df_final is None:
            st.error("Nenhum dado processado. Vá na página 'Relatório' e processe os dados.")
        else:
            df_table = st.session_state.df_final.copy()
            st.markdown("### Filtro Global (pesquisa em todas as colunas)")
            search_term = st.text_input("Digite o termo para filtrar:")
            if search_term:
                df_table = df_table[df_table.apply(lambda row: row.astype(str).str.contains(search_term, case=False).any(), axis=1)]
            st.dataframe(df_table)
            buffer = io.BytesIO()
            df_table.to_excel(buffer, index=False)
            st.download_button(label="Exportar Dados (Personalizado)",
                               data=buffer,
                               file_name=f"Status_Treinamento_Personalizado_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # ----- Página Uploads Salvos (Visualização de uploads anteriores) -----
    elif page == "Uploads Salvos":
        st.header("Uploads Salvos")
        upload_dir = "uploaded_files"
        if not os.path.exists(upload_dir):
            st.error("Nenhum upload salvo encontrado.")
        else:
            session_folders = [os.path.join(upload_dir, d) for d in os.listdir(upload_dir)
                               if os.path.isdir(os.path.join(upload_dir, d)) and os.path.exists(os.path.join(upload_dir, d, "final.xlsx"))]
            if not session_folders:
                st.info("Nenhum upload com relatório processado encontrado.")
            else:
                sessions_dict = {}
                for folder in session_folders:
                    session_name = os.path.basename(folder)
                    try:
                        session_date = datetime.strptime(session_name, "%Y%m%d%H%M%S")
                        sessions_dict[folder] = session_date.strftime("%Y-%m-%d %H:%M:%S")
                    except Exception:
                        sessions_dict[folder] = "Data desconhecida"
                
                selected_session = st.selectbox("Selecione o upload", list(sessions_dict.keys()), format_func=lambda x: sessions_dict[x])
                if selected_session:
                    final_file = os.path.join(selected_session, "final.xlsx")
                    if os.path.exists(final_file):
                        df_saved = pd.read_excel(final_file)
                        st.dataframe(df_saved)
                        buffer = io.BytesIO()
                        df_saved.to_excel(buffer, index=False)
                        st.download_button(label="Exportar Dados do Upload Selecionado",
                                           data=buffer,
                                           file_name=f"Relatorio_{os.path.basename(selected_session)}.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    else:
                        st.error("Arquivo final.xlsx não encontrado no upload selecionado.")
    
    # ----- Página Admin (Gerenciamento de Usuários) -----
    elif page == "Admin":
        st.header("Administração de Usuários")
        if st.session_state.username.lower() != "admin":
            st.error("Acesso restrito a administradores.")
        else:
            st.subheader("Cadastrar Novo Usuário")
            new_username = st.text_input("Novo Usuário", key="new_user")
            new_password = st.text_input("Nova Senha", key="new_pass", type="password")
            if st.button("Cadastrar Usuário"):
                if new_username and new_password:
                    add_user(new_username, new_password)
                    st.success("Usuário cadastrado com sucesso!")
                else:
                    st.error("Informe um nome de usuário e senha.")
            
            st.subheader("Lista de Usuários")
            df_users = get_all_users()
            st.dataframe(df_users)
            
            st.subheader("Excluir Usuário")
            users_list = df_users['username'].tolist()
            user_to_delete = st.selectbox("Selecione um usuário para excluir", users_list)
            if st.button("Excluir Usuário"):
                if user_to_delete.lower() == "admin":
                    st.error("Não é permitido excluir o usuário admin.")
                else:
                    delete_user(user_to_delete)
                    st.success(f"Usuário '{user_to_delete}' excluído com sucesso!")
                    st.experimental_rerun()

    # ----- Página Histórico (Logs dos relatórios) -----
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
