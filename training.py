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

# --- Configurar o layout para "wide" ---
st.set_page_config(page_title="Training Report - FPSO", layout="wide")

# ========================
# Email Settings (ajuste conforme necessário)
# ========================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USERNAME = "seuemail@gmail.com"      # Altere para o seu e-mail
SMTP_PASSWORD = "suasenha"                  # Altere para sua senha (ou app password)
EMAIL_RECIPIENT = "destinatario@exemplo.com"  # E-mail do destinatário

def send_email(subject, body, to_email, attachment_path=None):
    msg = MIMEMultipart()
    msg['From'] = SMTP_USERNAME
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    # Anexa o arquivo, se houver
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
        st.success("E-mail sent successfully!")
    except Exception as e:
        st.error(f"Error sending e-mail: {e}")

# ========================
# Display company logo
# ========================
st.image("logoYP.png", width=200, caption="Yinson Production")
st.sidebar.image("logoYP.png", width=200, caption="Yinson Production")

# ========================
# Database Configuration
# ========================
DB_PATH = "report_history.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    # Cria tabela report_history, se não existir
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
    # Cria tabela users, se não existir (sem a coluna last_access inicialmente)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password TEXT
    )
    """)
    # Verifica se a coluna last_access existe na tabela users
    cursor.execute("PRAGMA table_info(users)")
    columns = [row[1] for row in cursor.fetchall()]
    if "last_access" not in columns:
        cursor.execute("ALTER TABLE users ADD COLUMN last_access TEXT")
    
    # Insere usuários padrões (usando INSERT OR IGNORE para evitar duplicatas)
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

def log_report(report_type, file_name, filter_options="", user="Unknown"):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        INSERT INTO report_history (timestamp, report_type, file_name, filter_options, user)
        VALUES (?, ?, ?, ?, ?)
    """, (timestamp, report_type, file_name, filter_options, user))
    conn.commit()
    conn.close()

# Funções para administração de usuários (admin)
def add_user(username, password):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO users (username, password, last_access) VALUES (?, ?, ?)", (username, password, None))
        conn.commit()
    except Exception as e:
        st.error(f"Error registering user: {e}")
    finally:
        conn.close()

def delete_user(username):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM users WHERE username = ?", (username,))
        conn.commit()
    except Exception as e:
        st.error(f"Error deleting user: {e}")
    finally:
        conn.close()

def get_all_users():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT username, password, last_access FROM users", conn)
    conn.close()
    return df

# ========================
# Utility Functions
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
# Funções de Persistência para a Tabela VCP
# ========================
def load_vcp_data():
    """Carrega os dados do VCP de um arquivo CSV, se existir."""
    file_path = "vcp_data.csv"
    if os.path.exists(file_path):
        return pd.read_csv(file_path)
    else:
        return None

def save_vcp_data(df):
    """Salva os dados do VCP em um arquivo CSV."""
    file_path = "vcp_data.csv"
    df.to_csv(file_path, index=False)

# ========================
# Data Processing Function
# ========================
def process_data(team_file, train_file, control_file, training_type_file=None, unisea_file=None, fuzzy_threshold=80):
    try:
        # Lê o arquivo Team e separa as colunas de posição
        df_team = pd.read_excel(team_file)
        if "Position in Matrix" not in df_team.columns:
            st.error("Column 'Position in Matrix' not found in Team.xlsx.")
            return None
        df_team[['cargo_en_team', 'cargo_pt_team']] = df_team["Position in Matrix"].str.split("\n", n=1, expand=True)
        df_team['cargo_en_team'] = df_team['cargo_en_team'].str.strip()
        df_team['cargo_pt_team'] = df_team['cargo_pt_team'].str.strip()

        # Lê o arquivo Trainings
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

        # Lê o arquivo Control
        df_control = pd.read_excel(control_file)
        df_control['nome_padrao'] = df_control.iloc[:, 0].astype(str).str.upper().str.strip()
        df_control['procedimento_num_controle'] = df_control.iloc[:, 4].astype(str).str.strip()
        df_control['procedimento_nome_controle'] = df_control.iloc[:, 5].astype(str).str.upper().str.strip()
        df_control['rev'] = df_control['procedimento_nome_controle'].str[-7:]
        df_control['status'] = df_control.iloc[:, 8]
        df_control['control_data_completo'] = pd.to_datetime(df_control.iloc[:, 9], errors='coerce')
        df_result['nome_padrao'] = df_result['Unisea E-learning User'].astype(str).str.upper().str.strip()

        # Função para fazer match com o Control
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

        # Processa o arquivo Training Type (opcional)
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
        st.error(f"An error occurred while processing data: {e}")
        return None

# ========================
# Inicializa o Banco de Dados e Sistema de Login
# ========================
init_db()

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("Login")
    st.image("logoYP.png", width=200, caption="Yinson Production")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if check_login(username, password):
            st.session_state.logged_in = True
            st.session_state.username = username
            update_last_access(username)
            st.success("Login successful!")
        else:
            st.error("Invalid credentials!")

# ========================
# Aplicação Principal
# ========================
if st.session_state.get('logged_in'):
    st.title(f"Training Report - FPSO | Logged in as: {st.session_state.username}")
    
    # Implementando Tabs
    tabs = st.tabs(["Report", "Filters", "Visualization", "Full Table", "Saved Uploads", "History", "VCP", "Admin"] if st.session_state.username.lower() == "admin" else ["Report", "Filters", "Visualization", "Full Table", "Saved Uploads", "History", "VCP"])
    
    with tabs[0]:  # Tab "Report"
        st.header("Upload Files")
        # Conteúdo atual da seção Report...

    with tabs[1]:  # Tab "Filters"
        st.header("Advanced Filters")
        # Conteúdo atual da seção Filters...

    with tabs[2]:  # Tab "Visualization"
        st.header("Visualization Dashboard")
        # Conteúdo atual da seção Visualization...

    with tabs[3]:  # Tab "Full Table"
        st.header("Full Table")
        # Conteúdo atual da seção Full Table...

    with tabs[4]:  # Tab "Saved Uploads"
        st.header("Saved Uploads")
        # Conteúdo atual da seção Saved Uploads...

    with tabs[5]:  # Tab "History"
        st.header("Reports History")
        # Conteúdo atual da seção History...

    if st.session_state.username.lower() == "admin":
        with tabs[6]:  # Tab "Admin"
            st.header("User Administration")
            # Conteúdo atual da seção Admin...

    # ----- Página Report (Upload + Export + Email) -----
    if page == "Report":
        st.header("Upload Files")
        upload_option = st.radio("Select an option", ["New Upload", "Use Last Upload"])
        
        if upload_option == "New Upload":
            team_file = st.file_uploader("Team.xlsx", type=["xlsx"], key="team")
            train_file = st.file_uploader("Trainings.xlsx", type=["xlsx"], key="train")
            control_file = st.file_uploader("Control.xlsx", type=["xlsx"], key="control")
            training_type_file = st.file_uploader("Training Type Listing (optional)", type=["xlsx"], key="training_type")
            unisea_file = st.file_uploader("Unisea Sheet (optional)", type=["xlsx"], key="unisea")
            fuzzy_threshold = st.number_input("Fuzzy Threshold:", min_value=0, max_value=100, value=80)
            
            if st.button("Process Data"):
                if not (team_file and train_file and control_file):
                    st.error("You must upload the Team, Trainings, and Control files.")
                else:
                    with st.spinner("Processing data..."):
                        upload_dir = "uploaded_files"
                        if not os.path.exists(upload_dir):
                            os.makedirs(upload_dir)
                        timestamp_folder = datetime.now().strftime("%Y%m%d%H%M%S")
                        session_folder = os.path.join(upload_dir, timestamp_folder)
                        os.makedirs(session_folder)
                        
                        team_path = os.path.join(session_folder, "Team.xlsx")
                        with open(team_path, "wb") as f:
                            f.write(team_file.getbuffer())
                        
                        train_path = os.path.join(session_folder, "Trainings.xlsx")
                        with open(train_path, "wb") as f:
                            f.write(train_file.getbuffer())
                        
                        control_path = os.path.join(session_folder, "Control.xlsx")
                        with open(control_path, "wb") as f:
                            f.write(control_file.getbuffer())
                        
                        training_type_path = None
                        if training_type_file:
                            training_type_path = os.path.join(session_folder, "Training_Type_Listing.xlsx")
                            with open(training_type_path, "wb") as f:
                                f.write(training_type_file.getbuffer())
                        
                        unisea_path = None
                        if unisea_file:
                            unisea_path = os.path.join(session_folder, "Unisea_Sheet.xlsx")
                            with open(unisea_path, "wb") as f:
                                f.write(unisea_file.getbuffer())
                        
                        df_final = process_data(team_path, train_path, control_path, training_type_path, unisea_path, fuzzy_threshold)
                        
                        # Salva o DataFrame final para uso futuro
                        final_data_path = os.path.join(session_folder, "final.xlsx")
                        if df_final is not None:
                            df_final.to_excel(final_data_path, index=False)
                    
                    if df_final is not None:
                        st.session_state.df_final = df_final
                        st.success("Report processed successfully!")
                        st.write("Displaying first 5 records:")
                        st.dataframe(df_final.head())
                        
                        buffer = io.BytesIO()
                        df_final.to_excel(buffer, index=False)
                        st.download_button(label="Download Full Table",
                                           data=buffer,
                                           file_name=f"Training_Status_Full_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        
                        email_subject = "Training Report Finalized"
                        email_body = "The report was processed successfully. Attached is the final file."
                        send_email(email_subject, email_body, EMAIL_RECIPIENT, attachment_path=final_data_path)
        
        else:  # Use Last Upload with individual replacement
            upload_dir = "uploaded_files"
            if not os.path.exists(upload_dir):
                st.error("No saved upload found. Please do a new upload.")
            else:
                sessions = [os.path.join(upload_dir, d) for d in os.listdir(upload_dir) if os.path.isdir(os.path.join(upload_dir, d))]
                if not sessions:
                    st.error("No saved upload found. Please do a new upload.")
                else:
                    last_session = sorted(sessions)[-1]
                    last_session_name = os.path.basename(last_session)
                    try:
                        last_upload_date = datetime.strptime(last_session_name, "%Y%m%d%H%M%S")
                        last_upload_date_str = last_upload_date.strftime("%Y-%m-%d %H:%M:%S")
                    except Exception:
                        last_upload_date_str = "Unknown Date"
                    
                    st.info(f"Last upload made on: {last_upload_date_str}")
                    st.write("Files available in the last upload:")
                    files_available = os.listdir(last_session)
                    st.write(files_available)
                    
                    st.markdown("### Replace files (optional)")
                    team_file_new = st.file_uploader("Replace Team.xlsx", type=["xlsx"], key="team_replace")
                    train_file_new = st.file_uploader("Replace Trainings.xlsx", type=["xlsx"], key="train_replace")
                    control_file_new = st.file_uploader("Replace Control.xlsx", type=["xlsx"], key="control_replace")
                    training_type_file_new = st.file_uploader("Replace Training Type Listing (optional)", type=["xlsx"], key="training_type_replace")
                    unisea_file_new = st.file_uploader("Replace Unisea Sheet (optional)", type=["xlsx"], key="unisea_replace")
                    
                    fuzzy_threshold = st.number_input("Fuzzy Threshold:", min_value=0, max_value=100, value=80)
                    if st.button("Process Data from Last Upload"):
                        with st.spinner("Processing data..."):
                            team_path = os.path.join(last_session, "Team.xlsx")
                            if team_file_new is not None:
                                with open(team_path, "wb") as f:
                                    f.write(team_file_new.getbuffer())
                            train_path = os.path.join(last_session, "Trainings.xlsx")
                            if train_file_new is not None:
                                with open(train_path, "wb") as f:
                                    f.write(train_file_new.getbuffer())
                            control_path = os.path.join(last_session, "Control.xlsx")
                            if control_file_new is not None:
                                with open(control_path, "wb") as f:
                                    f.write(control_file_new.getbuffer())
                            training_type_path = os.path.join(last_session, "Training_Type_Listing.xlsx")
                            if training_type_file_new is not None:
                                with open(training_type_path, "wb") as f:
                                    f.write(training_type_file_new.getbuffer())
                            else:
                                if not os.path.exists(training_type_path):
                                    training_type_path = None
                            unisea_path = os.path.join(last_session, "Unisea_Sheet.xlsx")
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
                            st.success("Report processed successfully!")
                            st.write("Displaying first 5 records:")
                            st.dataframe(df_final.head())
                            
                            buffer = io.BytesIO()
                            df_final.to_excel(buffer, index=False)
                            st.download_button(label="Download Customized Data",
                                               data=buffer,
                                               file_name=f"Training_Status_Custom_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            
                            email_subject = "Training Report Finalized"
                            email_body = "The report was processed successfully. Attached is the final file."
                            send_email(email_subject, email_body, EMAIL_RECIPIENT, attachment_path=final_data_path)
    
    # ----- Página Filters (export customizado) -----
    elif page == "Filters":
        st.header("Advanced Filters")
        if st.session_state.df_final is None:
            st.error("No processed data available for filtering. Go to the 'Report' page and process the data.")
        else:
            df_final = st.session_state.df_final.copy()
            cargos = sorted(df_final['cargo_pt_team'].dropna().unique())
            cargo_selected = st.selectbox("Position", options=["All"] + cargos)
            status_selected = st.selectbox("Status", options=["All", "OK", "Retreinamento", "Not started"])
            data_inicial = st.date_input("Start Date")
            data_final = st.date_input("End Date")
            
            if cargo_selected != "All":
                df_final = df_final[df_final['cargo_pt_team'] == cargo_selected]
            if status_selected != "All":
                df_final = df_final[df_final['status_final'] == status_selected]
            if 'control_data_completo' in df_final.columns:
                df_final['control_data_completo'] = pd.to_datetime(df_final['control_data_completo'], errors='coerce')
                df_final = df_final[(df_final['control_data_completo'] >= pd.to_datetime(data_inicial)) & 
                                    (df_final['control_data_completo'] <= pd.to_datetime(data_final))]
            
            if df_final.empty:
                st.info("No records found with the applied filters.")
            else:
                st.dataframe(df_final)
                buffer = io.BytesIO()
                df_final.to_excel(buffer, index=False)
                st.download_button(label="Export Filtered Data",
                                   data=buffer,
                                   file_name=f"Training_Status_Filtered_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # ----- Página Visualization (Dashboard com Gráficos) -----
    elif page == "Visualization":
        st.header("Visualization Dashboard")
        if st.session_state.df_final is None:
            st.error("No processed data available for visualization. Go to the 'Report' page.")
        else:
            df_final = st.session_state.df_final
            # Pie Chart – Overall Status
            status_counts = df_final['status_final'].value_counts()
            labels = ['OK', 'Retreinamento', 'Not started']
            data = [status_counts.get(l, 0) for l in labels]
            fig1, ax1 = plt.subplots()
            ax1.pie(data, labels=labels, autopct='%1.1f%%', startangle=90)
            ax1.axis('equal')
            st.pyplot(fig1)
            
            # Bar Chart – Status by Position
            if 'cargo_pt_team' in df_final.columns and 'status_final' in df_final.columns:
                group = df_final.groupby(['cargo_pt_team', 'status_final']).size().unstack(fill_value=0)
                fig2, ax2 = plt.subplots(figsize=(8, 4))
                group.plot(kind='bar', ax=ax2)
                ax2.set_title("Status by Position")
                st.pyplot(fig2)
    
    # ----- Página Full Table (Exibição estilo Excel) -----
    elif page == "Full Table":
        st.header("Full Table")
        if st.session_state.df_final is None:
            st.error("No processed data. Go to the 'Report' page and process the data.")
        else:
            df_table = st.session_state.df_final.copy()
            st.markdown("### Global Filter (search all columns)")
            search_term = st.text_input("Enter search term:")
            if search_term:
                df_table = df_table[df_table.apply(lambda row: row.astype(str).str.contains(search_term, case=False).any(), axis=1)]
            st.dataframe(df_table)
            buffer = io.BytesIO()
            df_table.to_excel(buffer, index=False)
            st.download_button(label="Export Customized Data",
                               data=buffer,
                               file_name=f"Training_Status_Custom_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # ----- Página Saved Uploads (Uploads anteriores) -----
    elif page == "Saved Uploads":
        st.header("Saved Uploads")
        upload_dir = "uploaded_files"
        if not os.path.exists(upload_dir):
            st.error("No saved uploads found.")
        else:
            session_folders = [os.path.join(upload_dir, d) for d in os.listdir(upload_dir)
                               if os.path.isdir(os.path.join(upload_dir, d)) and os.path.exists(os.path.join(upload_dir, d, "final.xlsx"))]
            if not session_folders:
                st.info("No uploads with processed reports found.")
            else:
                sessions_dict = {}
                for folder in session_folders:
                    session_name = os.path.basename(folder)
                    try:
                        session_date = datetime.strptime(session_name, "%Y%m%d%H%M%S")
                        sessions_dict[folder] = session_date.strftime("%Y-%m-%d %H:%M:%S")
                    except Exception:
                        sessions_dict[folder] = "Unknown Date"
                
                selected_session = st.selectbox("Select an upload", list(sessions_dict.keys()), format_func=lambda x: sessions_dict[x])
                if selected_session:
                    final_file = os.path.join(selected_session, "final.xlsx")
                    if os.path.exists(final_file):
                        df_saved = pd.read_excel(final_file)
                        st.dataframe(df_saved)
                        buffer = io.BytesIO()
                        df_saved.to_excel(buffer, index=False)
                        st.download_button(label="Export Selected Upload Data",
                                           data=buffer,
                                           file_name=f"Report_{os.path.basename(selected_session)}.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    else:
                        st.error("final.xlsx file not found in the selected upload.")
    
    # ----- Página VCP (Persistent R & VCP Control) -----
    elifpage == "VCP":
        st.header("R & VCP Tracking")
        if st.session_state.df_final is None:
            st.error("No processed data available. Please process the report first in the 'Report' page.")
        else:
            # --- NOVA SEÇÃO: Importar tabela VCP via arquivo ---
            st.subheader("Importar Tabela VCP")
            st.markdown(
                "Caso você possua uma tabela com as colunas **Nome**, **Cargo**, **Procedimento** e **Data Concluída** (no formato YYYY-MM-DD), "
                "faça o upload do arquivo (CSV ou Excel). Essa tabela será utilizada para atualizar os campos de data e cargo."
            )
            vcp_table_file = st.file_uploader("Upload do arquivo (CSV ou XLSX)", type=["csv", "xlsx"], key="vcp_import")
            imported_df = None
            if vcp_table_file is not None:
                try:
                    if vcp_table_file.name.lower().endswith("csv"):
                        imported_df = pd.read_csv(vcp_table_file)
                    else:
                        imported_df = pd.read_excel(vcp_table_file)
                    
                    # Padroniza nomes das colunas removendo espaços em branco extras
                    imported_df.columns = [col.strip() for col in imported_df.columns]
                    # As colunas esperadas são: "Nome", "Cargo", "Procedimento", "Data Concluída"
                    required_columns = ["Nome", "Cargo", "Procedimento", "Data Concluída"]
                    if not all(col in imported_df.columns for col in required_columns):
                        st.error("O arquivo importado não contém todas as colunas necessárias: Nome, Cargo, Procedimento, Data Concluída")
                        imported_df = None
                    else:
                        st.success("Tabela importada com sucesso!")
                except Exception as e:
                    st.error(f"Erro ao processar o arquivo importado: {e}")
                    imported_df = None

            # --- Processamento dos dados para a aba VCP ---
            # Filtra os dados processados para registros contendo "R & VCP" (case-insensitive)
            df_vcp = st.session_state.df_final.copy()
            df_vcp = df_vcp[df_vcp['procedimento_nome'].str.contains(r"R\s*&\s*VCP", case=False, na=False)]
            if df_vcp.empty:
                st.info("No employees found for R & VCP.")
            else:
                # Cria DataFrame base a partir dos dados processados
                df_vcp_new = pd.DataFrame({
                    "Employee": df_vcp["Unisea E-learning User"],
                    "Position (English)": df_vcp.get("cargo_en_team", df_vcp["cargo_pt_team"]),
                    "Procedure Number": df_vcp["procedimento_num_assigned"],
                    "Date Completed": ""  # Campo que poderá ser atualizado manualmente ou via importação
                })
                df_vcp_new["Due Date"] = ""  # Calculada com base na "Date Completed"
                df_vcp_new["Reading"] = df_vcp["status_final"].apply(lambda x: "Completed" if str(x).lower() == "ok" else "Pending")
                df_vcp_new["Upload"] = ""  # Para informações de upload de arquivo

                # Se a tabela foi importada, atualiza os dados
                if imported_df is not None:
                    # Renomeia as colunas da tabela importada para casar com as colunas da tabela VCP:
                    # - "Nome" -> "Employee"
                    # - "Cargo" -> "Position (English)"
                    # - "Procedimento" -> "Procedure Number"
                    # - "Data Concluída" -> "Date Completed"
                    imported_df.rename(columns={
                        "Nome": "Employee",
                        "Cargo": "Position (English)",
                        "Procedimento": "Procedure Number",
                        "Data Concluída": "Date Completed"
                    }, inplace=True)
                    # Para simplificar o merge, definimos a combinação de Employee e Procedure Number como chave
                    df_vcp_new.set_index(["Employee", "Procedure Number"], inplace=True)
                    imported_df.set_index(["Employee", "Procedure Number"], inplace=True)
                    # Atualiza os valores existentes com os dados importados
                    df_vcp_new.update(imported_df)
                    df_vcp_new.reset_index(inplace=True)

                # Tenta carregar dados VCP previamente salvos (persistência externa)
                persisted_vcp = load_vcp_data()
                if persisted_vcp is not None:
                    df_vcp_new.set_index(["Employee", "Procedure Number"], inplace=True)
                    persisted_vcp.set_index(["Employee", "Procedure Number"], inplace=True)
                    merged_vcp = persisted_vcp.combine_first(df_vcp_new)
                    # Atualiza a coluna "Reading" com o valor calculado a partir dos dados processados
                    merged_vcp["Reading"] = df_vcp_new["Reading"]
                    merged_vcp.reset_index(inplace=True)
                    st.session_state.vcp_data = merged_vcp.copy()
                else:
                    st.session_state.vcp_data = df_vcp_new.copy()

                # Exibe a tabela para edição – o usuário pode alterar manualmente "Date Completed" (entre outras colunas)
                st.markdown("### R & VCP Table (edite 'Date Completed' conforme necessário no formato YYYY-MM-DD)")
                edited_df = st.data_editor(st.session_state.vcp_data, num_rows="dynamic", key="vcp_table")
                
                # Função para calcular a "Due Date" adicionando 730 dias à data concluída
                def calc_due_date(date_str):
                    try:
                        dt = datetime.strptime(date_str, "%Y-%m-%d")
                        due = dt + pd.Timedelta(days=730)
                        return due.strftime("%Y-%m-%d")
                    except Exception:
                        return ""
                
                edited_df["Due Date"] = edited_df["Date Completed"].apply(lambda x: calc_due_date(x) if x != "" else "")
                # Coluna "Status VCP": mostra "OK" se a "Due Date" ainda não venceu ou "Overdue" se já passou
                edited_df["Status VCP"] = edited_df["Due Date"].apply(
                    lambda d: "OK" if d != "" and datetime.strptime(d, "%Y-%m-%d").date() >= datetime.today().date() else ("Overdue" if d != "" else "")
                )
                
                st.markdown("### Updated R & VCP Table")
                st.dataframe(edited_df, height=500)
                
                # Seção para upload de arquivo para um empregado selecionado
                st.markdown("#### Upload File for Employee")
                selected_employee = st.selectbox("Select Employee", edited_df["Employee"].unique())
                uploaded_file = st.file_uploader("Drag and drop file here", type=["pdf", "docx", "xlsx"], key="vcp_upload")
                if uploaded_file is not None:
                    idx = edited_df.index[edited_df["Employee"] == selected_employee].tolist()
                    if idx:
                        edited_df.at[idx[0], "Upload"] = uploaded_file.name
                        st.success(f"File '{uploaded_file.name}' uploaded for {selected_employee}.")
                        st.session_state.vcp_data = edited_df.copy()
                
                # Botão para salvar alterações e persistir os dados no CSV
                if st.button("Save Table Changes"):
                    st.session_state.vcp_data = edited_df.copy()
                    save_vcp_data(edited_df)
                    st.success("Table changes saved!")
    
    # ----- Página Admin (Administração de Usuários) -----
    elif page == "Admin":
        st.header("User Administration")
        if st.session_state.username.lower() != "admin":
            st.error("Access restricted to administrators.")
        else:
            st.subheader("Register New User")
            new_username = st.text_input("New Username", key="new_user")
            new_password = st.text_input("New Password", key="new_pass", type="password")
            if st.button("Register User"):
                if new_username and new_password:
                    add_user(new_username, new_password)
                    st.success("User registered successfully!")
                else:
                    st.error("Please provide a username and password.")
            
            st.subheader("List of Users")
            df_users = get_all_users()
            st.dataframe(df_users)
            
            st.subheader("Delete User")
            users_list = df_users['username'].tolist()
            user_to_delete = st.selectbox("Select a user to delete", users_list)
            if st.button("Delete User"):
                if user_to_delete.lower() == "admin":
                    st.error("It is not allowed to delete the admin user.")
                else:
                    delete_user(user_to_delete)
                    st.success(f"User '{user_to_delete}' deleted successfully!")
                    st.experimental_rerun()

    # ----- Página History (Histórico de Relatórios) -----
    elif page == "History":
        st.header("Reports History")
        try:
            conn = sqlite3.connect(DB_PATH)
            df_history = pd.read_sql_query("SELECT * FROM report_history ORDER BY id DESC", conn)
            conn.close()
            if df_history.empty:
                st.info("No report logs found.")
            else:
                st.dataframe(df_history)
        except Exception as e:
            st.error(f"Error loading history: {e}")
