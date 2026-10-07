import streamlit as st
import pandas as pd
import unicodedata
import os
import smtplib
import pytz
import json
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Certificação Royal Canin São Paulo", page_icon="🐾", layout="centered", initial_sidebar_state="collapsed")

# CSS Customizado: Fundo branco, tema claro e tons azulados refinados
st.markdown("""
    <style>
        .stApp {
            background-color: #FFFFFF !important;
            color: #1E293B !important;
        }
        [data-testid="stSidebar"], [data-testid="stSidebarNav"], [data-testid="collapsedControl"] {
            display: none;
        }
        h1, h2, h3 {
            color: #E2001A !important;
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        }
        .stRadio label, .stCheckbox label, .stSelectbox label, p, span, div {
            color: #1E293B !important;
        }
        div[data-baseweb="select"] > div {
            background-color: #F8FAFC !important;
            color: #1E293B !important;
            border-color: #CBD5E1 !important;
            font-weight: bold !important;
        }
        div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"] {
            background-color: #FFFFFF !important;
        }
        div[data-baseweb="popover"] div[role="option"], ul[role="listbox"] li {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            font-weight: bold !important;
        }
        .custom-card {
            background-color: #F8FAFC;
            border-left: 5px solid #E2001A;
            padding: 14px;
            border-radius: 6px;
            margin-bottom: 15px;
            color: #1E293B;
        }
        .stButton>button {
            background-color: #E2001A !important;
            color: #FFFFFF !important;
            border-radius: 6px;
            border: none;
            font-weight: 600;
        }
        .stButton>button:hover {
            background-color: #B91C1C !important;
            color: #FFFFFF !important;
        }
    </style>
""", unsafe_allow_html=True)

menu = "📝 Área da Promotora"

# --- CARREGAR PLANILHA DE CLIENTES ---
@st.cache_data
def load_clients():
    try:
        return pd.read_excel("clientes.xlsx")
    except Exception as e:
        st.error(f"Erro ao carregar a planilha de clientes: {e}")
        return pd.DataFrame()

df_clientes = load_clients()

# --- CONEXÃO GOOGLE SHEETS ---
@st.cache_resource
def conectar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open("Historico_Auditorias_RoyalCanin")
    except Exception as e:
        print(f"Erro de conexão com Google Sheets: {e}")
        return None

def registrar_log_acesso(supervisor):
    try:
        wb = conectar_google_sheets()
        if wb:
            try:
                sheet_log = wb.worksheet("Logs_Acesso")
            except:
                sheet_log = wb.add_worksheet(title="Logs_Acesso", rows="100", cols="5")
                sheet_log.append_row(["Supervisor", "Data/Hora"])
            
            fuso_sp = pytz.timezone('America/Sao_Paulo')
            agora = datetime.now(fuso_sp).strftime("%d/%m/%Y %H:%M:%S")
            sheet_log.append_row([supervisor, agora])
    except Exception:
        pass

def listar_usuarios_cadastrados():
    usuarios_padrao = ["Benedito", "Caio", "Poli", "Rubens", "Daniel", "CDRC RIO PRETO", "CDRC SÃO JOÃO DA BOA VISTA"]
    try:
        wb = conectar_google_sheets()
        if wb:
            try:
                sheet_usr = wb.worksheet("Usuarios_Senhas")
                linhas = sheet_usr.get_all_values()
                usuarios_planilha = []
                for idx, linha in enumerate(linhas):
                    if idx == 0:
                        continue
                    if len(linha) >= 1 and linha[0].strip():
                        usuarios_planilha.append(linha[0].strip())
                if usuarios_planilha:
                    todos = list(dict.fromkeys(usuarios_padrao + usuarios_planilha))
                    return todos
            except:
                pass
    except:
        pass
    return usuarios_padrao

def buscar_senha_usuario(usuario):
    try:
        wb = conectar_google_sheets()
        if wb:
            try:
                sheet_usr = wb.worksheet("Usuarios_Senhas")
            except:
                return None
            
            linhas = sheet_usr.get_all_values()
            usuario_limpo = str(usuario).strip().lower()
            
            for idx, linha in enumerate(linhas):
                if idx == 0:
                    continue
                if len(linha) >= 2:
                    nome_planilha = str(linha[0]).strip().lower()
                    senha_planilha = str(linha[1]).strip()
                    if nome_planilha == usuario_limpo:
                        if senha_planilha and senha_planilha.lower() != "none" and senha_planilha != "":
                            return senha_planilha
    except Exception as e:
        st.error(f"Erro ao ler usuários: {e}")
    return None

def salvar_nova_senha(usuario, senha):
    try:
        wb = conectar_google_sheets()
        if wb:
            try:
                sheet_usr = wb.worksheet("Usuarios_Senhas")
            except:
                sheet_usr = wb.add_worksheet(title="Usuarios_Senhas", rows="50", cols="3")
                sheet_usr.append_row(["Usuário", "Senha"])
                sheet_usr = wb.worksheet("Usuarios_Senhas")
                
            linhas = sheet_usr.get_all_values()
            encontrado = False
            usuario_limpo = str(usuario).strip().lower()
            
            for idx, linha in enumerate(linhas, start=1):
                if idx == 1:
                    continue
                if len(linha) >= 1 and str(linha[0]).strip().lower() == usuario_limpo:
                    sheet_usr.update_cell(idx, 2, senha)
                    encontrado = True
                    break
            
            if not encontrado:
                sheet_usr.append_row([usuario.strip(), senha.strip()])
            return True
    except Exception as e:
        print(f"Erro ao salvar senha: {e}")
        return False
    return False

def enviar_email_interacao(loja, supervisor, tipo, comentario, data_auditoria):
    remetente = "beneditobandola@gmail.com"
    senha = "kfih ccqx cskn oito"
    destino = "benedito.bandola@minassal.com.br"
    
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destino
    msg['Subject'] = f"💬 Nova Interação da Gestão [{tipo}]: {loja}"
    
    corpo = f"""
    Olá!
    
    O supervisor {supervisor} registrou uma nova interação na loja {loja} (Certificação SP de {data_auditoria}).
    
    Tipo: {tipo}
    Comentário: "{comentario}"
    
    Atenciosamente,
    Sistema de Certificação Royal Canin SP.
    """
    msg.attach(MIMEText(corpo, 'plain'))
    try:
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(remetente, senha)
        s.sendmail(remetente, destino, msg.as_string())
        s.quit()
        return True
    except:
        return False

def salvar_comentario_planilha(loja, supervisor, comentario, tipo="Feedback", data_auditoria=""):
    try:
        wb = conectar_google_sheets()
        if wb:
            try:
                sheet_com = wb.worksheet("Comentarios_Gestao_SP")
            except:
                sheet_com = wb.add_worksheet(title="Comentarios_Gestao_SP", rows="100", cols="7")
                sheet_com.append_row(["Loja", "Supervisor", "Tipo", "Comentário", "Data_Auditoria", "Data/Hora_Interacao"])
            
            fuso_sp = pytz.timezone('America/Sao_Paulo')
            agora = datetime.now(fuso_sp).strftime("%d/%m/%Y %H:%M")
            sheet_com.append_row([loja, supervisor, tipo, comentario, data_auditoria, agora])
            enviar_email_interacao(loja, supervisor, tipo, comentario, data_auditoria)
            return True
    except:
        return False
    return False

def salvar_no_google_sheets(dados_auditoria):
    try:
        wb = conectar_google_sheets()
        if wb:
            try:
                sheet = wb.worksheet("Historico_CertificacaoSP")
            except:
                try:
                    sheet = wb.add_worksheet(title="Historico_CertificacaoSP", rows="100", cols="19")
                    sheet.append_row([
                        "Data/Hora", "CDRC", "Cidade", "Loja", "Nota Total", 
                        "Tipo Loja", "Plano Cão", "Plano Gato", "Plano Vet", 
                        "Super Premium Separada", "Super Premium Categoria", "Conservação", 
                        "Materiais JSON", "Pontos Extras", "Observações"
                    ])
                except:
                    sheet = wb.sheet1
            
            dados_str = [str(x) for x in dados_auditoria]
            sheet.append_row(dados_str)
            return True
    except Exception as e:
        print(f"Erro crítico ao salvar no Google Sheets: {e}")
        return False
    return False

def gerar_pdf_certificacao_sp(tipo_auditoria, promotora, loja, cidade, endereco, dados_completos, nota_total):
    loja_limpa = "".join([c for c in loja if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
    arq_completo = f"CERTIFICACAO_SP_{tipo_auditoria.upper().replace('-', '_')}_{loja_limpa}.pdf"

    estilos = getSampleStyleSheet()
    style_celula = ParagraphStyle('EstiloCelula', parent=estilos['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#1F2937'))
    style_celula_cab = ParagraphStyle('EstiloCelulaCab', parent=estilos['Normal'], fontSize=8, leading=10, textColor=colors.white, fontName="Helvetica-Bold")
    
    cor_cabecalho_principal = colors.HexColor('#990000') # Vermelho Royal Canin
    cor_cabecalho_secundario = colors.HexColor('#C53030') # Vermelho secundário
    
    fuso_sp = pytz.timezone('America/Sao_Paulo')
    agora = datetime.now(fuso_sp).strftime("%d/%m/%Y %H:%M")

    def construir_pdf():
        elem = []
        elem.append(Paragraph(f"<b>RELATÓRIO DE CERTIFICAÇÃO - {tipo_auditoria.upper()} (SÃO PAULO)</b>", estilos['Title']))
        elem.append(Spacer(1, 5))
        elem.append(Paragraph(f"<b>LOJA:</b> {loja} | <b>CIDADE:</b> {cidade}", estilos['Normal']))
        elem.append(Paragraph(f"<b>ENDEREÇO:</b> {endereco}", estilos['Normal']))
        elem.append(Paragraph(f"<b>CDRC RESPONSÁVEL:</b> {promotora} | <b>DATA/HORA:</b> {agora}", estilos['Normal']))
        elem.append(Paragraph(f"<b>NOTA FINAL DA CERTIFICAÇÃO:</b> <font color='#990000'><b>{nota_total:.2f} / 6.0 pts</b></font>", estilos['Heading2']))
        elem.append(Spacer(1, 10))

        # 1. PLANOGRAMAS E TIPO DE LOJA COM PONTOS GANHOS
        elem.append(Paragraph("<b>1. DADOS DA LOJA E PLANOGRAMAS</b>", estilos['Heading3']))
        
        p_cao_val = 1.0 if dados_completos.get('plano_cao') == "Sim" else 0.0
        p_gato_val = 1.0 if dados_completos.get('plano_gato') == "Sim" else 0.0
        p_
