import streamlit as st
import pandas as pd
import unicodedata
import os
import smtplib
import pytz
import json
import tempfile
from datetime import datetime
from PIL import Image as PILImage, ImageOps
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Certificação Royal Canin São Paulo", page_icon="🐾", layout="centered", initial_sidebar_state="collapsed")

# CSS Customizado: Força o modo claro e limpa a cor do menu suspenso (dropdown)
st.markdown("""
    <style>
        /* Fundo principal da aplicação e textos */
        .stApp {
            background-color: #FFFFFF !important;
            color: #1E293B !important;
        }
        [data-testid="stSidebar"], [data-testid="stSidebarNav"], [data-testid="collapsedControl"] {
            display: none;
        }
        /* Títulos principais em azul corporativo */
        h1, h2, h3 {
            color: #1E3A8A !important;
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        }
        /* Rótulos e textos perfeitamente legíveis */
        .stRadio label, .stCheckbox label, .stSelectbox label, p, span, div {
            color: #1E293B !important;
        }
        /* Caixa de seleção (Selectbox) - Força fundo branco e texto escuro */
        div[data-baseweb="select"] > div {
            background-color: #F8FAFC !important;
            color: #1E293B !important;
            border-color: #CBD5E1 !important;
        }
        /* Dropdown/Lista suspensa aberta - Fundo branco e texto visível */
        div[data-baseweb="popover"] div[role="listbox"] {
            background-color: #FFFFFF !important;
            color: #1E293B !important;
        }
        div[data-baseweb="popover"] div[role="option"] {
            background-color: #FFFFFF !important;
            color: #1E293B !important;
        }
        div[data-baseweb="popover"] div[role="option"]:hover {
            background-color: #E2E8F0 !important;
            color: #0F172A !important;
        }
        /* Cartões de destaque */
        .custom-card {
            background-color: #F8FAFC;
            border-left: 5px solid #2563EB;
            padding: 14px;
            border-radius: 6px;
            margin-bottom: 15px;
            color: #1E293B;
        }
        /* Botões em azul corporativo */
        .stButton>button {
            background-color: #1D4ED8 !important;
            color: #FFFFFF !important;
            border-radius: 6px;
            border: none;
            font-weight: 600;
        }
        .stButton>button:hover {
            background-color: #1E40AF !important;
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

# --- PREPARAR FOTO APENAS PARA O PDF (Alta Qualidade) ---
def preparar_foto_pdf(caminho):
    try:
        img = PILImage.open(caminho)
        img = ImageOps.exif_transpose(img)
        img.thumbnail((800, 800))
        caminho_novo = caminho.replace(".jpg", "_pdf.jpg")
        img.save(caminho_novo, "JPEG", quality=85)
        return caminho_novo
    except:
        return caminho

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
                    sheet = wb.add_worksheet(title="Historico_CertificacaoSP", rows="100", cols="20")
                    sheet.append_row([
                        "Data/Hora", "CDRC", "Cidade", "Loja", "Nota Total", 
                        "Tipo Loja", "Plano Cão", "Plano Gato", "Plano Vet", 
                        "Super Premium Separada", "Super Premium Categoria", "Conservação", 
                        "Materiais JSON", "Pontos Extras", "Observações", "Fotos JSON"
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

def gerar_pdf_certificacao_sp(tipo_auditoria, promotora, loja, cidade, endereco, dados_completos, nota_total, caminhos_fotos=None):
    loja_limpa = "".join([c for c in loja if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
    arq_completo = f"CERTIFICACAO_SP_{tipo_auditoria.upper().replace('-', '_')}_{loja_limpa}.pdf"

    estilos = getSampleStyleSheet()
    style_celula = ParagraphStyle('EstiloCelula', parent=estilos['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#1F2937'))
    style_celula_cab = ParagraphStyle('EstiloCelulaCab', parent=estilos['Normal'], fontSize=8, leading=10, textColor=colors.white, fontName="Helvetica-Bold")
    
    cor_cabecalho_principal = colors.HexColor('#1E3A8A')
    cor_cabecalho_secundario = colors.HexColor('#0284C7')
    
    fuso_sp = pytz.timezone('America/Sao_Paulo')
    agora = datetime.now(fuso_sp).strftime("%d/%m/%Y %H:%M")

    def criar_pagina_1():
        elem = []
        elem.append(Paragraph(f"<b>RELATÓRIO DE CERTIFICAÇÃO - {tipo_auditoria.upper()} (SÃO PAULO)</b>", estilos['Title']))
        elem.append(Spacer(1, 5))
        elem.append(Paragraph(f"<b>LOJA:</b> {loja} | <b>CIDADE:</b> {cidade}", estilos['Normal']))
        elem.append(Paragraph(f"<b>ENDEREÇO:</b> {endereco}", estilos['Normal']))
        elem.append(Paragraph(f"<b>CDRC RESPONSÁVEL:</b> {promotora} | <b>DATA/HORA:</b> {agora}", estilos['Normal']))
        elem.append(Paragraph(f"<b>NOTA FINAL DA CERTIFICAÇÃO:</b> <font color='#1E3A8A'><b>{nota_total:.2f} / 6.0 pts</b></font>", estilos['Heading2']))
        elem.append(Spacer(1, 10))

        elem.append(Paragraph("<b>1. DADOS DA LOJA E PLANOGRAMAS</b>", estilos['Heading3']))
        data_exec = [
            [Paragraph("<b>Indicador / Critério</b>", style_celula_cab), Paragraph("<b>Resultado Registrado</b>", style_celula_cab)],
            [Paragraph("Tipo de Estabelecimento", style_celula), Paragraph(str(dados_completos.get('tipo_loja', '')), style_celula)],
            [Paragraph("Planograma de Cão", style_celula), Paragraph(str(dados_completos.get('plano_cao', '')), style_celula)],
            [Paragraph("Planograma de Gato", style_celula), Paragraph(str(dados_completos.get('plano_gato', '')), style_celula)],
            [Paragraph("Planograma Veterinary", style_celula), Paragraph(str(dados_completos.get('plano_vet', '')), style_celula)],
        ]
        t_exec = Table(data_exec, colWidths=[200, 340])
        t_exec.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), cor_cabecalho_principal),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))
        elem.append(t_exec)
        elem.append(Spacer(1, 10))

        elem.append(Paragraph("<b>2. SEPARAÇÃO E CONSERVAÇÃO</b>", estilos['Heading3']))
        data_sep = [
            [Paragraph("<b>Critério de Execução</b>", style_celula_cab), Paragraph("<b>Status</b>", style_celula_cab)],
            [Paragraph("Super Premium Cat está separada da linha FHN?", style_celula), Paragraph(str(dados_completos.get('sep_fhn', '')), style_celula)],
            [Paragraph("Super Premium Cat está na categoria Super Premium?", style_celula), Paragraph(str(dados_completos.get('cat_sp', '')), style_celula)],
            [Paragraph("Materiais bem executados e em bom estado de conservação?", style_celula), Paragraph(str(dados_completos.get('conservacao', '')), style_celula)],
        ]
        t_sep = Table(data_sep, colWidths=[350, 190])
        t_sep.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), cor_cabecalho_secundario),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))
        elem.append(t_sep)
        elem.append(Spacer(1, 10))

        elem.append(Paragraph("<b>3. MERCHANDISING E PONTOS EXTRAS</b>", estilos['Heading3']))
        mat_ativos = dados_completos.get('materiais_ativos', [])
        materiais_str = ", ".join(mat_ativos) if mat_ativos else "Nenhum material assinalado"
        
        data_merch = [
            [Paragraph("<b>Indicador</b>", style_celula_cab), Paragraph("<b>Resultado / Detalhe</b>", style_celula_cab)],
            [Paragraph("Materiais de Merchandising Presentes", style_celula), Paragraph(materiais_str, style_celula)],
            [Paragraph("Qtd de Pontos Extras Encontrados", style_celula), Paragraph(str(dados_completos.get('qtd_extras', 0)), style_celula)],
        ]
        t_merch = Table(data_merch, colWidths=[180, 360])
        t_merch.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), cor_cabecalho_principal),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))
        elem.append(t_merch)
        elem.append(Spacer(1, 10))

        if dados_completos.get('observacoes'):
            elem.append(Paragraph("<b>4. OBSERVAÇÕES</b>", estilos['Heading3']))
            t_obs = Table([[Paragraph(dados_completos['observacoes'], style_celula)]], colWidths=[540])
            t_obs.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F8FAFC')), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#94A3B8')), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
            elem.append(t_obs)
            elem.append(Spacer(1, 10))

        return elem

    def criar_pagina_2():
        elem = []
        if caminhos_fotos:
            elem.append(PageBreak())
            elem.append(Paragraph("<b>ANEXO: EVIDÊNCIAS FOTOGRÁFICAS</b>", estilos['Heading3']))
            elem.append(Spacer(1, 10))
            linhas_tabela = []
            par_atual = []
            for caminho in caminhos_fotos:
                caminho_otimizado = preparar_foto_pdf(caminho)
                img_element = RLImage(caminho_otimizado, width=200, height=200, kind='proportional')
                par_atual.append(img_element)
                if len(par_atual) == 2:
                    linhas_tabela.append(par_atual)
                    par_atual = []
            if par_atual:
                par_atual.append("")
                linhas_tabela.append(par_atual)
            
            if linhas_tabela:
                t_fotos = Table(linhas_tabela, colWidths=[270, 270])
                t_fotos.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
                elem.append(t_fotos)

        return elem

    def adicionar_rodape(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(colors.HexColor('#64748B'))
        canvas.drawRightString(A4[0] - 25, 15, "Desenvolvido por Benedito Bandola")
        canvas.restoreState()

    elementos_pdf = criar_pagina_1()
    if caminhos_fotos:
        elementos_pdf += criar_pagina_2()

    doc_completo = SimpleDocTemplate(arq_completo, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    doc_completo.build(elementos_pdf, onFirstPage=adicionar_rodape, onLaterPages=adicionar_rodape)

    return arq_completo

def enviar_email_auditoria(assunto, pdf_paths, destinatarios, corpo_email=""):
    remetente = "beneditobandola@gmail.com"
    senha = "kfih ccqx cskn oito"

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = ", ".join(destinatarios)
    msg['Subject'] = assunto

    if not corpo_email:
        corpo_email = "Olá,\n\nSegue em anexo o relatório executivo e as comprovações fotográficas da Certificação Royal Canin São Paulo.\n\nAtenciosamente,\nSistema de Certificação SP."

    msg.attach(MIMEText(corpo_email, 'plain', 'utf-8'))

    try:
        for pdf_path in pdf_paths:
            with open(pdf_path, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(pdf_path))
                part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_path))
                msg.attach(part)
        
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(remetente, senha)
        s.sendmail(remetente, destinatarios, msg.as_string())
        s.quit()
        return True
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
        return False

# ==============================================================
# FLUXO 1: ÁREA DA PROMOTORA
# ==============================================================
if menu == "📝 Área da Promotora":
    st.title("🐾 Certificação Royal Canin São Paulo")
    st.markdown("---")
    
    if df_clientes.empty:
        st.warning("A planilha de clientes não foi encontrada ou está vazia.")
    else:
        def normalize(text):
            if pd.isna(text): return ""
            nfkd = unicodedata.normalize('NFKD', str(text))
            return "".join([c for c in nfkd if not unicodedata.combining(c)]).upper().strip()

        df_clientes['CIDADE_NORM'] = df_clientes['CIDADE'].apply(normalize)

        st.subheader("1. Identificação e Tipo de Registro")
        
        col_tipo, col_prom = st.columns(2)
        with col_tipo:
            tipo_auditoria = st.radio("Selecione o Tipo de Registro:", ["Pré-Auditoria", "Auditoria"], horizontal=True)
        with col_prom:
            promotora = st.selectbox("Selecione o CDRC / Responsável:", ["Selecione...", "CDRC RIO PRETO", "CDRC SÃO JOÃO DA BOA VISTA"])

        if promotora != "Selecione...":
            cidades_map = {
                "CDRC RIO PRETO": [
                    "SAO JOSE DO RIO PRETO", "MIRASSOL", "OLIMPIA", "IBIRA", "POTIRENDABA", 
                    "JOSE BONIFACIO", "NEVES PAULISTA", "TABAPUA", "POLONI", "CEDRAL", 
                    "TANABI", "JACI", "BADY BASSITT", "IPIGUA", "NOVA GRANADA", "GUAPIACU", 
                    "MONTE APRAZIVEL", "URUPES", "MENDONCA", "BALSAMO", "UCHOA", "CATIGUA", 
                    "MIRASSOLANDIA"
                ],
                "CDRC SÃO JOÃO DA BOA VISTA": [
                    "RIBEIRAO PRETO", "BATATAIS", "SERRANA", "JARDINOPOLIS", "SERTAOZINHO", 
                    "CRAVINHOS", "PONTAL", "SAO SIMAO", "PITANGUEIRAS", "BRODOWSKI", 
                    "LUIS ANTONIO", "DUMONT", "SALES OLIVEIRA", "BARRINHA", "SERRA AZUL", 
                    "SAO CARLOS", "MATAO", "ARARAQUARA", "IBATE", "BOA ESPERANCA DO SUL", 
                    "AMERICO BRASILIENSE", "RINCAO", "DOURADO", "NOVA EUROPA", "RIBEIRAO BONITO", 
                    "SANTA LUCIA", "TABATINGA", "SAO JOSE DO RIO PARDO", "VARGEM GRANDE DO SUL", 
                    "MOCOCA", "DIVINOLANDIA", "CASA BRANCA", "SAO SEBASTIAO DA GRAMA", 
                    "SAO JOAO DA BOA VISTA", "TAPIRATIBA", "ITOBI", "AGUAS PRATA", "ALTINOPOLIS"
                ]
            }

            cidades_alvo = cidades_map.get(promotora, [])
            df_filtrado = df_clientes[df_clientes['CIDADE_NORM'].isin(cidades_alvo)]
            lojas_lista = sorted(df_filtrado['NOME'].dropna().unique().tolist())

            st.subheader("2. Seleção da Loja")
            if len(lojas_lista) > 0:
                loja_selecionada = st.selectbox("Selecione o Cliente / Loja:", lojas_lista)
                
                dados_loja = df_filtrado[df_filtrado['NOME'] == loja_selecionada].iloc[0]
                cidade_loja = dados_loja.get('CIDADE', '')
                endereco_loja = str(dados_loja.get('ENDEREÇO', 'Endereço não informado'))
                
                doc_bruto = str(dados_loja.get('DOCUMENTO', '')).strip()
                if doc_bruto.endswith('.0'):
                    doc_bruto = doc_bruto[:-2]
                
                nums_doc = "".join([c for c in doc_bruto if c.isdigit()])
                
                if len(nums_doc) == 14:
                    cnpj_formatado = f"{nums_doc[:2]}.{nums_doc[2:5]}.{nums_doc[5:8]}/{nums_doc[8:12]}-{nums_doc[12:]}"
                elif len(nums_doc) > 0 and len(nums_doc) < 14:
                    nums_doc = nums_doc.zfill(14)
                    cnpj_formatado = f"{nums_doc[:2]}.{nums_doc[2:5]}.{nums_doc[5:8]}/{nums_doc[8:12]}-{nums_doc[12:]}"
                else:
                    cnpj_formatado = "Não informado"
                
                st.markdown(f"""
                <div class="custom-card">
                    <div style="font-size: 15px; margin-bottom: 4px;">
                        🏢 <b>CNPJ:</b> <span style="color: #1D4ED8; font-weight: 900; font-size: 17px; font-family: monospace;">{cnpj_formatado}</span>
                    </div>
                    <div style="font-size: 13.5px; color: #475569;">
                        📍 <b>Cidade:</b> {cidade_loja} | <b>Endereço:</b> {endereco_loja}
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("---")
                st.subheader("3. Tipos de Lojas")
                tipo_loja = st.radio(
                    "Selecione o Tipo de Estabelecimento:",
                    ["Pet Shop", "Clínica Veterinária", "Pet Shop + Consultório", "Clínica + Pet Shop", "Agropecuária"],
                    key="tipo_loja"
                )

                st.markdown("---")
                st.subheader("4. Planogramas e Separação (RTM SP)")

                plano_cao = st.radio("2. Planograma de Cão (Peso: 1.0 pt)", ["Sim", "Não"], key="plano_cao")
                plano_gato = st.radio("3. Planograma de Gato (Peso: 1.0 pt)", ["Sim", "Não"], key="plano_gato")
                plano_vet = st.radio("4. Planograma Veterinary (Peso: 1.0 pt)", ["Sim", "Não"], key="plano_vet")
                
                sep_fhn = st.radio("5. Super Premium Cat está separada da linha FHN? (Peso: 0,5 pt)", ["Sim", "Não"], key="sep_fhn")
                cat_sp = st.radio("6. Super Premium Cat está na categoria Super Premium? (Peso: 0,5 pt)", ["Sim", "Não"], key="cat_sp")
                conservacao = st.radio("7. Os materiais estão bem executados e em bom estado de conservação? (Peso: 0,25 pt)", ["Sim", "Não"], key="conservacao")

                st.markdown("---")
                st.subheader("8. Merchandising (Materiais POP)")
                st.info("💡 Regra: >= 3 materiais = 0,75 pt | 2 materiais = 0,50 pt | 1 material = 0,25 pt")
                materiais = [
                    "Faixa de Gôndola", "Bobina Forração", "Display Carona", 
                    "Cartazete precificador", "Base de Sacarias (can base)", 
                    "Totem Silhueta", "Cubo", "Clip Strip", "Stopper", "Outros materiais"
                ]
                mat_presenca = {mat: st.checkbox(mat, key=f"mat_{mat}") for mat in materiais}

                st.markdown("---")
                st.subheader("9. Pontos Extras Presentes")
                st.info("💡 Regra: 3 = 1,0 pt | 2 = 0,50 pt | 1 = 0,25 pt | 0 = 0 pt")
                qtd_pontos_extras = st.number_input("Quantidade de Pontos Extras encontrados:", min_value=0, max_value=3, value=0)

                st.markdown("---")
                st.subheader("10. Registro Fotográfico")
                st.info("💡 As fotos serão anexadas diretamente no Relatório PDF com alta qualidade.")
                arquivos_fotos = st.file_uploader("Envie as fotos da certificação:", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
                
                caminhos_temporarios = []
                if arquivos_fotos:
                    st.write("📷 **Pré-visualização das fotos:**")
                    cols = st.columns(3)
                    for idx, foto in enumerate(arquivos_fotos):
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as f:
                            f.write(foto.read())
                            caminho_temp = f.name
                            caminhos_temporarios.append(caminho_temp)
                        
                        caminho_tela = preparar_foto_pdf(caminho_temp)
                        with cols[idx % 3]:
                            st.image(caminho_tela, caption=f"Foto {idx+1}", use_container_width=True)

                st.markdown("---")
                st.subheader("11. Observações e Comentários")
                observacoes_promotora = st.text_area("Digite aqui qualquer observação relevante sobre o PDV:")

                st.markdown("---")
                st.subheader("12. Opções de Envio")
                opcao_envio = st.radio(
                    "Selecione quem deve receber o relatório completo (PDF):",
                    ["Somente Benedito", "Toda a Equipe (Benedito, Poli, Caio, Daniel, Rubens)"]
                )
                
                if st.button("Finalizar, Salvar e Enviar Certificação", type="primary"):
                    
                    if opcao_envio == "Somente Benedito":
                        destinatarios = ["benedito.bandola@minassal.com.br"]
                    else:
                        destinatarios = [
                            "benedito.bandola@minassal.com.br",
                            "poli@minassal.com.br",
                            "caio.poli@minassal.com.br",
                            "daniel.santini@minassal.com.br",
                            "rubens.porfirio@minassal.com.br"
                        ]

                    # --- CÁLCULO DA NOTA EXATO CONFORME REGRAS ---
                    nota_total = 0.0

                    if plano_cao == "Sim": nota_total += 1.0
                    if plano_gato == "Sim": nota_total += 1.0
                    if plano_vet == "Sim": nota_total += 1.0

                    if sep_fhn == "Sim": nota_total += 0.5
                    if cat_sp == "Sim": nota_total += 0.5

                    if conservacao == "Sim": nota_total += 0.25

                    materiais_ativos_lista = [m for m in materiais if mat_presenca[m]]
                    total_materiais = len(materiais_ativos_lista)
                    p_merch = 0.0
                    if total_materiais >= 3: p_merch = 0.75
                    elif total_materiais == 2: p_merch = 0.50
                    elif total_materiais == 1: p_merch = 0.25
                    nota_total += p_merch

                    p_extras = 0.0
                    if qtd_pontos_extras >= 3: p_extras = 1.0
                    elif qtd_pontos_extras == 2: p_extras = 0.50
                    elif qtd_pontos_extras == 1: p_extras = 0.25
                    nota_total += p_extras

                    dados_completos = {
                        'tipo_loja': tipo_loja,
                        'plano_cao': plano_cao, 'plano_gato': plano_gato, 'plano_vet': plano_vet,
                        'sep_fhn': sep_fhn, 'cat_sp': cat_sp, 'conservacao': conservacao,
                        'materiais_ativos': materiais_ativos_lista, 'qtd_extras': qtd_pontos_extras,
                        'observacoes': observacoes_promotora.strip()
                    }

                    fuso_sp = pytz.timezone('America/Sao_Paulo')
                    data_atual = datetime.now(fuso_sp).strftime("%d/%m/%Y %H:%M:%S")
                    
                    corpo_email = f"Olá,\n\nSegue o relatório executivo e as comprovações fotográficas referente à {tipo_auditoria.upper()} - Certificação Royal Canin SP da loja {loja_selecionada} ({cidade_loja}).\n\n"
                    corpo_email += f"📊 NOTA FINAL ATUAL: {nota_total:.2f} / 6.0 pts\n\n"
                    corpo_email += f"• CDRC Responsável: {promotora}\n"
                    corpo_email += f"• Tipo de Loja: {tipo_loja}\n"
                    corpo_email += f"• Planograma Cão: {plano_cao} | Gato: {plano_gato} | Vet: {plano_vet}\n"
                    corpo_email += f"• Super Premium Separada: {sep_fhn} | Categoria SP: {cat_sp}\n"
                    corpo_email += f"• Conservação: {conservacao} | Materiais POP: {total_materiais} | Pontos Extras: {qtd_pontos_extras}\n\n"
                    corpo_email += f"Observações: {observacoes_promotora.strip() if observacoes_promotora.strip() else 'Nenhuma.'}\n\n"
                    corpo_email += "Atenciosamente,\nSistema de Certificação Royal Canin SP."

                    with st.spinner("Gerando PDF e enviando e-mail..."):
                        pdf_path = gerar_pdf_certificacao_sp(
                            tipo_auditoria, promotora, loja_selecionada, cidade_loja, endereco_loja,
                            dados_completos, nota_total, caminhos_temporarios
                        )

                        dados_planilha = [
                            data_atual, promotora, cidade_loja, loja_selecionada, f"{nota_total:.2f}",
                            tipo_loja, plano_cao, plano_gato, plano_vet, sep_fhn, cat_sp, conservacao,
                            json.dumps(materiais_ativos_lista), qtd_pontos_extras,
                            observacoes_promotora.strip(), json.dumps([])
                        ]

                        salvar_no_google_sheets(dados_planilha)
                        
                        assunto = f"📋 Certificação SP [{tipo_auditoria}]: {loja_selecionada} ({cidade_loja}) - Nota {nota_total:.2f}"
                        enviado = enviar_email_auditoria(assunto, [pdf_path], destinatarios, corpo_email)

                        if enviado:
                            st.success(f"✅ Certificação finalizada com sucesso! Relatório enviado para os e-mails da gestão.")
                            with open(pdf_path, "rb") as f:
                                st.download_button("📥 Baixar PDF Gerado", data=f, file_name=os.path.basename(pdf_path), mime="application/pdf")
                        else:
                            st.warning("⚠️ Dados salvos na planilha, mas houve falha no envio do e-mail.")

# ==============================================================
# FLUXO 2: HISTÓRICO & FEEDBACKS
# ==============================================================
elif menu == "📋 Histórico & Feedbacks":
    st.title("📋 Histórico de Certificações & Feedbacks")
    st.markdown("---")
    
    wb = conectar_google_sheets()
    if wb:
        try:
            sheet = wb.worksheet("Historico_CertificacaoSP")
            dados = sheet.get_all_records()
            if dados:
                df_hist = pd.DataFrame(dados)
                st.dataframe(df_hist.tail(20), use_container_width=True)
            else:
                st.info("Nenhuma certificação registrada até o momento.")
        except Exception as e:
            st.error(f"Erro ao ler histórico: {e}")
    else:
        st.error("Não foi possível conectar ao Google Sheets.")

# ==============================================================
# FLUXO 3: PAINEL ADMIN
# ==============================================================
elif menu == "🔒 Painel Admin (Supervisores)":
    st.title("🔒 Painel Administrativo de Supervisores")
    st.markdown("---")
    
    usuarios_disp = listar_usuarios_cadastrados()
    usuario_sel = st.selectbox("Selecione seu Usuário:", ["Selecione..."] + usuarios_disp)
    
    if usuario_sel != "Selecione...":
        senha_correta = buscar_senha_usuario(usuario_sel)
        
        if not senha_correta:
            st.warning("Primeiro acesso deste usuário! Cadastre sua senha abaixo:")
            nova_senha = st.text_input("Defina sua nova senha:", type="password")
            confirma_senha = st.text_input("Confirme sua nova senha:", type="password")
            if st.button("Salvar Nova Senha"):
                if nova_senha and nova_senha == confirma_senha:
                    if salvar_nova_senha(usuario_sel, nova_senha):
                        st.success("Senha cadastrada com sucesso! Recarregue a página para fazer login.")
                    else:
                        st.error("Erro ao salvar a senha na planilha.")
                else:
                    st.error("As senhas digitadas não coincidem ou estão vazias.")
        else:
            senha_digitada = st.text_input("Digite sua senha:", type="password")
            if st.button("Entrar"):
                if senha_digitada == senha_correta:
                    st.session_state[f"auth_{usuario_sel}"] = True
                    registrar_log_acesso(usuario_sel)
                    st.success(f"Bem-vindo(a), {usuario_sel}!")
                else:
                    st.error("Senha incorreta.")

        if st.session_state.get(f"auth_{usuario_sel}", False):
            st.markdown("---")
            st.subheader("💬 Registrar Feedback / Ação para Loja")
            
            wb = conectar_google_sheets()
            if wb:
                try:
                    sheet = wb.worksheet("Historico_CertificacaoSP")
                    registros = sheet.get_all_records()
                    if registros:
                        df_aud = pd.DataFrame(registros)
                        lojas_aud = df_aud['Loja'].dropna().unique().tolist()
                        
                        loja_comentar = st.selectbox("Selecione a Loja Auditada:", lojas_aud)
                        tipo_interacao = st.selectbox("Tipo de Registro:", ["Feedback da Visita", "Plano de Ação Corretiva", "Observação Geral"])
                        comentario_texto = st.text_area("Descreva a orientação ou retorno para a equipe:")
                        
                        if st.button("Salvar e Notificar por E-mail"):
                            if comentario_texto.strip():
                                sucesso = salvar_comentario_planilha(loja_comentar, usuario_sel, comentario_texto, tipo_interacao)
                                if sucesso:
                                    st.success("✅ Feedback registrado e notificação por e-mail enviada com sucesso!")
                                else:
                                    st.error("Erro ao registrar feedback.")
                            else:
                                st.warning("Digite um comentário antes de enviar.")
                    else:
                        st.info("Nenhuma certificação cadastrada para receber feedbacks.")
                except Exception as e:
                    st.error(f"Erro ao buscar histórico: {e}")

# Assinatura de autoria na tela do programa (interface do Streamlit)
st.markdown("<br><hr><p style='text-align: center; color: #64748B; font-size: 11px;'>Desenvolvido por Benedito Bandola</p>", unsafe_allow_html=True)
