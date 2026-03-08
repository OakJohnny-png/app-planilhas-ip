import streamlit as st
import pandas as pd
import io
import os
import requests
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font as xlFont, PatternFill, Alignment
from reportlab.lib.pagesizes import A4, landscape as rl_landscape, portrait as rl_portrait
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- FUNÇÃO PARA BAIXAR E REGISTRAR FONTES ---
def configurar_fontes_google():
    # URL da fonte Roboto (Regular e Bold) do Google Fonts
    fontes = {
        "Roboto-Regular.ttf": "https://github.com/google/fonts/raw/main/apache/roboto/static/Roboto-Regular.ttf",
        "Roboto-Bold.ttf": "https://github.com/google/fonts/raw/main/apache/roboto/static/Roboto-Bold.ttf"
    }
    
    fontes_disponiveis = ["Helvetica", "Times-Roman", "Courier"]
    
    for nome_arquivo, url in fontes.items():
        if not os.path.exists(nome_arquivo):
            try:
                response = requests.get(url)
                with open(nome_arquivo, 'wb') as f:
                    f.write(response.content)
            except:
                continue
    
    if os.path.exists("Roboto-Regular.ttf"):
        pdfmetrics.registerFont(TTFont('Roboto', 'Roboto-Regular.ttf'))
        pdfmetrics.registerFont(TTFont('Roboto-Bold', 'Roboto-Bold.ttf'))
        fontes_disponiveis.append("Roboto")
        
    return fontes_disponiveis

# --- INÍCIO DO APP ---
st.set_page_config(page_title="V.A.G.A.L.U.M.E. Pro", layout="wide")
fontes_lista = configurar_fontes_google()

st.sidebar.header("⚙️ Filtros e Estilo")
ano_sel = st.sidebar.number_input("Ano do Relatório:", value=datetime.now().year)
fonte_sel = st.sidebar.selectbox("Fonte do PDF", fontes_lista)
cor_principal = st.sidebar.color_picker("Cor das Rotas", "#1E3A8A")

# Índices das Colunas (Ajustados conforme sua planilha)
COL_DATA = 7      # Coluna H
COL_PROBLEMA = 1  # Coluna B
COL_STATUS = 3    # Coluna D

routes = {
    "ROTA 1": ["CENTRO", "JARDIM AMERICA"],
    "ROTA 2": ["ALBERTINA", "LARANJEIRAS", "BOA VISTA", "EUGENIO SCHNEIDER"],
    "ROTA 3": ["FUNDO CANOAS", "CANOAS", "PROGRESSO", "PAMPLONA", "CANTA GALO"],
    "ROTA 4": ["BARRA DO TROMBUDO", "BARRAGEM", "BUDAG", "SUMARE"],
    "ROTA 5": ["SANTANA", "TABOAO", "BREMER", "BELA ALIANÇA"],
    "ROTA 6": ["BARRA DA ITOUPAVA", "NAVEGANTES", "SANTA RITA", "VALADA ITOUPAVA", "VALADA SÃO PAULO", "RAINHA"]
}

uploaded_file = st.file_uploader("📥 Envie a planilha Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 Processar e Gerar Arquivos"):
        try:
            xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
            abas_disponiveis = {nome.strip().upper(): nome for nome in xls.keys()}
            
            data_by_route = {}
            total_encontrado = 0

            for route, neighborhoods in routes.items():
                route_data = {}
                for neighborhood in neighborhoods:
                    nome_upper = neighborhood.upper()
                    if nome_upper in abas_disponiveis:
                        df = xls[abas_disponiveis[nome_upper]].copy()
                        
                        if df.shape[1] <= COL_DATA: continue

                        # Filtro robusto
                        df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors='coerce')
                        mask = (
                            (df[COL_DATA].dt.year == ano_sel) & 
                            (df[COL_STATUS].astype(str).str.strip().str.upper().isin(['NÃO REALIZADO', 'NÃO EXECUTADO', 'NAO REALIZADO', 'NAO EXECUTADO']))
                        )
                        
                        items = df[mask][COL_PROBLEMA].dropna().astype(str).str.strip().tolist()
                        if items:
                            route_data[neighborhood] = items
                            total_encontrado += len(items)
                
                if route_data:
                    data_by_route[route] = route_data

            if total_encontrado == 0:
                st.warning(f"Nenhum dado de {ano_sel} encontrado na Coluna H com status pendente.")
            else:
                # --- PDF GERADO COM A NOVA FONTE ---
                pdf_buffer = io.BytesIO()
                doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)
                
                # Lógica para Negrito
                f_bold = "Helvetica-Bold"
                if fonte_sel == "Roboto": f_bold = "Roboto-Bold"
                elif fonte_sel == "Times-Roman": f_bold = "Times-Bold"
                
                style_header = ParagraphStyle('H', fontName=f_bold, fontSize=14, textColor=HexColor("#FFFFFF"), backColor=HexColor(cor_principal), alignment=1, padding=5)
                style_bairro = ParagraphStyle('B', fontName=f_bold, fontSize=11, leftIndent=10, spaceBefore=5)
                style_item = ParagraphStyle('I', fontName=fonte_sel, fontSize=10, leftIndent=20)

                elements = [Paragraph(f"ORDENS DE SERVIÇO - {ano_sel}", style_header), Spacer(1, 15)]
                for r, bairros in data_by_route.items():
                    elements.append(Paragraph(r, style_header))
                    for b, probs in bairros.items():
                        elements.append(Paragraph(b, style_bairro))
                        for p in probs:
                            elements.append(Paragraph(f"• {p}", style_item))
                    elements.append(Spacer(1, 10))

                doc.build(elements)
                pdf_buffer.seek(0)
                
                st.success(f"Sucesso! {total_encontrado} itens encontrados.")
                st.download_button("📄 Baixar PDF com Fonte Customizada", data=pdf_buffer, file_name="Relatorio_Vagalume.pdf")

        except Exception as e:
            st.error(f"Erro: {e}")
