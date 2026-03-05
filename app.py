import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font as xlFont, PatternFill, Alignment
from reportlab.lib.pagesizes import A4, landscape as rl_landscape, portrait as rl_portrait
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm

st.set_page_config(page_title="V.A.G.A.L.U.M.E. Pro", layout="wide")

# --- CONFIGURAÇÕES LATERAIS ---
st.sidebar.header("⚙️ Configurações de Filtro")
ano_selecionado = st.sidebar.number_input("Filtrar por Ano:", min_value=2020, max_value=2030, value=datetime.now().year)

# Mapeamento de Colunas (Ajuste aqui se precisar mudar outras)
col_data = 7  # Coluna H
col_problema = 1 # Coluna B (Descrição)
col_status = 3 # Coluna D (Status)

st.sidebar.header("🎨 Estilização")
fonte_escolhida = st.sidebar.selectbox("Fonte", ["Helvetica", "Times-Roman", "Courier"])
cor_fundo_rota = st.sidebar.color_picker("Cor das Rotas", "#1E3A8A")
cor_txt_rota = st.sidebar.color_picker("Texto das Rotas", "#FFFFFF")
cor_fundo_bairro = st.sidebar.color_picker("Fundo dos Bairros", "#D3D3D3")

# --- REGRAS DE ROTAS ---
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
    if st.button("🚀 Processar e Gerar Downloads"):
        with st.spinner('Lendo Coluna H e filtrando dados...'):
            try:
                xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
                abas_disponiveis = {nome.strip().upper(): nome for nome in xls.keys()}
                
                data_by_route = {}
                total_itens = 0

                for route, neighborhoods in routes.items():
                    route_data = {}
                    for neighborhood in neighborhoods:
                        nome_aba_upper = neighborhood.upper()
                        if nome_aba_upper in abas_disponiveis:
                            df = xls[abas_disponiveis[nome_aba_upper]].copy()
                            
                            # Verifica se a planilha tem pelo menos até a coluna H (8 colunas)
                            if df.shape[1] <= col_data:
                                continue

                            # CONVERSÃO DA COLUNA H (Índice 7)
                            df[col_data] = pd.to_datetime(df[col_data], errors='coerce')
                            
                            # FILTRAGEM
                            mask = (
                                (df[col_data].dt.year == ano_selecionado) & 
                                (df[col_status].astype(str).str.strip().str.upper().isin(['NÃO REALIZADO', 'NÃO EXECUTADO', 'NAO REALIZADO', 'NAO EXECUTADO'])) &
                                (df[col_problema].notna())
                            )
                            
                            problems = df[mask][col_problema].astype(str).str.strip().tolist()
                            
                            if problems:
                                route_data[neighborhood] = problems
                                total_itens += len(problems)
                    
                    if route_data:
                        data_by_route[route] = route_data

                if total_itens == 0:
                    st.warning(f"⚠️ Nenhum dado encontrado na Coluna H para o ano {ano_selecionado}.")
                else:
                    # --- GERAÇÃO DOS ARQUIVOS (EXCEL E PDF) ---
                    # (Mesma lógica de buffer e download anterior)
                    excel_buffer = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    row_idx = 1
                    for r_name, b_dict in data_by_route.items():
                        c = ws.cell(row=row_idx, column=1, value=r_name)
                        c.font = xlFont(bold=True, color=cor_txt_rota.replace('#','')); c.fill = PatternFill(start_color=cor_fundo_rota.replace('#',''), fill_type="solid")
                        row_idx += 1
                        for b_name, p_list in b_dict.items():
                            ws.cell(row=row_idx, column=1, value=f"📍 {b_name}").font = xlFont(bold=True)
                            row_idx += 1
                            for p in p_list:
                                ws.cell(row=row_idx, column=1, value=f"    • {p}")
                                row_idx += 1
                    wb.save(excel_buffer)
                    excel_buffer.seek(0)

                    pdf_buffer = io.BytesIO()
                    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)
                    style_r = ParagraphStyle('R', fontSize=14, backColor=HexColor(cor_fundo_rota), textColor=HexColor(cor_txt_rota), alignment=1)
                    elements = [Paragraph(f"RELATÓRIO {ano_selecionado}", style_r), Spacer(1, 10)]
                    for r_name, b_dict in data_by_route.items():
                        elements.append(Paragraph(r_name, style_r))
                        for b_name, p_list in b_dict.items():
                            elements.append(Paragraph(f"<b>{b_name}</b>", ParagraphStyle('B', leftIndent=10)))
                            for p in p_list:
                                elements.append(Paragraph(f"• {p}", ParagraphStyle('I', leftIndent=20)))
                    doc.build(elements)
                    pdf_buffer.seek(0)

                    st.success(f"✅ {total_itens} chamados encontrados!")
                    st.download_button("📥 Baixar Excel", data=excel_buffer, file_name="Relatorio.xlsx")
                    st.download_button("📄 Baixar PDF", data=pdf_buffer, file_name="Relatorio.pdf")

            except Exception as e:
                st.error(f"Erro: {e}")
