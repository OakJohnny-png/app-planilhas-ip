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
# Permitir que o usuário escolha o ano (evita arquivos em branco)
ano_selecionado = st.sidebar.number_input("Filtrar por Ano:", min_value=2020, max_value=2030, value=datetime.now().year)

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
        with st.spinner('Extraindo e validando dados...'):
            try:
                # Lendo o Excel
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
                            
                            if df.shape[1] < 4: continue

                            # CONVERSÃO SEGURA DE DATA
                            df[0] = pd.to_datetime(df[0], errors='coerce')
                            
                            # FILTRAGEM CORRIGIDA (Usando .str para evitar o erro de 'Series')
                            # Verificamos: Ano, Status (sem espaços e maiúsculo) e se a descrição existe
                            mask = (
                                (df[0].dt.year == ano_selecionado) & 
                                (df[3].astype(str).str.strip().str.upper().isin(['NÃO REALIZADO', 'NÃO EXECUTADO', 'NAO REALIZADO', 'NAO EXECUTADO'])) &
                                (df[1].notna())
                            )
                            
                            problems = df[mask][1].astype(str).str.strip().tolist()
                            
                            if problems:
                                route_data[neighborhood] = problems
                                total_itens += len(problems)
                    
                    if route_data:
                        data_by_route[route] = route_data

                if total_itens == 0:
                    st.warning(f"⚠️ Nenhum dado encontrado para o ano {ano_selecionado}. Tente alterar o ano na barra lateral ou verifique o status na planilha.")
                else:
                    # --- EXCEL ---
                    excel_buffer = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Pendências"
                    
                    row_idx = 1
                    for r_name, b_dict in data_by_route.items():
                        cell = ws.cell(row=row_idx, column=1, value=r_name)
                        cell.font = xlFont(bold=True, color=cor_txt_rota.replace('#',''))
                        cell.fill = PatternFill(start_color=cor_fundo_rota.replace('#',''), fill_type="solid")
                        row_idx += 1
                        
                        for b_name, p_list in b_dict.items():
                            c_b = ws.cell(row=row_idx, column=1, value=f"📍 {b_name}")
                            c_b.font = xlFont(bold=True)
                            c_b.fill = PatternFill(start_color=cor_fundo_bairro.replace('#',''), fill_type="solid")
                            row_idx += 1
                            for p in p_list:
                                ws.cell(row=row_idx, column=1, value=f"    • {p}")
                                row_idx += 1
                        row_idx += 1
                    
                    wb.save(excel_buffer)
                    excel_buffer.seek(0)

                    # --- PDF ---
                    pdf_buffer = io.BytesIO()
                    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, margin=15*mm)
                    
                    styles = ParagraphStyle('Normal', fontName=fonte_escolhida, fontSize=10)
                    style_r = ParagraphStyle('Rota', fontName=f"{fonte_escolhida}-Bold" if fonte_escolhida != "Helvetica" else "Helvetica-Bold", 
                                            fontSize=14, backColor=HexColor(cor_fundo_rota), textColor=HexColor(cor_txt_rota), 
                                            alignment=1, spaceAfter=10, padding=5)
                    style_b = ParagraphStyle('Bairro', fontName=f"{fonte_escolhida}-Bold" if fonte_escolhida != "Helvetica" else "Helvetica-Bold", 
                                            fontSize=11, backColor=HexColor(cor_fundo_bairro), leftIndent=5, spaceBefore=5)

                    elements = [Paragraph(f"RELATÓRIO DE MANUTENÇÃO - {ano_selecionado}", style_r), Spacer(1, 10)]
                    for r_name, b_dict in data_by_route.items():
                        elements.append(Paragraph(r_name, style_r))
                        for b_name, p_list in b_dict.items():
                            elements.append(Paragraph(b_name, style_b))
                            for p in p_list:
                                elements.append(Paragraph(f"&nbsp;&nbsp;&nbsp;&nbsp;• {p}", styles))
                            elements.append(Spacer(1, 4))
                    
                    doc.build(elements)
                    pdf_buffer.seek(0)

                    st.success(f"✅ Sucesso! {total_itens} chamados encontrados.")
                    col1, col2 = st.columns(2)
                    st.download_button("📥 Baixar Excel", data=excel_buffer, file_name=f"Manutencao_{ano_selecionado}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.download_button("📄 Baixar PDF", data=pdf_buffer, file_name=f"Manutencao_{ano_selecionado}.pdf", mime="application/pdf")

            except Exception as e:
                st.error(f"Erro no processamento: {e}")
