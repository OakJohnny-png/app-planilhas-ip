import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font as xlFont, PatternFill, Border, Side, Alignment
from reportlab.lib.pagesizes import A4, landscape as rl_landscape, portrait as rl_portrait
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm

# Configuração Inicial
st.set_page_config(page_title="V.A.G.A.L.U.M.E. Pro", layout="wide")
ano_atual = datetime.now().year

st.title("💡 V.A.G.A.L.U.M.E. - Relatórios de Iluminação")
st.markdown(f"Gerador de ordens de serviço para o ano de **{ano_atual}**.")

# --- BARRA LATERAL APERFEIÇOADA ---
st.sidebar.header("🎨 Configurações de Formatação")

with st.sidebar.expander("📝 Estilo de Texto", expanded=True):
    fonte_escolhida = st.sidebar.selectbox("Fonte", ["Helvetica", "Times-Roman", "Courier"])
    tam_fonte_rota = st.sidebar.slider("Tamanho Fonte Rota", 12, 24, 16)
    tam_fonte_bairro = st.sidebar.slider("Tamanho Fonte Bairro", 10, 18, 13)
    tam_fonte_item = st.sidebar.slider("Tamanho Fonte Itens", 8, 14, 11)

with st.sidebar.expander("🌈 Cores do Relatório", expanded=True):
    cor_fundo_rota = st.sidebar.color_picker("Fundo das Rotas", "#1E3A8A")
    cor_txt_rota = st.sidebar.color_picker("Texto das Rotas", "#FFFFFF")
    cor_fundo_bairro = st.sidebar.color_picker("Fundo dos Bairros", "#F3F4F6")

with st.sidebar.expander("📄 Configuração da Página"):
    orientacao = st.sidebar.radio("Orientação do PDF", ["Retrato", "Paisagem"])
    margem_val = st.sidebar.number_input("Margens (mm)", value=15)

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
    if st.button("🚀 Processar e Gerar Arquivos"):
        with st.spinner('Lendo dados e formatando documentos...'):
            try:
                # 1. Leitura dos Dados
                xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
                abas_disponiveis = {nome.strip().upper(): nome for nome in xls.keys()}
                data_by_route = {r: {} for r in routes}

                for route, neighborhoods in routes.items():
                    for neighborhood in neighborhoods:
                        nome_aba_upper = neighborhood.upper()
                        if nome_aba_upper in abas_disponiveis:
                            df = xls[abas_disponiveis[nome_aba_upper]]
                            if df.shape[1] < 4: continue

                            df[0] = pd.to_datetime(df[0], errors='coerce')
                            mask = (df[0].dt.year == ano_atual) & \
                                   (df[3].astype(str).str.upper().isin(['NÃO REALIZADO', 'NÃO EXECUTADO']))
                            
                            problems = df[mask][1].dropna().tolist()
                            if problems:
                                data_by_route[route][neighborhood] = problems

                # 2. GERAÇÃO DO EXCEL
                excel_buffer = io.BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Pendências"
                
                # Estilos Excel
                ft_rota = xlFont(name=fonte_escolhida, size=tam_fonte_rota, bold=True, color=cor_txt_rota.replace('#',''))
                fill_rota = PatternFill(start_color=cor_fundo_rota.replace('#',''), fill_type="solid")
                ft_bairro = xlFont(name=fonte_escolhida, size=tam_fonte_bairro, bold=True)
                fill_bairro = PatternFill(start_color=cor_fundo_bairro.replace('#',''), fill_type="solid")
                
                row = 1
                for route, neighborhoods in data_by_route.items():
                    if not neighborhoods: continue
                    c = ws.cell(row=row, column=1, value=route)
                    c.font = ft_rota; c.fill = fill_rota; row += 1
                    
                    for bairro, problems in neighborhoods.items():
                        c = ws.cell(row=row, column=1, value=f"  {bairro}")
                        c.font = ft_bairro; c.fill = fill_bairro; row += 1
                        for p in problems:
                            ws.cell(row=row, column=1, value=f"    • {p}")
                            row += 1
                    row += 1 # Espaço entre rotas

                wb.save(excel_buffer)
                excel_buffer.seek(0)

                # 3. GERAÇÃO DO PDF
                pdf_buffer = io.BytesIO()
                pg_size = rl_landscape(A4) if orientacao == "Paisagem" else rl_portrait(A4)
                doc = SimpleDocTemplate(pdf_buffer, pagesize=pg_size, 
                                        leftMargin=margem_val*mm, rightMargin=margem_val*mm, 
                                        topMargin=margem_val*mm, bottomMargin=margem_val*mm)
                
                # Estilos PDF
                style_rota = ParagraphStyle('Rota', fontName=f"{fonte_escolhida}-Bold" if fonte_escolhida != "Helvetica" else "Helvetica-Bold", 
                                            fontSize=tam_fonte_rota, textColor=HexColor(cor_txt_rota), 
                                            backColor=HexColor(cor_fundo_rota), alignment=1, spaceAfter=10, padding=5)
                
                style_bairro = ParagraphStyle('Bairro', fontName=f"{fonte_escolhida}-Bold" if fonte_escolhida != "Helvetica" else "Helvetica-Bold", 
                                              fontSize=tam_fonte_bairro, backColor=HexColor(cor_fundo_bairro), spaceBefore=5, leftIndent=10)
                
                style_item = ParagraphStyle('Item', fontName=fonte_escolhida, fontSize=tam_fonte_item, leftIndent=20, spaceAfter=2)

                elements = [Paragraph(f"RELATÓRIO DE MANUTENÇÃO - {ano_atual}", style_rota), Spacer(1, 10)]

                for route, neighborhoods in data_by_route.items():
                    if not neighborhoods: continue
                    elements.append(Paragraph(route, style_rota))
                    for bairro, problems in neighborhoods.items():
                        elements.append(Paragraph(bairro, style_bairro))
                        for p in problems:
                            elements.append(Paragraph(f"• {p}", style_item))
                        elements.append(Spacer(1, 5))

                doc.build(elements)
                pdf_buffer.seek(0)

                # 4. EXIBIÇÃO DOS BOTÕES
                st.success("✅ Relatórios processados com sucesso!")
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(label="📥 Baixar Excel", data=excel_buffer, 
                                       file_name=f"Vagalume_{ano_atual}.xlsx", 
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with col2:
                    st.download_button(label="📄 Baixar PDF", data=pdf_buffer, 
                                       file_name=f"Vagalume_{ano_atual}.pdf", 
                                       mime="application/pdf")

            except Exception as e:
                st.error(f"Ocorreu um erro no processamento: {e}")

