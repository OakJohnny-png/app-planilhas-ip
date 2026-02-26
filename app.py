import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font as xlFont, PatternFill, Border, Side, Alignment
from reportlab.lib.pagesizes import A4, landscape as rl_landscape, portrait as rl_portrait
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm

st.set_page_config(page_title="V.A.G.A.L.U.M.E.", layout="wide")

# --- LOGO RESPONSIVA NO LUGAR DO TÍTULO ---
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception:
        st.title("💡 V.A.G.A.L.U.M.E.")

# --- BARRA LATERAL PARA FORMATAÇÕES ---
st.sidebar.header("🎨 Cores e Fontes")
fonte_escolhida = st.sidebar.selectbox("Estilo da Fonte", ["Helvetica", "Times-Roman", "Courier"])

st.sidebar.subheader("1. Rotas")
cor_fundo_rota = st.sidebar.color_picker("Cor de Fundo (Rotas)", "#FF0000")
cor_fonte_rota = st.sidebar.color_picker("Cor da Fonte (Rotas)", "#FFFFFF")
tamanho_rota = st.sidebar.number_input("Tamanho da Fonte (Rotas)", min_value=8, max_value=36, value=16)

st.sidebar.subheader("2. Bairros")
cor_fundo_bairro = st.sidebar.color_picker("Cor de Fundo (Bairros)", "#D3D3D3")
cor_fonte_bairro = st.sidebar.color_picker("Cor da Fonte (Bairros)", "#000000")
tamanho_bairro = st.sidebar.number_input("Tamanho da Fonte (Bairros)", min_value=8, max_value=36, value=14)

st.sidebar.subheader("3. Problemas")
cor_fundo_prob = st.sidebar.color_picker("Cor de Fundo (Problemas)", "#FFFFFF")
cor_fonte_prob = st.sidebar.color_picker("Cor da Fonte (Problemas)", "#000000")
tamanho_prob = st.sidebar.number_input("Tamanho da Fonte (Problemas)", min_value=8, max_value=36, value=12)

# --- CONFIGURAÇÕES DO PDF ---
st.sidebar.header("📄 Configurações do PDF")
orientacao_pdf = st.sidebar.radio("Orientação da Página", ["Retrato (Em pé)", "Paisagem (Deitada)"])
margem_sup = st.sidebar.number_input("Margem Superior (mm)", value=20)
margem_inf = st.sidebar.number_input("Margem Inferior (mm)", value=20)
margem_esq = st.sidebar.number_input("Margem Esquerda (mm)", value=15)
margem_dir = st.sidebar.number_input("Margem Direita (mm)", value=15)

# --- ROTAS DEFINIDAS PELA REGRA 2 ---
routes = {
    "ROTA 1": ["CENTRO", "JARDIM AMERICA"],
    "ROTA 2": ["ALBERTINA", "LARANJEIRAS", "BOA VISTA", "EUGENIO SCHNEIDER"],
    "ROTA 3": ["FUNDO CANOAS", "CANOAS", "PROGRESSO", "PAMPLONA", "CANTA GALO"],
    "ROTA 4": ["BARRA DO TROMBUDO", "BARRAGEM", "BUDAG", "SUMARE"],
    "ROTA 5": ["SANTANA", "TABOAO", "BREMER", "BELA ALIANÇA"],
    "ROTA 6": ["BARRA DA ITOUPAVA", "NAVEGANTES", "SANTA RITA", "VALADA ITOUPAVA", "VALADA SÃO PAULO", "RAINHA"]
}

st.write("Faça o upload da planilha Excel com as abas e o sistema irá gerar seu relatório e um Gráfico Interativo no Excel.")

uploaded_file = st.file_uploader("📥 Envie sua planilha Excel (.xlsx) aqui", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀 Processar e Gerar Relatórios"):
        with st.spinner('Lendo dados e gerando relatórios...'):
            try:
                xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, dtype=str)
                abas_disponiveis = {nome.strip().upper(): nome for nome in xls.keys()}
                
                data_by_route = {r: {} for r in routes}
                rotas_com_problemas = []

                # Filtragem
                for route, neighborhoods in routes.items():
                    tem_problema_na_rota = False
                    for neighborhood in neighborhoods:
                        nome_aba_upper = neighborhood.upper()
                        if nome_aba_upper in abas_disponiveis:
                            nome_real_aba = abas_disponiveis[nome_aba_upper]
                            df = xls[nome_real_aba]
                            if len(df.columns) >= 4:
                                df_filtered = df[(df[3].isin(['NÃO REALIZADO', 'NÃO EXECUTADO'])) & (df[1].notna()) & (df[1].str.strip() != "")]
                                problems = df_filtered[1].tolist()
                                if problems:
                                    data_by_route[route][neighborhood] = problems
                                    tem_problema_na_rota = True
                    if tem_problema_na_rota:
                        rotas_com_problemas.append(route)

                # --- GERAÇÃO DO EXCEL ---
                wb = Workbook()
                ws = wb.active
                ws.title = "Chamados Pendentes"
                ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

                font_rota = xlFont(name="Helvetica", size=tamanho_rota, color=cor_fonte_rota.replace('#', ''), bold=True)
                fill_rota = PatternFill(start_color=cor_fundo_rota.replace('#', ''), end_color=cor_fundo_rota.replace('#', ''), fill_type="solid")
                font_bairro = xlFont(name="Helvetica", size=tamanho_bairro, color=cor_fonte_bairro.replace('#', ''), bold=True)
                fill_bairro = PatternFill(start_color=cor_fundo_bairro.replace('#', ''), end_color=cor_fundo_bairro.replace('#', ''), fill_type="solid")
                font_prob = xlFont(name="Helvetica", size=tamanho_prob, color=cor_fonte_prob.replace('#', ''))
                fill_prob = PatternFill(start_color=cor_fundo_prob.replace('#', ''), end_color=cor_fundo_prob.replace('#', ''), fill_type="solid")
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                current_row = 1
                for route, neighborhoods in data_by_route.items():
                    if not neighborhoods: continue
                    cell = ws.cell(row=current_row, column=1, value=route)
                    cell.font = font_rota; cell.fill = fill_rota; cell.border = thin_border; cell.alignment = Alignment(wrap_text=True)
                    current_row += 1
                    
                    first_bairro = True
                    for bairro, problems in neighborhoods.items():
                        if not first_bairro: current_row += 1
                        first_bairro = False
                        cell = ws.cell(row=current_row, column=1, value=bairro)
                        cell.font = font_bairro; cell.fill = fill_bairro; cell.border = thin_border; cell.alignment = Alignment(wrap_text=True)
                        current_row += 1
                        for problem in problems:
                            # Coluna A = Problema Visível
                            cell = ws.cell(row=current_row, column=1, value=str(problem).strip())
                            cell.font = font_prob; cell.border = thin_border; cell.alignment = Alignment(wrap_text=True)
                            if cor_fundo_prob != "#FFFFFF": cell.fill = fill_prob
                            
                            # Coluna B = Rota Oculta (Usada para o gráfico contar sozinho)
                            ws.cell(row=current_row, column=2, value=route)
                            current_row += 1

                ws.column_dimensions['A'].width = 150
                ws.column_dimensions['B'].hidden = True # Esconde a coluna B de controle

                # --- CRIANDO DADOS PARA O GRÁFICO NATIVO ---
                if rotas_com_problemas:
                    # Cria uma aba oculta para armazenar o resumo dinâmico
                    ws_resumo = wb.create_sheet(title="Resumo_Grafico")
                    ws_resumo.sheet_state = 'hidden'
                    
                    for idx, route in enumerate(rotas_com_problemas, start=1):
                        ws_resumo.cell(row=idx, column=1, value=route) # Nome da Rota
                        # Fórmula NATIVA do Excel que conta os problemas automaticamente
                        ws_resumo.cell(row=idx, column=2, value=f"=COUNTIF('Chamados Pendentes'!B:B, A{idx})")

                    # Criando o gráfico de Barras
                    chart = BarChart()
                    chart.type = "col"
                    chart.style = 10
                    chart.title = "Prioridade: Quantidade de Chamados por Rota"
                    chart.y_axis.title = "Nº de Chamados Pendentes"
                    chart.width = 18
                    chart.height = 9
                    chart.legend = None # Remove a legenda que fica redundante

                    # Linkando o gráfico com as fórmulas ocultas
                    data = Reference(ws_resumo, min_col=2, min_row=1, max_row=len(rotas_com_problemas))
                    cats = Reference(ws_resumo, min_col=1, min_row=1, max_row=len(rotas_com_problemas))
                    chart.add_data(data, titles_from_data=False)
                    chart.set_categories(cats)

                    # Anexando o gráfico na planilha principal
                    celula_ancora = f"A{current_row + 2}"
                    ws.add_chart(chart, celula_ancora)

                excel_output = io.BytesIO()
                wb.save(excel_output)
                excel_output.seek(0)

                # --- GERAÇÃO DO PDF (LIMPO, SEM GRÁFICO) ---
                pdf_output = io.BytesIO()
                pagesize = rl_landscape(A4) if "Paisagem" in orientacao_pdf else rl_portrait(A4)
                
                doc = SimpleDocTemplate(
                    pdf_output, pagesize=pagesize,
                    rightMargin=margem_dir * mm, leftMargin=margem_esq * mm,
                    topMargin=margem_sup * mm, bottomMargin=margem_inf * mm
                )

                story = []
                
                if fonte_escolhida == "Times-Roman": fonte_negrito = "Times-Bold"
                elif fonte_escolhida == "Courier": fonte_negrito = "Courier-Bold"
                else: fonte_negrito = "Helvetica-Bold"

                estilo_rota = ParagraphStyle('Rota', fontName=fonte_negrito, fontSize=tamanho_rota, 
                                             textColor=HexColor(cor_fonte_rota), backColor=HexColor(cor_fundo_rota), 
                                             spaceAfter=6, spaceBefore=12, padding=5)
                
                estilo_bairro = ParagraphStyle('Bairro', fontName=fonte_negrito, fontSize=tamanho_bairro, 
                                               textColor=HexColor(cor_fonte_bairro), backColor=HexColor(cor_fundo_bairro), 
                                               spaceAfter=3, spaceBefore=6, padding=4)
                
                estilo_prob = ParagraphStyle('Prob', fontName=fonte_escolhida, fontSize=tamanho_prob, 
                                             textColor=HexColor(cor_fonte_prob), backColor=HexColor(cor_fundo_prob) if cor_fundo_prob != "#FFFFFF" else None,
                                             spaceAfter=2, leading=tamanho_prob + 4)

                for route, neighborhoods in data_by_route.items():
                    if not neighborhoods: continue
                    story.append(Paragraph(route, estilo_rota))
                    
                    for bairro, problems in neighborhoods.items():
                        story.append(Paragraph(bairro, estilo_bairro))
                        for problem in problems:
                            story.append(Paragraph(f"• {str(problem).strip()}", estilo_prob))
                        story.append(Spacer(1, 4 * mm))

                doc.build(story)
                pdf_output.seek(0)

                st.success("✅ Tudo pronto! O Gráfico Inteligente foi anexado no final do Excel.")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button("📥 Baixar Planilha (Excel Inteligente)", data=excel_output, file_name="Chamados_Prioridades.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with col2:
                    st.download_button("📄 Baixar Relatório (PDF Limpo)", data=pdf_output, file_name="Chamados_Prioridades.pdf", mime="application/pdf")

            except Exception as e:
                st.error(f"❌ Ocorreu um erro ao processar: {e}")
              
