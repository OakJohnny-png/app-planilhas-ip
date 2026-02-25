import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font as xlFont, PatternFill, Border, Side, Alignment
from reportlab.lib.pagesizes import A4, landscape as rl_landscape, portrait as rl_portrait
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm

st.set_page_config(page_title="Chamados IP", layout="wide")
st.title("💡 Gerador de Relatórios de Chamados")

# --- BARRA LATERAL PARA FORMATAÇÕES (CUSTOMIZAÇÃO) ---
st.sidebar.header("🎨 Cores e Fontes")
fonte_escolhida = st.sidebar.selectbox("Estilo da Fonte", ["Arial","Arial Black","Bahnschrift","Calibri","Cambria","Candara","Comic Sans MS","Consolas","Constantia","Corbel","Courier New","Ebrima","Franklin Gothic Medium","Gabriola","Gadugi","Georgia","Impact","Ink Free","Javanese Text","Leelawadee UI","Lucida Console","Lucida Sans Unicode","Malgun Gothic","Microsoft Himalaya","Microsoft JhengHei","Microsoft New Tai Lue","Microsoft PhagsPa","Microsoft Tai Le","Microsoft YaHei","Microsoft Yi Baiti","MingLiU-ExtB","Mongolian Baiti","MS Gothic","MS UI Gothic","MV Boli","Myanmar Text","Nirmala UI","Palatino Linotype","Segoe MDL2 Assets","Segoe Print","Segoe Script","Segoe UI","Segoe UI Historic","Segoe UI Emoji","Segoe UI Symbol","SimSun","Sitka","Sylfaen","Symbol","Tahoma","Times New Roman","Trebuchet MS","Verdana","Webdings","Wingdings","Yu Gothic",])

st.sidebar.subheader("1. Rotas")
cor_fundo_rota = st.sidebar.color_picker("Cor de Fundo (Rotas)", "#FF0000")
cor_fonte_rota = st.sidebar.color_picker("Cor da Fonte (Rotas)", "#FFFFFF")
tamanho_rota = st.sidebar.number_input("Tamanho da Fonte (Rotas)", min_value=8, max_value=36, value=16)

st.sidebar.subheader("2. Bairros")
cor_fundo_bairro = st.sidebar.color_picker("Cor de Fundo (Bairros)", "#D3D3D3")
cor_fonte_bairro = st.sidebar.color_picker("Cor da Fonte (Bairros)", "#000000")
tamanho_bairro = st.sidebar.number_input("Tamanho da Fonte (Bairros)", min_value=8, max_value=36, value=14)

st.sidebar.subheader("3. Chamados")
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

st.write("Faça o upload da planilha Excel com as abas e o sistema irá processar, formatar e gerar o Excel e o PDF (com gráfico).")

# --- ÁREA DE UPLOAD ---
uploaded_file = st.file_uploader("📥 Envie sua planilha Excel (.xlsx) aqui", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀Gerar Relatórios"):
        with st.spinner('Lendo dados e desenhando relatórios...'):
            try:
                # 1. Leitura do arquivo enviado
                xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, dtype=str)
                abas_disponiveis = {nome.strip().upper(): nome for nome in xls.keys()}
                
                data_by_route = {r: {} for r in routes}
                contagem_por_rota = {r: 0 for r in routes}

                # 2. Processamento e Filtragem
                for route, neighborhoods in routes.items():
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
                                    contagem_por_rota[route] += len(problems)

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
                            cell = ws.cell(row=current_row, column=1, value=str(problem).strip())
                            cell.font = font_prob; cell.border = thin_border; cell.alignment = Alignment(wrap_text=True)
                            if cor_fundo_prob != "#FFFFFF": cell.fill = fill_prob
                            current_row += 1
                             
                ws.column_dimensions['A'].width = 150
                excel_output = io.BytesIO()
                wb.save(excel_output)
                excel_output.seek(0)

                # --- GERAÇÃO DO GRÁFICO (MATPLOTLIB) ---
                rotas_nomes = [r for r, qtd in contagem_por_rota.items() if qtd > 0]
                rotas_qtds = [qtd for r, qtd in contagem_por_rota.items() if qtd > 0]
                
                grafico_buffer = io.BytesIO()
                if rotas_nomes:
                    fig, ax = plt.subplots(figsize=(8, 5))
                    ax.bar(rotas_nomes, rotas_qtds, color=cor_fundo_rota)
                    ax.set_title("Prioridade: Quantidade de Chamados por Rota", fontsize=16, fontweight='bold')
                    ax.set_ylabel("Nº de Chamados Pendentes")
                    for i, v in enumerate(rotas_qtds):
                        ax.text(i, v + 0.5, str(v), ha='center', fontweight='bold')
                    plt.xticks(rotation=45, ha='right')
                    plt.tight_layout()
                    plt.savefig(grafico_buffer, format='png')
                    plt.close(fig)
                grafico_buffer.seek(0)

                # --- GERAÇÃO DO PDF (REPORTLAB) ---
                pdf_output = io.BytesIO()
                pagesize = rl_landscape(A4) if "Paisagem" in orientacao_pdf else rl_portrait(A4)
                
                doc = SimpleDocTemplate(
                    pdf_output, 
                    pagesize=pagesize,
                    rightMargin=margem_dir * mm, leftMargin=margem_esq * mm,
                    topMargin=margem_sup * mm, bottomMargin=margem_inf * mm
                )

                story = []
                
                estilo_rota = ParagraphStyle('Rota', fontName=f"{fonte_escolhida}-Bold", fontSize=tamanho_rota, 
                                             textColor=HexColor(cor_fonte_rota), backColor=HexColor(cor_fundo_rota), 
                                             spaceAfter=6, spaceBefore=12, padding=5)
                
                estilo_bairro = ParagraphStyle('Bairro', fontName=f"{fonte_escolhida}-Bold", fontSize=tamanho_bairro, 
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

                if rotas_nomes:
                    story.append(Spacer(1, 10 * mm))
                    story.append(Paragraph("Resumo de Prioridades", estilo_rota))
                    story.append(Spacer(1, 5 * mm))
                    img_largura = doc.width
                    img_altura = img_largura * 0.6
                    story.append(RLImage(grafico_buffer, width=img_largura, height=img_altura))

                doc.build(story)
                pdf_output.seek(0)

                st.success("✅ Tudo pronto! Seus relatórios foram gerados com sucesso.")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button("📥 Baixar Planilha (Excel)", data=excel_output, file_name="Chamados_Prioridades.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with col2:
                    st.download_button("📄 Baixar Relatório (PDF)", data=pdf_output, file_name="Chamados_Prioridades.pdf", mime="application/pdf")

            except Exception as e:
                st.error(f"❌ Ocorreu um erro ao processar: {e}")
              
