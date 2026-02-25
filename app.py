import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Processador de Chamados IP", layout="wide")

st.title("💡 Processador de Chamados de Iluminação Pública")
st.write("Faça o upload da planilha Excel contendo as abas dos bairros e personalize a formatação do arquivo final.")

# --- BARRA LATERAL PARA FORMATAÇÕES (CUSTOMIZAÇÃO) ---
st.sidebar.header("🎨 Personalização")

fonte_escolhida = st.sidebar.selectbox("Estilo da Fonte", ["Helvetica", "Arial", "Calibri", "Times New Roman"])

st.sidebar.subheader("1. Formatação das Rotas")
cor_fundo_rota = st.sidebar.color_picker("Cor de Fundo (Rotas)", "#FF0000")
cor_fonte_rota = st.sidebar.color_picker("Cor da Fonte (Rotas)", "#000000")
tamanho_rota = st.sidebar.number_input("Tamanho da Fonte (Rotas)", min_value=8, max_value=36, value=16)

st.sidebar.subheader("2. Formatação dos Bairros")
cor_fundo_bairro = st.sidebar.color_picker("Cor de Fundo (Bairros)", "#D3D3D3")
cor_fonte_bairro = st.sidebar.color_picker("Cor da Fonte (Bairros)", "#000000")
tamanho_bairro = st.sidebar.number_input("Tamanho da Fonte (Bairros)", min_value=8, max_value=36, value=16)

st.sidebar.subheader("3. Formatação dos Problemas")
cor_fundo_prob = st.sidebar.color_picker("Cor de Fundo (Problemas)", "#FFFFFF")
cor_fonte_prob = st.sidebar.color_picker("Cor da Fonte (Problemas)", "#000000")
tamanho_prob = st.sidebar.number_input("Tamanho da Fonte (Problemas)", min_value=8, max_value=36, value=14)

# --- ROTAS DEFINIDAS PELA REGRA 2 ---
routes = {
    "ROTA 1": ["CENTRO", "JARDIM AMERICA"],
    "ROTA 2": ["ALBERTINA", "LARANJEIRAS", "BOA VISTA", "EUGENIO SCHNEIDER"],
    "ROTA 3": ["FUNDO CANOAS", "CANOAS", "PROGRESSO", "PAMPLONA", "CANTA GALO"],
    "ROTA 4": ["BARRA DO TROMBUDO", "BARRAGEM", "BUDAG", "SUMARE"],
    "ROTA 5": ["SANTANA", "TABOAO", "BREMER", "BELA ALIANÇA"],
    "ROTA 6": ["BARRA DA ITOUPAVA", "NAVEGANTES", "SANTA RITA", "VALADA ITOUPAVA", "VALADA SÃO PAULO", "RAINHA"]
}

# --- ÁREA PRINCIPAL ---
uploaded_file = st.file_uploader("Envie sua planilha Excel (.xlsx) aqui", type=["xlsx"])

if uploaded_file is not None:
    if st.button("Processar Planilha"):
        with st.spinner('Lendo e processando os dados...'):
            try:
                xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, dtype=str)
                abas_disponiveis = {nome.strip().upper(): nome for nome in xls.keys()}
                data_by_route = {r: {} for r in routes}

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

                wb = Workbook()
                ws = wb.active
                ws.title = "Chamados Pendentes"
                ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

                hex_bg_rota = cor_fundo_rota.replace('#', '')
                hex_fg_rota = cor_fonte_rota.replace('#', '')
                hex_bg_bairro = cor_fundo_bairro.replace('#', '')
                hex_fg_bairro = cor_fonte_bairro.replace('#', '')
                hex_bg_prob = cor_fundo_prob.replace('#', '')
                hex_fg_prob = cor_fonte_prob.replace('#', '')

                font_rota = Font(name=fonte_escolhida, size=tamanho_rota, color=hex_fg_rota, bold=True)
                fill_rota = PatternFill(start_color=hex_bg_rota, end_color=hex_bg_rota, fill_type="solid")
                font_bairro = Font(name=fonte_escolhida, size=tamanho_bairro, color=hex_fg_bairro, bold=True)
                fill_bairro = PatternFill(start_color=hex_bg_bairro, end_color=hex_bg_bairro, fill_type="solid")
                font_prob = Font(name=fonte_escolhida, size=tamanho_prob, color=hex_fg_prob)
                fill_prob = PatternFill(start_color=hex_bg_prob, end_color=hex_bg_prob, fill_type="solid")

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
                            if hex_bg_prob != "FFFFFF": cell.fill = fill_prob
                            current_row += 1

                ws.column_dimensions['A'].width = 150
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success("✅ Concluído!")
                st.download_button(label="📥 Baixar Planilha Pronta", data=output, file_name="Chamados_Pendentes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ Erro: {e}")
              
