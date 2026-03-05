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
ano_atual = datetime.now().year

st.title("💡 V.A.G.A.L.U.M.E. - Relatórios de Iluminação")

# --- CONFIGURAÇÕES LATERAIS ---
st.sidebar.header("🎨 Configurações")
fonte_escolhida = st.sidebar.selectbox("Fonte", ["Helvetica", "Times-Roman", "Courier"])
cor_fundo_rota = st.sidebar.color_picker("Cor das Rotas", "#1E3A8A")
cor_txt_rota = st.sidebar.color_picker("Texto das Rotas", "#FFFFFF")

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
    if st.button("🚀 Processar Dados"):
        with st.spinner('Extraindo dados...'):
            try:
                # Lendo o Excel (dtype=str evita conversões erradas iniciais)
                xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
                abas_disponiveis = {nome.strip().upper(): nome for nome in xls.keys()}
                
                data_by_route = {}
                encontrou_dados = False

                for route, neighborhoods in routes.items():
                    route_data = {}
                    for neighborhood in neighborhoods:
                        nome_aba_upper = neighborhood.upper()
                        if nome_aba_upper in abas_disponiveis:
                            df = xls[abas_disponiveis[nome_aba_upper]].copy()
                            
                            if df.shape[1] < 4: continue

                            # --- LIMPEZA DOS DADOS (O SEGREDO) ---
                            # 1. Converte Coluna 0 para Data (forçando o erro virar nulo)
                            df[0] = pd.to_datetime(df[0], errors='coerce')
                            
                            # 2. Limpa espaços e coloca em maiúsculo a Coluna 3 (Status)
                            df[3] = df[3].astype(str).str.strip().upper()
                            
                            # 3. Filtro: Ano Corrente E (Não Realizado OU Não Executado)
                            # Removi o filtro de ano momentaneamente para teste se o seu arquivo for antigo, 
                            # mas vou manter como solicitado.
                            mask = (
                                (df[0].dt.year == ano_atual) & 
                                (df[3].isin(['NÃO REALIZADO', 'NÃO EXECUTADO', 'NAO REALIZADO', 'NAO EXECUTADO']))
                            )
                            
                            # 4. Pega a descrição (Coluna 1)
                            filt_df = df[mask]
                            problems = filt_df[1].dropna().astype(str).str.strip().tolist()
                            
                            if problems:
                                route_data[neighborhood] = problems
                                encontrou_dados = True
                    
                    if route_data:
                        data_by_route[route] = route_data

                if not encontrou_dados:
                    st.warning(f"⚠️ Atenção: Nenhuma linha encontrada para o ano {ano_atual} com status 'NÃO REALIZADO'. Verifique se a data na primeira coluna é de {ano_atual}.")
                else:
                    # --- GERAÇÃO EXCEL ---
                    excel_buffer = io.BytesIO()
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Pendências"
                    
                    row_num = 1
                    for r_name, b_dict in data_by_route.items():
                        cell = ws.cell(row=row_num, column=1, value=r_name)
                        cell.font = xlFont(bold=True, color=cor_txt_rota.replace('#','')); cell.fill = PatternFill(start_color=cor_fundo_rota.replace('#',''), fill_type="solid")
                        row_num += 1
                        for b_name, p_list in b_dict.items():
                            ws.cell(row=row_num, column=1, value=f"📍 {b_name}").font = xlFont(bold=True)
                            row_num += 1
                            for p in p_list:
                                ws.cell(row=row_num, column=1, value=f"  - {p}")
                                row_num += 1
                        row_num += 1
                    
                    wb.save(excel_buffer)
                    excel_buffer.seek(0)

                    # --- GERAÇÃO PDF ---
                    pdf_buffer = io.BytesIO()
                    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)
                    
                    # Estilos
                    style_r = ParagraphStyle('R', fontSize=14, backColor=HexColor(cor_fundo_rota), textColor=HexColor(cor_txt_rota), alignment=1, spaceBefore=10)
                    style_b = ParagraphStyle('B', fontSize=12, fontName="Helvetica-Bold", leftIndent=10, spaceBefore=5)
                    style_i = ParagraphStyle('I', fontSize=10, leftIndent=20)
                    
                    elements = [Paragraph(f"ORDENS PENDENTES - {ano_atual}", style_r), Spacer(1, 10)]
                    for r_name, b_dict in data_by_route.items():
                        elements.append(Paragraph(r_name, style_r))
                        for b_name, p_list in b_dict.items():
                            elements.append(Paragraph(b_name, style_b))
                            for p in p_list:
                                elements.append(Paragraph(f"• {p}", style_i))
                    
                    doc.build(elements)
                    pdf_buffer.seek(0)

                    st.success(f"✅ Processado! {sum(len(p) for r in data_by_route.values() for p in r.values())} itens encontrados.")
                    st.download_button("📥 Baixar Excel", data=excel_buffer, file_name="Lista_Manutencao.xlsx")
                    st.download_button("📄 Baixar PDF", data=pdf_buffer, file_name="Lista_Manutencao.pdf")

            except Exception as e:
                st.error(f"Erro no processamento: {e}")
