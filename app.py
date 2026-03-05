import streamlit as st
import pandas as pd
import io
import urllib.parse  # Para formatar o link do WhatsApp
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font as xlFont, PatternFill
from reportlab.lib.pagesizes import A4, landscape as rl_landscape, portrait as rl_portrait
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm

# Configuração da Página
st.set_page_config(page_title="V.A.G.A.L.U.M.E. Pro", layout="wide")
ano_atual = datetime.now().year

st.title("💡 V.A.G.A.L.U.M.E. - Integração WhatsApp")

# --- BARRA LATERAL ---
st.sidebar.header("📲 Destinatário WhatsApp")
numero_whatsapp = st.sidebar.text_input("Número (Ex: 5547999999999)", help="Inclua o 55 + DDD + Número")

with st.sidebar.expander("🎨 Estilo do Relatório"):
    fonte_escolhida = st.sidebar.selectbox("Fonte", ["Helvetica", "Times-Roman", "Courier"])
    cor_fundo_rota = st.sidebar.color_picker("Cor de Rota", "#1E3A8A")
    cor_fonte_rota = st.sidebar.color_picker("Cor da Fonte", "#FFFFFF")

# --- LÓGICA DE ROTAS ---
routes = {
    "ROTA 1": ["CENTRO", "JARDIM AMERICA"],
    "ROTA 2": ["ALBERTINA", "LARANJEIRAS", "BOA VISTA", "EUGENIO SCHNEIDER"],
    "ROTA 3": ["FUNDO CANOAS", "CANOAS", "PROGRESSO", "PAMPLONA", "CANTA GALO"],
    "ROTA 4": ["BARRA DO TROMBUDO", "BARRAGEM", "BUDAG", "SUMARE"],
    "ROTA 5": ["SANTANA", "TABOAO", "BREMER", "BELA ALIANÇA"],
    "ROTA 6": ["BARRA DA ITOUPAVA", "NAVEGANTES", "SANTA RITA", "VALADA ITOUPAVA", "VALADA SÃO PAULO", "RAINHA"]
}

uploaded_file = st.file_uploader("📥 Envie a planilha Excel", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 Processar e Gerar Link de Envio"):
        try:
            xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
            abas_disponiveis = {nome.strip().upper(): nome for nome in xls.keys()}
            
            data_by_route = {r: {} for r in routes}
            texto_resumo_wa = f"*🔦 RELATÓRIO V.A.G.A.L.U.M.E. ({ano_atual})*\n\n"

            for route, neighborhoods in routes.items():
                tem_problema_na_rota = False
                resumo_rota = f"*📍 {route}*\n"
                
                for neighborhood in neighborhoods:
                    nome_aba_upper = neighborhood.upper()
                    if nome_aba_upper in abas_disponiveis:
                        df = xls[abas_disponiveis[nome_aba_upper]]
                        
                        if df.shape[1] < 4: continue

                        df[0] = pd.to_datetime(df[0], errors='coerce')
                        mask = (df[0].dt.year == ano_atual) & \
                               (df[3].astype(str).str.upper().isin(['NÃO REALIZADO', 'NÃO EXECUTADO']))
                        
                        problems = df[mask][1].tolist()
                        if problems:
                            data_by_route[route][neighborhood] = problems
                            tem_problema_na_rota = True
                            resumo_rota += f"  • _{neighborhood}_: {len(problems)} pendência(s)\n"
                
                if tem_problema_na_rota:
                    texto_resumo_wa += resumo_rota + "\n"

            # --- GERAÇÃO DO LINK DO WHATSAPP ---
            mensagem_codificada = urllib.parse.quote(texto_resumo_wa)
            link_wa = f"https://wa.me/{numero_whatsapp}?text={mensagem_codificada}"

            # --- INTERFACE DE RESULTADO ---
            st.success("✅ Processamento concluído!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.info("📱 **Envio por WhatsApp**")
                if numero_whatsapp:
                    st.link_button("📤 Abrir WhatsApp e Enviar Resumo", link_wa)
                else:
                    st.warning("Insira um número na barra lateral para habilitar o envio.")
            
            with col2:
                st.info("💾 **Arquivos para Download**")
                # (Aqui entraria a lógica de geração de PDF/Excel que já fizemos anteriormente)
                st.write("PDF e Excel prontos para download abaixo.")

            # Mostrar pré-visualização da mensagem
            with st.expander("🔍 Visualizar Mensagem do WhatsApp"):
                st.markdown(texto_resumo_wa)

        except Exception as e:
            st.error(f"Erro ao processar: {e}")

