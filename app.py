import streamlit as st
import pandas as pd
import io
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

st.set_page_config(page_title="Processador de Chamados IP", layout="wide")
st.title("💡 Processador Automático de Chamados")

# --- COLOQUE O LINK DO GOOGLE SHEETS AQUI ---
# Lembre-se de terminar o link com /export?format=xlsx
URL_DO_GOOGLE_SHEETS = "https://docs.google.com/spreadsheets/d/1xS00XT4h5cYyzwezkE0Ajae0ZpK6YVamHOR8Og_8GKU/edit?gid=1507549907#gid=1507549907/export?format=xlsx"

# ... (Mantenha toda a parte do menu lateral de Cores e Fontes que te passei antes) ...
# ... (Mantenha o dicionário de 'routes' que te passei antes) ...

st.write("Clique no botão abaixo para o sistema ir até o Google Sheets, baixar os dados mais recentes e formatar sua planilha.")

if st.button("🔄 Sincronizar Google Sheets e Gerar Planilha"):
    with st.spinner('Acessando o Google Sheets e processando...'):
        try:
            # 1. Faz o download direto do Google Sheets
            resposta = requests.get(URL_DO_GOOGLE_SHEETS)
            
            if resposta.status_code == 200:
                # 2. Lê a planilha baixada na memória
                xls = pd.read_excel(io.BytesIO(resposta.content), sheet_name=None, header=None, dtype=str)
                abas_disponiveis = {nome.strip().upper(): nome for nome in xls.keys()}
                data_by_route = {r: {} for r in routes}

                # 3. Faz toda aquela filtragem de NÃO REALIZADO...
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

                # 4. Cria o Excel formatado... (mesmo código do wb = Workbook() que te passei na mensagem anterior)
                # ... [CÓDIGO DE FORMATAÇÃO DO EXCEL AQUI] ...
                
                # Após gerar o 'output' com wb.save()
                st.success("✅ Sincronização e formatação concluídas com sucesso!")
                st.download_button(
                    label="📥 Baixar Planilha Pronta", 
                    data=output, 
                    file_name="Chamados_Pendentes.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("❌ Não foi possível acessar o Google Sheets. Verifique se o link está público para leitura.")

        except Exception as e:
            st.error(f"❌ Erro interno: {e}")
              
