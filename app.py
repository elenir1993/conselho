import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import io

def processar_planilhas(arquivo_mapao, arquivo_prova):
    # Lendo os arquivos direto do upload
    df_mapao = pd.read_excel(arquivo_mapao, skiprows=8) 
    df_prova = pd.read_excel(arquivo_prova)

    # Tratamento do Mapão
    df_mapao = df_mapao.dropna(subset=['Nome do Aluno'])
    df_mapao['Nome do Aluno'] = df_mapao['Nome do Aluno'].str.strip().str.upper()
    
    colunas_padrao = ['Nº', 'Nome do Aluno', 'Sit.']
    disciplinas_mapao = [col for col in df_mapao.columns if col not in colunas_padrao and 'Unnamed' not in col]
    df_alunos = df_mapao[colunas_padrao].copy()

    # Tratamento da Prova Paulista
    df_prova['Nome do Aluno'] = df_prova['Nome do Aluno'].str.strip().str.upper()
    coluna_nota_prova = 'Nota' # Ajuste aqui se a coluna tiver outro nome
    df_prova_reduzido = df_prova[['Nome do Aluno', coluna_nota_prova]].rename(columns={coluna_nota_prova: 'Prova Paulista'})

    # Cruzamento dos Dados
    df_final = pd.merge(df_alunos, df_prova_reduzido, on='Nome do Aluno', how='left')

    for disciplina in disciplinas_mapao:
        df_final[disciplina] = ""
    df_final['Observações'] = ""

    colunas_finais = ['Nº', 'Nome do Aluno', 'Sit.', 'Prova Paulista'] + disciplinas_mapao + ['Observações']
    df_final = df_final[colunas_finais]

    # Formatação com openpyxl
    wb = Workbook()
    ws = wb.active
    ws.title = "Conselho de Classe"

    escola_nome = "ESCOLA ESTADUAL AMERICO BRASILIENSE DOUTOR"
    ws.append([f"{escola_nome}  ·  9° Ano A – Tarde Anual  ·  Conselho de Classe — 1º Bimestre / 2026"])
    ws.append(["Tipo de Ensino: Ensino Fundamental de 9 Anos"])
    
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(colunas_finais))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(colunas_finais))

    for row in ws.iter_rows(min_row=1, max_row=2):
        for cell in row:
            cell.font = Font(bold=True, size=12)
            cell.alignment = Alignment(horizontal='center', vertical='center')

    ws.append(colunas_finais)
    
    for cell in ws[3]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r in dataframe_to_rows(df_final, index=False, header=False):
        ws.append(r)

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 45
    ws.column_dimensions['C'].width = 8
    ws.column_dimensions['D'].width = 15
    for i in range(5, len(colunas_finais)):
        col_letra = chr(64 + i) if i <= 26 else chr(64 + (i // 26)) + chr(64 + (i % 26))
        ws.column_dimensions[col_letra].width = 12
    ws.column_dimensions[chr(64 + len(colunas_finais))].width = 30

    # Salvando em memória para o botão de download
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Interface Web com Streamlit ---
st.set_page_config(page_title="Gerador de Planilha - Conselho de Classe", layout="centered")

st.title("📚 Gerador de Planilha para o Conselho de Classe")
st.write("Faça o upload do Mapão da SED e dos Resultados da Prova Paulista para gerar a planilha formatada.")

mapao_file = st.file_uploader("1. Faça o upload do Mapão (Excel)", type=["xlsx", "xls"])
prova_file = st.file_uploader("2. Faça o upload das Notas da Prova Paulista (Excel)", type=["xlsx", "xls"])

if mapao_file and prova_file:
    if st.button("Gerar Planilha do Conselho"):
        try:
            planilha_pronta = processar_planilhas(mapao_file, prova_file)
            
            st.success("Planilha gerada com sucesso!")
            st.download_button(
                label="⬇️ Baixar Planilha Pronta",
                data=planilha_pronta,
                file_name="Conselho_Classe_Preenchido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar os arquivos. Verifique se o modelo está correto. Detalhe: {e}")
