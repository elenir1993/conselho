import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import io

def processar_multiplas_planilhas(arquivos_mapao, arquivos_prova, turno_selecionado):
    # 1. Unificar as notas da Prova Paulista (seja 1 ou vários arquivos)
    df_provas_lista = []
    for arq_prova in arquivos_prova:
        df = pd.read_excel(arq_prova)
        df_provas_lista.append(df)
    
    df_prova_unificado = pd.concat(df_provas_lista, ignore_index=True)
    df_prova_unificado['Nome do Aluno'] = df_prova_unificado['Nome do Aluno'].astype(str).str.strip().str.upper()
    
    coluna_nota_prova = 'Nota' # Ajuste se no BI Educação vier com outro nome
    
    # Se a coluna existir, filtra. Se não, avisa (evita quebrar o app)
    if coluna_nota_prova in df_prova_unificado.columns:
        df_prova_reduzido = df_prova_unificado[['Nome do Aluno', coluna_nota_prova]].rename(columns={coluna_nota_prova: 'Prova Paulista'})
        df_prova_reduzido = df_prova_reduzido.drop_duplicates(subset=['Nome do Aluno'])
    else:
        df_prova_reduzido = pd.DataFrame(columns=['Nome do Aluno', 'Prova Paulista'])

    # 2. Criar o arquivo Excel mestre
    wb = Workbook()
    wb.remove(wb.active) # Remove a aba inicial vazia

    escola_nome = "ESCOLA ESTADUAL AMERICO BRASILIENSE DOUTOR"

    # 3. Processar cada Mapão (cada um vira uma aba)
    for arq_mapao in arquivos_mapao:
        # Pega o nome do arquivo enviado para dar nome à aba (Ex: "9 ANO A")
        nome_aba = arq_mapao.name.replace('.xlsx', '').replace('.xls', '')[:31]
        
        df_mapao = pd.read_excel(arq_mapao, skiprows=8) 
        df_mapao = df_mapao.dropna(subset=['Nome do Aluno'])
        df_mapao['Nome do Aluno'] = df_mapao['Nome do Aluno'].astype(str).str.strip().str.upper()
        
        colunas_padrao = ['Nº', 'Nome do Aluno', 'Sit.']
        disciplinas_mapao = [col for col in df_mapao.columns if col not in colunas_padrao and 'Unnamed' not in col]
        
        df_alunos = df_mapao[colunas_padrao].copy()

        # Cruzamento
        df_final = pd.merge(df_alunos, df_prova_reduzido, on='Nome do Aluno', how='left')

        for disciplina in disciplinas_mapao:
            df_final[disciplina] = ""
        df_final['Observações'] = ""

        colunas_finais = ['Nº', 'Nome do Aluno', 'Sit.', 'Prova Paulista'] + disciplinas_mapao + ['Observações']
        df_final = df_final[colunas_finais]

        # 4. Criar e formatar a Aba da Turma
        ws = wb.create_sheet(title=nome_aba)
        
        titulo_linha1 = f"{escola_nome}  ·  {nome_aba} – {turno_selecionado}  ·  Conselho de Classe — 1º Bimestre / 2026"
        ws.append([titulo_linha1])
        ws.append(["Tipo de Ensino: Ensino Fundamental de 9 Anos / Ensino Médio"])
        
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

        # Largura das colunas
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 45
        ws.column_dimensions['C'].width = 8
        ws.column_dimensions['D'].width = 15
        for i in range(5, len(colunas_finais)):
            col_letra = chr(64 + i) if i <= 26 else chr(64 + (i // 26)) + chr(64 + (i % 26))
            ws.column_dimensions[col_letra].width = 12
        ws.column_dimensions[chr(64 + len(colunas_finais))].width = 30

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Interface Web ---
st.set_page_config(page_title="Gerador de Planilha - Conselho em Lote", layout="centered")

st.title("📚 Gerador de Conselho (Múltiplas Salas)")
st.write("Faça o upload de vários mapões ao mesmo tempo. O sistema gerará um único arquivo Excel com várias abas.")

turno = st.selectbox("1. Selecione o Turno:", ["Manhã", "Tarde", "Noite", "Integral"])

# accept_multiple_files=True permite selecionar as 9 salas de uma vez!
mapoes_files = st.file_uploader("2. Faça o upload dos Mapões da SED (Selecione vários de uma vez)", type=["xlsx", "xls"], accept_multiple_files=True)
provas_files = st.file_uploader("3. Faça o upload das Notas da Prova Paulista (Pode ser 1 arquivo com toda a escola ou vários)", type=["xlsx", "xls"], accept_multiple_files=True)

if mapoes_files and provas_files:
    if st.button("Gerar Planilhão do Conselho"):
        with st.spinner('Processando todas as salas...'):
            try:
                planilha_pronta = processar_multiplas_planilhas(mapoes_files, provas_files, turno)
                
                st.success(f"Planilha gerada com sucesso! ({len(mapoes_files)} salas processadas)")
                st.download_button(
                    label="⬇️ Baixar Planilhão Pronta",
                    data=planilha_pronta,
                    file_name=f"Conselho_Classe_{turno}_Bim1.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erro ao processar: {e}")
