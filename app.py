import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import io
import unicodedata

# 1. FUNÇÃO CHAVE: Limpa nomes para cruzamento exato
def normalizar_nome(nome):
    if pd.isna(nome): return ""
    nome = str(nome).strip().upper()
    nome = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn')
    return nome

# 2. LEITOR DO NOVO MAPÃO DA SED (Com Presença Completa e Filtro de Ativos)
def ler_novo_mapao(arquivo):
    df = pd.read_excel(arquivo, header=None)
    
    turma = arquivo.name.replace('.xls', '').replace('.xlsx', '')[:31]
    linha_cabecalho_1 = -1
    linha_cabecalho_2 = -1
    
    # Encontrar a Turma e a linha de Cabeçalhos
    for i, row in df.head(20).iterrows():
        for j, cell in enumerate(row):
            if isinstance(cell, str) and "Turma:" in cell:
                if cell.strip() == "Turma:" and j+1 < len(row):
                    turma = str(row[j+1]).strip()
                else:
                    turma = cell.replace("Turma:", "").strip()
            
            if isinstance(cell, str) and cell.strip().upper() == "ALUNO":
                linha_cabecalho_1 = i
                linha_cabecalho_2 = i + 1
                
    if linha_cabecalho_1 == -1:
        raise ValueError("Não foi possível encontrar a tabela de alunos neste arquivo.")
        
    row_1 = df.iloc[linha_cabecalho_1].fillna('').tolist()
    row_2 = df.iloc[linha_cabecalho_2].fillna('').tolist()
    
    disciplinas = []
    idx_aluno = -1
    idx_sit = -1
    col_primeira_disciplina = -1
    
    # Identificar colunas do ALUNO, SITUAÇÃO e as DISCIPLINAS
    for i, val in enumerate(row_1):
        v = str(val).strip()
        if v.upper() == 'ALUNO':
            idx_aluno = i
        elif v.upper() in ['SITUAÇÃO', 'SITUACAO', 'SIT']:
            idx_sit = i
        elif v and v.upper() != 'TOTAL':
            # Extrai apenas o nome da disciplina (Ex: 'ARTE\n1813' vira 'ARTE')
            subj_name = v.split('\n')[0].strip()
            disciplinas.append(subj_name)
            if col_primeira_disciplina == -1:
                col_primeira_disciplina = i
                
    # Identificar colunas de Frequência e Faltas (Quadro TOTAL)
    idx_tf, idx_fre, idx_ft_an, idx_fre_an = -1, -1, -1, -1
    for i, val in enumerate(row_2):
        v = str(val).strip().upper()
        if v == 'TF': idx_tf = i
        elif v in ['FRE(%)', 'FRE (%)']: idx_fre = i
        elif v in ['FT AN', 'FT AN.', 'FT. AN.']: idx_ft_an = i
        elif v in ['FRE AN(%)', 'FRE.AN(%)', 'FRE AN (%)']: idx_fre_an = i
            
    alunos = []
    for i in range(linha_cabecalho_2 + 1, len(df)):
        row = df.iloc[i]
        nome = str(row[idx_aluno]).strip()
        if nome.lower() in ['nan', 'none', '']: continue
        
        sit = str(row[idx_sit]).strip() if idx_sit != -1 else ""
        
        # FILTRO IMPORTANTE: Retirar alunos não ativos
        if 'ATIVO' not in sit.upper() and sit.upper() != 'AT':
            continue
            
        # O número da chamada (Nº) fica exatamente abaixo do nome da 1ª disciplina
        num = str(row[col_primeira_disciplina]).strip() if pd.notna(row[col_primeira_disciplina]) else ""
        if num.lower() in ["nan", "none"]: num = ""
        
        # Bloco de Presença e Faltas
        val_tf = str(row[idx_tf]).strip() if idx_tf != -1 else ""
        if val_tf.lower() in ["nan", "none"]: val_tf = "-"
        
        val_fre = str(row[idx_fre]).strip() if idx_fre != -1 else ""
        if val_fre.lower() in ["nan", "none"]: val_fre = "-"
        
        val_ft_an = str(row[idx_ft_an]).strip() if idx_ft_an != -1 else ""
        if val_ft_an.lower() in ["nan", "none"]: val_ft_an = "-"
        
        val_fre_an = str(row[idx_fre_an]).strip() if idx_fre_an != -1 else ""
        if val_fre_an.lower() in ["nan", "none"]: val_fre_an = "-"
        
        alunos.append({
            'Nº': num,
            'Nome do Aluno': nome,
            'Sit.': sit,
            'TF': val_tf,
            'Fre(%)': val_fre,
            'FT An': val_ft_an,
            'Fre An(%)': val_fre_an
        })
        
    return pd.DataFrame(alunos), disciplinas, turma

# --- INÍCIO DO APP STREAMLIT ---
st.set_page_config(page_title="Gerador de Conselho", layout="centered")

st.title("📚 Sistema do Conselho de Classe")
st.write("**ESCOLA ESTADUAL AMERICO BRASILIENSE DOUTOR**")

turno = st.selectbox("Turno:", ["Manhã", "Tarde", "Noite", "Integral"])

st.subheader("Passo 1: Carregar Turmas (Mapões da SED)")
mapoes_files = st.file_uploader("Suba os arquivos de Mapão aqui", type=["xlsx", "xls"], accept_multiple_files=True)

if mapoes_files:
    st.success(f"✅ {len(mapoes_files)} arquivo(s) carregado(s). Alunos inativos serão ocultados automaticamente.")
    st.divider()
    
    st.subheader("Passo 2: Configurar Prova Paulista")
    tem_prova = st.radio("As turmas acima possuem nota da Prova Paulista neste conselho?", ["Não", "Sim"])
    
    provas_files = []
    if tem_prova == "Sim":
        st.info("📌 Suba o(s) arquivo(s) 'RESULTADOS DA TURMA' extraído(s) do BI Educação.")
        provas_files = st.file_uploader("Área de Upload Seguro - Prova Paulista", type=["xlsx", "xls"], accept_multiple_files=True)
        
    st.divider()
    
    pode_gerar = True
    if tem_prova == "Sim" and not provas_files:
        st.warning("⚠️ Aguardando arquivos: Você marcou que há Prova Paulista, mas ainda não subiu as planilhas do BI.")
        pode_gerar = False

    if pode_gerar and st.button("Validar e Gerar Planilha Oficial", type="primary"):
        with st.spinner('Lendo dados, calculando presenças e cruzando notas...'):
            try:
                # TRATAR PROVA PAULISTA
                df_prova_reduzido = pd.DataFrame(columns=['Nome_Chave', 'Prova Paulista'])
                
                if tem_prova == "Sim" and provas_files:
                    df_provas_lista = []
                    for arq_prova in provas_files:
                        df_p = pd.read_excel(arq_prova)
                        df_provas_lista.append(df_p)
                    
                    df_prova_unificado = pd.concat(df_provas_lista, ignore_index=True)
                    col_nome_prova = 'Nome' if 'Nome' in df_prova_unificado.columns else df_prova_unificado.columns[1]
                    df_prova_unificado['Nome_Chave'] = df_prova_unificado[col_nome_prova].apply(normalizar_nome)
                    
                    coluna_nota = None
                    for col in df_prova_unificado.columns:
                        if '(%) de acertos' in col.lower() or 'nota' in col.lower() or 'percentual' in col.lower():
                            coluna_nota = col
                            break
                            
                    if coluna_nota:
                        df_prova_reduzido = df_prova_unificado[['Nome_Chave', coluna_nota]].rename(columns={coluna_nota: 'Prova Paulista'})
                        df_prova_reduzido = df_prova_reduzido.drop_duplicates(subset=['Nome_Chave'])
                        
                        df_prova_reduzido['Prova Paulista'] = pd.to_numeric(df_prova_reduzido['Prova Paulista'], errors='coerce')
                        
                        if df_prova_reduzido['Prova Paulista'].max() <= 1.0:
                            df_prova_reduzido['Prova Paulista'] = df_prova_reduzido['Prova Paulista'] * 10
                            
                        df_prova_reduzido['Prova Paulista'] = df_prova_reduzido['Prova Paulista'].apply(
                            lambda x: f"{x:.1f}".replace('.', ',') if pd.notnull(x) else ""
                        )

                # INICIAR EXCEL OFICIAL
                wb = Workbook()
                wb.remove(wb.active)
                escola_nome = "ESCOLA ESTADUAL AMERICO BRASILIENSE DOUTOR"

                for i, arq_mapao in enumerate(mapoes_files):
                    # LEITURA DO NOVO MODELO
                    df_mapao, disciplinas, nome_turma = ler_novo_mapao(arq_mapao)
                    
                    if df_mapao is None or df_mapao.empty:
                        st.error(f"Erro: A turma do arquivo {arq_mapao.name} não gerou alunos (talvez todos estejam inativos?).")
                        continue
                        
                    df_mapao['Nome_Chave'] = df_mapao['Nome do Aluno'].apply(normalizar_nome)

                    if tem_prova == "Sim" and not df_prova_reduzido.empty:
                        df_final = pd.merge(df_mapao, df_prova_reduzido, on='Nome_Chave', how='left')
                    else:
                        df_final = df_mapao.copy()
                        df_final['Prova Paulista'] = ""
                        
                    if 'Prova Paulista' in df_final.columns:
                        df_final['Prova Paulista'] = df_final['Prova Paulista'].fillna("-")
                        
                    for disciplina in disciplinas:
                        df_final[disciplina] = ""
                    df_final['Observações'] = ""

                    # Remontando colunas com o quadro TOTAL completo
                    colunas_finais = ['Nº', 'Nome do Aluno', 'Sit.', 'TF', 'Fre(%)', 'FT An', 'Fre An(%)', 'Prova Paulista'] + disciplinas + ['Observações']
                    df_final = df_final[colunas_finais]

                    nome_aba = nome_turma[:31] 
                    ws = wb.create_sheet(title=nome_aba)
                    
                    titulo_linha1 = f"{escola_nome}  ·  {nome_turma} – {turno}  ·  Conselho de Classe — 1º Bimestre / 2026"
                    ws.append([titulo_linha1])
                    ws.append(["Tipo de Ensino: Ensino Fundamental de 9 Anos / Ensino Médio / EJA"])
                    
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
                        
                    # Ajuste Dinâmico e Limpo das larguras
                    for j, col_name in enumerate(colunas_finais, start=1):
                        col_letra = get_column_letter(j)
                        if col_name == 'Nº': ws.column_dimensions[col_letra].width = 5
                        elif col_name == 'Nome do Aluno': ws.column_dimensions[col_letra].width = 45
                        elif col_name == 'Sit.': ws.column_dimensions[col_letra].width = 10
                        elif col_name == 'TF': ws.column_dimensions[col_letra].width = 6
                        elif col_name == 'Fre(%)': ws.column_dimensions[col_letra].width = 10
                        elif col_name == 'FT An': ws.column_dimensions[col_letra].width = 8
                        elif col_name == 'Fre An(%)': ws.column_dimensions[col_letra].width = 12
                        elif col_name == 'Prova Paulista': ws.column_dimensions[col_letra].width = 15
                        elif col_name == 'Observações': ws.column_dimensions[col_letra].width = 30
                        else: ws.column_dimensions[col_letra].width = 12 # Colunas das disciplinas

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success("Sucesso! Planilha gerada com o quadro completo de presenças e alunos inativos removidos.")
                st.download_button(
                    label="⬇️ Baixar Planilha Oficial do Conselho",
                    data=output,
                    file_name=f"Conselho_Classe_{turno}_Consolidado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Ocorreu um problema ao gerar: {e}")
