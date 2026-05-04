import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import unicodedata

# 1. FUNÇÃO CHAVE: Limpa nomes para evitar erros
def normalizar_nome(nome):
    if pd.isna(nome): return ""
    nome = str(nome).strip().upper()
    nome = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn')
    return nome

# 2. LEITOR DA SED: Resolve o problema do Excel HTML e puxa a Turma
def ler_mapao_robusto(arquivo):
    try:
        tabelas = pd.read_html(arquivo)
        df_bruto = tabelas[0]
        
        linha_disciplinas, linha_nome = -1, -1
        nome_turma = arquivo.name.replace('.xls', '').replace('.xlsx', '')[:31] # Nome padrão se falhar
        
        # Caçador de Turma e Cabeçalhos
        for i, row in df_bruto.iterrows():
            row_str_full = row.dropna().astype(str).tolist()
            for cell in row_str_full:
                if "Turma:" in cell:
                    nome_turma = cell.replace("Turma:", "").strip()
                    
            row_str = row.astype(str).str.lower()
            if 'aluno' in row_str.values: linha_disciplinas = i
            if 'nome' in row_str.values and ('sit' in row_str.values or 'situação' in row_str.values):
                linha_nome = i
                break
                
        disciplinas = []
        if linha_disciplinas != -1:
            for val in df_bruto.iloc[linha_disciplinas]:
                val_str = str(val).strip()
                if val_str.lower() not in ['nan', 'aluno', 'total', 'none'] and val_str not in disciplinas:
                    disciplinas.append(val_str)
                    
        alunos = []
        if linha_nome != -1:
            row_nome = df_bruto.iloc[linha_nome].astype(str).str.lower()
            idx_nome, idx_sit, idx_num = -1, -1, -1
            
            # Trava as colunas corretas para preservar o Número de Chamada (Nº) fiel ao mapão
            for col_idx, val in enumerate(row_nome):
                if val == 'nome' and idx_nome == -1: idx_nome = col_idx
                if val in ['sit', 'situação'] and idx_sit == -1: idx_sit = col_idx
                if val in ['nº', 'numero', 'chamada', 'n'] and idx_num == -1: idx_num = col_idx
                
            if idx_num == -1 and idx_sit != -1: idx_num = idx_sit + 1
            
            for i in range(linha_nome + 1, len(df_bruto)):
                row = df_bruto.iloc[i]
                nome = str(row[idx_nome]).strip()
                if nome.lower() in ['nan', 'none', '']: continue
                if len(nome) < 3: continue 
                
                sit = str(row[idx_sit]).strip() if idx_sit != -1 else ""
                # Preserva o número de chamada exatamente como está (ex: "01")
                num = str(row[idx_num]).strip() if idx_num != -1 else ""
                if num == "nan" or num == "None": num = ""
                
                alunos.append({'Nº': num, 'Nome do Aluno': nome, 'Sit.': sit})
                
        return pd.DataFrame(alunos), disciplinas, nome_turma
    except Exception:
        # Fallback (para arquivos XLS/XLSX verdadeiros)
        df = pd.read_excel(arquivo, header=None)
        nome_turma = arquivo.name.replace('.xls', '').replace('.xlsx', '')[:31]
        
        linha_cabecalho = 0
        for i, linha in df.iterrows():
            linha_str_full = linha.dropna().astype(str).tolist()
            for cell in linha_str_full:
                if "Turma:" in cell:
                    nome_turma = cell.replace("Turma:", "").strip()
                    
            linha_str = linha.astype(str).str.lower()
            if linha_str.str.contains('nome do aluno|nome|aluno', na=False).any():
                linha_cabecalho = i
                break
                
        df = pd.read_excel(arquivo, skiprows=linha_cabecalho)
        df.columns = df.columns.astype(str).str.strip()
        
        for col in df.columns:
            if 'nome' in col.lower() or 'aluno' in col.lower():
                df.rename(columns={col: 'Nome do Aluno'}, inplace=True)
                break
        
        colunas_padrao = ['Nº', 'Nome do Aluno', 'Sit.']
        for col in df.columns:
            if 'sit' in col.lower(): df.rename(columns={col: 'Sit.'}, inplace=True)
            if 'nº' in col.lower() or 'numero' in col.lower() or 'chamada' in col.lower(): df.rename(columns={col: 'Nº'}, inplace=True)
        
        if 'Sit.' not in df.columns: df['Sit.'] = ""
        if 'Nº' not in df.columns: df['Nº'] = ""
        
        # Converte número para texto para não perder zero à esquerda
        df['Nº'] = df['Nº'].astype(str).str.replace(r'\.0$', '', regex=True)
        df.loc[df['Nº'] == 'nan', 'Nº'] = ""
        
        disciplinas = [col for col in df.columns if col not in colunas_padrao and 'unnamed' not in col.lower()]
        return df[colunas_padrao].dropna(subset=['Nome do Aluno']), disciplinas, nome_turma

# --- INÍCIO DO APP STREAMLIT ---
st.set_page_config(page_title="Gerador de Conselho", layout="centered")

st.title("📚 Sistema do Conselho de Classe")
st.write("**ESCOLA ESTADUAL AMERICO BRASILIENSE DOUTOR**")

turno = st.selectbox("Turno:", ["Manhã", "Tarde", "Noite", "Integral"])

st.subheader("Passo 1: Carregar Turmas (Mapões da SED)")
mapoes_files = st.file_uploader("Suba os arquivos de Mapão aqui", type=["xlsx", "xls"], accept_multiple_files=True)

if mapoes_files:
    # Exibe no painel apenas a quantidade de arquivos para manter limpo
    st.success(f"✅ {len(mapoes_files)} arquivo(s) carregado(s). O sistema vai ler a turma oficial de dentro de cada um.")
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
        with st.spinner('Lendo turmas da SED e cruzando notas de forma segura...'):
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
                    # AGORA EXTRAIMOS A TURMA REAL DE DENTRO DO MAPÃO
                    df_mapao, disciplinas, nome_turma = ler_mapao_robusto(arq_mapao)
                    
                    if df_mapao is None or df_mapao.empty:
                        st.error(f"Erro ao ler os dados do arquivo {arq_mapao.name}. Verifique o formato.")
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

                    colunas_finais = ['Nº', 'Nome do Aluno', 'Sit.', 'Prova Paulista'] + disciplinas + ['Observações']
                    df_final = df_final[colunas_finais]

                    # Aba do Excel limitada a 31 caracteres, mas usando o nome real da turma
                    nome_aba = nome_turma[:31] 
                    
                    ws = wb.create_sheet(title=nome_aba)
                    
                    # Cabeçalho usando o nome EXATO da turma que estava no Mapão
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
                        
                    ws.column_dimensions['A'].width = 5
                    ws.column_dimensions['B'].width = 45
                    ws.column_dimensions['C'].width = 8
                    ws.column_dimensions['D'].width = 15
                    for j in range(5, len(colunas_finais)):
                        col_letra = chr(64 + j) if j <= 26 else chr(64 + (j // 26)) + chr(64 + (j % 26))
                        ws.column_dimensions[col_letra].width = 12
                    ws.column_dimensions[chr(64 + len(colunas_finais))].width = 30

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success("Sucesso! As turmas foram nomeadas corretamente com os números de chamada idênticos à SED.")
                st.download_button(
                    label="⬇️ Baixar Planilha Oficial do Conselho",
                    data=output,
                    file_name=f"Conselho_Classe_{turno}_Consolidado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Ocorreu um problema ao gerar: {e}")
