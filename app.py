import streamlit as st
import pandas as pd
import io
import unicodedata
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak
from reportlab.lib import colors

# 1. FUNÇÃO CHAVE: Limpa nomes para cruzamento exato
def normalizar_nome(nome):
    if pd.isna(nome): return ""
    nome = str(nome).strip().upper()
    nome = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn')
    return nome

# 2. ABREVIAÇÃO INTELIGENTE: Necessário para caber no A4
def abreviar_disciplina(nome):
    n = str(nome).upper().strip()
    if 'PORTUGUESA' in n: return 'PORT'
    if 'FISICA' in n and 'EDUCACAO' not in n and 'EDUCAÇÃO' not in n: return 'FÍS'
    if 'EDUCACAO FISICA' in n or 'EDUCAÇÃO FÍSICA' in n: return 'ED.FÍS'
    if 'MATEMATICA' in n or 'MATEMÁTICA' in n: return 'MAT'
    if 'BIOLOGIA' in n: return 'BIO'
    if 'HISTORIA' in n or 'HISTÓRIA' in n: return 'HIST'
    if 'GEOGRAFIA' in n: return 'GEO'
    if 'FILOSOFIA' in n: return 'FILO'
    if 'SOCIOLOGIA' in n: return 'SOC'
    if 'QUIMICA' in n or 'QUÍMICA' in n: return 'QUÍM'
    if 'ARTE' in n: return 'ARTE'
    if 'INGLESA' in n or 'INGLÊS' in n: return 'INGL'
    if 'FINANCEIRA' in n: return 'ED.FIN'
    if 'REDAÇ' in n or 'REDAC' in n: return 'RED'
    if 'ATUALIDADE' in n: return 'ATUAL'
    if 'LIDERANÇA' in n or 'ORATÓRIA' in n: return 'ORAT'
    if 'PROJETO' in n and 'VIDA' in n: return 'P.VID'
    return n[:5]

# 3. LEITOR DO NOVO MAPÃO DA SED
def ler_novo_mapao(arquivo):
    df = pd.read_excel(arquivo, header=None)
    
    turma = arquivo.name.replace('.xls', '').replace('.xlsx', '')[:31]
    linha_cabecalho_1, linha_cabecalho_2 = -1, -1
    
    for i, row in df.head(20).iterrows():
        for j, cell in enumerate(row):
            if isinstance(cell, str) and "Turma:" in cell:
                if cell.strip() == "Turma:" and j+1 < len(row):
                    turma = str(row[j+1]).strip()
                else:
                    turma = cell.replace("Turma:", "").strip()
            
            if isinstance(cell, str) and cell.strip().upper() == "ALUNO":
                linha_cabecalho_1, linha_cabecalho_2 = i, i + 1
                
    if linha_cabecalho_1 == -1:
        raise ValueError("Não foi possível encontrar a tabela de alunos neste arquivo.")
        
    row_1 = df.iloc[linha_cabecalho_1].fillna('').tolist()
    row_2 = df.iloc[linha_cabecalho_2].fillna('').tolist()
    
    disciplinas = []
    idx_aluno, idx_sit, col_primeira_disciplina = -1, -1, -1
    
    for i, val in enumerate(row_1):
        v = str(val).strip()
        if v.upper() == 'ALUNO': idx_aluno = i
        elif v.upper() in ['SITUAÇÃO', 'SITUACAO', 'SIT']: idx_sit = i
        elif v and v.upper() != 'TOTAL':
            subj_name = v.split('\n')[0].strip()
            disciplinas.append(subj_name)
            if col_primeira_disciplina == -1: col_primeira_disciplina = i
                
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
        if 'ATIVO' not in sit.upper() and sit.upper() != 'AT': continue
            
        num = str(row[col_primeira_disciplina]).strip() if pd.notna(row[col_primeira_disciplina]) else ""
        if num.lower() in ["nan", "none"]: num = ""
        
        val_tf = str(row[idx_tf]).strip() if idx_tf != -1 else "-"
        val_fre = str(row[idx_fre]).strip() if idx_fre != -1 else "-"
        val_ft_an = str(row[idx_ft_an]).strip() if idx_ft_an != -1 else "-"
        val_fre_an = str(row[idx_fre_an]).strip() if idx_fre_an != -1 else "-"
        
        alunos.append({
            'Nº': num, 'Nome do Aluno': nome, 'Sit.': sit,
            'TF': val_tf, 'Fre(%)': val_fre, 'FT An': val_ft_an, 'Fre An(%)': val_fre_an
        })
        
    return pd.DataFrame(alunos), disciplinas, turma

# --- INÍCIO DO APP STREAMLIT ---
st.set_page_config(page_title="Gerador de Conselho", layout="centered")

st.title("📄 Sistema do Conselho de Classe (PDF)")
st.write("**ESCOLA ESTADUAL AMERICO BRASILIENSE DOUTOR**")

turno = st.selectbox("Turno:", ["Manhã", "Tarde", "Noite", "Integral"])

st.subheader("Passo 1: Carregar Turmas (Mapões da SED)")
mapoes_files = st.file_uploader("Suba os arquivos de Mapão aqui", type=["xlsx", "xls"], accept_multiple_files=True)

if mapoes_files:
    st.success(f"✅ {len(mapoes_files)} arquivo(s) carregado(s).")
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
        st.warning("⚠️ Aguardando arquivos da Prova Paulista.")
        pode_gerar = False

    if pode_gerar and st.button("Gerar Caderno do Conselho (PDF)", type="primary"):
        with st.spinner('Desenhando páginas A4 e cruzando dados...'):
            try:
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

                # CONFIGURAÇÃO DO PDF
                pdf_buffer = io.BytesIO()
                # A4 Paisagem com margens de 20 pontos
                doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
                elementos_pdf = []
                escola_nome = "E.E. Dr. Américo Brasiliense"

                for arq_mapao in mapoes_files:
                    df_mapao, disciplinas, nome_turma = ler_novo_mapao(arq_mapao)
                    
                    if df_mapao is None or df_mapao.empty: continue
                        
                    df_mapao['Nome_Chave'] = df_mapao['Nome do Aluno'].apply(normalizar_nome)

                    if tem_prova == "Sim" and not df_prova_reduzido.empty:
                        df_final = pd.merge(df_mapao, df_prova_reduzido, on='Nome_Chave', how='left')
                    else:
                        df_final = df_mapao.copy()
                        df_final['Prova Paulista'] = ""
                        
                    if 'Prova Paulista' in df_final.columns:
                        df_final['Prova Paulista'] = df_final['Prova Paulista'].fillna("-")

                    # Montagem da Tabela para o PDF
                    colunas_finais = ['Nº', 'Nome', 'Sit.', 'TF', 'Fre(%)', 'FT An', 'Fre An(%)', 'Prova'] + [abreviar_disciplina(d) for d in disciplinas] + ['Obs.']
                    
                    # Cálculo matemático perfeito para preencher a largura do A4 (802 pontos)
                    fixed_widths = [18, 140, 30, 20, 32, 28, 42, 28] # Total = 338
                    obs_width = 50
                    rem_width = 802 - sum(fixed_widths) - obs_width
                    disc_width = rem_width / max(len(disciplinas), 1)
                    widths = fixed_widths + [disc_width]*len(disciplinas) + [obs_width]

                    title = f"{escola_nome}  |  {nome_turma} – {turno}  |  Conselho 1º Bimestre / 2026"
                    
                    data_table = []
                    data_table.append([title] + [''] * (len(colunas_finais) - 1))
                    data_table.append(colunas_finais)

                    df_final = df_final.fillna("-")
                    for _, row in df_final.iterrows():
                        # Corta o nome um pouco para não quebrar a tabela se o nome for gigante
                        nome_trunc = str(row['Nome do Aluno'])[:38] 
                        sit_trunc = str(row['Sit.'])[:5] # Ativo -> Ativo
                        
                        linha = [
                            str(row['Nº']), nome_trunc, sit_trunc,
                            str(row['TF']).replace("nan", "-"), str(row['Fre(%)']).replace("nan", "-"), 
                            str(row['FT An']).replace("nan", "-"), str(row['Fre An(%)']).replace("nan", "-"),
                            str(row['Prova Paulista'])
                        ] + ["" for _ in disciplinas] + [""]
                        data_table.append(linha)

                    # Desenho e Estilo da Tabela
                    t = Table(data_table, colWidths=widths, repeatRows=2)
                    t.setStyle(TableStyle([
                        # Cabeçalho Principal (Mesclado)
                        ('SPAN', (0,0), (-1,0)),
                        ('ALIGN', (0,0), (-1,0), 'CENTER'),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,0), 9),
                        ('BOTTOMPADDING', (0,0), (-1,0), 6),

                        # Linha de Colunas
                        ('BACKGROUND', (0,1), (-1,1), colors.HexColor("#4F81BD")),
                        ('TEXTCOLOR', (0,1), (-1,1), colors.whitesmoke),
                        ('ALIGN', (0,1), (-1,1), 'CENTER'),
                        ('VALIGN', (0,1), (-1,1), 'MIDDLE'),
                        ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,1), (-1,1), 6.5),

                        # Dados dos Alunos
                        ('FONTNAME', (0,2), (-1,-1), 'Helvetica'),
                        ('FONTSIZE', (0,2), (-1,-1), 6.5), # Fonte pequena para caber
                        ('ALIGN', (0,2), (0,-1), 'CENTER'), # Centraliza Nº
                        ('ALIGN', (1,2), (1,-1), 'LEFT'),   # Esquerda no Nome
                        ('ALIGN', (2,2), (-1,-1), 'CENTER'), # Centraliza o resto
                        ('VALIGN', (0,2), (-1,-1), 'MIDDLE'),

                        # Bordas e Cores alternadas (Zebra) para facilitar leitura na régua
                        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                        ('ROWBACKGROUNDS', (0,2), (-1,-1), [colors.white, colors.HexColor("#F2F2F2")])
                    ]))
                    
                    elementos_pdf.append(t)
                    elementos_pdf.append(PageBreak()) # Garante que a próxima turma comece em uma página NOVA

                # Finaliza a montagem do documento PDF
                doc.build(elementos_pdf)
                
                pdf_buffer.seek(0)
                st.success("Caderno do conselho gerado! Pronto para impressão em A4.")
                st.download_button(
                    label="⬇️ Baixar Caderno Oficial (PDF)",
                    data=pdf_buffer,
                    file_name=f"Conselho_Classe_{turno}_Caderno.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"Ocorreu um problema ao gerar o PDF: {e}")
