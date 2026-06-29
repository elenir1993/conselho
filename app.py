import streamlit as st
import pandas as pd
import io
import unicodedata
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Spacer, Paragraph, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet
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

# 3. LEITOR DO NOVO MAPÃO DA SED (Com Radar de Notas Vermelhas)
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
    subject_m_cols = {}
    idx_aluno, idx_sit, col_primeira_disciplina = -1, -1, -1
    
    for i, val in enumerate(row_1):
        v = str(val).strip()
        if v.upper() == 'ALUNO': idx_aluno = i
        elif v.upper() in ['SITUAÇÃO', 'SITUACAO', 'SIT']: idx_sit = i
        elif v and v.upper() != 'TOTAL':
            subj_name = v.split('\n')[0].strip()
            disciplinas.append(subj_name)
            subject_m_cols[subj_name] = i + 1 
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
        
        aluno_data = {
            'Nº': num, 'Nome do Aluno': nome, 'Sit.': sit,
            'TF': val_tf, 'Fre(%)': val_fre, 'FT An': val_ft_an, 'Fre An(%)': val_fre_an
        }
        
        for subj_name, m_col_idx in subject_m_cols.items():
            m_val = str(row[m_col_idx]).strip().replace(',', '.')
            is_red = False
            try:
                if float(m_val) < 5.0:
                    is_red = True
            except ValueError:
                pass 
                
            aluno_data[subj_name] = "RED" if is_red else ""
            
        alunos.append(aluno_data)
        
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
        with st.spinner('Aplicando inteligência de dados, radar de faltas e notas...'):
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
                doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
                elementos_pdf = []
                escola_nome = "E.E. Dr. Américo Brasiliense"
                styles_text = getSampleStyleSheet()

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

                    colunas_finais = ['Nº', 'Nome', 'Sit.', 'TF', 'Fre(%)', 'FT An', 'Fre An(%)', 'Prova'] + [abreviar_disciplina(d) for d in disciplinas] + ['Obs.']
                    
                    # Leve ajuste na largura da Obs para acomodar a nova frase
                    fixed_widths = [18, 140, 30, 20, 32, 28, 42, 28] 
                    obs_width = 54
                    rem_width = 802 - sum(fixed_widths) - obs_width
                    disc_width = rem_width / max(len(disciplinas), 1)
                    widths = fixed_widths + [disc_width]*len(disciplinas) + [obs_width]

                    title = f"{escola_nome}  |  {nome_turma} – {turno}  |  Conselho 2º Bimestre / 2026"
                    
                    data_table = []
                    data_table.append([title] + [''] * (len(colunas_finais) - 1))
                    data_table.append(colunas_finais)

                    custom_styles = [
                        ('SPAN', (0,0), (-1,0)),
                        ('ALIGN', (0,0), (-1,0), 'CENTER'),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,0), (-1,0), 9),
                        ('BOTTOMPADDING', (0,0), (-1,0), 6),
                        ('BACKGROUND', (0,1), (-1,1), colors.HexColor("#4F81BD")),
                        ('TEXTCOLOR', (0,1), (-1,1), colors.whitesmoke),
                        ('ALIGN', (0,1), (-1,1), 'CENTER'),
                        ('VALIGN', (0,1), (-1,1), 'MIDDLE'),
                        ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
                        ('FONTSIZE', (0,1), (-1,1), 6.5),
                        ('FONTNAME', (0,2), (-1,-1), 'Helvetica'),
                        ('FONTSIZE', (0,2), (-1,-1), 6.5),
                        ('ALIGN', (0,2), (0,-1), 'CENTER'), 
                        ('ALIGN', (1,2), (1,-1), 'LEFT'),   
                        ('ALIGN', (2,2), (-1,-1), 'CENTER'), 
                        ('VALIGN', (0,2), (-1,-1), 'MIDDLE'),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                    ]

                    df_final = df_final.fillna("-")
                    for row_idx, (_, row) in enumerate(df_final.iterrows()):
                        nome_trunc = str(row['Nome do Aluno'])[:38] 
                        sit_trunc = str(row['Sit.'])[:5] 
                        
                        # Tratamento da Frequência para cor vermelha
                        fre_str = str(row['Fre(%)']).replace("nan", "-")
                        fre_num_str = fre_str.replace('%', '').strip()
                        is_low_freq = False
                        try:
                            if float(fre_num_str) < 75.0:
                                is_low_freq = True
                        except ValueError:
                            pass
                            
                        linha = [
                            str(row['Nº']), nome_trunc, sit_trunc,
                            str(row['TF']).replace("nan", "-"), fre_str, 
                            str(row['FT An']).replace("nan", "-"), str(row['Fre An(%)']).replace("nan", "-"),
                            str(row['Prova Paulista'])
                        ]
                        
                        # Pinta Frequência de vermelho claro se for menor que 75%
                        if is_low_freq:
                            pdf_r = row_idx + 2
                            pdf_c_fre = 4 # Posição da coluna Fre(%)
                            custom_styles.append(('BACKGROUND', (pdf_c_fre, pdf_r), (pdf_c_fre, pdf_r), colors.HexColor("#FFCCCC")))
                            custom_styles.append(('TEXTCOLOR', (pdf_c_fre, pdf_r), (pdf_c_fre, pdf_r), colors.HexColor("#CC0000")))
                        
                        # Contagem do Radar de Notas e preenchimento dos quadradinhos
                        red_count = 0
                        for subj_idx, subj in enumerate(disciplinas):
                            val = row.get(subj, "")
                            if val == "RED":
                                red_count += 1
                                linha.append("........") 
                                pdf_r = row_idx + 2
                                pdf_c = 8 + subj_idx
                                custom_styles.append(('BACKGROUND', (pdf_c, pdf_r), (pdf_c, pdf_r), colors.HexColor("#EAEAEA")))
                                custom_styles.append(('TEXTCOLOR', (pdf_c, pdf_r), (pdf_c, pdf_r), colors.HexColor("#A0A0A0")))
                            else:
                                linha.append("")
                        
                        # Lógica da Observação (> 60% com pontilhado vermelho)
                        obs_text = ""
                        if len(disciplinas) > 0 and (red_count / len(disciplinas)) > 0.6:
                            obs_text = "Baixo desemp.\nbimestral"
                            
                        linha.append(obs_text) 
                        data_table.append(linha)

                    t = Table(data_table, colWidths=widths, repeatRows=2)
                    t.setStyle(TableStyle(custom_styles))
                    elementos_pdf.append(t)
                    
                    bloco_final = []
                    bloco_final.append(Spacer(1, 20))
                    bloco_final.append(Paragraph("<b>Perfil da Turma / Decisões do Conselho:</b>", styles_text['Normal']))
                    bloco_final.append(Spacer(1, 8))
                    
                    linhas_perfil = [[""] for _ in range(6)]
                    t_perfil = Table(linhas_perfil, colWidths=[800], rowHeights=[18]*6)
                    t_perfil.setStyle(TableStyle([
                        ('LINEBELOW', (0,0), (-1,-1), 0.5, colors.gray)
                    ]))
                    bloco_final.append(t_perfil)
                    
                    bloco_final.append(Spacer(1, 25))
                    bloco_final.append(Paragraph("<b>Assinaturas dos Professores:</b>", styles_text['Normal']))
                    bloco_final.append(Spacer(1, 10))
                    
                    assinaturas_data = []
                    chunk_size = 4
                    abbrev_disciplinas = [abreviar_disciplina(d) for d in disciplinas]
                    for idx_chunk in range(0, len(abbrev_disciplinas), chunk_size):
                        chunk = abbrev_disciplinas[idx_chunk : idx_chunk+chunk_size]
                        linha_assinatura = [f"{d}: ______________________________" for d in chunk]
                        while len(linha_assinatura) < chunk_size:
                            linha_assinatura.append("") 
                        assinaturas_data.append(linha_assinatura)
                        
                    t_assinaturas = Table(assinaturas_data, colWidths=[200]*4, rowHeights=[30]*len(assinaturas_data))
                    t_assinaturas.setStyle(TableStyle([
                        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
                        ('FONTSIZE', (0,0), (-1,-1), 8),
                        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
                    ]))
                    bloco_final.append(t_assinaturas)

                    elementos_pdf.append(KeepTogether(bloco_final))
                    elementos_pdf.append(PageBreak()) 

                doc.build(elementos_pdf)
                pdf_buffer.seek(0)
                
                st.success("Ata Oficial gerada! Análise de frequência e baixo desempenho já aplicadas nas observações.")
                st.download_button(
                    label="⬇️ Baixar Ata Oficial (PDF)",
                    data=pdf_buffer,
                    file_name=f"Conselho_Classe_{turno}_Ata_Oficial.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"Ocorreu um problema ao gerar o PDF: {e}")
