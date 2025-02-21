import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def criar_documento_word(arquivo_excel, nome_planilha):
    # Lê o arquivo Excel
    df = pd.read_excel(arquivo_excel, sheet_name=nome_planilha, header=None)

    # Remove linhas sem semestre (coluna I = índice 8)
    df = df.dropna(subset=[8])

    # Converte a coluna de semestre para string
    df[8] = df[8].astype(str)
    
    # Remove disciplinas duplicadas, considerando nome e semestre
    df_disciplinas = df[[7, 8, 12, 15]].drop_duplicates(subset=[7, 8])
    
    # Cria o documento Word
    documento = Document()
    
    # Define o estilo padrão
    estilo_padrao = documento.styles['Normal']
    fonte = estilo_padrao.font
    fonte.name = 'Calibri'
    fonte.size = Pt(11)
    fonte.color.rgb = RGBColor(0, 0, 0)

    # Obtém os semestres únicos e ordenados
    semestres = sorted(df_disciplinas[8].unique())

    for semestre in semestres:
        if semestre == '0':
            continue  # Pula as optativas por enquanto
        
        # Adiciona o título do semestre
        paragrafo_semestre = documento.add_heading(level=1)
        paragrafo_semestre.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragrafo_semestre.add_run(f"{semestre}° Semestre")
        run.font.bold = True

        # Filtra disciplinas do semestre
        disciplinas = df_disciplinas[df_disciplinas[8] == semestre]

        for _, disciplina in disciplinas.iterrows():
            adicionar_disciplina_no_documento(documento, df, disciplina)

    # Adiciona as disciplinas optativas no final
    optativas = df_disciplinas[df_disciplinas[8] == '0']
    if not optativas.empty:
        paragrafo_optativas = documento.add_heading(level=1)
        paragrafo_optativas.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragrafo_optativas.add_run("Disciplinas Optativas")
        run.font.bold = True

        for _, disciplina in optativas.iterrows():
            adicionar_disciplina_no_documento(documento, df, disciplina)
    
    # Salva o documento
    arquivo_saida = "finalizado.docx"
    documento.save(arquivo_saida)
    print(f"Documento gerado: {arquivo_saida}")

def adicionar_disciplina_no_documento(documento, df, disciplina):
    """Adiciona uma disciplina ao documento Word."""
    paragrafo_disciplina = documento.add_paragraph()
    run_disciplina = paragrafo_disciplina.add_run(f"Nome da disciplina: {disciplina[7]}")
    run_disciplina.bold = True
    
    paragrafo_ementa = documento.add_paragraph()
    run_ementa_title = paragrafo_ementa.add_run("Ementa:")
    run_ementa_title.bold = True
    ementa_texto = disciplina[12] if pd.notna(disciplina[12]) else "Nenhuma informação disponível."
    documento.add_paragraph(ementa_texto)

    # Bibliografias
    bibliografias = df[df[7] == disciplina[7]]
    adicionar_bibliografia(documento, bibliografias, "Básica")
    adicionar_bibliografia(documento, bibliografias, "Complementar")

    # Adequação Referendado
    paragrafo_adequacao = documento.add_paragraph()
    run_adequacao_title = paragrafo_adequacao.add_run("Adequação Referendado:")
    run_adequacao_title.bold = True
    documento.add_paragraph(disciplina[15] if pd.notna(disciplina[15]) else "Nenhuma informação disponível.")

def adicionar_bibliografia(documento, df, tipo):
    """Adiciona bibliografia ao documento."""
    paragrafo_bib = documento.add_paragraph()
    run_bib = paragrafo_bib.add_run(f"Bibliografia {tipo}:")
    run_bib.bold = True
    bibliografias = df[df[10].str.contains(tipo, na=False)][13].dropna().drop_duplicates()
    
    if not bibliografias.empty:
        for texto in bibliografias:
            documento.add_paragraph(texto)
    else:
        documento.add_paragraph("Nenhuma bibliografia disponível.")

# Caminho do arquivo Excel
arquivo_excel = "ARTES VISUAIS.xlsx"  # Substitua pelo caminho correto
nome_planilha = "RelatorioCursoBibliografia"  # Nome da planilha correta

# Chamada da função
criar_documento_word(arquivo_excel, nome_planilha)

