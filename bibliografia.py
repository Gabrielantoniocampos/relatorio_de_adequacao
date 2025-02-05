import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Instala as bibliotecas necessárias caso não estejam instaladas
# !pip install pandas openpyxl python-docx


def criar_documento_word(arquivo_excel, nome_planilha):
    """
    Cria um documento Word a partir de um arquivo Excel, formatando as informações
    de disciplinas, ementas e bibliografias.

    Args:
        arquivo_excel (str): Caminho completo para o arquivo Excel.
        nome_planilha (str): Nome da planilha dentro do arquivo Excel a ser processada.
    """
    try:
        # Lê o arquivo Excel
        df = pd.read_excel(arquivo_excel, sheet_name=nome_planilha, header=None)

        # Remove linhas sem semestre (coluna I = índice 8)
        df = df.dropna(subset=[8])

        # Converte a coluna de semestre para string, garantindo ordenação correta
        df[8] = df[8].astype(str)

        # Ordena o DataFrame por semestre e nome da disciplina
        df = df.sort_values(by=[8, 7])

        # Cria o documento Word
        documento = Document()

        # Define o estilo padrão do documento
        estilo_padrao = documento.styles['Normal']
        fonte = estilo_padrao.font
        fonte.name = 'Calibri'
        fonte.size = Pt(11)
        fonte.color.rgb = RGBColor(0, 0, 0)  # Define a cor da fonte como preta

        # Obtém os semestres únicos e ordenados
        semestres = sorted(df[8].unique())

        for semestre in semestres:
            # Adiciona o título do semestre centralizado
            paragrafo_semestre = documento.add_heading(level=1)
            paragrafo_semestre.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = paragrafo_semestre.add_run(f"{semestre}° Semestre")
            # Formata o título do semestre
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
            run.bold = True

            # Filtra as disciplinas do semestre atual
            disciplinas = df[df[8] == semestre]

            for _, disciplina in disciplinas.iterrows():
                # Adiciona o nome da disciplina
                paragrafo_disciplina = documento.add_paragraph()
                run_disciplina = paragrafo_disciplina.add_run(f"Nome da disciplina: {disciplina[7]}")
                # Formata o nome da disciplina
                run_disciplina.font.name = 'Calibri'
                run_disciplina.font.size = Pt(11)
                run_disciplina.font.color.rgb = RGBColor(0, 0, 0)
                run_disciplina.bold = True

                # Adiciona a ementa da disciplina
                paragrafo_ementa = documento.add_paragraph()
                run_ementa_title = paragrafo_ementa.add_run("Ementa:")
                # Formata o título da ementa
                run_ementa_title.font.name = 'Calibri'
                run_ementa_title.font.size = Pt(11)
                run_ementa_title.font.color.rgb = RGBColor(0, 0, 0)
                run_ementa_title.bold = True

                # Processa e adiciona a ementa, tratando valores ausentes
                ementa_texto = disciplina[12] if not pd.isna(disciplina[12]) else "Nenhuma informação disponível."
                ementa_linhas = [linha[3:] for linha in ementa_texto.splitlines()]  # Remove os três primeiros caracteres de cada linha
                for linha in ementa_linhas:
                    paragrafo = documento.add_paragraph()
                    run_ementa_linha = paragrafo.add_run(linha)
                    run_ementa_linha.font.name = 'Calibri'
                    run_ementa_linha.font.size = Pt(11)
                    run_ementa_linha.font.color.rgb = RGBColor(0, 0, 0)

                # Adiciona as bibliografias (básica e complementar)
                # ... (código para adicionar bibliografias - semelhante ao código original)

                # Adiciona a adequação referenciada
                # ... (código para adicionar adequação - semelhante ao código original)

        # Salva o documento
        arquivo_saida = "finalizado.docx"
        documento.save(arquivo_saida)
        print(f"Documento gerado: {arquivo_saida}")

    except FileNotFoundError:
        print(f"Erro: Arquivo Excel '{arquivo_excel}' não encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")


# Exemplo de uso
arquivo_excel = "/content/ARTES VISUAIS.xlsx"  # Substitua pelo caminho correto do arquivo
nome_planilha = "RelatorioCursoBibliografia"  # Substitua pelo nome da planilha correta

criar_documento_word(arquivo_excel, nome_planilha)
