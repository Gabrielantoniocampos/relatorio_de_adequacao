# Documentação Técnica: Geração de Documento Word a Partir de Planilha Excel

## Visão Geral
Este script automatiza a criação de um documento no formato **Word (.docx)** a partir de dados contidos em uma planilha **Excel (.xlsx)**.
Ele filtra, organiza e estrutura informações de disciplinas e suas respectivas bibliografias em um documento final chamado **"finalizado.docx"**.

O objetivo principal é gerar um documento estruturado com as disciplinas de um curso, separadas por semestre, incluindo ementas, bibliografias e adequação referendada.

---

## Requisitos

### Bibliotecas Utilizadas
```python
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
```

Caso essas bibliotecas ainda não estejam instaladas, use o seguinte comando para instalá-las:

```bash
pip install pandas python-docx
```

---

## Entrada Esperada
O script utiliza um arquivo **Excel (.xlsx)** que deve conter:

- **Nome da disciplina** (coluna 7)
- **Semestre da disciplina** (coluna 8)
- **Ementa** (coluna 12)
- **Bibliografia** (coluna 13)
- **Tipo de bibliografia** (coluna 10) - "Básica" ou "Complementar"
- **Adequação referendada** (coluna 15)

---

## Saída Gerada
O script gera um documento chamado **"finalizado.docx"**, contendo:
- Disciplinas organizadas por semestre.
- Disciplinas optativas no final do documento.
- Informações detalhadas sobre cada disciplina, incluindo ementa, bibliografia e adequação referendada.

---

## Estrutura do Script
```python
def criar_documento_word(arquivo_excel, nome_planilha):
    df = pd.read_excel(arquivo_excel, sheet_name=nome_planilha, header=None)
    df = df.dropna(subset=[8])
    df[8] = df[8].astype(str)
    df_disciplinas = df[[7, 8, 12, 15]].drop_duplicates(subset=[7, 8])
    
    documento = Document()
    fonte = documento.styles['Normal'].font
    fonte.name = 'Calibri'
    fonte.size = Pt(11)
    fonte.color.rgb = RGBColor(0, 0, 0)
    
    semestres = sorted(df_disciplinas[8].unique())
    for semestre in semestres:
        if semestre == '0':
            continue
        paragrafo_semestre = documento.add_heading(level=1)
        paragrafo_semestre.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragrafo_semestre.add_run(f"{semestre}° Semestre")
        run.font.bold = True
    
    documento.save("finalizado.docx")
    print("Documento gerado: finalizado.docx")
```

---

## Como Utilizar o Script

1. Certifique-se de que o arquivo Excel está no diretório correto e que a planilha tem os dados esperados.
2. Altere o nome do arquivo e da planilha no código:

```python
arquivo_excel = "ARTES VISUAIS.xlsx"
nome_planilha = "RelatorioCursoBibliografia"
criar_documento_word(arquivo_excel, nome_planilha)
```

3. Execute o script no terminal ou em um ambiente Python:

```bash
python nome_do_script.py
```

4. O arquivo **"finalizado.docx"** será gerado no mesmo diretório.

---

## Possíveis Melhorias Futuras
- Tornar o script mais flexível para diferentes estruturas de planilhas.
- Adicionar opção para selecionar a saída do arquivo Word.
- Melhorar o tratamento de erros caso os dados estejam incompletos.



- Tornar o script mais flexível para diferentes estruturas de planilhas.
- Adicionar opção para selecionar a saída do arquivo Word.
- Melhorar o tratamento de erros caso os dados estejam incompletos.
Conclusão
Este script automatiza a geração de documentos a partir de dados estruturados no Excel, facilitando a criação de relatórios organizados sem a necessidade de formatação manual.
#
