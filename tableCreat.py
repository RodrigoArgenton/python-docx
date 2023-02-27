from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


def tableStyle(s):
    style = (
        'Table Grid',        # [0]
        'Light Shading',     # [1]
        'Light List',        # [2]
        'Medium Shading 1',  # [3]
        'Medium Shading 2',  # [4]
        'Medium List 1',     # [5]
        'Medium List 2',     # [6]
        'Medium Grid 1',     # [7]
        'Medium Grid 2',     # [8]
        'Medium Grid 3',     # [9]
    )
    for i, v in enumerate(style):
        if s == i:
            return v
    return 'Table Grid'


def newTable(styleTable=0, condHeaderCenter=False, condTableCenter=False,):
    items = (
        (1, 'rua barão, 987', 'Marcio da silva'),
        (2, 'rua trad, 67', 'José silva'),
        (3, 'rua marcinho, 09', 'Maria silva'),
    )

    document = Document()
    table = document.add_table(rows=1, cols=3)
    table.style = tableStyle(styleTable)

    head_cell = table.rows[0].cells
    head_cell[0].text = 'Number'
    head_cell[1].text = 'Endereço'
    head_cell[2].text = 'Nome completo'

    for item in items:
        cells = table.add_row().cells
        cells[0].text = str(item[0])
        cells[1].text = item[1]
        cells[2].text = item[2]

    if condHeaderCenter is True:
        for cell in table.rows[0].cells:
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_TABLE_ALIGNMENT.CENTER

    if condTableCenter is True:
        # Percorra todas as células da tabela
        for row in table.rows:
            for cell in row.cells:
                # Defina o alinhamento horizontal da célula como centralizado
                cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    document.save('teste.docx')
