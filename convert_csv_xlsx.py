import csv
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

def csv_para_xlsx(csv_path, xlsx_path):
    """
    Converte CSV para XLSX de forma confiável usando apenas bibliotecas padrão.
    """
    # Lê o arquivo CSV
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    # Prepara os dados XML escapando caracteres especiais
    rows_xml = []
    for i, row in enumerate(rows, 1):
        cells = []
        for j, cell in enumerate(row, 1):
            # Escapa caracteres especiais no conteúdo das células
            escaped_cell = escape(str(cell))
            cells.append(f'<c t="inlineStr"><is><t>{escaped_cell}</t></is></c>')
        rows_xml.append(f'<row r="{i}">{"".join(cells)}</row>')
    
    # Cria a estrutura completa do XLSX
    with zipfile.ZipFile(xlsx_path, 'w') as zf:
        # Adiciona os arquivos necessários para um XLSX válido
        zf.writestr('[Content_Types].xml', """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>""")

        zf.writestr('_rels/.rels', """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""")

        zf.writestr('xl/workbook.xml', """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Planilha1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>""")

        zf.writestr('xl/_rels/workbook.xml.rels', """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""")

        zf.writestr('xl/worksheets/sheet1.xml', f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        {"".join(rows_xml)}
    </sheetData>
</worksheet>""")

def converter_todos_csv():
    """Converte todos os arquivos CSV na pasta atual para XLSX"""
    # Cria pasta de saída se não existir, altere o nome da pasta conforme a necessidade.
    output_dir = Path("polos_prova_excel")
    output_dir.mkdir(exist_ok=True)
    
    # Processa cada arquivo CSV
    for csv_file in Path.cwd().glob("*.csv"):
        xlsx_file = output_dir / f"{csv_file.stem}.xlsx"
        csv_para_xlsx(csv_file, xlsx_file)
        print(f"Convertido: {csv_file.name} → {xlsx_file.name}")

if __name__ == "__main__":
    converter_todos_csv()