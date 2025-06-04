import sys
import os
import xml.etree.ElementTree as ET
import zipfile

def _load_shared_strings(z):
    """Retorna lista de shared strings a partir de um ZipFile aberto."""
    try:
        xml = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    strings = []
    for si in xml.findall("m:si", ns):
        t = si.find(".//m:t", ns)
        strings.append(t.text if t is not None else "")
    return strings


def _column_index(cell_ref):
    """Converte referência de coluna Excel (ex: 'A', 'AB') para índice zero-based."""
    col = ''.join(filter(str.isalpha, cell_ref))
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return idx - 1


def _read_sheet(z, sheet_path, shared_strings):
    """Retorna lista de rows, onde cada row é uma lista de valores."""
    xml = ET.fromstring(z.read(sheet_path))
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows = []
    for row in xml.findall(".//m:sheetData/m:row", ns):
        row_values = {}
        for c in row.findall("m:c", ns):
            ref = c.get("r")
            idx = _column_index(ref)
            v = c.find("m:v", ns)
            if v is None:
                val = ""
            elif c.get("t") == "s":
                val = shared_strings[int(v.text)]
            else:
                val = v.text
            row_values[idx] = val
        if row_values:
            max_idx = max(row_values.keys())
            row_list = [row_values.get(i, "") for i in range(max_idx + 1)]
            rows.append(row_list)
    return rows


def read_entregas_sheet(xlsx_path):
    """
    Lê a sheet 'ENTREGAS' do xlsx_path e retorna (headers, data_rows).
    O cabeçalho é considerado na linha de índice 2 (3a linha do Excel),
    e as linhas de dados a partir da 4a linha.
    """
    with zipfile.ZipFile(xlsx_path) as z:
        shared_strings = _load_shared_strings(z)

        # Procura, dentro de workbook.xml, pela sheet chamada "ENTREGAS"
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        rid = None
        for sheet in wb.findall(".//m:sheets/m:sheet", ns):
            if sheet.get("name") == "ENTREGAS":
                rid = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                break
        if not rid:
            raise ValueError("Sheet 'ENTREGAS' não encontrada")

        # Agora encontramos, em workbook.xml.rels, qual o Target para esse rId
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        nsr = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
        target = None
        for rel in rels.findall("r:Relationship", nsr):
            if rel.get("Id") == rid:
                target = rel.get("Target")
                break
        if not target:
            raise ValueError("Relacionamento para sheet 'ENTREGAS' não encontrado")

        sheet_path = os.path.join("xl", target)
        rows = _read_sheet(z, sheet_path, shared_strings)
        if len(rows) < 4:
            return [], []
        headers = rows[2]    # cabeçalho na 3a linha (índice 2)
        data_rows = rows[3:] # dados a partir da 4a linha (índice ≥3)
        return headers, data_rows


def _escape(text):
    return text.replace("\\", "\\\\") \
               .replace("(", "\\(") \
               .replace(")", "\\)") \
               .replace("\r", "") \
               .replace("\n", "\\n")


def write_pdf(path, lines):
    """Cria um PDF mínimo em `path` com as linhas contidas em `lines`."""
    # Monta o conteúdo da página (em coordenadas “y” decrescentes):
    y = 750
    content_lines = ["BT", "/F1 12 Tf"]
    for line in lines:
        content_lines.append(f"72 {y} Td ({_escape(line)}) Tj")
        y -= 14
    content_lines.append("ET")
    content_stream = "\n".join(content_lines) + "\n"

    # Define os objetos básicos do PDF
    objects = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
        "/Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        f"<< /Length {len(content_stream.encode('utf-8'))} >>\nstream\n"
        f"{content_stream}endstream",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]

    offsets = []
    offset = len("%PDF-1.4\n")
    for obj in objects:
        offsets.append(offset)
        offset += len(f"{objects.index(obj)+1} 0 obj\n{obj}\nendobj\n".encode("utf-8"))
    xref_offset = offset

    with open(path, "wb") as f:
        f.write(b"%PDF-1.4\n")
        for i, obj in enumerate(objects, start=1):
            f.write(f"{i} 0 obj\n".encode("utf-8"))
            f.write(obj.encode("utf-8"))
            f.write(b"\nendobj\n")
        f.write(f"xref\n0 {len(objects)+1}\n".encode("utf-8"))
        f.write(b"0000000000 65535 f \n")
        for off in offsets:
            f.write(f"{off:010d} 00000 n \n".encode("utf-8"))
        f.write(b"trailer\n")
        f.write(f"<< /Size {len(objects)+1} /Root 1 0 R >>\n".encode("utf-8"))
        f.write(b"startxref\n")
        f.write(f"{xref_offset}\n".encode("utf-8"))
        f.write(b"%%EOF")


def main():
    # Monta o caminho para “data/PROJECT MANAGEMENT.xlsx” de forma dinâmica
    script_dir = os.path.dirname(__file__)
    xlsx_path = os.path.join(script_dir, "data", "PROJECT MANAGEMENT.xlsx")

    # Caso o arquivo não exista, imprime erro e sai
    if not os.path.isfile(xlsx_path):
        print(f"ERRO: não encontrei o arquivo em: {xlsx_path!r}")
        sys.exit(1)

    # Pasta de saída (será criada dentro de "data/pdf_output")
    out_dir = os.path.join(script_dir, "data", "pdf_output")
    os.makedirs(out_dir, exist_ok=True)

    # Lê cabeçalhos e linhas de dados da sheet ENTREGAS
    headers, rows = read_entregas_sheet(xlsx_path)
    if not headers:
        print("Nenhum dado encontrado na aba 'ENTREGAS'.")
        return

    # Para cada linha, monta as linhas "Header: Valor" e cria um PDF
    for idx, row in enumerate(rows, start=1):
        lines = []
        for h, v in zip(headers, row):
            lines.append(f"{h}: {v}")
        pdf_name = os.path.join(out_dir, f"row_{idx}.pdf")
        write_pdf(pdf_name, lines)
        print(f"Wrote {pdf_name}")


if __name__ == "__main__":
    main()
