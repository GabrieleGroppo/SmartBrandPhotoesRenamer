"""Lettore XLSX minimale basato solo sulla standard library.

Evita la dipendenza da pandas/openpyxl: un file .xlsx e' uno zip di XML,
quindi basta zipfile + xml.etree per estrarre fogli, intestazioni e righe.
"""
import zipfile
from xml.etree import ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _col_to_idx(cell_ref):
    letters = "".join(c for c in cell_ref if c.isalpha())
    idx = 0
    for c in letters:
        idx = idx * 26 + (ord(c.upper()) - ord("A") + 1)
    return idx - 1


def _load_shared_strings(zf):
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    return ["".join(t.text or "" for t in si.findall(".//m:t", NS)) for si in root.findall("m:si", NS)]


def _sheet_name_to_path(zf):
    root = ET.fromstring(zf.read("xl/workbook.xml"))
    names = [s.get("name") for s in root.findall(".//m:sheets/m:sheet", NS)]
    sheet_files = sorted(
        n for n in zf.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
    )
    return {name: sheet_files[i] for i, name in enumerate(names) if i < len(sheet_files)}


def _read_sheet_rows(zf, sheet_path, shared):
    root = ET.fromstring(zf.read(sheet_path))
    rows = []
    for row in root.findall(".//m:sheetData/m:row", NS):
        cells = {}
        for c in row.findall("m:c", NS):
            idx = _col_to_idx(c.get("r"))
            cell_type = c.get("t")
            v = c.find("m:v", NS)
            if v is not None:
                cells[idx] = shared[int(v.text)] if cell_type == "s" else v.text
            else:
                is_node = c.find("m:is", NS)
                if is_node is not None:
                    cells[idx] = "".join(t.text or "" for t in is_node.findall(".//m:t", NS))
        rows.append(cells)
    return rows


def read_workbook(path):
    """Ritorna {nome_foglio: [ {intestazione: valore, ...}, ... ]} per ogni foglio.

    La prima riga di ogni foglio e' trattata come intestazione. I valori sono
    stringhe (eventualmente None se la cella e' vuota), stessa convenzione di
    ``pandas.read_excel(dtype=str)`` ma senza introdurre notazione scientifica
    o suffissi ".0" sui codici numerici lunghi (es. EAN).
    """
    with zipfile.ZipFile(path) as zf:
        shared = _load_shared_strings(zf)
        sheet_map = _sheet_name_to_path(zf)

        workbook = {}
        for sheet_name, sheet_path in sheet_map.items():
            raw_rows = _read_sheet_rows(zf, sheet_path, shared)
            if not raw_rows:
                workbook[sheet_name] = []
                continue

            header_row = raw_rows[0]
            # Alcune colonne dati non hanno intestazione (nessuna cella nella riga 1,
            # es. colonne aggiunte in fondo al foglio senza etichetta): calcolare
            # max_col solo dalla riga di intestazione le farebbe scartare in
            # silenzio. Si considera quindi la larghezza massima su TUTTE le righe.
            max_col = max((max(r.keys(), default=-1) for r in raw_rows), default=-1)
            headers = [header_row.get(i) or f"__col{i}__" for i in range(max_col + 1)]

            records = []
            for raw in raw_rows[1:]:
                if not raw:
                    continue
                record = {}
                for i, name in enumerate(headers):
                    val = raw.get(i)
                    record[name] = val.strip() if isinstance(val, str) else val
                records.append(record)
            workbook[sheet_name] = records
        return workbook
