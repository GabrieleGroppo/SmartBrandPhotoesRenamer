"""Helper condivisi per individuare cartelle/file di una coppia (stagione, brand)."""
import os
import re
import sys

PHOTO_EXTENSIONS = (".jpg", ".jpeg")


def brand_dir_path(season, brand):
    return os.path.join("assets", season, brand)


def require_brand_dir(season, brand):
    path = brand_dir_path(season, brand)
    if not os.path.isdir(path):
        sys.exit(
            f"Cartella non trovata: {path}\n"
            f"Struttura attesa: assets/<stagione>/<brand>/ con dentro le foto .jpg e il file .xlsx"
        )
    return path


def find_excel(brand_dir):
    # Estensione case-insensitive: capita che arrivi ".XLSX" maiuscolo (guess ai26).
    candidates = sorted(
        os.path.join(brand_dir, f)
        for f in os.listdir(brand_dir)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    )
    if not candidates:
        sys.exit(f"Nessun file .xlsx trovato in {brand_dir}")
    if len(candidates) > 1:
        print(f"ATTENZIONE: più file .xlsx trovati, uso il primo: {candidates}", file=sys.stderr)
    return candidates[0]


def list_photos(brand_dir):
    return sorted(f for f in os.listdir(brand_dir) if f.lower().endswith(PHOTO_EXTENSIONS))


def stem_tokens(filename):
    """Spezza il nome file (senza estensione) nei token separati da _ - o spazi.

    Il codice prodotto puo' occupare un solo token (es. "91010339X95219_SN...")
    oppure piu' token consecutivi (es. "E1M50120101_352_0" = modello + colore,
    poi indice foto): per questo il matching esatto prova le concatenazioni dei
    primi k token invece di assumere che il codice sia sempre il primo token.
    """
    stem = os.path.splitext(filename)[0]
    return [t for t in re.split(r"[_\-\s]+", stem) if t]


def code_token(filename):
    """Primo token del nome file: solo per la vista d'insieme di inspect_brand.py."""
    tokens = stem_tokens(filename)
    return tokens[0] if tokens else ""


def natural_sort_key(filename):
    """Chiave di ordinamento che tratta le sequenze di cifre come numeri (2 < 10),
    cosi' "foto_2.jpg" precede "foto_10.jpg" invece di seguirla come farebbe
    l'ordinamento lessicografico puro. Usata per decidere qual e' la prima/ultima
    foto di una serie prodotto."""
    stem = os.path.splitext(filename)[0].lower()
    return [int(part) if part.isdigit() else part for part in re.split(r"(\d+)", stem)]
