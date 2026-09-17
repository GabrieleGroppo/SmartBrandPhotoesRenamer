#!/usr/bin/env python3
"""Rinomina le foto prodotto con il relativo codice EAN, deducendo la
corrispondenza foto<->riga Excel dal codice prodotto contenuto nel nome file.

Uso:
  rename_photos.py <stagione> <brand> [--dry-run] [--skip-compress]
                    [--ean-column NOME] [--code-columns "A,B,C"]

Se --ean-column/--code-columns non sono passati, viene usata la
configurazione salvata in brand_configs/<brand>.json (vedi save_brand_config.py).

Logica di matching per ogni riga (deduplicata per EAN):
  1. match ESATTO: il nome file viene spezzato in token separati da "_"/"-"/
     spazi; si prova a concatenare i primi k token (k=1,2,3,...) finche' il
     risultato coincide (case-insensitive) con la concatenazione di TUTTE le
     colonne codice non vuote per quella riga. Serve perche' il codice puo'
     stare in un solo token (es. "91010339X95219_SN...") oppure su piu' token
     consecutivi (es. "E1M50120101_352_0" = modello + colore, poi indice foto).
  2. se nessun match esatto, fallback PARZIALE: tutte le colonne codice non
     vuote compaiono come sottostringa nel nome file (matching per substring,
     come nello script originale).
  3. RED FLAG: se il match e' parziale (fallback substring, oppure riga con
     colonne codice mancanti) e i file candidati sono più di 7, non si
     rinomina nulla per quell'EAN: probabilmente le colonne scelte sono
     troppo generiche. Viene riportato in un report dedicato per revisione
     manuale, invece di rinominare a caso.

Le foto rinominate con successo vengono spostate nella sottocartella
"rinominate/" (dentro la cartella del brand), cosi' restano separate dagli
originali non ancora processati e da eventuali foto senza corrispondenza.
Le foto senza nessuna riga EAN corrispondente vengono spostate in
"non_rinominate/" (mai cancellate).

Foto "campione colore" (senza il prodotto): alcuni brand includono, per ogni
prodotto, una foto che mostra solo il colore/materiale e non la borsa/gadget.
Vengono escluse dalla rinomina e spostate in "scartate/" (mai cancellate) in
due modi:
  a) parola chiave nel nome file (swatch/colore/color/colour) - deterministico;
  b) se il brand ha "swatch_position": "first"/"last" in config (dedotto una
     tantum da Claude guardando visivamente un campione di prodotto), si
     scarta sempre la prima/ultima foto di ogni serie - ma solo se restano
     almeno 2 foto dopo aver tolto quelle già escluse per parola chiave (mai
     azzerare le foto di un prodotto).

Le foto rinominate e spostate in "rinominate/" sopra 1MB vengono infine
compresse con ImageMagick (comando "magick"): prima si tenta di ridurre solo
la qualita' JPEG (jpeg:extent) restando il piu' vicino possibile a ~1MB; se
questo farebbe scendere la qualita' sotto una soglia minima (si vedrebbero
gli artefatti di compressione), si riduce anche la risoluzione finche' non si
trova una combinazione qualita'/dimensione accettabile.
"""
import argparse
import csv
import os
import re
import subprocess
import sys
import tempfile
import time
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from brand_config import load_config
from discovery import find_excel, list_photos, natural_sort_key, require_brand_dir, stem_tokens
from xlsx_reader import read_workbook

RED_FLAG_THRESHOLD = 7
MAX_SIZE_MB = 1
RENAMED_SUBDIR = "rinominate"
DISCARDED_SUBDIR = "scartate"
NON_RENAMED_SUBDIR = "non_rinominate"
COMPRESS_TARGET_KB = 1000
COMPRESS_MIN_QUALITY = 70
COMPRESS_RESIZE_DIMS = (2600, 2400, 2200, 2000, 1800, 1600, 1400, 1200, 1000)
SWATCH_KEYWORDS = ("swatch", "colore", "color", "colour")


def is_keyword_swatch(filename):
    """True se un token del nome file indica esplicitamente un campione colore."""
    tokens = [t.lower() for t in stem_tokens(filename)]
    return any(keyword in token for token in tokens for keyword in SWATCH_KEYWORDS)


def split_swatches(candidates, swatch_position):
    """Ritorna (foto_prodotto, foto_scartate_con_motivo).

    Rimuove prima le foto con parola chiave colore nel nome, poi - se
    swatch_position e' "first"/"last" e restano almeno 2 foto - anche quella
    in quella posizione della serie (ordinata con natural sort).
    """
    ordered = sorted(candidates, key=natural_sort_key)
    keep = []
    discarded = []
    for f in ordered:
        if is_keyword_swatch(f):
            discarded.append((f, "parola chiave colore nel nome file"))
        else:
            keep.append(f)

    if swatch_position in ("first", "last") and len(keep) > 1:
        picked = keep.pop(0) if swatch_position == "first" else keep.pop(-1)
        discarded.append((picked, f"posizione '{swatch_position}' configurata per il brand"))

    return keep, discarded


def _magick_quality(path):
    result = subprocess.run(["identify", "-format", "%Q", path], capture_output=True, text=True)
    try:
        return int(result.stdout.strip())
    except ValueError:
        return 0


def compress_image_magick(file_path):
    """Comprime un JPG sopra MAX_SIZE_MB con ImageMagick (comando "magick"),
    restando il piu' vicino possibile a ~COMPRESS_TARGET_KB senza scendere
    sotto COMPRESS_MIN_QUALITY: prima prova a ridurre solo la qualita' JPEG
    (jpeg:extent); se questo costringerebbe la qualita' troppo in basso (si
    vedrebbe), riduce anche la risoluzione finche' la qualita' non torna
    accettabile. Ritorna (bytes_prima, bytes_dopo, qualita_finale) oppure None
    se il file non e' stato toccato (gia' piccolo, o magick non disponibile)."""
    if os.path.getsize(file_path) <= MAX_SIZE_MB * 1024 * 1024:
        return None

    extent = f"{COMPRESS_TARGET_KB}KB"
    # Il file temporaneo deve stare sullo stesso filesystem della destinazione:
    # os.replace() e' un rename atomico e fallisce con "Invalid cross-device
    # link" se /tmp e la cartella assets/ sono su mount diversi (es. tmpfs).
    fd, tmp_path = tempfile.mkstemp(suffix=os.path.splitext(file_path)[1], dir=os.path.dirname(file_path))
    os.close(fd)
    try:
        subprocess.run(
            ["magick", file_path, "-define", f"jpeg:extent={extent}", tmp_path],
            capture_output=True, text=True,
        )
        quality = _magick_quality(tmp_path)

        if quality < COMPRESS_MIN_QUALITY:
            for dim in COMPRESS_RESIZE_DIMS:
                subprocess.run(
                    ["magick", file_path, "-resize", f"{dim}x{dim}>", "-define", f"jpeg:extent={extent}", tmp_path],
                    capture_output=True, text=True,
                )
                quality = _magick_quality(tmp_path)
                if quality >= COMPRESS_MIN_QUALITY:
                    break

        before = os.path.getsize(file_path)
        after = os.path.getsize(tmp_path) if os.path.exists(tmp_path) else 0
        if 0 < after < before:
            os.replace(tmp_path, file_path)
            return before, after, quality
        return None
    except FileNotFoundError:
        return None  # magick non installato: compressione facoltativa, si salta
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def compress_images_in_folder(folder_path):
    compressed = 0
    total_before = 0
    total_after = 0
    for filename in list_photos(folder_path):
        result = compress_image_magick(os.path.join(folder_path, filename))
        if result:
            before, after, _ = result
            compressed += 1
            total_before += before
            total_after += after
    if compressed:
        saved_mb = (total_before - total_after) / (1024 * 1024)
        print(f"Compresse {compressed} immagini (>{MAX_SIZE_MB}MB) con ImageMagick (magick), risparmiati {saved_mb:.1f}MB.")
    return compressed


def resolve_config(brand, cli_ean_column, cli_code_columns):
    saved = load_config(brand) or {}
    ean_column = cli_ean_column or saved.get("ean_column")
    code_columns = cli_code_columns or saved.get("code_columns")
    swatch_position = saved.get("swatch_position")

    if not ean_column or not code_columns:
        sys.exit(
            f"Configurazione mancante per il brand '{brand}'.\n"
            f"Esegui prima inspect_brand.py per individuare le colonne, poi "
            f"save_brand_config.py per salvarle (oppure passa --ean-column e --code-columns)."
        )
    return ean_column, code_columns, swatch_position


CODE_COLUMN_SLICE_RE = re.compile(r"^(.+)\[(-?\d*):(-?\d*)\]$")


def parse_code_column(col):
    """Supporta la sintassi 'NomeColonna[start:end]' per usare solo un
    sottoinsieme dei token separati da spazio del valore di una colonna.

    Serve per i campi "composti" (piu' identificatori in un'unica cella, es.
    'CODICE SOTTOCODICE STAGIONE') in cui uno dei token (tipicamente la
    stagione, costante per tutto il file) non compare nei nomi foto e quindi
    andrebbe escluso dal confronto. Torna (nome_colonna_reale, slice|None).
    """
    m = CODE_COLUMN_SLICE_RE.match(col)
    if not m:
        return col, None
    name, start, end = m.groups()
    return name, slice(int(start) if start else None, int(end) if end else None)


def validate_columns(workbook, ean_column, code_columns):
    all_columns = set()
    for records in workbook.values():
        if records:
            all_columns.update(records[0].keys())

    code_column_names = [parse_code_column(c)[0] for c in code_columns]
    missing = [c for c in [ean_column, *code_column_names] if c not in all_columns]
    if missing:
        sys.exit(
            f"Colonne non trovate nel file Excel: {missing}\n"
            f"Colonne disponibili: {sorted(all_columns)}"
        )


def iter_unique_ean_rows(workbook, ean_column):
    """Itera tutte le righe di tutti i fogli, deduplicando per EAN (prima occorrenza vince)."""
    seen = set()
    for sheet_name, records in workbook.items():
        for row in records:
            ean = row.get(ean_column)
            if ean in (None, ""):
                continue
            ean = str(ean).strip()
            if ean in seen:
                continue
            seen.add(ean)
            yield sheet_name, ean, row


def build_code(row, code_columns):
    """Ritorna (codice_concatenato, segments[(colonna, valore_pulito)], is_partial)."""
    segments = []
    for col in code_columns:
        name, word_slice = parse_code_column(col)
        val = row.get(name)
        if val not in (None, ""):
            if word_slice is not None:
                val = " ".join(str(val).split()[word_slice])
            # rimuove anche gli spazi interni (es. "E1 M50 12 01 01" -> "E1M50120101"):
            # nei file Excel il codice e' spesso "impaginato" con spazi che non compaiono nel nome file
            cleaned = "".join(str(val).split())
            if cleaned:
                segments.append((col, cleaned))
    full_code = "".join(v for _, v in segments)
    is_partial = len(segments) < len(code_columns)
    return full_code, segments, is_partial


def matches_exact(filename, full_code_lower):
    """Prova a concatenare i primi k token del nome file finche' non combacia col codice."""
    tokens = stem_tokens(filename)
    joined = ""
    for token in tokens:
        joined += token.lower()
        if joined == full_code_lower:
            return True
        if len(joined) > len(full_code_lower):
            return False
    return False


def find_candidates(full_code, segments, remaining_files):
    """Ritorna (candidati, tipo_match)."""
    if not full_code:
        return [], None

    full_code_lower = full_code.lower()
    exact = [f for f in remaining_files if matches_exact(f, full_code_lower)]
    if exact:
        return sorted(exact), "esatto"

    if segments:
        values = [v.lower() for _, v in segments]
        partial = [f for f in remaining_files if all(v in f.lower() for v in values)]
        if partial:
            return sorted(partial), "parziale(substring)"

    return [], None


def write_csv_report(path, fieldnames, rows):
    if not rows:
        return None
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    return path


def run(season, brand, ean_column, code_columns, swatch_position, dry_run, skip_compress):
    brand_dir = require_brand_dir(season, brand)
    excel_path = find_excel(brand_dir)
    workbook = read_workbook(excel_path)
    validate_columns(workbook, ean_column, code_columns)

    renamed_dir = os.path.join(brand_dir, RENAMED_SUBDIR)
    discarded_dir = os.path.join(brand_dir, DISCARDED_SUBDIR)
    non_renamed_dir = os.path.join(brand_dir, NON_RENAMED_SUBDIR)

    print(f"Cartella: {brand_dir}")
    print(f"Excel: {os.path.basename(excel_path)} (fogli: {list(workbook.keys())})")
    print(f"Colonna EAN: {ean_column!r}  |  Colonne codice: {code_columns}")
    print(f"Posizione campione colore (swatch_position): {swatch_position!r}")
    print(f"Le foto rinominate verranno spostate in: {renamed_dir}")
    print(f"Le foto scartate (campione colore) verranno spostate in: {discarded_dir}")
    print(f"Le foto senza EAN corrispondente verranno spostate in: {non_renamed_dir}")
    if dry_run:
        print(">>> MODALITA' DRY-RUN: nessun file verra' rinominato, spostato o scartato <<<")
    else:
        os.makedirs(renamed_dir, exist_ok=True)
        os.makedirs(discarded_dir, exist_ok=True)
        os.makedirs(non_renamed_dir, exist_ok=True)

    photos = list_photos(brand_dir)
    print(f"Foto trovate: {len(photos)}")
    remaining_files = set(photos)

    renamed_rows = []
    unmatched_rows = []
    red_flag_rows = []
    discarded_rows = []
    renamed_count = 0

    for sheet_name, ean, row in iter_unique_ean_rows(workbook, ean_column):
        full_code, segments, is_partial = build_code(row, code_columns)
        used_columns = [c for c, _ in segments]
        candidates, match_type = find_candidates(full_code, segments, remaining_files)

        if not candidates:
            unmatched_rows.append({
                "foglio": sheet_name,
                "ean": ean,
                "codice_atteso": full_code,
                "colonne_usate": ",".join(used_columns),
                "motivo": "nessuna foto corrispondente",
            })
            continue

        is_generic_match = is_partial or match_type != "esatto"
        if is_generic_match and len(candidates) > RED_FLAG_THRESHOLD:
            red_flag_rows.append({
                "foglio": sheet_name,
                "ean": ean,
                "codice_atteso": full_code,
                "colonne_usate": ",".join(used_columns),
                "tipo_match": match_type,
                "n_candidati": len(candidates),
                "esempio_candidati": ";".join(candidates[:10]),
            })
            continue

        product_photos, swatch_photos = split_swatches(candidates, swatch_position)

        for old_name, motivo in swatch_photos:
            discarded_rows.append({
                "vecchio_nome": old_name,
                "ean": ean,
                "foglio": sheet_name,
                "motivo": motivo,
            })
            if not dry_run:
                os.rename(os.path.join(brand_dir, old_name), os.path.join(discarded_dir, old_name))
            remaining_files.discard(old_name)

        for i, old_name in enumerate(product_photos, start=1):
            ext = os.path.splitext(old_name)[1]
            new_name = f"{ean}-{i}{ext}"
            renamed_rows.append({
                "vecchio_nome": old_name,
                "nuovo_nome": new_name,
                "ean": ean,
                "foglio": sheet_name,
                "tipo_match": match_type,
                "colonne_usate": ",".join(used_columns),
            })
            if not dry_run:
                os.rename(os.path.join(brand_dir, old_name), os.path.join(renamed_dir, new_name))
            remaining_files.discard(old_name)
            renamed_count += 1

    reports_dir = os.path.join(brand_dir, "reports")
    os.makedirs(reports_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    prefix = "dryrun_" if dry_run else ""

    renamed_report = write_csv_report(
        os.path.join(reports_dir, f"{prefix}rinominati_{brand}_{ts}.csv"),
        ["vecchio_nome", "nuovo_nome", "ean", "foglio", "tipo_match", "colonne_usate"],
        renamed_rows,
    )
    unmatched_report = write_csv_report(
        os.path.join(reports_dir, f"{prefix}non_trovati_{brand}_{ts}.csv"),
        ["foglio", "ean", "codice_atteso", "colonne_usate", "motivo"],
        unmatched_rows,
    )
    red_flag_report = write_csv_report(
        os.path.join(reports_dir, f"{prefix}RED_FLAG_da_rivedere_{brand}_{ts}.csv"),
        ["foglio", "ean", "codice_atteso", "colonne_usate", "tipo_match", "n_candidati", "esempio_candidati"],
        red_flag_rows,
    )
    orphan_photo_rows = [{"nome_file": f, "token_iniziali": " ".join(stem_tokens(f))} for f in sorted(remaining_files)]
    orphan_photo_report = write_csv_report(
        os.path.join(reports_dir, f"{prefix}foto_senza_ean_{brand}_{ts}.csv"),
        ["nome_file", "token_iniziali"],
        orphan_photo_rows,
    )
    discarded_report = write_csv_report(
        os.path.join(reports_dir, f"{prefix}scartate_{brand}_{ts}.csv"),
        ["vecchio_nome", "ean", "foglio", "motivo"],
        discarded_rows,
    )

    if not dry_run:
        for old_name in sorted(remaining_files):
            os.rename(os.path.join(brand_dir, old_name), os.path.join(non_renamed_dir, old_name))

        if not skip_compress:
            compress_images_in_folder(renamed_dir)

    print()
    print("=== Riepilogo ===")
    verb = "rinominabili (dry-run)" if dry_run else f"rinominate e spostate in {renamed_dir}"
    print(f"Foto {verb}: {renamed_count}")
    disc_verb = "da scartare (dry-run)" if dry_run else f"scartate (campione colore) e spostate in {discarded_dir}"
    print(f"Foto {disc_verb}: {len(discarded_rows)}" + (f" -> {discarded_report}" if discarded_report else ""))
    orphan_verb = "senza nessuna riga EAN corrispondente (dry-run)" if dry_run else f"senza nessuna riga EAN corrispondente e spostate in {non_renamed_dir}"
    print(f"Foto {orphan_verb}: {len(remaining_files)}" + (f" -> {orphan_photo_report}" if orphan_photo_report else ""))
    print(f"EAN senza foto corrispondente: {len(unmatched_rows)}" + (f" -> {unmatched_report}" if unmatched_report else ""))
    print(f"RED FLAG (match troppo generico, >{RED_FLAG_THRESHOLD} candidati): {len(red_flag_rows)}" + (f" -> {red_flag_report}" if red_flag_report else ""))
    if renamed_report:
        print(f"Report rinomine: {renamed_report}")

    if red_flag_rows:
        print(
            "\nATTENZIONE: alcuni EAN hanno prodotto troppi candidati con un match parziale.\n"
            "Probabilmente le code_columns configurate sono troppo generiche: valuta di "
            "aggiungerne una (es. una colonna colore/taglia in più) e ripeti il dry-run."
        )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("season")
    parser.add_argument("brand")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--skip-compress", action="store_true", help="non comprimere le foto rinominate sopra 1MB")
    parser.add_argument("--ean-column")
    parser.add_argument("--code-columns", help="lista separata da virgole, in ordine")
    args = parser.parse_args()

    cli_code_columns = None
    if args.code_columns:
        cli_code_columns = [c.strip() for c in args.code_columns.split(",") if c.strip()]

    ean_column, code_columns, swatch_position = resolve_config(args.brand, args.ean_column, cli_code_columns)

    start = time.time()
    run(args.season, args.brand, ean_column, code_columns, swatch_position, args.dry_run, args.skip_compress)
    print(f"\nCompletato in {time.time() - start:.2f}s")


if __name__ == "__main__":
    main()
