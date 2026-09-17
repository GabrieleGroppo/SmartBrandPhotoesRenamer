#!/usr/bin/env python3
"""Predispone assets/<stagione>/<brand>/ prima di iniziare la rinomina.

Le foto consegnate dal cliente arrivano spesso dentro archivi zip estratti in
sottocartelle annidate (una gerarchia a piacere sotto il brand, es.
"<hash>/<collezione>/<taglia>/foto.jpg" oppure "<collezione>/foto.jpg") invece
di stare allo stesso livello del file .xlsx. Questo script porta tutte le
foto trovate in QUALSIASI sottocartella direttamente al livello di
assets/<stagione>/<brand>/ (lo stesso livello del file Excel), poi ripulisce:

  1. sposta ogni foto annidata nella cartella del brand;
  2. se una foto con lo stesso nome esiste gia' a quel livello, la confronta
     (hash MD5): se identica la considera un doppione e la scarta, se diversa
     la lascia dov'e' e segnala un CONFLITTO (mai sovrascrive silenziosamente);
  3. rimuove le sottocartelle rimaste vuote dopo lo spostamento;
  4. elimina gli archivi .zip originali che stanno al livello del brand, ma
     SOLO se non sono rimasti conflitti irrisolti (altrimenti l'archivio puo'
     ancora servire per recuperare l'originale in caso di dubbio).

E' idempotente: si puo' rilanciare sempre come primo passo, anche se la
cartella e' gia' nella struttura attesa (in quel caso non fa nulla).

Uso:
  prepare_folder.py <stagione> <brand> [--dry-run]
"""
import argparse
import hashlib
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from discovery import PHOTO_EXTENSIONS, brand_dir_path, require_brand_dir  # noqa: E402


def md5sum(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def find_nested_photos(brand_dir):
    nested = []
    for root, _dirs, files in os.walk(brand_dir):
        if os.path.abspath(root) == os.path.abspath(brand_dir):
            continue
        for name in files:
            if name.lower().endswith(PHOTO_EXTENSIONS):
                nested.append(os.path.join(root, name))
    return sorted(nested)


def flatten_photos(brand_dir, dry_run):
    moved, duplicates_removed, conflicts = [], [], []
    for filepath in find_nested_photos(brand_dir):
        dest = os.path.join(brand_dir, os.path.basename(filepath))
        if not os.path.exists(dest):
            moved.append((filepath, dest))
            if not dry_run:
                os.rename(filepath, dest)
        elif md5sum(filepath) == md5sum(dest):
            duplicates_removed.append(filepath)
            if not dry_run:
                os.remove(filepath)
        else:
            conflicts.append(filepath)
    return moved, duplicates_removed, conflicts


def remove_empty_dirs(brand_dir, dry_run):
    # os.walk(topdown=False) precalcola dirs/files per ogni cartella prima di
    # iniziare la discesa: se una sottocartella viene rimossa durante lo stesso
    # walk, la cartella padre continua comunque a "vederla" nella sua lista
    # dirs (stale) e non risulterebbe mai vuota. Si ricontrolla quindi lo stato
    # reale su disco con os.listdir invece di fidarsi delle liste di os.walk.
    removed = []
    for root, _dirs, _files in os.walk(brand_dir, topdown=False):
        if os.path.abspath(root) == os.path.abspath(brand_dir):
            continue
        if not os.listdir(root):
            removed.append(root)
            if not dry_run:
                os.rmdir(root)
    return removed


def remove_top_level_zips(brand_dir, dry_run, skip):
    if skip:
        return []
    removed = []
    for name in os.listdir(brand_dir):
        if name.lower().endswith(".zip"):
            path = os.path.join(brand_dir, name)
            removed.append(path)
            if not dry_run:
                os.remove(path)
    return removed


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("season")
    parser.add_argument("brand")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    brand_dir = require_brand_dir(args.season, args.brand)
    print(f"Cartella: {brand_dir}")

    moved, duplicates_removed, conflicts = flatten_photos(brand_dir, args.dry_run)
    removed_dirs = remove_empty_dirs(brand_dir, args.dry_run)
    removed_zips = remove_top_level_zips(brand_dir, args.dry_run, skip=bool(conflicts))

    verb_moved = "da spostare" if args.dry_run else "spostate"
    verb_dup = "da scartare (doppioni identici)" if args.dry_run else "scartate (doppioni identici)"
    verb_dir = "da rimuovere" if args.dry_run else "rimosse"
    verb_zip = "da eliminare" if args.dry_run else "eliminati"

    print("=== Riepilogo preparazione cartella ===")
    print(f"Foto annidate {verb_moved} al livello del brand: {len(moved)}")
    print(f"Foto {verb_dup}: {len(duplicates_removed)}")
    print(f"Sottocartelle vuote {verb_dir}: {len(removed_dirs)}")
    print(f"Archivi .zip {verb_zip}: {len(removed_zips)}")
    if conflicts:
        print(f"\nATTENZIONE - CONFLITTI ({len(conflicts)}): stesso nome file ma contenuto diverso "
              f"da quello gia' presente al livello del brand. Lasciati dove sono, NON spostati/cancellati "
              f"automaticamente. Vanno risolti a mano prima di procedere:")
        for path in conflicts:
            print(f"  - {path}")
    else:
        print("\nNessun conflitto: cartella pronta.")

    if args.dry_run and (moved or duplicates_removed or removed_dirs or removed_zips):
        print("\n(--dry-run: nessuna modifica applicata davvero)")


if __name__ == "__main__":
    main()
