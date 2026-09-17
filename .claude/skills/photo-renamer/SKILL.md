---
name: photo-renamer
description: Rinomina in batch le foto prodotto con il relativo codice EAN, deducendo dal nome dei file quali colonne dell'Excel compongono il codice prodotto e quale colonna è l'EAN/barcode. Usa questa skill quando l'utente chiede di rinominare/abbinare foto con i codici EAN a partire da una cartella assets/<stagione>/<brand>/ (es. "rinomina le foto di borbonese per ai26", "/photo-renamer ai26 borbonese").
---

# Photo Renamer (MSAFR)

Rinomina le foto di un brand con il codice EAN del prodotto corrispondente,
leggendo i dati da un file Excel multi-foglio. Sostituisce il vecchio script
`multi_sheet_photo_renamer.py` con mappature di colonne hard-coded per brand:
qui le colonne vengono **dedotte** osservando i nomi dei file e i dati Excel,
e la deduzione viene salvata per essere riusata nelle stagioni successive.

## Struttura cartelle attesa

```
assets/<stagione>/<brand>/
├── *.xlsx              <- un solo file Excel, uno o più fogli
├── foto1.jpg            <- originali non ancora processati
├── foto2.jpg
├── rinominate/          <- creata dallo script: foto rinominate (e compresse) spostate qui
├── scartate/            <- creata dallo script: foto "campione colore" spostate qui
├── non_rinominate/      <- creata dallo script: foto senza EAN corrispondente spostate qui
└── reports/             <- creata dallo script: report CSV
```

Questa è la struttura **dopo** il Passo 0: il materiale grezzo consegnato dal
cliente ha spesso le foto dentro sottocartelle annidate (zip estratti) invece
di stare già a questo livello — `prepare_folder.py` la sistema in automatico.

`assets/` è nel `.gitignore`: è materiale del cliente, non va versionato.
Le foto rinominate con successo vengono **spostate** in `rinominate/`, non
solo rinominate sul posto: così restano separate dagli originali non ancora
matchati e da eventuali foto senza corrispondenza EAN. Le foto "campione
colore" (solo un colore/materiale, senza borsa/gadget visibile) vengono
spostate in `scartate/` — **mai cancellate**. Le foto per cui non è stata
trovata nessuna riga EAN corrispondente vengono spostate in
`non_rinominate/` — anche queste **mai cancellate**, restano lì per
revisione manuale (es. materiale di un'altra stagione finito nella cartella
sbagliata, o varianti fuori listino).

## Invocazione

L'utente specifica **stagione** e **brand** (es. `ai26 borbonese`), corrispondenti
a `assets/ai26/borbonese/`. Se manca uno dei due, chiedilo prima di procedere:
non indovinare quale sottocartella di `assets/` usare se ce n'è più di una.

## Passo 0 — Predisponi la cartella (sempre, prima di ogni altra cosa)

Il materiale consegnato dal cliente arriva spesso con le foto dentro
sottocartelle annidate (zip estratti con struttura tipo
`<hash>/<collezione>/<taglia>/foto.jpg`) invece che già allo stesso livello
del file Excel. Prima di passare al Passo 1, esegui **sempre**:

```
python3 .claude/skills/photo-renamer/scripts/prepare_folder.py <stagione> <brand>
```

Lo script è idempotente (se la cartella è già nella struttura attesa non fa
nulla): porta tutte le foto annidate al livello del brand (stesso livello del
.xlsx), scarta i doppioni identici (stesso contenuto, verificato per hash,
mai per nome soltanto), rimuove le sottocartelle rimaste vuote ed elimina gli
archivi `.zip` originali al livello del brand — ma solo se non sono rimasti
conflitti. Se lo script segnala **CONFLITTI** (stesso nome file ma contenuto
diverso), non decidere da solo quale tenere: mostra la lista all'utente e
chiedi come procedere prima di continuare con il Passo 1.

Controlla se esiste già `.claude/skills/photo-renamer/brand_configs/<brand>.json`
(nomi file in minuscolo). Se esiste, salta al Passo 3 usando quella
configurazione — ma tratta l'esito del dry-run come verifica: se il file Excel
di questa stagione ha una struttura diversa (colonne mancanti, troppi RED FLAG),
torna al Passo 2 per ri-dedurre la configurazione e sovrascrivila.

## Passo 2 — Deduzione colonna EAN e colonne codice (solo la prima volta per brand)

Esegui:
```
python3 .claude/skills/photo-renamer/scripts/inspect_brand.py <stagione> <brand>
```

Questo stampa, senza decidere nulla: intestazioni di ogni foglio, alcune righe
complete di esempio, statistiche per colonna (% di valori solo numerici,
lunghezze più comuni, esempi di valori), e un campione di nomi foto con il
"codice candidato" (la parte prima del primo `_`/`-`).

Analizza l'output tu stesso (ragionamento, non uno script) per dedurre:

1. **Colonna EAN**: di solito il nome contiene "EAN", "BARCODE", "GTIN" o
   "CODICE A BARRE"; i valori sono quasi sempre stringhe **solo numeriche di
   lunghezza 8, 12, 13 o 14** (EAN-8/UPC-12/EAN-13/ITF-14). Usa nome colonna +
   statistiche insieme, non uno dei due da solo.

   Se **nessuna colonna con nome sensato sembra un EAN** (capitato con Liujo:
   file chiamato letteralmente `BARCODE LIUJO AI26.xlsx` ma senza nessuna
   colonna "Barcode"/"EAN" visibile in `inspect_brand.py`), controlla se
   `inspect_brand.py` mostra colonne chiamate `__colN__`: sono colonne che
   hanno dati ma **nessuna intestazione** nella riga 1 dell'Excel (tipico di
   colonne aggiunte in fondo al foglio senza etichetta) — spesso è proprio lì
   che si nasconde l'EAN vero. Prima di concludere che l'EAN non esiste,
   guarda le statistiche di *tutte* le `__colN__`, non solo delle colonne con
   nome. Se anche lì non c'è nulla che sembri un EAN, non indovinare: chiedi
   conferma all'utente (magari il file giusto è un altro, o quel brand/ordine
   non ha barcode per alcuni prodotti — capita, e va bene lasciare quelle
   righe senza foto rinominata: `iter_unique_ean_rows` salta già da sola le
   righe con EAN vuoto).

2. **Colonne codice** (`code_columns`, in ordine): il sottoinsieme di colonne
   la cui concatenazione (valori ripuliti, maiuscolo/minuscolo indifferente)
   riproduce il "codice candidato" visto nei nomi foto. Sono in genere campi
   tipo Modello/Codice/Parte/Colore/Variante — **non** descrizioni testuali
   libere, indirizzi, clienti, pesi, prezzi o tariffe doganali. Verifica
   l'ipotesi concatenando a mano 3-5 righe campione e confrontando col codice
   dei file (puoi scrivere un piccolo snippet Python usando
   `xlsx_reader.read_workbook(...)` per farlo velocemente, invece di fidarti
   a occhio).

   Tieni conto che per alcuni prodotti non tutti i campi sono valorizzati:
   è normale e lo script gestisce già il matching parziale (colonne vuote
   vengono saltate nella concatenazione). Non serve che tu forzi un match
   perfetto al 100% delle righe.

   **Colonne "composte"**: a volte una singola colonna contiene più
   identificatori uniti da spazio in un'unica cella (es. Liujo: `Articolo` =
   `"2F6001 MA29S AI26"`, dove il terzo token è la stagione, ripetuta
   identica su ogni riga del file e **assente** dai nomi foto). Concatenare
   l'intero valore fallirebbe sia il match esatto sia quello per sottostringa
   (la stagione non compare mai nel nome file). In questo caso usa la
   sintassi `NomeColonna[start:end]` in `--code-columns` per restringere il
   valore ai soli token (separati da spazio, slice Python standard su
   `valore.split()`) che servono, es. `--code-columns "Articolo[:2],Cod Col"`
   per prendere solo i primi due token di `Articolo` e scartare il terzo.

3. Se non riesci a determinare un sottoinsieme di colonne che spieghi la
   maggioranza dei codici visti nei file, chiedi conferma all'utente prima
   di proseguire piuttosto che indovinare.

Salva la configurazione dedotta:
```
python3 .claude/skills/photo-renamer/scripts/save_brand_config.py <brand> \
  --ean-column "NOME_COLONNA_EAN" \
  --code-columns "Colonna1,Colonna2,Colonna3" \
  --notes "breve nota su come hai dedotto lo schema"
```

## Passo 2bis — Foto "campione colore" (solo la prima volta per brand)

Alcuni brand includono, per ogni prodotto, una foto che mostra solo il
colore/materiale (niente borsa/gadget) — di solito la prima o l'ultima della
serie. Va esclusa dalla rinomina, non trattata come una foto prodotto.

1. **Parola chiave nel nome file** (automatico, niente da fare): se i nomi
   contengono già un indizio testuale tipo `_SWATCH`, `_COLORE`, `_COLOR` (es.
   coccinelle: `E1M50120101_352_SWATCH.jpg`), lo script la riconosce da solo e
   la sposta in `scartate/`. Non serve configurare nulla per questo caso.

2. **Nessun indizio nel nome** (es. borbonese: solo `..._CLOSEUP001.jpg` fino
   a `..._CLOSEUP00N.jpg`, nessuna parola chiave): verifica **una sola volta**
   se il brand ha comunque questo problema, guardando le immagini vere e
   proprie:
   - scegli UN prodotto con più foto (es. dal report `rinominati_*.csv` di un
     dry-run, o contando i file candidati per un codice nell'output di
     `inspect_brand.py`);
   - apri con lo strumento di lettura immagini la prima e l'ultima foto di
     quella serie (ordinate su base numerica, non lessicografica: usa
     `natural_sort_key` se scrivi uno snippet, o semplicemente guarda l'indice
     più alto);
   - se una delle due mostra solo un colore/materiale piatto senza il
     prodotto, quella posizione ("first" o "last") vale per **tutti** i
     prodotti della collezione — non serve controllare ogni prodotto, il
     numero di foto per prodotto varia ma la posizione (prima/ultima) è
     costante. Salvala:
     ```
     python3 .claude/skills/photo-renamer/scripts/save_brand_config.py <brand> --swatch-position first
     # oppure: --swatch-position last
     ```
   - se invece sia la prima sia l'ultima mostrano il prodotto, registra
     esplicitamente che è stato verificato e non serve nulla:
     ```
     python3 .claude/skills/photo-renamer/scripts/save_brand_config.py <brand> --swatch-position none
     ```
   - se il campione scelto è ambiguo (es. mostra un dettaglio ravvicinato e
     non è chiaro se sia "prodotto" o "colore"), guardane un secondo prima di
     decidere da solo; se resta ambiguo, chiedi conferma all'utente mostrando
     le immagini invece di indovinare — è un'operazione che sposta file reali.

La regola di posizione scarta **al massimo una foto per prodotto**, e solo se
dopo aver tolto le foto già escluse per parola chiave ne restano almeno 2
(non viene mai lasciato un prodotto senza nessuna foto).

## Passo 3 — Dry-run e controllo dei RED FLAG

```
python3 .claude/skills/photo-renamer/scripts/rename_photos.py <stagione> <brand> --dry-run
```

Lo script:
- prova prima un match **esatto**: spezza il nome file in token (separatori
  `_`/`-`/spazio) e concatena i primi k token finché non combacia con la
  concatenazione di tutte le colonne codice non vuote (gli spazi interni ai
  valori Excel vengono rimossi prima del confronto). Serve perché il codice
  può stare in un solo token (`91010339X95219_SN...`) o su più token
  consecutivi (`E1M50120101_352_0` = modello + colore, poi indice foto);
- se non trova nulla, ripiega su un match **parziale per sottostringa**
  (come il vecchio script);
- **RED FLAG**: se un match parziale/generico produce più di 7 file
  candidati per lo stesso EAN, NON rinomina nulla per quell'EAN e lo segna
  in un report `RED_FLAG_da_rivedere_*.csv` — troppi candidati con colonne
  parziali è quasi sempre segno che le `code_columns` scelte sono troppo
  generiche (regex/colonne del codice non abbastanza specifiche).

Leggi il riepilogo stampato a schermo (foto rinominabili, foto scartate come
campione colore → `scartate_*.csv`, foto senza EAN corrispondente →
`foto_senza_ean_*.csv`, EAN senza foto → `non_trovati_*.csv`, RED FLAG →
`RED_FLAG_da_rivedere_*.csv`):
- Se **ci sono RED FLAG**: guarda `RED_FLAG_da_rivedere_*.csv`, capisci quale
  colonna aggiuntiva serve per disambiguare (es. una colonna Taglia/Variante
  mancante), aggiorna la config col Passo 2 e ripeti il dry-run finché i RED
  FLAG sono spariti o comunque ridotti a casi genuinamente ambigui da
  segnalare all'utente.
- Se il numero di foto "rinominabili" è molto basso rispetto al totale delle
  foto trovate, **non dare per scontato che la config sia sbagliata**: apri
  `foto_senza_ean_*.csv` e verifica a campione se i codici dedotti dai nomi
  file esistono davvero nell'Excel (potrebbero essere colori/varianti fuori
  listino, foto di una stagione diversa finite nella cartella sbagliata,
  ecc. — un mismatch reale tra materiale fotografico e Excel, non un bug di
  matching). Solo se i pochi match trovati sono anche palesemente sbagliati
  (EAN abbinato al file sbagliato), torna al Passo 2 e correggi la config.
  Segnala comunque all'utente una percentuale di match anomala prima di
  procedere al Passo 4, così può decidere se procedere o controllare prima
  il materiale con il brand.

  Caso frequente verificato più volte (mandarina e liujo ai26): l'Excel è un
  **ordine per pochi negozi/clienti specifici**, non il catalogo completo —
  copre solo una minima parte degli articoli/colori effettivamente
  fotografati. In quel caso un tasso di match anche sotto il 5% è normale e
  atteso (verificalo scegliendo 2-3 codici da `foto_senza_ean_*.csv` e
  controllando che quell'Articolo non compaia proprio in nessuna riga
  dell'Excel, non solo con quel colore): non è un errore di config, va solo
  segnalato all'utente com'è per farlo decidere.

## Passo 4 — Conferma ed esecuzione reale

Mostra all'utente il riepilogo del dry-run (quante foto verrebbero
rinominate, quanti EAN senza foto, eventuali RED FLAG residui) e chiedi
conferma esplicita prima di rinominare davvero i file — è un'operazione che
tocca file reali del cliente. Dopo conferma:

```
python3 .claude/skills/photo-renamer/scripts/rename_photos.py <stagione> <brand>
```

(stessa invocazione, senza `--dry-run`). Dopo la rinomina, lo script:
- sposta le foto rimaste senza EAN corrispondente in `non_rinominate/`;
- comprime con ImageMagick (`magick`) le foto in `rinominate/` sopra 1MB
  (vedi Note sotto).

Riporta all'utente il riepilogo finale e il percorso dei report in
`assets/<stagione>/<brand>/reports/` (mapping vecchio→nuovo nome, EAN senza
foto, eventuali RED FLAG non risolti).

## Note

- **Compressione**: le foto in `rinominate/` sopra 1MB vengono compresse con
  ImageMagick (comando `magick`, mai `convert` che è deprecato in IMv7):
  prima si tenta di ridurre solo la qualità JPEG (`-define jpeg:extent=1000KB`)
  restando il più vicino possibile a ~1MB; se la qualità risultante scende
  sotto 70/100 (si vedrebbe), si ridimensiona anche l'immagine finché non si
  trova una combinazione qualità/risoluzione accettabile — evita di
  massacrare la qualità di scan ad altissima risoluzione pur di rientrare
  nel limite di dimensione. È "best effort": se `magick` non è installato
  viene saltata senza errori. Con `--skip-compress` la si disattiva
  esplicitamente. Non fa backup degli originali prima di comprimere (la
  compressione è lossy e irreversibile): se serve conservarli, copiali altrove
  prima di eseguire lo script senza `--dry-run`. Il file temporaneo usato da
  ImageMagick viene creato nella stessa cartella del file finale (non in
  `/tmp`): se fosse su un filesystem diverso, lo swap finale (`os.replace`)
  fallirebbe con "Invalid cross-device link" — capitato la prima volta su
  `rinominate/` di mandarina.
- Nessuna dipendenza esterna (niente pandas/openpyxl): la lettura xlsx è
  fatta con `xlsx_reader.py`, solo standard library. Legge le colonne anche
  se non hanno intestazione nella riga 1 (le espone come `__colN__`,
  `N`=indice colonna 0-based): la larghezza del foglio viene calcolata su
  *tutte* le righe, non solo sulla riga di intestazione, altrimenti colonne
  dati aggiunte in fondo al foglio senza etichetta (visto con l'EAN di
  Liujo) verrebbero scartate in silenzio.
- Le righe Excel duplicate per lo stesso EAN (tipico quando lo stesso
  prodotto compare per più clienti/filiali) vengono deduplicate
  automaticamente: vince la prima occorrenza.
- `foto_senza_ean_*.csv` elenca le foto che non hanno trovato nessuna riga
  Excel corrispondente: non è necessariamente un errore di configurazione,
  spesso sono foto di varianti/colori non più a listino o materiale di
  un'altra stagione finito nella cartella sbagliata.
