# AT614 Configuration Editor

Bootstrap project for the AT614 Configuration Editor.

Current scope:

- project scaffold;
- Python package entry point;
- first fixture files based on real project data;
- core domain parsers and reference index;
- first Qt shell UI with reusable structural components;
- BAS schema extractor producing `schemas/auto/*.yaml`;
- smoke and domain tests.

The implementation roadmap is described in:

- `PIANO_SVILUPPO.md`
- `BACKLOG_IMPLEMENTAZIONE.md`
- `LAYOUT_UX.md`

## Running the BAS schema extractor

The extractor reads `Module_TEST_*.bas` files and emits one YAML per test
handler with parameter index, name, raw enum identifier and inline comment.

```powershell
$env:PYTHONPATH = "src"
python -m at614_editor.domain.bas_extractor .\BAS .\schemas\auto
```

Modules without a `Private Enum eTestParameter` block (e.g. the dispatcher
`Module_TEST_MANAGER.bas`) are skipped. Files in `schemas/auto/` are
regenerated each run; manual refinements belong in `schemas/overrides/`.

## File format inventory

| Path pattern | Parser | Editor (shell) | CRUD scope |
|---|---|---|---|
| `DISTRIBUTORE/*.cfg` | `domain.parsers.distributore` | `DistEditor` | crea, apri, duplica, elimina |
| `TEST/*.csv` | `domain.parsers.test_csv` | `SeqEditor` | crea, apri, duplica, elimina |
| `CURVE_COMANDO/*.csv` | `domain.parsers.point_series` | `CurveEditor` | crea, apri, duplica, elimina |
| `CURVE_LIMITE/*.csv` | `domain.parsers.point_series` | `LimitEditor` | crea, apri, duplica, elimina |
| `RAMPE_XY/*.csv` | `domain.parsers.point_series` | `RampEditor` | crea, apri, duplica, elimina |
| `SETTAGGI PROGRAMMA/<profilo>/CE16/*.cfg` | `domain.parsers.ce16` | `Ce16Editor` | crea, apri, duplica (nuovo slot `CE16_<next>`) |
| `SETTAGGI PROGRAMMA/<profilo>/MMS2218/*.csv` | `domain.parsers.mms2218` | `MmsEditor` | crea, apri, duplica (nuovo canale `ADC<next>`) |
| `*.txt` esterni referenziati dai test | lettura testo, encoding cp1252/utf-8 | `ExternalFileEditor` | crea, apri, salva (modalità testo assistito), duplica |
| `GRAPH/**`, `RAMPE_XY_LAST/**`, archivi esterni | `domain.output_archive` (metadati lazy + cache) | `OutputViewer` (sola lettura, filtro stringa/data, scelta assi X/Y) | nessun authoring; azioni `Esporta`, `Confronta`, `Stampa`, `Apri cartella` |
| `BAS/Module_TEST_*.bas` | `domain.bas_extractor` | — | sola sorgente: produce `schemas/auto/*.yaml` |

## Duplica programma (M5.7)

Dalla dashboard il pulsante `Duplica programma` apre un wizard che:

- raccoglie il distributore sorgente e tutte le risorse referenziate da
  sezioni e righe TEST (curve, limiti, rampe, file esterni);
- per ogni risorsa permette di scegliere fra **Duplica** (crea una copia con
  nome derivato) o **Riusa** (mantiene il riferimento all'originale);
- mostra un riepilogo `[nuovo]` / `[duplicato da …]` / `[riusato]` prima della
  conferma;
- al click su `Duplica`, esegue il piano: crea i file fisici dei duplicati,
  riscrive le sequenze TEST aggiornando i riferimenti interni, salva il nuovo
  distributore con le sezioni puntate ai nuovi TEST e atterra nell'editor del
  nuovo programma.

Convenzione di naming: lo stem del file sorgente viene affiancato dal nuovo
codice distributore (`<stem>_<new_code><suffix>`); se lo stem coincide con il
codice sorgente, il file destinazione usa solo il nuovo codice
(`15.1001.356_C.csv` → `15.1001.999.csv`).

## Viewer output banco (M6)

L'OutputViewer è la vista per consultare archivi di output del banco. Si apre
dalla categoria *Output banco* nell'albero risorse. Caratteristiche:

- **Archivio configurabile**: di default usa la cartella `GRAPH/` del progetto
  se esiste; il pulsante `Configura archivio output` permette di scegliere
  qualunque cartella esterna.
- **Indice lazy con cache persistente**: la scansione iniziale legge solo i
  metadati dei `.csv` (nome, cartella, data creazione, data modifica,
  dimensione), mai il contenuto; l'indice è cachato in `.at614_output_index.json`
  nella root dell'archivio e ri-scansionato solo se rileva file nuovi o
  modificati.
- **Filtri**: stringa libera (nome o cartella) e range di date di creazione.
- **Apertura on-demand**: il CSV viene letto solo quando l'utente seleziona
  una riga; encoding auto-detect tra utf-8 e cp1252.
- **Scelta assi**: dropdown asse X + lista multi-select assi Y popolati dalle
  colonne del CSV selezionato.
- **Sola lettura**: `ReadOnlyBanner` permanente in cima; le uniche azioni
  permesse sono `Esporta`, `Confronta`, `Stampa`, `Apri cartella`. Nessun
  pulsante `Salva`/`Nuovo`/`Modifica`.
