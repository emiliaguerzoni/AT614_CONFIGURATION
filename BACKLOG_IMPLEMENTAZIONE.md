# Backlog implementazione — AT614 Configuration Editor

Documento operativo che traduce [PIANO_SVILUPPO.md](PIANO_SVILUPPO.md) e
[LAYOUT_UX.md](LAYOUT_UX.md) in attivita implementabili.

Questo file non ridefinisce il prodotto. Serve a stabilire ordine di esecuzione,
dipendenze, priorita e Definition of Done per le prossime fasi di sviluppo.

---

## 1. Regole di esecuzione

- il brief UX autoritativo e [LAYOUT_UX.md](LAYOUT_UX.md);
- nessun editor UI entra in sviluppo senza parser e fixture del formato che deve
  modificare;
- prima si consolidano le risorse sorgente e le loro relazioni, poi si affronta
  la UX avanzata dei test dinamici;
- il viewer output banco resta separato dal dominio di authoring;
- ogni milestone deve lasciare il repository in uno stato verificabile.

---

## 2. Ordine delle milestone

| Milestone | Obiettivo | Priorita |
|---|---|---|
| M0 | baseline progetto, fixture, inventario formati | P0 |
| M1 | parser, modelli e indice riferimenti | P0 |
| M2 | shell UI, resource tree e componenti base | P0 |
| M3 | editor risorse base non-TEST | P0 |
| M4 | editor sequenza TEST e schema dinamico | P1 |
| M5 | validazione, rename, duplica, delete sicuri | P1 |
| M6 | viewer output banco scalabile | P2 |
| M7 | packaging, logging, documentazione | P2 |

---

## 3. Backlog per milestone

### M0 — Baseline progetto e inventario formati

Deliverable:

- struttura repository iniziale;
- fixture campione per ogni tipologia sorgente;
- matrice tipo file -> parser -> editor -> CRUD.

Task:

1. creare `pyproject.toml` con dipendenze minime e script di avvio;
2. creare struttura `src/at614_editor/` e `tests/`;
3. copiare fixture reali per distributore, test, curva comando, curva limite,
   rampa XY, CE16, MMS2218, file esterni;
4. scrivere `docs` o tabella interna che classifica i formati del progetto;
5. implementare `bas_extractor.py` con output YAML minimo per i moduli TEST.

DoD:

- il repository avvia un entry point vuoto;
- le fixture coprono tutte le tipologie sorgente principali;
- gli YAML base dei `Module_TEST_*.bas` vengono generati.

### M1 — Core domain, parser e riferimenti

Deliverable:

- parser/serializer round-trip per tutte le risorse sorgente;
- indice riferimenti centralizzato;
- primi test automatici di round-trip.

Task:

1. definire `ProjectResource`, `DistributoreConfig`, `TestSequence`,
   `PointSeriesResource`, `ExternalConfigResource`;
2. implementare parser `distributore.py`;
3. implementare parser `test_csv.py`;
4. implementare parser `curva_comando.py`, `curva_limite.py`, `rampa_xy.py`;
5. implementare parser `ce16.py` e `mms2218.py`;
6. implementare parser `external_config.py` con fallback best-effort;
7. implementare `project.py` con scansione risorse e indice riferimenti;
8. scrivere test di round-trip per ogni tipologia;
9. scrivere test dell'indice riferimenti.

DoD:

- i file sorgente supportati fanno round-trip senza perdita di compatibilita;
- l'indice risponde a "chi uso" e "chi mi usa";
- i test base del dominio sono verdi.

### M2 — Shell UI e componenti strutturali

Deliverable:

- finestra principale navigabile;
- layout 3 colonne operativo;
- componenti base del brief UX riusabili.

Task:

1. creare `main_window.py` e `workspace_home.py`;
2. implementare layout editing fisso `240 | fluid | 260`;
3. implementare `resource_tree.py` con 9 categorie fisse nell'ordine del brief;
4. implementare `RPanel` con sezioni fisse `Riferimenti`, `Validazione`,
   `Azioni rapide`;
5. implementare `FileWidget`;
6. implementare `ResourceCard`;
7. implementare `ChartStage` e `ReadOnlyBanner`;
8. applicare microcopy e stati definiti in [LAYOUT_UX.md](LAYOUT_UX.md).

DoD:

- dashboard e shell UI sono navigabili;
- ogni schermata editor usa il layout corretto;
- i componenti base sono pronti per essere riusati dagli editor.

### M3 — Editor risorse base non-TEST

Deliverable:

- editor dedicati per tutte le risorse sorgente principali non-TEST.

Task:

1. implementare `DistEditor` con 5 sezioni e 5 slot CE16 sempre visibili;
2. implementare `CurveEditor`;
3. implementare `LimitEditor`;
4. implementare `RampEditor`;
5. implementare `Ce16Editor`;
6. implementare `MmsEditor`;
8. implementare `ExternalFileEditor` con doppia modalita:
  strutturata se il formato e riconosciuto, testo assistito con warning se non
  lo e;
8. collegare gli editor al `resource_tree`;
9. aggiungere azioni `Nuovo`, `Duplica`, `Apri`, `Elimina`, `Usi`.

DoD:

- tutte le risorse non-TEST principali si possono aprire, salvare e duplicare;
- le azioni distruttive aprono una modale con impatto riferimenti;
- l'utente non deve usare il filesystem per il CRUD base.

### M4 — Editor sequenza TEST e schema dinamico

Deliverable:

- `SeqEditor` con tabella test e dettaglio dinamico per ID test.

Task:

1. implementare tabella righe con drag and drop;
2. implementare reindicizzazione di `Parameter(0)`;
3. implementare loader `schemas.py` con merge auto + override;
4. creare renderer dinamico campi per ID test basato sul catalogo schema;
5. integrare `FileWidget` nei parametri file reference;
6. gestire righe aggiungi/duplica/elimina;
7. aggiungere stato riga e badge validazione;
8. predisporre fallback editor generico per ID ancora incompleti o schema
  parziale.

DoD:

- una sequenza TEST reale si apre, si modifica e si salva;
- le righe con file referenziati aprono o creano risorse figlie;
- gli ID non rifiniti restano modificabili con fallback sicuro.

### M5 — Validazione e operazioni sicure

Deliverable:

- validazione live;
- rename, duplica e delete coerenti con i riferimenti.

Task:

1. implementare validazione live warning;
2. implementare validazione on-save errori bloccanti;
3. implementare pannello riferimenti mancanti;
4. implementare comportamento `Usi` secondo la decisione chiusa:
  popover inline se `n <= 10`, tab `Riferimenti` se `n > 10`;
5. implementare dialog di rename con impatto a cascata e aggiornamento atomico
  di tutti i riferimenti del progetto;
6. implementare dialog di delete con `ImpactList`;
7. ✅ clone programma completo con scelta `riusa/duplica` — `domain/program_clone.py`
   + `ui/dialogs/clone_program_dialog.py`, integrato nella dashboard.

DoD:

- il sistema segnala riferimenti mancanti prima del salvataggio;
- rename e delete non possono rompere silenziosamente le dipendenze;
- il clone programma genera un set coerente di file.

### M6 — Viewer output banco ✅

Deliverable:

- ✅ browser output sola lettura, scalabile, con indice lazy.

Task:

1. ✅ modello dati metadati archivio output — `OutputFile`, `OutputArchive` in
   `domain/output_archive.py`.
2. ✅ scansione iniziale solo metadati (stat su `.csv`, mai `read_bytes`).
3. ✅ cache persistente indice metadati — `.at614_output_index.json` con
   ri-scansione automatica su mtime cambiati o file aggiunti.
4. ✅ browser risultati con filtro stringa libera + range date di creazione
   nell'`OutputViewer`.
5. ✅ apertura on-demand del CSV selezionato — `read_csv_columns(path)`,
   encoding auto-detect utf-8/cp1252.
6. ✅ scelta asse X e uno o più assi Y — combo X + list multi-select Y popolati
   dall'header del file scelto.
7. ✅ `OutputViewer` con `ReadOnlyBanner` permanente.
8. ✅ azioni limitate a `Esporta`, `Confronta`, `Stampa`, `Apri cartella`;
   `Esporta` funzionante via `QFileDialog`.

DoD:

- ✅ il viewer naviga archivi grandi senza caricare tutti i CSV all'avvio
  (solo metadati);
- ✅ il file viene parsato solo all'apertura del risultato;
- ✅ il grafico consente la scelta esplicita degli assi.

### M7 — Packaging e documentazione

Deliverable:

- build Windows distribuibile;
- manuale essenziale;
- logging minimo.

Task:

1. creare spec PyInstaller;
2. aggiungere logging con rotazione;
3. redigere manuale breve in italiano;
4. aggiungere checklist di collaudo manuale prima della release.

DoD:

- l'applicazione si avvia su Windows senza Python preinstallato;
- esiste documentazione minima per produzione e ufficio tecnico.

---

## 4. Dipendenze chiave

| Da | A | Motivo |
|---|---|---|
| M0 | M1 | servono fixture e inventario prima dei parser |
| M1 | M2 | la UI deve appoggiarsi a modelli e indice riferimenti |
| M2 | M3 | gli editor usano componenti base e shell UI |
| M1 + M2 | M4 | il `SeqEditor` dipende da parser TEST e componenti UI |
| M4 | M5 | non si puo validare bene senza editor sequenza operativo |
| M2 | M6 | il viewer output riusa shell UI e componenti grafici |

---

## 5. Decisioni chiuse

Queste decisioni sono state fissate e non sono piu considerate bloccanti.

1. `SeqEditor`: catalogo schema come fonte autoritativa, fallback generico per
  ID non rifiniti o parziali.
2. `Rinomina`: sempre globale sulla risorsa, con aggiornamento atomico dei
  riferimenti; nessuna opzione "rinomina solo qui".
3. `ExternalFileEditor`: vista strutturata se riconosciuta, altrimenti editor
  testo assistito con warning e preservazione del contenuto.
4. `Usi`: popover inline se `n <= 10`, tab `Riferimenti` della `RPanel` se
  `n > 10`.
5. `Validazione`: live per warning, on-save per errori bloccanti, comando
  esplicito `Validazione` per scansione completa.
6. `CE16` parziale: slot non compilati in stato `opz.` con testo
  `— non configurato —`.
7. `Elimina`: nessun cestino in v1, solo modale di impatto e conferma esplicita.
8. `Confronta`: picker modale filtrato per stessa tipologia, con overlay nello
  stesso editor/viewer.

---

## 6. Prossimo sprint consigliato

Sprint immediato suggerito: M0 + inizio M1.

Obiettivi dello sprint:

- creare struttura progetto;
- raccogliere fixture reali;
- implementare parser `distributore.py` e `test_csv.py`;
- impostare `project.py` con indice riferimenti minimo;
- preparare test round-trip iniziali.

Uscita attesa dello sprint:

- base tecnica verificabile;
- primi parser solidi;
- una direzione chiara per la UI senza ancora scrivere editor complessi.