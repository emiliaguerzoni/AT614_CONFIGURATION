# Audit implementazione — M0, M1, M2

Audit indipendente dello stato del codice al termine della validazione M2,
confrontato con la Definition of Done dichiarata in
[BACKLOG_IMPLEMENTAZIONE.md](BACKLOG_IMPLEMENTAZIONE.md) e con i contratti di
[PIANO_SVILUPPO.md](PIANO_SVILUPPO.md) e [LAYOUT_UX.md](LAYOUT_UX.md).

Esito generale: **53/53 test verdi**, struttura coerente con il brief UX,
M0 chiusa, M1 e M2 sostanzialmente complete con debiti tecnici tracciati.

> **Aggiornamento M0** — task 4 (matrice formati) e 5 (`bas_extractor.py`)
> chiusi successivamente: la cartella `schemas/auto/` ora contiene 15 YAML
> generati dai `Module_TEST_*.bas`, la matrice è documentata nel README.

---

## 1. Snapshot del repository

```
src/at614_editor/
├── __init__.py                       (__version__ = "0.1.0")
├── __main__.py                       (CLI con --version, --project-root, --smoke-ui)
├── domain/
│   ├── models.py                     (8 dataclass)
│   ├── project.py                    (AT614Project, load_project, indice riferimenti)
│   ├── test_schema.py                (estrazione file references — hard-coded 3 ID)
│   └── parsers/
│       ├── distributore.py
│       ├── test_csv.py
│       ├── point_series.py
│       ├── ce16.py
│       └── mms2218.py
└── ui/
    ├── main_window.py                (QStackedWidget Home/Shell + toolbar)
    ├── workspace_home.py             (dashboard con summary)
    ├── editor_shell.py               (layout 3 colonne, dispatch editor)
    ├── resource_tree.py              (9 categorie fisse)
    ├── components/
    │   ├── rpanel.py                 (3 tab + 6 azioni rapide)
    │   ├── file_widget.py            (icona + nome + 5 azioni)
    │   ├── resource_card.py          (titolo, file, meta, badge, 5 azioni)
    │   ├── chart_stage.py            (placeholder grafico + toolbar)
    │   └── read_only_banner.py
    └── editors/                      (M3 — già avviato in anticipo)
        ├── dist_editor.py
        ├── point_series_editors.py   (Curve/Limit/Ramp/CE16/MMS)
        └── external_file_editor.py

tests/                                (9 file di test, 32 test totali)
└── fixtures/                         (2 dist, 2 test, 14 series, 6 CE16, 2 MMS)
```

Mancano nel filesystem rispetto al piano:
- `schemas/auto/` e `schemas/overrides/`
- `domain/bas_extractor.py`
- `domain/parsers/external_config.py`

---

## 2. M0 — Baseline progetto e inventario formati

### 2.1 Task vs. realizzato

| Task M0 | Stato | Note |
|---|---|---|
| 1. `pyproject.toml` con dipendenze e script di avvio | ✅ Fatto | dipendenze ridotte a `PySide6` + `pytest` + `PyYAML` (dev), no ruff/mypy/pre-commit (era opzionale nel piano) |
| 2. struttura `src/at614_editor/` e `tests/` | ✅ Fatto | layout coerente con piano §3.2 |
| 3. fixture campione per ogni tipologia sorgente | ⚠️ Parziale | mancano fixture `file_esterni` (.txt referenziati) — tracciati solo come riferimenti irrisolti |
| 4. matrice tipo file → parser → editor → CRUD | ✅ Fatto | documentata nella sezione "File format inventory" del README |
| 5. `bas_extractor.py` con output YAML | ✅ Fatto | `src/at614_editor/domain/bas_extractor.py` + 15 file `schemas/auto/*.yaml` generati, 10 test unitari verdi |

### 2.2 DoD M0

- ✅ Repository avvia un entry point: `python -m at614_editor` funziona, smoke-ui test passa
- ⚠️ Fixture: principali sì, mancano `.txt` esterni e una varietà di formati legacy (`OLD`, EVO config)
- ✅ YAML base dei `Module_TEST_*.bas` generati: 15 file in `schemas/auto/`, conteggio parametri allineato al catalogo PIANO §1.3 (ACQUISIZIONE_PRE 17, RISPOSTE_GRADINO 12, ACQUISIZIONE_CAN 7, ASSEGNA_NODEID 1, CALIBRAZIONE 5, ecc.)

### 2.3 Stato del catalogo schema (M0.5 chiuso)

Il piano (§2.3) e il backlog (decisione chiusa #1) prevedono:
*catalogo schema come fonte autoritativa, fallback generico per ID non rifiniti*.

Lo scanner BAS ora popola `schemas/auto/*.yaml` con 15 voci. Il dict hard-coded
in [test_schema.py:8-20](src/at614_editor/domain/test_schema.py#L8-L20) resta
per ora come unica fonte usata a runtime (copre 3 ID), ma può essere
ri-espresso come `schemas/overrides/*.yaml` quando M4 introdurrà il loader
schema-driven nel `SeqEditor`.

**Anomalie note nei nomi dei file generati** (da risolvere via overrides in M4):
- `CICLICA_CAN.yaml` → l'ID effettivo nel csv è `TEST_CICLICA_CAN`
- `SCRITTURA_SERIAL_NUMBER.yaml` → l'ID effettivo nel csv è `SERIAL_NUMBER`

Lo scanner strippa fedelmente solo `Module_TEST_`, lasciando agli override il
compito di gestire le inconsistenze storiche di naming nel software VB6.

---

## 3. M1 — Core domain, parser e riferimenti

### 3.1 Task vs. realizzato

| Task M1 | Stato | Note |
|---|---|---|
| 1. modelli risorse (`ProjectResource`, ecc.) | ⚠️ Parziale | dataclass per ogni risorsa OK, ma manca classe base `ProjectResource` comune; `ExternalConfigResource` assente |
| 2. parser `distributore.py` | ✅ Fatto | round-trip byte-identico testato |
| 3. parser `test_csv.py` | ✅ Fatto | round-trip OK con trailing fields preservati a livello riga (vedi rischio §3.4) |
| 4. parser `curva_comando.py`, `curva_limite.py`, `rampa_xy.py` | ⚠️ Unificato | un solo `point_series.py` serve tutti e tre — semanticamente OK ma diverge dal nome dei file richiesti dal backlog |
| 5. parser `ce16.py` e `mms2218.py` | ✅ Fatto | wrapper su `point_series.py` con estrazione profilo/slot dal path |
| 6. parser `external_config.py` con fallback best-effort | ❌ **Mancante** | nessun parser per i `.txt` esterni; `ExternalFileEditor` mostra solo placeholder |
| 7. `project.py` con indice riferimenti | ✅ Fatto | `uses`/`used_by` bidirezionali, `unresolved_file_references` tracciati |
| 8. test round-trip per ogni tipologia | ✅ Fatto | un round-trip byte-identico per ognuno dei parser principali |
| 9. test dell'indice riferimenti | ✅ Fatto | 7 test in `test_project_index.py`, coprono dist→test, dist→CE16 con alias profilo, test→risorse, riferimenti irrisolti |

### 3.2 DoD M1

- ✅ Round-trip senza perdita: confermato per tutti i parser (test verdi)
- ✅ Indice risponde a "chi uso" (`get_uses`) e "chi mi usa" (`get_used_by`)
- ✅ Test base verdi: 32/32

### 3.3 Punti positivi

- **Indice riferimenti con alias multi-livello** (`_build_resource_file_index`,
  [project.py:61-82](src/at614_editor/domain/project.py#L61-L82)): registra
  ogni file con nome, path relativo, e per i `SETTAGGI PROGRAMMA` anche con
  alias `<profilo>/<file>` e `<profilo>/<gruppo>/<file>`. Questo risolve
  correttamente il caso `MERLO/CE16_1.cfg` evitando l'ambiguità con
  `DEFAULT/CE16_1.cfg` ed è testato.
- **Encoding handling consistente**: ogni parser legge bytes, prova UTF-8 e
  fallback su cp1252, preserva line ending originale e `endswith_newline`. Il
  round-trip è byte-identico nei fixture testati.
- **Risoluzione "unique-or-skip"**: `_pick_unique_path`
  ([project.py:85-89](src/at614_editor/domain/project.py#L85-L89)) ritorna
  `None` se l'alias matcha più di un file, evitando link spuri. Sensato.

### 3.4 Rischi e debiti tecnici M1

1. **Righe vuote in mezzo al `.csv` test → silenziosamente droppate.**
   [test_csv.py:36-37](src/at614_editor/domain/parsers/test_csv.py#L36-L37):
   `if not raw_line.strip(): continue`. Se un file legacy ha una riga vuota
   intermedia, viene parsata e poi non riemessa al save. I fixture attuali non
   contengono questo caso, ma è una rottura potenziale del round-trip su
   produzione. **Mitigazione:** registrare le righe vuote come `TestRow`
   "sentinella" (analogamente a `DistributoreLine` che già conserva tutte le
   righe), oppure aggiungere test esplicito che fallisca su un fixture con
   blank line.

2. **`TestSequence` non memorizza il numero di trailing columns dell'header.**
   Il piano §4 (modello `TestSequence`) prevede un campo `trailing_columns: int`.
   L'implementazione lo gestisce de facto preservando `parameters` per riga
   (lunghezze indipendenti). Funziona oggi perché i fixture hanno righe
   coerenti col proprio numero di `;`, ma se si **aggiunge** una riga via UI
   sarà difficile sapere quanti trailing empty mettere senza un riferimento
   esplicito. **Mitigazione:** aggiungere `header_trailing_empty: int` o
   memorizzare `expected_field_count` dall'header.

3. **`test_csv.serialize` ignora il numero di campi dell'header per le righe.**
   Conseguenza diretta di (2): una row con 6 parametri produce sempre 6 trailing
   campi indipendentemente da quanti ne avesse l'header. Per i fixture testati
   le righe erano già "lunghe quanto l'header" alla lettura, quindi
   round-trip OK. Non visibile finché non si edita davvero.

4. **`ProjectResource` base class mancante.** Task M1.1 la cita esplicitamente.
   Oggi ogni risorsa è una dataclass slegata. Quando arriverà M5 (validazione,
   delete, rename con `ImpactList`) servirà un protocollo comune (es. `path`,
   `display_name`, `category_key`, `unresolved_references()`). Aggiungerlo ora
   è poco costoso, dopo si paga col refactoring di 6 editor.

5. **`ExternalConfigResource` non esiste.** Il piano §2.4 e il backlog M3.7
   prevedono `ExternalFileEditor` "doppia modalità: strutturata se il formato è
   riconosciuto, testo assistito con warning". Oggi `ExternalFileEditor` mostra
   solo un placeholder statico, non legge nemmeno il contenuto del file. Per la
   strada strutturata mancano sia il parser sia il modello.

6. **`discover_project_root` è fragile.**
   [main_window.py:14-21](src/at614_editor/ui/main_window.py#L14-L21) scansiona
   solo `cwd` e i suoi figli diretti. Se lanciato dalla cartella `src/`, non
   trova mai la root. Non è un blocco ma è facile da migliorare risalendo i
   parent.

7. **`_detect_encoding` duplicato in 3 parser.** Stesso codice in
   `distributore.py`, `test_csv.py`, `point_series.py`. Va estratto in
   `io/encoding.py` come previsto dal piano §3.2.

---

## 4. M2 — Shell UI e componenti strutturali

### 4.1 Task vs. realizzato

| Task M2 | Stato | Note |
|---|---|---|
| 1. `main_window.py` e `workspace_home.py` | ✅ Fatto | `QStackedWidget` Home/Shell + toolbar Dashboard/Shell |
| 2. layout editing fisso `240 \| fluid \| 260` | ✅ Fatto | `setFixedWidth(240)` per albero, `setFixedWidth(260)` per RPanel — più stretto del `≈` del brief, ma rispetta la grandezza |
| 3. `resource_tree.py` con 9 categorie fisse nell'ordine del brief | ✅ Fatto | `CATEGORY_DEFINITIONS` 9 elementi, ordine identico a LAYOUT_UX §3.3 |
| 4. `RPanel` con `Riferimenti`, `Validazione`, `Azioni rapide` | ✅ Fatto | 3 tab nell'ordine corretto, focus automatico su Validazione se ci sono avvisi |
| 5. `FileWidget` | ✅ Fatto | icona + nome + 5 azioni (`Seleziona`, `Nuovo`, `Duplica`, `Apri`, `Usi`) + stato "riferimento mancante" con colore rosso |
| 6. `ResourceCard` | ✅ Fatto | titolo + file + meta + badge stato + 5 azioni |
| 7. `ChartStage` e `ReadOnlyBanner` | ✅ Fatto | placeholder grafico con toolbar configurabile + banner giallo "Sola lettura" |
| 8. microcopy e stati definiti in LAYOUT_UX.md | ✅ Fatto | italiano coerente, badge `stabile`/`opz.`/`opzionale`, microcopy "Sezione non attiva · Attiva sezione" |

### 4.2 DoD M2

- ✅ Dashboard e shell UI navigabili (smoke-ui test passa)
- ✅ Ogni schermata editor usa il layout corretto (verificato via
  `test_main_window_builds_shell_layout`)
- ✅ Componenti base pronti per essere riusati dagli editor

### 4.3 Punti positivi

- **Test UI robusti**: `test_ui_shell.py` verifica con assert puntuali sia il
  layout (larghezze, ordine tab, etichette categorie) sia il routing (apertura
  di DistEditor/RampEditor/Ce16Editor/ExternalFileEditor). Sono test che
  prevengono regressioni reali, non smoke vuoti.
- **`offscreen` Qt platform forzato in `conftest.py`**: tests UI girabili in CI
  senza display. Buona ergonomia.
- **`activate_path` e `activate_first_resource`**: API navigazionale pulita,
  testabile, usata sia dal main window sia dai test.
- **`RPanel.set_context(focus_validation=...)`**: cambia automaticamente tab su
  Validazione se ci sono unresolved references — implementa direttamente la
  decisione chiusa #5 (validazione live per warning).

### 4.4 Punti di attenzione e debiti M2

1. **`Sostituisci` nelle azioni rapide non è previsto dal brief.** LAYOUT_UX
   §4.3 elenca per `Azioni rapide` "Nuovo", "Duplica", "Apri", "Elimina", "Usi"
   (5 azioni); l'implementazione ne ha 6 aggiungendo `Sostituisci` e usa
   `Nuovo collegato`/`Duplica collegato` invece dei nomi del brief. Va
   allineato con LAYOUT_UX o aggiornato il brief — al momento c'è una
   divergenza non giustificata.

2. **Header dell'editor centrale non rispetta lo schema della scheda.**
   LAYOUT_UX §3.4 indica per il centro: titolo + breadcrumb + badge stato +
   tab interne. L'implementazione ha titolo + sottotitolo (non breadcrumb) e
   un solo badge `stabile` cablato sempre verde. Quando si arriverà alla
   validazione (M5) andrà rivisto.

3. **Icone testuali come stringhe (`□`, `▲`, `○`, `◇`).** Funzionali, ma non
   accessibili a screen reader e fragili in font diversi. È esplicitamente un
   placeholder, non un problema bloccante.

4. **`open_first_available_resource` dipende dall'ordine alfabetico di
   `_build_grouped_selections`.** Se domani il primo distributore in ordine
   alfabetico non esiste più, il test `test_main_window_opens_first_resource`
   passerà ancora ma punterà a un'altra risorsa. Test brittle.

5. **Header dell'editor mostra "Dashboard" come titolo di default**: confonde
   perché in `MainWindow` "Dashboard" è il nome dell'altra pagina (Home). La
   `show_empty_state()` setta titolo "Dashboard" — semanticamente sbagliato per
   lo stato vuoto della shell.

6. **`workspace_home._build_summary_lines` non elenca file esterni risolti**,
   solo gli "irrisolti". Manca la voce "File esterni" totali — coerente col
   fatto che `ExternalConfigResource` non esiste ancora.

---

## 5. M3 — avvio in anticipo

Il backlog colloca M3 dopo la validazione M2. Tuttavia sono già implementati 7
editor M3:

| Editor M3 | Implementato | Stato |
|---|---|---|
| `DistEditor` | ✅ | 5 sezioni + 5 slot CE16 + tabella parametri extra (sola lettura) |
| `CurveEditor` | ✅ | ChartStage + tabella punti (sola lettura) |
| `LimitEditor` | ✅ | ChartStage + tabella punti (sola lettura) |
| `RampEditor` | ✅ | ChartStage + tabella punti (sola lettura) |
| `Ce16Editor` | ✅ | con badge `opzionale` |
| `MmsEditor` | ✅ | con toolbar `Range ADC` |
| `ExternalFileEditor` | ⚠️ | solo placeholder testuale, no doppia modalità |

**Conseguenza positiva:** la shell M2 è già stata stressata da editor reali,
non solo da widget vuoti. Il routing in `EditorShell.show_selection` è già
operativo per tutte le categorie.

**Conseguenze problematiche:**
- Nessuna delle azioni `Nuovo` / `Duplica` / `Salva` / `Elimina` è funzionale.
  Tutti i bottoni sono dichiarati ma scollegati. La DoD M3 (*"tutte le risorse
  non-TEST principali si possono aprire, salvare e duplicare"*) **non è ancora
  raggiunta**.
- Tutte le tabelle sono `NoEditTriggers`, quindi non c'è editing reale.
- Il dispatch `editor_shell._show_test_sequence` è uno **stub**: elenca 4
  `FileWidget` con i primi riferimenti file della sequenza. Non è un
  `SeqEditor`, ed è chiaro nei commenti.

Va deciso esplicitamente se M3 procede in parallelo all'audit M2 o se questi 7
editor restano come "scaffolding pre-M3" e verranno cablati come parte di M3
piena.

---

## 6. Test e qualità

- **32 test, 100% pass**, runtime ~0.9 s
- Distribuzione: 7 UI shell + 7 project index + 3 test_schema + 3 test_csv +
  3 distributore + 3 point_series + 2 mms2218 + 2 ce16 + 2 smoke
- **Coverage non misurata** (il piano §1 dice "coverage > 85% del modulo
  `domain/`"). Non c'è `pytest-cov` configurato.
- **Mancano test negativi** che falliscano deliberatamente su:
  - encoding non gestibile (né UTF-8 né cp1252)
  - blank line in mezzo al file TEST (rischio §3.4.1)
  - distributore con header non riconosciuto
  - alias non univoco nell'indice risorse
- **Mancano test sul round-trip dei distributori che hanno commenti**:
  `15.1001.310_RC.cfg` ha commenti `//` e righe vuote. C'è
  `test_parse_distributore_base_fields` ma il round-trip byte-identico è
  testato solo su `15.1001.356_C.cfg`. Vale la pena duplicare l'asserzione di
  byte-equality anche sull'altro.

---

## 7. Conformità ai brief

### vs. PIANO_SVILUPPO.md

| Punto piano | Realtà | Verdetto |
|---|---|---|
| Stack: Python+PySide6, PyQtGraph, pydantic, PyInstaller | PySide6 ✓, **PyQtGraph mancante** (ChartStage è placeholder QLabel), **pydantic non usato** (scelta dichiarata: dataclass), PyInstaller non ancora applicabile (M7) | Allineato con la nota §3.1 "privilegiare la soluzione più semplice" |
| Schema parametri: YAML auto + override | Dict hard-coded a 3 ID | **Divergenza** rispetto alla decisione chiusa #1 |
| Sub-editor curve/rampe con plot interattivi | Tabella read-only + placeholder grafico | OK per M2/early-M3, ma è la feature differenziante: andrà completata in M3 piena |

### vs. LAYOUT_UX.md

| Componente | Brief | Implementato |
|---|---|---|
| Layout 240/fluid/260 | "≈" approssimato | `setFixedWidth` esatto |
| 9 categorie nell'ordine: Distributori, Sequenze test, Curve comando, Curve limite, Rampe XY, Calibrazioni CE16, Calibrazioni MMS2218, File esterni, Output banco | identico | ✅ |
| `Azioni rapide`: Nuovo, Duplica, Apri, Elimina, Usi | "Nuovo collegato", "Duplica collegato", "Sostituisci", "Apri", "Elimina", "Usi" | ⚠️ Divergenza nei nomi e numero |
| Tab `Riferimenti` / `Validazione` / `Azioni rapide` | identico ordine | ✅ |
| `ReadOnlyBanner` testo "Sola lettura — questo file è un output di banco" | identico | ✅ |
| Microcopy in italiano | ✅ | ✅ |

### vs. BACKLOG_IMPLEMENTAZIONE.md (DoD)

- **M0 DoD**: 2/3 soddisfatti, **YAML moduli TEST non generati**
- **M1 DoD**: 3/3 sostanzialmente soddisfatti (con i rischi §3.4)
- **M2 DoD**: 3/3 soddisfatti

---

## 8. Verdetto sintetico

| Milestone | Esito | Note |
|---|---|---|
| M0 | 🟢 **Completo** | scanner BAS + 15 YAML generati + matrice formati documentata; resta solo l'aggiunta opzionale di fixture `.txt` esterni |
| M1 | 🟢 **Completo** | con 5-7 debiti tecnici tracciati ma non bloccanti |
| M2 | 🟢 **Completo** | piccole divergenze di microcopy/azioni da risolvere con il maintainer del LAYOUT_UX |

## 9. Azioni raccomandate prima di aprire M3 a pieno regime

In ordine di priorità:

1. **Chiudere M0.5** — implementare `bas_extractor.py` e generare
   `schemas/auto/*.yaml` per i 15 moduli TEST. Senza schema autoritativo il
   `SeqEditor` di M4 non può partire pulito.
2. **Allineare azioni rapide** con LAYOUT_UX: o aggiornare il brief
   formalmente, o riportare le 6 azioni alle 5 previste. Decidere
   sull'aggiunta di `Sostituisci` con una nota nelle "Decisioni chiuse".
3. **Aggiungere classe base `ProjectResource`** (anche minimale: `path`,
   `category`, `display_name`) per non pagare il refactoring in M5.
4. **Estrarre `io/encoding.py`** consolidando i tre `_detect_encoding`
   duplicati e `_detect_line_ending`.
5. **Conservare le blank line in `test_csv.parse`** o aggiungere un test che
   protegga dalla loss. Idem byte-identical round-trip anche sul distributore
   con commenti.
6. **Aggiungere `ExternalConfigResource` + parser fallback** (anche solo
   conservazione bytes con metadata), così `ExternalFileEditor` può uscire
   dallo stato placeholder.
7. **Configurare `pytest-cov`** e fissare la soglia 85% su `domain/` come da
   piano.
8. **Decidere lo status dei 7 editor M3 già abbozzati**: continuare in M3
   piena oppure renderli read-only ufficiali fino al via libera M3.

Nessuna di queste azioni invalida quanto fatto fino a M2: sono debiti tracciati
con costo basso di rientro se affrontati ora, alto se rimandati a M4-M5.
