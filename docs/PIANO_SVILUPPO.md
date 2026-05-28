# Piano di sviluppo — AT614 Configuration Editor

> Editor grafico in italiano per creare, modificare, duplicare ed eliminare i
> file di configurazione del banco AT614, pensato per utenti non programmatori
> e compatibile con la logica VB6 esistente.

---

## 1. Obiettivo e confini

### 1.1 Obiettivo di prodotto

L'obiettivo del progetto e fornire uno strumento che permetta a un utente non
programmatore di costruire un nuovo programma di collaudo senza intervenire sui
sorgenti VB6. L'editor deve:

- creare un nuovo programma da zero o partendo da un template esistente;
- modificare tutte le risorse di configurazione coinvolte;
- duplicare, rinominare ed eliminare in modo sicuro file collegati tra loro;
- evidenziare subito riferimenti mancanti o incoerenti;
- salvare file compatibili con il software legacy;
- presentare una UI solo in italiano.

Per la parte UX/UI fa fede [LAYOUT_UX.md](LAYOUT_UX.md). Il piano di sviluppo e
[PROPOSTA_FLUSSO_GUI_UX.md](PROPOSTA_FLUSSO_GUI_UX.md) devono restare coerenti
con quel brief operativo e non ridefinirne microcopy, stati o componenti base.

### 1.2 Vincoli di progetto

| Aspetto | Scelta |
|---|---|
| Lingua UI | Solo italiano |
| Priorita | Produttivita utente e facilita di implementazione |
| Piattaforma | Desktop Windows, offline, single-user |
| Compatibilita | Nessuna modifica alla logica VB6; l'editor produce solo file compatibili |
| Approccio | Parser semplici, round-trip affidabile, validazione cross-file, editor dedicati solo dove portano valore reale |

### 1.3 Confini del dominio

| Tipo | Ruolo | Da interpretare/modificare | Note |
|---|---|---|---|
| `DISTRIBUTORE/*.cfg` | configurazione sorgente | si | definisce sezioni, calibrazioni CE16, limiti olio e pressione |
| `TEST/*.csv` | configurazione sorgente | si | sequenza test per modulo |
| `CURVE_COMANDO/*.csv` | configurazione sorgente | si | curva comando |
| `CURVE_LIMITE/*.csv` | configurazione sorgente | si | limiti di accettazione |
| `RAMPE_XY/*.csv` | configurazione sorgente | si | profili tempo -> tensione |
| `SETTAGGI PROGRAMMA/<computer>/CE16/*.cfg` | configurazione sorgente | si | curva del sensore CE16, opzionale |
| `SETTAGGI PROGRAMMA/<computer>/MMS2218/*.csv` | configurazione sorgente | si | curve di calibrazione ADC |
| file esterni referenziati dai test (`.txt`, file EVO, altri setup) | configurazione sorgente | si | almeno apertura, validazione e duplicazione assistita |
| `GRAPH/**` | output del banco | no | esempio di output interno al progetto; il viewer deve supportare anche archivi output parametrizzabili esterni |
| `RAMPE_XY_LAST/**` | output del banco | no | esempio di output interno al progetto; solo consultazione, fuori dal modello di configurazione |
| cartelle `OLD`, `OLD FILE` | archivio storico | lettura/import | escluse dal flusso standard di authoring |

### 1.4 Nota funzionale sul CE16

Il CE16 e un sensore aggiuntivo che il software EOL puo usare oppure no per
effettuare un'acquisizione piu precisa del sistema. L'editor deve quindi:

- trattare i file CE16 come risorsa opzionale di configurazione;
- rendere evidente quando un distributore o un test li usa;
- non assumere che siano sempre presenti.

---

## 2. Analisi del dominio

### 2.1 Relazioni principali tra file

```
DISTRIBUTORE/<codice>.cfg
  ├─ sezione1..5 -> TEST/<codice modulo>.csv
  ├─ calibrazioneCE16_1..5 -> path relativo a file CE16
  ├─ Temperatura_olio_min/max
  └─ Pressure_min/max

TEST/<codice modulo>.csv
  ├─ righe: NOME TEST ; ID TEST ; PARAMETRO 0..N
  ├─ Parameter(0) = indice del test
  ├─ Parameter(1..N) = parametri specifici del test
  └─ puo referenziare curve, rampe, limiti, file esterni, settaggi CE16/MMS2218

CURVE_COMANDO/<file>.csv
CURVE_LIMITE/<file>.csv
RAMPE_XY/<file>.csv
SETTAGGI PROGRAMMA/<computer>/CE16/<file>.cfg
SETTAGGI PROGRAMMA/<computer>/MMS2218/<file>.csv
  └─ risorse di configurazione editabili dal tool

GRAPH/**
RAMPE_XY_LAST/**
  └─ file di uscita del banco; non partecipano all'authoring del programma
```

Nota: il viewer output non deve assumere che i file stiano dentro la root del
progetto. L'archivio puo vivere in uno o piu percorsi esterni parametrizzabili,
tipicamente organizzati in molte cartelle, una per salvataggio o prodotto, con
un file CSV grafico all'interno.

Nota: i file di impostazione grafica usati da alcuni test avanzati sono file di
configurazione esterni referenziati dai parametri del test; non coincidono con
la cartella `GRAPH/`, che contiene output.

### 2.2 Parsing del CSV TEST

Riferimento: `BAS/Module_TEST_MANAGER.bas` (`ApriFileTest`).

- Separatore `;`, prima riga = header
- `dati(0)` = nome test
- `dati(1)` = ID test
- `dati(2..N)` = `Parameter(0..N-2)`
- `Parameter(0)` e sempre l'indice del test e supporta i token:
  - `i`
  - `i++`
  - `++i`
- I parametri successivi cambiano significato in base all'ID test
- La serializzazione deve preservare encoding, line ending, colonne trailing e
  formattazione compatibile con i file legacy

### 2.3 Strategia di supporto agli ID TEST

L'editor non deve reimplementare la logica di esecuzione del banco. Deve
conoscere solo:

- lo schema dei parametri;
- il tipo di editor da aprire per i parametri che referenziano altri file;
- le regole minime di validazione.

Per produttivita conviene usare un approccio ibrido:

- schema auto-generato dai `Module_TEST_*.bas` per nomi e indici dei parametri;
- override manuali per etichette, tipi, descrizioni e file picker;
- fallback generico per i test rari o non ancora rifiniti.

### 2.4 Copertura funzionale minima degli ID principali

| ID TEST | Gestione prevista |
|---|---|
| `ASSEGNA_NODEID`, `LETTURA_CFG_HW`, `SERIAL_NUMBER` | form semplice tipizzato |
| `SCRITTURA_PARAMETRI` | picker verso file esterno + validazione esistenza + azioni duplica/apri |
| `CALIBRAZIONE` | form tipizzato con valori numerici |
| `TEST_CICLICA_CAN` | form tipizzato con numeri e riferimenti |
| `ACQUISIZIONE_CAN` | form tipizzato + integrazione con editor curva/rampa/limiti |
| `RISPOSTE_GRADINO` | form schema-driven con riferimenti file |
| `ACQUISIZIONE_PRE`, `ACQUISIZIONE_GENERICA`, `ACQUISIZIONE_POST` | form schema-driven + picker per risorse + eventuale editor dedicato successivo |
| `TEST_CICLICA_CAN_EVO`, `ACQUISIZIONE_PRE_EVO` | picker verso file esterno + viewer/editor assistito per la configurazione referenziata |
| `SEQUENZA_PLC`, `POST_PROCESSING` | apertura sicura e modifica tramite schema-driven editor; editor specializzato solo se davvero necessario |

---

## 3. Architettura del codice

### 3.1 Scelte tecniche

| Aspetto | Scelta |
|---|---|
| Stack UI | Python 3.11+ con PySide6/Qt6 |
| Grafici | PyQtGraph |
| Gestione file dati | standard library (`csv`, `pathlib`) + parser custom per i formati legacy |
| Modelli/validazione | `pydantic` o `dataclasses` con validator leggeri; privilegiare la soluzione piu semplice |
| Schema parametri TEST | YAML auto-generato + override manuali |
| Packaging | PyInstaller |
| Test | pytest con round-trip su fixture reali |

Scelta esplicita di semplificazione:

- niente database;
- niente backend separato;
- niente `pandas` salvo necessita reale dimostrata;
- logica di salvataggio centrata su parser deterministici e non su trasformazioni
  invasive dei file.

### 3.2 Struttura del progetto

```
at614_editor/
├── pyproject.toml
├── README.md
├── schemas/
│   ├── auto/
│   └── overrides/
├── src/at614_editor/
│   ├── __main__.py
│   ├── app.py
│   ├── domain/
│   │   ├── models.py
│   │   ├── resource_types.py
│   │   ├── project.py
│   │   ├── references.py
│   │   ├── validators.py
│   │   ├── bas_extractor.py
│   │   ├── schemas.py
│   │   ├── refactor.py
│   │   └── parsers/
│   │       ├── distributore.py
│   │       ├── test_csv.py
│   │       ├── curva_comando.py
│   │       ├── curva_limite.py
│   │       ├── rampa_xy.py
│   │       ├── ce16.py
│   │       ├── mms2218.py
│   │       └── external_config.py
│   ├── ui/
│   │   ├── main_window.py
│   │   ├── workspace_home.py
│   │   ├── resource_tree.py
│   │   ├── editors/
│   │   │   ├── distributore_editor.py
│   │   │   ├── test_sequence_editor.py
│   │   │   ├── curva_editor.py
│   │   │   ├── limite_editor.py
│   │   │   ├── rampa_editor.py
│   │   │   ├── ce16_editor.py
│   │   │   ├── mms2218_editor.py
│   │   │   ├── external_config_editor.py
│   │   │   └── acquisition_viewer.py
│   │   ├── widgets/
│   │   │   ├── file_ref_field.py
│   │   │   ├── reference_panel.py
│   │   │   └── plot_widget.py
│   │   └── dialogs/
│   │       ├── new_resource_dialog.py
│   │       ├── clone_program_dialog.py
│   │       ├── delete_resource_dialog.py
│   │       └── new_distributore_wizard.py
│   └── io/
│       ├── encoding.py
│       └── csv_dialect.py
└── tests/
    ├── fixtures/
    ├── test_parsers.py
    ├── test_round_trip.py
    ├── test_references.py
    ├── test_refactor.py
    └── test_schema_merge.py
```

### 3.3 Principi architetturali

- il progetto e centrato sulle risorse, non solo sui file `TEST`;
- ogni tipo di file sorgente ha parser dedicato e editor dedicato o assistito;
- i riferimenti tra file sono indicizzati una sola volta in `AT614Project`;
- operazioni come duplica, rinomina ed elimina passano da servizi centralizzati
  che aggiornano o verificano le dipendenze;
- `GRAPH` e `RAMPE_XY_LAST` restano fuori dal modello di configurazione.

---

## 4. Modello dati e metadati

### 4.1 Modelli core

I modelli core devono rappresentare sia il contenuto dei file sia le relazioni
tra risorse. I tipi minimi sono:

- `ProjectResource`: tipo risorsa, path, stato, lista dei referenti;
- `DistributoreConfig`: sezioni, riferimenti CE16, parametri aggiuntivi;
- `TestRow` e `TestSequence`: righe test, header, encoding, trailing columns;
- `PointSeriesResource`: serie di punti per curve, rampe, limiti, CE16, MMS2218;
- `ExternalConfigResource`: contenuto e metadati dei file esterni non ancora
  modellati in modo piu specifico;
- `ReferenceIndex`: mappa bidirezionale fra file che referenziano e file
  referenziati.

### 4.2 Schema parametri TEST

Lo schema YAML resta una buona idea, ma va chiarito che serve solo per i test
`TEST/*.csv`, non per tutte le altre tipologie di file.

Esempio minimo:

```yaml
id: ACQUISIZIONE_CAN
display_name: Acquisizione CAN
parameters:
  - index: 1
    name: NomeFileCurvaComando
    label: File curva comando
    type: file_ref
    resource_type: curva_comando
    editor: curva
    required: true
  - index: 2
    name: NomeFileRampaXY
    label: File rampa XY
    type: file_ref
    resource_type: rampa_xy
    editor: rampa
    required: true
  - index: 6
    name: LimiteInferiore
    label: Curva limite inferiore
    type: file_ref
    resource_type: curva_limite
    editor: limite
```

Regole di merge:

- `schemas/auto/*.yaml` generati dallo scanner BAS;
- `schemas/overrides/*.yaml` per affinare tipi, label, descrizioni e picker;
- in assenza di metadati completi, il parametro ricade su editor testuale
  semplice ma salva comunque in modo compatibile.

### 4.3 Tipologie di risorse da modellare in modo esplicito

Per l'MVP reale vanno modellati in modo esplicito almeno questi formati:

- distributore;
- sequenza test;
- curva comando;
- curva limite;
- rampa XY;
- curva CE16;
- curva MMS2218;
- file esterni referenziati dai test.

---

## 5. Editor grafici e operazioni richieste

### 5.1 Regola generale

Per ogni tipologia di file di configurazione realmente usata deve esistere una
UI dedicata che consenta almeno:

- nuovo file;
- apertura file;
- duplica da file esistente della stessa tipologia;
- rinomina;
- elimina con conferma e verifica riferimenti;
- salva e salva con nome;
- visualizzazione dei file che usano la risorsa;
- apertura rapida dai parametri che la referenziano.

### 5.2 Elenco editor da realizzare

| Tipo file | Editor previsto | Operazioni principali |
|---|---|---|
| `DISTRIBUTORE/*.cfg` | form strutturato per sezioni, CE16 e parametri extra | nuovo, duplica distributore, rinomina, elimina, mostra dipendenze |
| `TEST/*.csv` | tabella sequenza + form dinamico parametri | nuovo, duplica sequenza, drag-drop, reindicizzazione, template righe |
| `CURVE_COMANDO/*.csv` | grafico + tabella sincronizzata | nuovo, duplica, editing punti, validazione monotonia se richiesta |
| `CURVE_LIMITE/*.csv` | editor curva singola con apertura accoppiata INF/SUP | nuovo, duplica, confronto banda |
| `RAMPE_XY/*.csv` | editor tempo/tensione con punti trascinabili | nuovo, duplica, snap e ordinamento punti |
| `SETTAGGI PROGRAMMA/.../CE16/*.cfg` | editor grafico punti/posizione | nuovo, duplica, selezione profilo banco |
| `SETTAGGI PROGRAMMA/.../MMS2218/*.csv` | editor grafico calibrazione ADC | nuovo, duplica, controllo ordinamento e range |
| file esterni `.txt` o setup non strutturati | editor assistito testuale/chiave-valore | apri, duplica, rinomina, validazione base |
| archivi output banco (`GRAPH`, `RAMPE_XY_LAST` o root esterne) | viewer opzionale con browser dedicato | solo lettura, filtro stringa/data, scelta asse X e Y, confronto, export, stampa, apri cartella |

### 5.3 UX dei riferimenti file

Ogni parametro che punta a un'altra risorsa deve offrire un widget unico con:

`[Seleziona] [Nuovo] [Duplica] [Apri] [Usi]`

Comportamento richiesto:

- `Seleziona` sceglie una risorsa esistente compatibile con il parametro;
- `Nuovo` crea subito un file del tipo giusto e lo collega al parametro;
- `Duplica` clona un file esistente della stessa tipologia e lo collega;
- `Apri` porta all'editor dedicato;
- `Usi` mostra chi referenzia il file selezionato.

### 5.4 Flussi utente principali

I flussi minimi da supportare sono:

1. creare un nuovo distributore e associare le sue sequenze test;
2. duplicare un programma esistente rinominando anche le risorse collegate;
3. modificare una sequenza test e aprire al volo curve, rampe, limiti, CE16 e
   MMS2218;
4. individuare riferimenti mancanti prima del salvataggio;
5. eliminare una risorsa sapendo esattamente quali file la usano.

---

## 6. UI e organizzazione del workspace

### 6.1 Schermata principale

UI master-detail a tre aree, ma centrata sulle risorse del progetto:

- sinistra: albero risorse raggruppato per tipologia;
- centro: editor della risorsa selezionata;
- destra: pannello riferimenti, validazione e azioni rapide.

Per le schermate di editing il layout e fissato a tre colonne:

- albero risorse circa 240 px;
- editor centrale fluido;
- pannello destro contestuale circa 260 px.

Eccezioni ammesse: dashboard e wizard.

In alto conviene aggiungere tre azioni immediate:

- `Nuovo programma di collaudo`
- `Duplica programma esistente`
- `Apri risorsa`

### 6.2 Regole di usabilita

- solo terminologia italiana coerente con il dominio aziendale;
- errori spiegati in modo operativo, non tecnico;
- nessuna esposizione di concetti VB6 non indispensabili;
- nessuna esposizione di path filesystem grezzi nel flusso principale;
- salvataggio sicuro con conferma chiara quando un'azione impatta altri file;
- layout orientato alla modifica rapida, non all'esplorazione generica del file
  system.

### 6.4 Vincoli UI operativi

Il piano assume come invarianti queste decisioni gia fissate nel brief UX:

- albero risorse per tipologia funzionale, con 9 categorie fisse e ordine fisso;
- pannello destro con tre sezioni sempre nello stesso ordine: Riferimenti,
  Validazione, Azioni rapide;
- 5 sezioni distributore sempre visibili;
- 5 slot CE16 sempre visibili;
- 7 soli stati visivi di riferimento, senza tassonomie parallele nei mock-up o
  nella UI finale;
- componenti UI base nominati e riusabili: `FileWidget`, `ResourceCard`,
  `RPanel`, `ChartStage`, `ReadOnlyBanner`.

### 6.3 Funzioni trasversali

- dirty tracking;
- undo/redo;
- ricerca rapida per codice file;
- filtri per tipologia di risorsa;
- pannello riferimenti mancanti;
- report di validazione prima del salvataggio.

---

## 7. Roadmap di sviluppo

### Fase 0 — Inventario formati e bootstrap (2 giorni)

- classificare in modo definitivo file sorgente vs file di output;
- preparare fixture reali per ogni tipologia sorgente;
- creare `pyproject.toml`, struttura repo, lint e test base;
- implementare `bas_extractor.py` per gli schemi dei test.

**DoD:** esiste una matrice completa `tipo file -> parser -> editor -> CRUD` e
lo scanner genera gli YAML base per i moduli `Module_TEST_*.bas`.

### Fase 1 — Parser e round-trip di tutte le risorse sorgente (4 giorni)

- parser/serializer per distributore, test, curve comando, limiti, rampe,
  CE16, MMS2218 e file esterni minimi;
- gestione encoding cp1252/utf-8 e line ending preservati;
- `AT614Project` con indice riferimenti tra file.

**DoD:** round-trip byte-identico sui file campione di tutte le tipologie
sorgente supportate.

### Fase 2 — Resource workbench e CRUD generale (4 giorni)

- albero risorse per tipologia, con 9 categorie fisse nell'ordine del brief UX;
- open/save/save-as;
- nuovo/duplica/rinomina/elimina per ogni risorsa sorgente;
- dialog di delete con analisi referenze.

**DoD:** ogni file sorgente puo essere creato, duplicato e cancellato da UI
senza usare Esplora risorse, e il layout editing 3 colonne e operativo.

### Fase 3 — Editor grafici delle risorse base (5 giorni)

- implementazione componenti UI base riusabili del brief (`FileWidget`,
  `ResourceCard`, `RPanel`, `ChartStage`, `ReadOnlyBanner`);
- editor distributore;
- editor curva comando;
- editor limite;
- editor rampa XY;
- editor CE16;
- editor MMS2218.

**DoD:** tutte le tipologie sorgente non-TEST hanno un editor dedicato e
utilizzabile da un utente non tecnico.

### Fase 4 — Editor sequenza TEST e builder del programma (5 giorni)

- tabella sequenza con reindicizzazione automatica di `Parameter(0)`;
- form parametri dinamico guidato dagli schemi;
- widget file reference con `Seleziona/Nuovo/Duplica/Apri/Usi`;
- wizard nuovo distributore e clone programma completo.

**DoD:** un utente puo creare un nuovo programma di collaudo completo partendo
da template, senza editare file a mano.

### Fase 5 — Validazione cross-file e operazioni sicure (3 giorni)

- evidenziazione riferimenti mancanti;
- validazione tipi e range dei parametri principali;
- rinomina/duplica con aggiornamento riferimenti dove applicabile;
- report errori e warning pre-save.

**DoD:** il sistema impedisce o segnala in modo chiaro le incoerenze piu comuni
prima del salvataggio finale.

### Fase 6 — Viewer output, packaging e documentazione (3 giorni)

- viewer opzionale per archivi output banco con percorso parametrizzabile,
  anche esterno al progetto;
- filtri per stringa libera, date e metadati file/cartella;
- scelta asse X e di uno o piu assi Y nel grafico;
- indicizzazione lazy con cache metadati e caricamento CSV on-demand;
- viewer solo lettura con azioni consentite limitate a export, confronto,
  stampa e apertura cartella;
- build PyInstaller;
- manuale utente breve in italiano;
- logging con rotazione.

**DoD:** eseguibile Windows pronto all'uso, viewer output capace di navigare
archivi grandi con filtri e selezione assi, e documentazione minima per utenti
di produzione e ufficio tecnico.

---

## 8. Rischi e decisioni aperte

| Rischio | Mitigazione |
|---|---|
| Encoding misto cp1252/utf-8 | auto-detect al parse e riuso dell'encoding originale al save |
| Varianti legacy in `OLD` e `OLD FILE` | parse best-effort in sola lettura; non bloccare il flusso standard |
| File esterni referenziati ma non presenti nel workspace | validazione esplicita + fallback a editor testuale o placeholder |
| Parametri avanzati poco documentati | schema auto + override progressivi; fallback editor generico |
| Eliminazione o rinomina di file condivisi | dialog con impatto riferimenti e opzione annulla |
| Confusione fra input e output del banco | classificazione esplicita: `GRAPH` e `RAMPE_XY_LAST` viewer-only |
| CE16 presente solo in alcuni casi | modello opzionale, mai obbligatorio |
| Numero molto alto di file output | indice metadati lazy, cache persistente, refresh incrementale e parsing CSV solo on-demand |

Decisioni esplicite da mantenere ferme:

- il software resta solo in italiano;
- `GRAPH` e `RAMPE_XY_LAST` non fanno parte del dominio di authoring;
- il viewer output usa percorsi parametrizzabili e non dipende dalla sola root
  del progetto;
- il valore principale dell'MVP e il CRUD sicuro di tutte le risorse sorgente,
  non la copertura totale della logica runtime VB6;
- gli editor specializzati per test rari arrivano solo dopo che tutte le
  tipologie file sorgente hanno un flusso usabile.
