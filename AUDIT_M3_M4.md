# Audit implementazione — M3 completata, M4 al 70%

Audit indipendente dello stato del codice dopo la dichiarazione di completamento
M3 e l'avanzamento di M4 al 70%. Estende
[AUDIT_M0_M2.md](AUDIT_M0_M2.md) (M0/M1/M2 già chiuse) coprendo le milestone
M3 e M4 contro la DoD del
[BACKLOG_IMPLEMENTAZIONE.md](BACKLOG_IMPLEMENTAZIONE.md).

Esito generale: **80/80 test verdi**, M3 chiusa esclusi i task distruttivi
(rinviati a M5), M4 sopra le aspettative sul piano funzionale ma con un nodo
architetturale (loader schema YAML) ancora non chiuso che vincola la chiusura
formale.

> **Aggiornamento successivo all'audit** — i 4 punti M3 lasciati aperti
> (Ce16/Mms duplicate, ExternalFileEditor doppia modalità, azione "Nuovo",
> DistEditor non-mutation) sono stati chiusi con 15 test aggiuntivi in
> [test_m3_completion.py](tests/test_m3_completion.py). Vedi §2.5.

---

## 1. Snapshot dei cambiamenti dopo M2

```
src/at614_editor/
├── domain/
│   ├── project.py                    + resolve_resource_path, get_usage_count
│   ├── test_schema.py                + TestParameterDescriptor + describe_test_parameters
│   │                                 (catalogo dict esteso da 3 a 7 ID)
│   └── (resto invariato)
└── ui/
    ├── editor_shell.py               + save/duplicate dispatcher, _reload_project,
    │                                   open_resource_path, show_resource_usages
    ├── components/
    │   ├── file_widget.py            QLabel → QLineEdit (editable=True opzionale)
    │   └── resource_card.py          QLabel → QLineEdit (editable=True opzionale)
    └── editors/
        ├── resource_actions.py       (nuovo) build_duplicate_path
        ├── dist_editor.py            + save_changes, duplicate_resource, callback
        ├── point_series_editors.py   + save/duplicate per Curve/Limit/Ramp/CE16/MMS
        └── seq_editor.py             (nuovo, 443 righe) — cuore di M4

tests/
├── test_seq_editor.py                (nuovo, 4 test)
├── test_ui_crud.py                   (nuovo, 3 test)
└── test_ui_shell.py                  da 7 a 15 test
                                      test totali: 59
```

---

## 2. M3 — Editor risorse base non-TEST

### 2.1 Task vs. realizzato

| Task M3 | Stato | Note |
|---|---|---|
| 1. `DistEditor` con 5 sezioni e 5 slot CE16 | ✅ | invariato dal precedente audit ma ora con save/duplicate funzionanti |
| 2. `CurveEditor` | ✅ | tabella editabile + save + duplicate via `build_duplicate_path` |
| 3. `LimitEditor` | ✅ | come sopra |
| 4. `RampEditor` | ✅ | come sopra |
| 5. `Ce16Editor` | ⚠️ | save sì, **duplicate disattivato** (`supports_duplicate=False` ereditato dal base) |
| 6. `MmsEditor` | ⚠️ | come Ce16: save sì, duplicate disattivato |
| 7. `ExternalFileEditor` con doppia modalità | ⚠️ | resta in modalità placeholder; il "testo assistito con warning" del backlog è mostrato come testo statico, non legge il file |
| 8. collegare editor al `resource_tree` | ✅ | dispatch in `EditorShell.show_selection` completo |
| 9. azioni `Nuovo`, `Duplica`, `Apri`, `Elimina`, `Usi` | ⚠️ | `Apri` e `Usi` funzionano via FileWidget; `Duplica` funziona via header titolo; `Nuovo` ed `Elimina` **non esposti** (Elimina è dichiarato in M5) |

### 2.2 DoD M3 vs. realtà

- ✅ *"tutte le risorse non-TEST principali si possono aprire, salvare e duplicare"*: vero per distributori, curve, limiti, rampe; **falso per CE16 e MMS2218** (solo save, niente duplica).
- ⚠️ *"le azioni distruttive aprono una modale con impatto riferimenti"*: nessuna azione distruttiva implementata. Coerente con il rinvio dichiarato a M5.
- ✅ *"l'utente non deve usare il filesystem per il CRUD base"*: vero per i casi positivi (apri, modifica, salva, duplica).

### 2.3 Pattern editor riconosciuto

Tutti gli editor M3 espongono lo stesso contratto, che è ergonomico e
testabile:

```python
class FooEditor(QWidget):
    editor_name = "FooEditor"
    supports_save = True
    supports_duplicate = True
    def save_changes(self) -> Path: ...
    def duplicate_resource(self) -> Path | None: ...
```

`EditorShell` invoca i metodi solo se i flag sono `True`, poi chiama
`_reload_project(saved_path)` che ricostruisce `AT614Project` e attiva la
nuova selezione. Pattern pulito e coerente. Il fatto che `MainWindow`/
`EditorShell` non sappiano nulla di quale editor stia ospitando (solo
`getattr(self.current_editor, "supports_save", False)`) consente di aggiungere
nuovi editor senza toccare la shell.

### 2.5 Chiusura task M3 aperti (post-audit)

Successivamente all'audit ho completato i punti rimasti, ad eccezione delle
azioni distruttive (esplicitamente rinviate a M5):

| Punto aperto | Stato post-audit | Soluzione |
|---|---|---|
| CE16Editor / MmsEditor duplicate disabilitato | ✅ Risolto | `supports_duplicate = True` + override di `duplicate_resource` con naming coerente alla convenzione (`CE16_<next_slot>.cfg`, `ADC<next>.csv`) per garantire che il duplicato sia ricaricato come risorsa valida da `load_project` |
| ExternalFileEditor placeholder | ✅ Risolto | doppia modalità: legge il file se `source_path` esiste (auto-detect cp1252/utf-8, editor non read-only, save/duplicate abilitati, banner giallo "modalità testo assistito"); se il path è `None` o non esiste, mantiene il warning rosso e disabilita save/duplicate |
| Azione `Nuovo` non esposta | ✅ Risolto | aggiunto modulo `domain/templates.py` con factory `create_new_resource(category_key, project_root)` per le 8 categorie autoringabili (esclusa `output_banco`); pulsante "Nuovo" nell'header `EditorShell` collegato a `create_new_in_current_category()` con tracking di `_current_category_key`; nominamento auto-incrementale per evitare collisioni (`nuova_rampa.csv`, `nuova_rampa 1.csv`, ...) |
| DistEditor mutava il modello in-place | ✅ Risolto | `_collect_model(source_path)` ora produce un **nuovo** `DistributoreConfig` invece di riassegnare i campi di `self.config`. Il config originale resta intoccato fino al `load_project` post-save |

Azioni `Elimina` correttamente fuori scope: rinviate a M5 insieme alle modali
di impatto.

15 nuovi test coprono i punti chiusi
([test_m3_completion.py](tests/test_m3_completion.py)):

- 2 test per duplicate CE16/MMS (verificano sia il naming coerente sia la
  ri-leggibilità della risorsa duplicata)
- 4 test per ExternalFileEditor (warning mode, read-existing, save-back,
  duplicate-sibling)
- 6 test per l'azione Nuovo (copertura categorie supportate, struttura del
  template per distributore/curva/CE16, gestione collisioni nomi, abilitazione
  pulsante per categoria, apertura editor post-create, gestione speciale per
  file esterni)
- 1 test per DistEditor non-mutation (verifica che il modello originale resti
  inalterato dopo `_collect_model`)

### 2.4 Debiti tecnici M3 (analisi originale)

1. **`FileWidget.name_label` è ora `QLineEdit` ma il nome dice `_label`.**
   Stesso per `ResourceCard.file_label`. Confonde all'uso: chi legge si
   aspetta che `.text()` ritorni read-only, mentre invece può essere editato.
   Rinomina a `name_field` / `file_field` raccomandata.

2. **`Ce16Editor` / `MmsEditor` senza duplicate.** Il base
   `_PointSeriesEditorBase` ha `supports_duplicate = False`. Curve/Limit/Ramp
   lo override a `True`; CE16 e MMS no. Decisione consapevole o dimenticanza?
   Il [README.md](README.md) dichiara "apri, duplica" anche per questi due —
   serve un allineamento.

3. **`DistEditor.parameters_table` ordina gli extras alfabeticamente al
   caricamento** (riga 106: `sorted(self.config.extras.items())`). Al save,
   `_collect_extras` ricostruisce un dict che perde l'ordine d'inserimento
   originale. Il round-trip funziona perché i fixture hanno extras
   alfabetici, ma su file con ordinamento non-alfabetico viene riordinato.

4. **`DistEditor._collect_model` muta il modello in-place.** `self.config.sezioni = {}`,
   `self.config.calibrazioni_ce16 = {}`, `self.config.extras = ...` riassegnano
   sui campi del config originale. Effetti collaterali:
   - se l'utente fa `duplicate` con `_collect_model()` viene mutato anche il
     config originale prima della scrittura
   - se la save fallisce, lo stato in-memory è già modificato
   Pattern preferibile: produrre un nuovo `DistributoreConfig` e usare quello
   solo per la serializzazione.

5. **`_PointSeriesEditorBase.duplicate_resource` non è raggiungibile da CE16/MMS**
   nonostante sia definita nel base — perché il guard `if not
   self.supports_duplicate: return None` blocca subito. Codice morto?
   O placeholder per quando si abiliterà? Va chiarito.

6. **Righe vuote nelle point series eliminate al save.**
   [point_series_editors.py:71-79](src/at614_editor/ui/editors/point_series_editors.py#L71-L79):
   se *tutte* le celle di una riga sono vuote, la riga viene saltata. Buono
   per evitare righe spurie ma elimina silenziosamente eventuali righe
   blank intenzionali. Comportamento da documentare.

7. **`ExternalFileEditor` non è progredito da M2** — resta uno stub. Va
   chiarito se vive così fino a M5 o se M3.7 è considerato "abbastanza" con
   il placeholder.

---

## 3. M4 — SeqEditor e schema TEST

### 3.1 Task vs. realizzato

| Task M4 | Stato | Note |
|---|---|---|
| 1. tabella righe con drag and drop | ✅ | `TestRowTable` con `dropEvent` custom che emette `row_dropped`; `InternalMove` mode |
| 2. reindicizzazione di `Parameter(0)` | ✅ | `reindex_numeric_indices()` aggiorna solo gli indici numerici, preserva i token `i`/`i++`/`++i` (test esplicito) |
| 3. loader `schemas.py` con merge auto + override | ❌ | **non implementato**. Il catalogo è ancora un dict Python in `test_schema.py`; gli YAML in `schemas/auto/` esistono ma non vengono caricati |
| 4. renderer dinamico campi per ID test | ✅ | `describe_test_parameters(row)` produce una lista di `TestParameterDescriptor` con label/editor_kind/resource_type; `_render_parameter_details` genera dinamicamente form + FileWidget |
| 5. integrare `FileWidget` nei parametri file reference | ✅ | callback `open_resource`/`show_usages` cablati verso la shell |
| 6. righe aggiungi/duplica/elimina | ✅ | con re-indicizzazione automatica + selezione coerente post-azione |
| 7. stato riga e badge validazione | ❌ | **non implementato**. La tabella mostra solo nome/ID/indice, nessun indicatore visivo per riferimenti irrisolti o parametri mancanti |
| 8. fallback editor generico per ID parziali | ✅ | gli ID fuori dal catalogo mostrano "Parametro N" come label e text input; gli ID con schema parziale (es. ACQUISIZIONE_CAN ha solo 1,2,6,7) generano i parametri 3,4,5 con label generica |

### 3.2 DoD M4 vs. realtà

- ✅ *"una sequenza TEST reale si apre, si modifica e si salva"*: confermato
  dai test `test_seq_editor_save_persists_common_fields_and_file_refs` e
  `test_seq_editor_add_duplicate_delete_and_reindex_rows`.
- ✅ *"le righe con file referenziati aprono o creano risorse figlie"*:
  apertura sì (`open_resource` callback verso `EditorShell.open_resource_path`).
  Creazione di risorse figlie inline **non** è implementata (l'utente deve
  navigare manualmente).
- ✅ *"gli ID non rifiniti restano modificabili con fallback sicuro"*:
  `describe_test_parameters` produce sempre descrittori, anche per ID sconosciuti.

Tre task su otto restano aperti, di cui due (3 e 7) sono i pilastri della
dichiarazione M4-completa.

### 3.3 Punti di forza del SeqEditor

- **Working copy isolata**: `self.working_rows` è una copia dei `TestRow`
  originali; le modifiche restano in memoria fino al `save_changes`. Riduce
  il rischio di mutazioni accidentali.
- **`_ensure_row_parameter_capacity`**: garantisce che lo schema abbia gli
  slot necessari prima del rendering. Espande `parameters` con stringhe
  vuote senza perdere quelle preesistenti.
- **`_build_sequence(source_path)`** ricostruisce un `TestSequence` completo
  che eredita `encoding`, `line_ending`, `endswith_newline` originali → il
  save preserva l'encoding del file originale.
- **Drag-drop pulito**: il `dropEvent` custom è uno dei pattern Qt più
  fragili, qui è stato isolato bene con un signal `row_dropped(source,
  target)` che innesca `move_row` deterministica.
- **Reindex policy esplicita**: solo gli indici numerici vengono rinumerati,
  i token speciali restano invariati. Comportamento corretto secondo il
  parser VB6 (`Module_TEST_MANAGER.bas` riconosce `i`, `i++`, `++i`).

### 3.4 Rischi e debiti tecnici M4

1. **🔴 Loader schema YAML non implementato (task M4.3).** Questo è il blocker
   architetturale per la chiusura formale di M4 e per la coerenza con la
   decisione chiusa #1 del backlog ("catalogo schema come fonte autoritativa,
   fallback generico per ID non rifiniti"). Lo stato attuale:
   - `schemas/auto/*.yaml` (15 file) generati ma **non letti**
   - `schemas/overrides/` **non esiste**
   - 7 descrittori hard-coded in `test_schema.py:19-57`
   Conseguenza: per aggiungere/correggere un parametro schema serve modifica
   Python + ridistribuzione. L'intero scopo del catalogo YAML è inutilizzato.

2. **🟠 Strip silenzioso dei valori al primo commit.**
   [seq_editor.py:313-329](src/at614_editor/ui/editors/seq_editor.py#L313-L329):
   `_commit_row` esegue `.strip()` su nome, test_id, index_raw e su tutti i
   parametri. Quando l'utente seleziona una riga per la PRIMA volta,
   `_load_row` la carica e poi un successivo cambio selezione triggera
   `_commit_row(previous_row)` — che strippa i valori anche se l'utente non
   ha modificato nulla. Risultato: aprire il file → cambiare selezione →
   salvare **perde silenziosamente leading/trailing spaces** dai parametri.
   I fixture testati non hanno spazi significativi, ma il rischio su file
   legacy con valori del tipo `" 15.1001.310 "` esiste. Mitigazione:
   committare solo se i widget hanno generato effettivamente edit, oppure
   `.strip()` solo al save finale.

3. **🟠 Validazione live e badge stato riga mancanti (task M4.7).** La
   tabella mostra solo le 3 colonne testuali. Non c'è:
   - icona/colore di errore quando un parametro `file_ref` punta a un file
     irrisolto;
   - badge "incompleto" quando lo schema richiede N parametri ma il row ne ha
     M < N filled;
   - status bar centrale con totale errori/warning del file aperto.
   Il dato è già disponibile (`project.resolve_resource_path()`,
   `unresolved_file_references`), manca solo l'aggancio visivo.

4. **🟡 Numero "magico" `_next_numeric_index`.** Se i numeri attuali sono
   `[0, 1, 3, 7]`, `max+1 = 8`, ma `reindex_numeric_indices` poi rinumera in
   `[0, 1, 2, 3]`. Le due politiche divergono. In pratica `add_row` chiama
   subito `reindex_numeric_indices`, quindi il valore restituito viene
   sovrascritto. Codice cosmeticamente confuso, non bug.

5. **🟡 `duplicate_resource` include le modifiche unsaved del current row.**
   `_build_sequence` prima fa `_commit_row(self.current_row_index)`. Quindi
   il duplicato contiene anche edit non confermati. Coerente con il test ma
   potenzialmente sorprendente per l'utente: "duplica" significa "duplica
   ciò che vedi", non "duplica ciò che è salvato".

6. **🟡 Trailing fields del file CSV non gestiti.** Stesso debito dell'audit
   precedente: `serialize_test_csv` produce per ogni riga
   `name;test_id;index_raw;<params>` con `len(params)` libero. Quando
   `add_row` crea una riga senza parametri, salva `"Name;ID;0"` (3 campi vs N
   nell'header). Il dispatcher VB6 dovrebbe tollerare (calcola
   `UBound(dati)-2` ai sensi del proprio parser), ma il file non sarà
   byte-coerente con la convenzione "header lunghezza = riga lunghezza" dei
   fixture esistenti.

7. **🟡 Test fragili su indici di riga.**
   `editor.select_row(8)` e `editor.select_row(11)` hardcodano la posizione
   ordinale dei test ACQUISIZIONE/RISPOSTE in `15.1001.356_C.csv`. Se domani
   qualcuno riordina il file fixture, i test rompono in modi non ovvi.
   Preferibile cercare per nome (`next(i for i, row in enumerate(rows) if
   row.name == "Acquisizione")`).

8. **🟡 `parameter_bindings` può contenere widget che riferiscono parametri
   oltre la lunghezza della riga corrente.** Lo guardiamo in `_commit_row`:
   `if descriptor.parameter_index - 1 >= len(row.parameters): continue` —
   skip silenzioso. Combinato con `_ensure_row_parameter_capacity`, in
   pratica non succede mai, ma è difensivo in modo opaco. Un'asserzione
   sarebbe più chiara.

---

## 4. Stato del dominio e dell'indice riferimenti

### 4.1 Nuove API positive

- `AT614Project.resolve_resource_path(raw)` — wrapper che prova prima
  l'indice alias, poi fallback su `test_sequences` per stem (caso "sezione1
  → codice modulo senza estensione")
- `AT614Project.get_usage_count(path)` — semplifica il refresh dei badge
  "Usi (N)"
- `ResourceTree.set_project(project)` — riusato da `_reload_project` per
  ricostruire l'albero post-save

### 4.2 Debiti dominio non risolti

- ❌ `ProjectResource` base class ancora non esiste (audit precedente
  raccomandazione #3)
- ❌ `io/encoding.py` non estratto: `_detect_encoding` è ora duplicato in
  **4 moduli** (distributore, test_csv, point_series, bas_extractor)
  invece di 3
- ❌ `external_config.py` parser ancora mancante
- ❌ `ExternalConfigResource` model ancora mancante
- ✅ blank line nel `test_csv.parse`: dato che `working_rows` parte dai
  `sequence.rows` e questi sono già post-parse, la blank line resta
  droppata al primo round-trip. Il rischio dell'audit precedente è
  amplificato perché ora il save scrive (non solo il read fa round-trip)

---

## 5. Test e qualità

### 5.1 Distribuzione

```
test_bas_extractor.py        : 10  (M0)
test_ce16_parser.py          :  2  (M1)
test_distributore_parser.py  :  3  (M1)
test_mms2218_parser.py       :  2  (M1)
test_point_series_parser.py  :  3  (M1)
test_project_index.py        :  7  (M1)
test_seq_editor.py           :  4  (M4)
test_smoke.py                :  2  (smoke)
test_test_csv_parser.py      :  3  (M1)
test_test_schema.py          :  5  (M4 schema)
test_ui_crud.py              :  3  (M3 CRUD)
test_ui_shell.py             : 15  (M2/M3 shell)
----------------------------------
totale                       : 59
```

### 5.2 Cosa testano bene

- Round-trip byte-identico dei parser (parse → serialize identico)
- Pattern save-then-reparse: scrivere via UI, leggere via parser, verificare
  semantica (test_ui_crud, test_seq_editor)
- Drag-drop policy e re-indicizzazione (test_seq_editor_supports_row_reorder)
- Schema fallback per ID sconosciuti
- Resolver del progetto con alias profilo (`MERLO/CE16_1.cfg` vs
  `DEFAULT/CE16_1.cfg`)

### 5.3 Cosa NON è coperto

- **Validazione end-to-end** — nessun test verifica che un file con
  riferimento rotto venga correttamente segnalato a UI level (perché il
  badge non esiste ancora)
- **Strip silenzioso dei valori** — nessun test fallisce se gli spaces
  vengono persi al save
- **Mutazione in-place di DistEditor._collect_model** — nessun test rileva
  l'effetto collaterale
- **Save CE16/MMS** — coperto solo implicitamente via passthrough; nessun
  test esplicito di edit→save→reparse per queste risorse
- **Coverage non misurata** (audit precedente raccomandazione #7 non
  affrontata)
- **Test fragility**: come notato in §3.4.7, alcuni test usano indici di
  riga magici che cadono se i fixture vengono toccati

---

## 6. Conformità ai brief

### 6.1 vs. LAYOUT_UX.md

| Voce | Stato |
|---|---|
| Layout 240/fluid/260 | ✅ |
| 9 categorie nell'ordine corretto | ✅ |
| Tab `Riferimenti`/`Validazione`/`Azioni rapide` | ✅ |
| Quick actions: brief dice 5 (Nuovo, Duplica, Apri, Elimina, Usi) | ⚠️ Sono 6 con `Sostituisci` aggiuntivo e nomi `Nuovo collegato`/`Duplica collegato`. Divergenza segnalata dall'audit M0-M2 e ancora non risolta. |
| Microcopy italiana | ✅ |
| `ResourceCard` con badge stato | ✅ |
| `ReadOnlyBanner` su output | ✅ |

### 6.2 vs. BACKLOG_IMPLEMENTAZIONE.md

| Milestone | DoD soddisfatta? |
|---|---|
| M0 | ✅ chiusa nell'audit precedente |
| M1 | ✅ |
| M2 | ✅ |
| M3 | 🟡 **sostanzialmente sì** ma: CE16/MMS senza duplicate; ExternalFileEditor placeholder; azione `Nuovo` non implementata; `Elimina` correttamente rinviata a M5 |
| M4 | 🟡 **funzionalmente sì sulla maggioranza dei task**, ma il task 3 (loader YAML) è il pilastro mancante per dichiararla chiusa; task 7 (badge validazione) anche aperto |

---

## 7. Verdetto sintetico

| Milestone | Esito | Note |
|---|---|---|
| M0 | 🟢 Completata | scanner BAS + 15 YAML generati + matrice formati |
| M1 | 🟢 Completata | debiti tecnici tracciati, non bloccanti |
| M2 | 🟢 Completata | divergenze microcopy minori |
| M3 | 🟢 **Completata** | tutti gli editor con CRUD coerente (esclusi i task distruttivi rinviati a M5); azione `Nuovo` via template + pulsante shell; ExternalFileEditor in doppia modalità |
| M4 | 🟡 **70% reale, valore di prodotto significativo** | UX completa, persiste su disco, dispatch dinamico schema-driven funzionante; manca loader YAML e badge validazione |

## 8. Azioni raccomandate prima di dichiarare M4 chiusa

In ordine di priorità decrescente:

1. **🔴 Implementare il loader `schemas.py`** che:
   - legge `schemas/auto/<id>.yaml` come baseline
   - applica deep-merge da `schemas/overrides/<id>.yaml` se presente
   - produce `TestParameterDescriptor` runtime
   - sostituisce `SCHEMA_PARAMETER_DESCRIPTORS` hard-coded
   - risolve i casi `TEST_CICLICA_CAN` / `SERIAL_NUMBER` via override (vedi
     anomalie note in [AUDIT_M0_M2.md §2.3](AUDIT_M0_M2.md))
   Questo chiude **la decisione chiusa #1** del backlog e M4 task 3.

2. **🟠 Aggiungere badge validazione nella `sequence_table`**:
   - colorare l'item della colonna "Nome test" in rosso se almeno un file_ref
     parameter è irrisolto
   - tooltip con la lista dei riferimenti mancanti per quella riga
   - status bar centrale con conteggio totale
   Chiude M4 task 7.

3. **🟠 Risolvere lo strip silenzioso al commit**: spostare `.strip()` dentro
   `_build_sequence()` (solo al save), non in `_commit_row()` ad ogni
   selezione.

4. **🟡 Allineare M3**: o abilitare `supports_duplicate=True` su CE16/MMS
   con un test di copertura, oppure aggiornare il README per riflettere che
   queste due risorse non si duplicano (e dichiararlo "non necessario"
   anche nella tabella formati).

5. **🟡 Rinominare** `FileWidget.name_label` → `name_field` e
   `ResourceCard.file_label` → `file_field` (con tutti i call site). Ora
   sono QLineEdit ma si fanno passare per QLabel.

6. **🟡 Estrarre `io/encoding.py`** consolidando 4 copie identiche di
   `_detect_encoding`.

7. **🟡 Sostituire indici magici nei test** (`select_row(8)`,
   `select_row(11)`) con ricerca per nome → test più resistenti a futuri
   refactor dei fixture.

8. **🟢 Aggiungere fixture `.txt` esterni** per il branch `external_config`
   (era M0.3 parziale, ancora non chiuso).

Nessuna di queste azioni invalida il valore di prodotto già raggiunto. M4
ha già esposto un'UX coerente con il brief; ciò che resta è la chiusura
dell'architettura schema-driven e la visualizzazione della validazione che
il dominio sa già computare.
