\# AT614 Configuration Editor — Design Brief operativo



Documento di riferimento per qualunque LLM (o designer / sviluppatore) chiamato

a produrre o evolvere il mock-up dell'AT614 Configuration Editor.



Questo brief \*\*non sostituisce\*\* `PROPOSTA\_FLUSSO\_GUI\_UX.md`. Quel documento

descrive \*cosa\* il software deve fare. Questo descrive \*come\* deve apparire,

quali decisioni di UI sono già state prese e non vanno rinegoziate, e quali

sono ancora aperte.



Se questo brief e il documento di proposta divergono, vince \*\*questo\*\*.



\---



\## 0. Come usare questo documento



\- Sezione 1–2: contesto rapido, da leggere sempre.

\- Sezione 3: regole invarianti — non violare.

\- Sezione 4–6: vocabolario UI riutilizzabile (componenti, stati, microcopy).

\- Sezione 7: mappa risorsa → editor.

\- Sezione 8: anti-pattern.

\- Sezione 9: decisioni operative chiuse — applicale senza ridefinirle.



Per generare un mock-up: leggi 1, 3, 4, 5, 7, 8.

Per scrivere codice di produzione: leggi tutto.



\---



\## 1. Contesto in 5 righe



\- Software desktop Windows, \*\*solo italiano\*\*.

\- Utente: tecnico di collaudo / ufficio prove, \*\*non programmatore\*\*.

\- Scopo: creare / modificare programmi di collaudo per banco EOL.

\- Un programma di collaudo è un \*\*insieme di risorse collegate\*\*

&#x20; (distributore, sequenze test, curve, rampe, limiti, calibrazioni, file esterni).

\- Tono: \*\*strumento tecnico-industriale guidato\*\*, non IDE, non consumer.



\---



\## 2. Modello mentale delle risorse



```

Distributore (DST)              ← punto di ingresso del programma

&#x20;├─ Sezione 1 → Sequenza TEST   ← fino a 5 sezioni

&#x20;├─ Sezione 2 → Sequenza TEST

&#x20;├─ …

&#x20;└─ Calibrazioni CE16 (0–5)     ← OPZIONALI

&#x20;      e parametri generali



Sequenza TEST (TST)             ← lista ordinata di test

&#x20;└─ Riga test

&#x20;    ├─ ID test (determina i campi del dettaglio)

&#x20;    └─ Riferimenti a file: Curva CMD, Curva LIM, Rampa XY,

&#x20;                           CAL CE16, CAL MMS2218, file esterni



Output banco                    ← NON è configurazione, è output

&#x20;├─ Graph

&#x20;└─ Rampe XY Last

```



Regole conseguenti:



\- Il distributore è la radice logica del programma, non un file qualunque.

\- Una risorsa figlia può essere \*\*condivisa\*\* tra più sequenze / distributori

&#x20; → ogni editor deve poter rispondere alla domanda \*"chi mi usa?"\*.

\- CE16 è opzionale: l'assenza non è un errore, è uno stato esplicito.

\- Output banco è \*\*sempre e solo\*\* in sola lettura nel flusso principale.



\---



\## 3. Regole invarianti (NON violare)



1\. \*\*L'albero risorse è per tipologia funzionale\*\*, mai per filesystem reale.

&#x20;  Le 9 categorie fisse, in quest'ordine:

&#x20;  Distributori · Sequenze test · Curve comando · Curve limite · Rampe XY ·

&#x20;  Calibrazioni CE16 · Calibrazioni MMS2218 · File esterni · Output banco.

2\. \*\*Layout 3 colonne\*\* in ogni schermata di editing:

&#x20;  `albero (≈240px) | editor (fluid) | pannello contestuale (≈260px)`.

&#x20;  Dashboard e wizard sono eccezioni esplicite.

3\. \*\*Pannello destro a 3 sezioni fisse\*\*: \*Riferimenti\*, \*Validazione\*,

&#x20;  \*Azioni rapide\*. Stesso ordine ovunque.

4\. \*\*Mai esporre path filesystem\*\* nel flusso principale.

&#x20;  I file si scelgono per nome logico o tramite dialog.

5\. \*\*Ogni azione distruttiva\*\* (Elimina, Sostituisci, sovrascrivi) passa per

&#x20;  una modale che mostra \*chi usa la risorsa\* prima di confermare.

6\. \*\*Output banco\*\*: nessun pulsante `Salva`, `Nuovo`, `Modifica`. Solo

&#x20;  `Esporta`, `Confronta`, `Stampa`, `Apri cartella`.

7\. \*\*Lingua: italiano\*\*, sempre. Non mescolare con inglese, neanche nei

&#x20;  placeholder. Eccezione: codici file e ID test (es. `T\_CRV\_12`).

8\. \*\*CE16 ha 5 slot fissi\*\*, sempre visibili. Slot vuoti sono mostrati come

&#x20;  "non configurato", mai nascosti.

9\. \*\*Distributore ha 5 sezioni\*\*, sempre visibili. Sezioni non attive sono

&#x20;  mostrate disabilitate con pulsante "Attiva sezione", mai nascoste.

10\. \*\*Tema chiaro\*\* come default. Non proporre dark mode in v1.



\---



\## 4. Vocabolario di componenti riutilizzabili



Questi componenti hanno un nome. Usalo. Non re-inventarli.



\### 4.1 `FileWidget` — selettore file con 5 azioni



Usato \*\*ovunque\*\* un campo punta a un file (sequenza, curva, rampa, limite,

CE16, MMS2218, file esterno).



```

┌─────────────────────────────────────────────────────────────────┐

│ \[icona tipo] NOME\_FILE.ext        \[Seleziona]\[Nuovo]\[Duplica]\[Apri]\[Usi]│

└─────────────────────────────────────────────────────────────────┘

```



\- \*\*Seleziona\*\* → apre dialog di scelta da risorse esistenti dello stesso tipo

\- \*\*Nuovo\*\* → apre mini-dialog con nome suggerito, poi entra subito nell'editor

\- \*\*Duplica\*\* → clona il file referenziato, lo collega, apre editor

\- \*\*Apri\*\* → apre l'editor della risorsa referenziata

\- \*\*Usi\*\* → mostra in popup chi altro usa quella risorsa (numero in badge: `Usi (4)`)



Stato vuoto: testo `— riferimento mancante —` in rosso, mostra solo

`Seleziona` + `Nuovo` + `Duplica da …` (se contestualmente sensato).



\### 4.2 `ResourceCard` — card di sezione / modulo



Usata per ogni sezione del distributore e in altri elenchi compatti.



```

┌──────────────────────────┐

│ Sezione N      \[semaforo]│

│ NOME\_FILE.tst            │

│ 14 test · note breve     │

│ \[Sel]\[Nuovo]\[Dup]\[Apri]\[Usi]

└──────────────────────────┘

```



\### 4.3 `RPanel` — pannello destro contestuale



3 tab fisse: \*\*Riferimenti\*\* · \*\*Validazione\*\* · \*\*Azioni rapide\*\*.



\- Tab attiva di default = quella più rilevante al contesto (di solito

&#x20; \*Riferimenti\*; \*Validazione\* se ci sono errori bloccanti).

\- \*Azioni rapide\* contiene SEMPRE almeno:

&#x20; `Nuovo collegato` · `Duplica collegato` · `Sostituisci` · `Apri` · `Elimina` · `Usi`



\### 4.4 `Stepper` — wizard



Step numerati, separati da una linea tratteggiata. Stati:

`done` (✓ verde) · `on` (numero in giallo evidenziato) · `idle` (numero grigio).

Click su step `done` = torna a quello step. Step `idle` non cliccabili.



\### 4.5 `ImpactList` — lista impatto eliminazione



Usata in ogni modale di eliminazione / sostituzione. Mostra elenco di

file+riga che referenziano la risorsa, con badge rosso `romperà` accanto.

Max 200px di altezza, scrollabile.



\### 4.6 `ChartStage` — grafico schizzato



Componente grafico standard per curve, limiti, rampe, output banco.



\- Titolo in alto a sinistra, legenda in alto a destra.

\- Toolbar locale \*\*sopra\*\* il grafico, non dentro.

\- Punto selezionato evidenziato in giallo, punto in errore in rosso.

\- Per \*curve limite\*: tasto `Mostra opposta` per overlay tratteggiato.

\- Per \*rampe\*: tasti `Snap griglia` + `Inserisci punto intermedio`.

\- Per \*output banco\*: nessuna interazione di edit, solo zoom/pan/export.



\### 4.7 `ReadOnlyBanner`



Banner orizzontale ambra/sabbia in cima a qualsiasi viewer di output banco.

Testo: \*"Sola lettura — questo file è un output di banco. Non è un file di

configurazione."\* Sempre con icona lucchetto.



\---



\## 5. Stati visivi — tassonomia chiusa



Esattamente 7 stati. Non inventarne altri.



| Stato                          | Simbolo | Colore  | Quando                                |

|--------------------------------|---------|---------|---------------------------------------|

| salvato / ok                   | ✓       | verde   | file su disco coincide con memoria    |

| modificato non salvato         | ●       | giallo  | edit in memoria non ancora persistito |

| warning non bloccante          | !       | ambra   | validazione produce avvertimento      |

| errore bloccante               | ✕       | rosso   | il programma non è eseguibile         |

| opzionale non configurato      | opz.    | grigio  | risorsa opzionale lasciata vuota      |

| condivisa da più file          | i       | blu     | usata da ≥2 referenziatori            |

| sola lettura (output banco)    | 🔒      | ambra   | risorsa di output                     |



Convenzioni:



\- Il \*\*badge\*\* appare nel nodo dell'albero, nella card, nella titlebar

&#x20; dell'editor e nella riga della tabella test.

\- Gli stati possono coesistere (es. `●` + `!`). Mostra entrambi i badge.

\- I colori esatti sono definiti in §6.



\---



\## 6. Direzione visiva



\### Palette (lo-fi wireframe e hi-fi devono restare riconoscibili)



| Ruolo            | Hex        |

|------------------|------------|

| Sfondo carta     | `#f6f2e7`  |

| Sfondo elemento  | `#fffdf6`  |

| Inchiostro       | `#1b1a17`  |

| Inchiostro soft  | `#3a3833`  |

| Tratto regola    | `#2a2722`  |

| Stato rosso      | `#c44a3a`  |

| Stato verde      | `#3f7a4a`  |

| Stato ambra      | `#c98a2b`  |

| Stato blu        | `#3a6098`  |

| Highlight giallo | `#fff5a8`  |



In hi-fi: sfondo si schiarisce verso bianco `#fafaf7`, gli stessi accenti

vengono mantenuti (eventualmente +5% saturazione). Niente gradienti

decorativi. Niente glassmorphism.



\### Tipografia



\- Hi-fi: una sans neutra industriale (es. \*Inter Tight\*, \*IBM Plex Sans\*,

&#x20; \*Source Sans 3\*). Non usare Roboto/Arial generici.

\- Mono per nomi file e ID: \*JetBrains Mono\* o \*IBM Plex Mono\*.

\- Lo-fi: si conserva l'aspetto handwritten (\*Caveat\* + \*Kalam\*).



\### Densità



\- Touch target ufficio: minimo \*\*28×28 px\*\*.

\- Padding interno card: 12–16 px.

\- Tabelle test: righe 32–36 px, mai più strette.

\- Gap default tra elementi: 8 px (tight), 14 px (default), 22 px (sezione).



\### Iconografia



\- Set monocromatico, tratto 1.5–2 px, geometrico.

\- Le 4 forme di tipo risorsa sono fisse e non vanno sostituite con altre:

&#x20; - □ Distributore / file esterno / CE16 / MMS2218

&#x20; - ▲ Sequenza TEST

&#x20; - ○ Curva (comando o limite)

&#x20; - ◇ Rampa XY

\- Niente emoji nella UI di produzione (eccezione: lucchetto 🔒 sola lettura

&#x20; se non si ha alternativa pittografica).



\---



\## 7. Mappa risorsa → editor



| Risorsa            | Editor                | Layout specifico                                  |

|--------------------|-----------------------|---------------------------------------------------|

| Distributore       | `DistEditor`          | 5 ResourceCard moduli + CE16 grid (5) + Params    |

| Sequenza TEST      | `SeqEditor`           | Tabella test + dettaglio dinamico per ID test     |

| Curva comando      | `CurveEditor`         | ChartStage + tabella punti + toolbar locale       |

| Curva limite       | `LimitEditor`         | come Curve + toggle "Mostra opposta INF/SUP"      |

| Rampa XY           | `RampEditor`          | come Curve + Snap + punto intermedio              |

| CE16               | `Ce16Editor`          | ChartStage + tabella punti + badge "opzionale"    |

| MMS2218            | `MmsEditor`           | come CE16 + range / ordinamento ADC               |

| File esterno       | `ExternalFileEditor`  | tabella chiave/valore o testo assistito           |

| Output Graph       | `OutputViewer`        | ReadOnlyBanner + ChartStage + riassunto metriche  |

| Output Rampe Last  | `OutputViewer`        | come sopra                                        |



Ogni editor ha:



\- una \*\*titlebar\*\* con icona tipo + nome + sottotitolo + pill stato validazione

\- pulsanti \*\*Duplica\*\* e \*\*Salva\*\* sulla destra della titlebar

\- la `RPanel` a destra coerente con la risorsa attiva



Il \*\*dettaglio dinamico per ID test\*\* della sequenza è il punto più

delicato: i campi mostrati dipendono dall'ID. Vedi open question §9.1.



\---



\## 8. Anti-pattern (cosa NON fare)



\- ❌ Look da IDE (sidebar nera, badge a forma di branch, terminale integrato).

\- ❌ Look da app creativa (gradienti, blur, illustrazioni decorative).

\- ❌ Dark mode in v1.

\- ❌ Path filesystem nel flusso (es. `C:\\PROGRAMMI\\…` visibile all'utente).

\- ❌ Inglese mescolato nella UI (`Save`, `Cancel`, `Delete`).

\- ❌ Eliminazioni senza modale di impatto.

\- ❌ Output banco con qualsiasi pulsante di scrittura.

\- ❌ Nascondere sezioni o slot CE16 vuoti (vanno mostrati disabilitati).

\- ❌ Tooltip come unico veicolo di un'informazione critica.

\- ❌ Pulsanti senza etichetta testuale, eccetto le 4 forme icona tipo risorsa.

\- ❌ Inventare nuovi stati oltre i 7 di §5.

\- ❌ Layout a 2 colonne negli editor (l'RPanel destra è obbligatoria).

\- ❌ Wizard senza step "riepilogo file che verranno creati" prima della conferma.



\---



\## 9. Decisioni operative chiuse



\### 9.1 — Dettaglio dinamico per ID test

Decisione:

- la fonte autoritativa del dettaglio dinamico è il catalogo schema
	`{ID test → schema campi}` derivato dal merge `schemas/auto + schemas/overrides`;
- il `SeqEditor` mostra sempre i campi comuni `Nome test`, `ID test`, `Indice`;
- per gli ID presenti nel catalogo iniziale del piano, i campi vengono renderizzati
	in ordine con il tipo definito nello schema;
- se uno schema è parziale, i campi noti sono tipizzati e i restanti vengono
	mostrati come `Parametro N` in fallback generico;
- se un ID non ha ancora schema rifinito, il `SeqEditor` usa un editor generico
	testuale che preserva ordine e cardinalità dei parametri.



\### 9.2 — Comportamento di `Usi`

Decisione:

- `Usi (0)` è disabilitato;
- click su `Usi (n)` apre un popover inline se `n <= 10`;
- se `n > 10`, il click attiva la tab *Riferimenti* della `RPanel` e mette a
	fuoco la lista completa;
- ogni voce della lista apre direttamente la risorsa o la riga test che la
	referenzia.



\### 9.3 — Rename a cascata

Decisione:

- `Rinomina` su una risorsa è sempre una rinomina globale della risorsa stessa;
- prima della conferma compare una modale con l'impatto sui riferimenti;
- la conferma aggiorna automaticamente tutti i riferimenti compatibili nello
	stesso progetto in modo atomico;
- l'opzione "rinomina solo qui" non esiste in v1;
- se l'utente vuole divergere solo in un punto, il flusso corretto è
	`Duplica` + `Sostituisci` sulla referenza desiderata.



\### 9.4 — CE16 parziale dopo wizard

Decisione:

- gli slot CE16 non compilati restano in stato `opz.`;
- il testo mostrato è `— non configurato —`;
- non vengono convertiti in warning o errore solo perché sono vuoti.



\### 9.5 — File esterno non riconosciuto

Decisione:

- `ExternalFileEditor` ha due modalità:
	- vista strutturata se il formato è riconosciuto come chiave/valore o tabella;
	- editor testo assistito se il formato non è riconosciuto;
- in fallback testuale il file resta modificabile;
- il fallback mostra un warning esplicito che la modifica assistita non è
	disponibile;
- vanno preservati encoding, line ending e contenuto non interpretato.



\### 9.6 — Cestino del progetto

Decisione:

- in v1 non esiste un cestino del progetto;
- `Elimina` è definitiva dopo modale di impatto e conferma esplicita;
- la strategia di sicurezza in v1 è data da `Duplica`, `Sostituisci` e dalla
	modale di impatto, non da un flusso di ripristino.



\### 9.7 — Confronto banda / output

Decisione:

- il pulsante `Confronta` apre sempre un picker modale filtrato per la stessa
	tipologia di risorsa;
- per le curve limite INF/SUP resta disponibile anche l'azione dedicata
	`Mostra opposta`;
- il drag\&drop sul grafico non è previsto in v1;
- dopo la selezione, il confronto si apre nello stesso editor/viewer come overlay.



\### 9.8 — Validazione: quando?

Decisione:

- validazione live con debounce durante la modifica: produce warning e feedback
	locali, non blocca il lavoro;
- al `Salva`: validazione bloccante della risorsa attiva e dei riferimenti diretti
	necessari a garantirne la coerenza;
- al click su `Validazione`: scansione completa del programma o della risorsa
	attiva, con report esplicito nel pannello `Validazione`.



\---



\## 10. Microcopy italiana di riferimento



Tutta la UI deve usare ESATTAMENTE queste etichette. Non sinonimi.



\### 10.1 Pulsanti



| Funzione           | Etichetta da usare        | Da NON usare                   |

|--------------------|---------------------------|--------------------------------|

| crea da zero       | `Nuovo`                   | Crea, Aggiungi, Add            |

| crea programma     | `Nuovo programma`         | Nuovo collaudo, Crea programma |

| clona              | `Duplica`                 | Copia, Clona                   |

| apri editor        | `Apri`                    | Modifica, Edit                 |

| sostituisci rif.   | `Sostituisci`             | Cambia, Sposta                 |

| chi usa            | `Usi`                     | Riferimenti, Used by           |

| salva              | `Salva`                   | Conferma, Applica              |

| salva tutto        | `Salva tutto`             | Salva progetto                 |

| annulla            | `Annulla`                 | Indietro, Esci                 |

| elimina            | `Elimina`                 | Cancella, Rimuovi              |

| validazione        | `Validazione`             | Controlla, Verifica            |

| seleziona file     | `Seleziona`               | Scegli, Sfoglia (eccetto dialog) |

| sfoglia (dialog)   | `Sfoglia…`                | Browse, Apri…                  |

| esporta            | `Esporta`                 | Salva come                     |



\### 10.2 Stati nei badge



`ok` · `non salvato` · `warning` · `errore` · `opz.` · `condiviso` · `sola lettura`



\### 10.3 Frasi standard



\- Modale elimina: \*"Elimino NOME? Questa risorsa è referenziata da N righe in M sequenze TEST. Eliminarla senza sostituirla lascerà riferimenti mancanti."\*

\- Banner sola lettura: \*"Sola lettura — questo file è un output di banco. Non è un file di configurazione."\*

\- Wizard riepilogo: \*"Riepilogo file che verranno creati"\* + tree con marker `\[nuovo]` / `\[duplicato da …]` / `\[riusato]`.

\- Riferimento mancante: \*"— riferimento mancante —"\* (rosso).

\- CE16 vuoto: \*"— non configurato —"\* (grigio).



\---



\## 11. Output che ci si aspetta dall'LLM



Quando passi questo brief a un LLM di design, chiedi:



1\. Un \*\*HTML statico\*\* (o jsx) con le 7 schermate di §7 + Dashboard + Wizard,

&#x20;  navigabili tramite tab in alto.

2\. Vocabolario di componenti di §4 effettivamente riutilizzato (non

&#x20;  ridisegnato ogni volta).

3\. Stati di §5 visivamente coerenti tra schermate.

4\. Microcopy esatta di §10.

5\. Una nota a margine dove l'LLM ha dovuto fare scelte non coperte dal

&#x20;  brief (servirà per aggiornarlo).



Non chiedere:



\- Hi-fi rendering finale prima di aver validato la struttura.

\- Animazioni complesse.

\- Interazioni reali con file system.

\- Logica di business (validazione vera, parsing file, ecc.).



\---



\## 12. Glossario rapido



\- \*\*DST\*\* — file distributore, punto di ingresso del programma di collaudo.

\- \*\*TST / SEQ\*\* — sequenza TEST, ordinata, lista di righe test.

\- \*\*CRV\*\* — curva comando (input pilotato al banco).

\- \*\*LIM\*\* — curva limite (banda di accettazione, INF e SUP).

\- \*\*XY\*\* — rampa tempo/tensione punti trascinabili.

\- \*\*CE16\*\* — sensore opzionale, fino a 5 calibrazioni per distributore.

\- \*\*MMS2218\*\* — calibrazione ADC del profilo banco.

\- \*\*Graph / Rampe XY Last\*\* — cartelle di output banco, sola lettura.

\- \*\*Programma di collaudo\*\* — l'insieme `DST + sue risorse referenziate`.



\---



\*Fine del brief. Per qualunque modifica strutturale: aggiornare prima

questo documento, poi i mock-up.\*



