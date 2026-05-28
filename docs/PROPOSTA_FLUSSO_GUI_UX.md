# Proposta Flusso GUI/UX per Mock-up LLM

Questo documento descrive il flusso della GUI proposto per l'AT614
Configuration Editor ed e pensato per essere dato in input a un LLM di design
per ottenere un primo mock-up UX coerente con il piano di sviluppo.

Per layout invarianti, componenti nominati, microcopy, tassonomia degli stati e
direzione visiva fa fede [LAYOUT_UX.md](LAYOUT_UX.md). Questo documento resta
complementare e descrive soprattutto flussi, schermate e comportamento atteso.

---

## 1. Obiettivo del mock-up

Il mock-up deve rappresentare un'applicazione desktop Windows in italiano,
orientata a utenti non programmatori che devono creare o modificare programmi di
collaudo partendo da file esistenti o da template.

Il mock-up non deve sembrare un IDE o un editor generico di testo. Deve sembrare
uno strumento operativo di configurazione tecnica, guidato, leggibile e sicuro.

Obiettivi UX principali:

- permettere di creare un nuovo programma di collaudo senza toccare i file a mano;
- rendere visibili le relazioni tra distributore, test e file collegati;
- offrire editor dedicati per le risorse grafiche e configurabili;
- prevenire errori prima del salvataggio;
- rendere semplici le operazioni di duplica, rinomina, sostituzione ed eliminazione.

---

## 2. Utente target

Utente principale:

- tecnico di collaudo;
- tecnico ufficio prove;
- persona con conoscenza del dominio ma non necessariamente di VB6 o della
  struttura interna dei file.

Competenze attese:

- conosce il significato funzionale di distributore, curva, rampa, limite,
  calibrazione CE16, MMS2218 e sequenza test;
- non vuole ricordare percorsi file o convenzioni interne;
- si aspetta pulsanti chiari, nomi italiani e flussi guidati.

---

## 3. Concetti UX da rendere espliciti

Il mock-up deve comunicare chiaramente questi concetti:

- il programma di collaudo e un insieme di risorse collegate;
- il file DISTRIBUTORE e il punto di ingresso logico del programma;
- ogni sezione del distributore punta a una sequenza TEST;
- le righe TEST possono aprire o creare curve, rampe, limiti e altri file;
- CE16 e una risorsa opzionale;
- gli output banco possono vivere in uno o piu percorsi configurabili, anche
  esterni al progetto; la struttura tipica e una cartella per salvataggio o
  prodotto con un file CSV grafico al suo interno;
- questi output sono solo consultabili, non modificabili nel flusso principale.

---

## 4. Architettura della GUI

### 4.1 Struttura generale

Applicazione desktop con layout a tre colonne nelle schermate di editing:

- barra superiore con azioni globali;
- colonna sinistra con albero risorse del progetto;
- area centrale con editor della risorsa selezionata;
- pannello destro con riferimenti, validazione e azioni contestuali.

Eccezioni esplicite: dashboard iniziale e wizard.

### 4.2 Barra superiore

Azioni principali sempre visibili:

- Nuovo programma
- Duplica programma
- Apri risorsa
- Salva
- Salva tutto
- Validazione
- Cerca

Azioni secondarie:

- Impostazioni
- Aiuto
- Visualizza output
- Configura archivio output

### 4.3 Albero risorse a sinistra

L'albero non deve mostrare il filesystem grezzo. Deve essere organizzato per
tipologia funzionale:

- Distributori
- Sequenze test
- Curve comando
- Curve limite
- Rampe XY
- Calibrazioni CE16
- Calibrazioni MMS2218
- File esterni
- Output banco

Ogni nodo deve mostrare:

- nome risorsa;
- icona tipologia;
- eventuale stato errore/warning;
- eventuale marker di file modificato non salvato.

Per Output banco non va caricato un albero completo all'avvio: la UX deve
aprire un browser di ricerca con filtri e risultati caricati in modo lazy.

### 4.4 Area centrale

L'area centrale cambia in base al tipo risorsa selezionata.

Tipi di contenuto principali:

- dashboard iniziale;
- editor distributore;
- editor sequenza test;
- editor curva comando;
- editor curva limite;
- editor rampa;
- editor CE16;
- editor MMS2218;
- editor file esterno;
- viewer output banco.

### 4.5 Pannello destro

Pannello sempre contestuale con tre schede o sezioni fisse, sempre nello stesso
ordine:

- Riferimenti
- Validazione
- Azioni rapide

Contenuto tipico:

- chi usa questa risorsa;
- quali file sono referenziati da questa risorsa;
- errori e warning;
- pulsanti rapidi Nuovo collegato, Duplica collegato, Sostituisci, Apri,
  Elimina, Usi.

---

## 5. Schermate da includere nel mock-up

### 5.1 Home / Dashboard progetto

Scopo:

- fornire un punto di ingresso semplice;
- far scegliere rapidamente se creare, duplicare o aprire un programma.

Elementi UI:

- titolo applicazione;
- card grandi con:
  - Nuovo programma di collaudo
  - Duplica programma esistente
  - Apri distributore esistente
  - Apri una risorsa singola
- elenco recente con ultimi distributori o programmi aperti;
- piccolo riquadro stato progetto con errori aperti o file non salvati.

### 5.2 Wizard Nuovo programma

Scopo:

- guidare la creazione di un nuovo programma senza esporre dettagli tecnici.

Step proposti:

1. codice nuovo distributore;
2. scelta template di partenza oppure creazione vuota;
3. numero sezioni attive;
4. scelta se copiare o no file collegati;
5. configurazione CE16 opzionale;
6. riepilogo finale con elenco file che verranno creati.

Output del wizard:

- nuovo file distributore;
- nuove sequenze test o collegamenti a sequenze esistenti;
- eventuali copie di curve, rampe, limiti e file esterni.

### 5.3 Wizard Duplica programma

Scopo:

- clonare un programma esistente con controllo fine sulle risorse figlie.

Step proposti:

1. selezione distributore sorgente;
2. nuovo codice destinazione;
3. tabella risorse collegate con checkbox:
  - riusa
  - duplica
  - duplica e rinomina automaticamente
4. anteprima riferimenti finali;
5. conferma.

### 5.4 Editor Distributore

Scopo:

- trattare il distributore come risorsa principale del programma di collaudo.

Layout proposto:

- testata con nome distributore, stato validazione e azioni Salva/Duplica;
- sezione Moduli del distributore con 5 card o righe, una per ogni sezione;
- sezione Calibrazioni CE16 con 5 campi opzionali;
- sezione Parametri generali con temperatura olio e pressione;
- sezione Collegamenti rapidi per aprire le sequenze test associate.

Per ogni sezione mostrare:

- nome file TEST associato;
- pulsanti Seleziona, Nuovo, Duplica, Apri, Usi;
- semaforo stato validazione.

### 5.5 Editor Sequenza TEST

Scopo:

- modificare la lista dei test di una sezione in modo visuale e guidato.

Layout proposto:

- parte alta con titolo sequenza, distributore associato, stato validazione;
- tabella centrale con righe test;
- pannello inferiore o laterale con dettagli della riga selezionata.

Tabella test:

- colonna indice;
- colonna nome test;
- colonna ID test;
- colonna stato;
- colonna riferimenti;
- drag and drop per riordinare;
- pulsanti aggiungi riga, duplica riga, elimina riga.

Dettaglio riga selezionata:

- campi generati dinamicamente in base all'ID test;
- ogni file referenziato usa un widget composto da:
  - selettore file
  - Nuovo
  - Duplica
  - Apri
  - Usi

### 5.6 Editor Curve Comando

Layout proposto:

- grafico a sinistra;
- tabella punti a destra;
- toolbar locale con Aggiungi punto, Elimina punto, Ordina, Duplica file;
- legenda minima e stato validazione.

### 5.7 Editor Curve Limite

Layout proposto:

- grafico principale;
- tabella punti;
- toggle per visualizzare anche la curva opposta INF o SUP;
- azione confronto banda;
- pannello info su chi usa il limite.

### 5.8 Editor Rampa XY

Layout proposto:

- grafico tempo/tensione con punti trascinabili;
- tabella sotto o a destra;
- pulsanti snap, inserisci punto intermedio, ordina, duplica.

### 5.9 Editor CE16

Scopo:

- configurare il sensore CE16 quando presente.

Layout proposto:

- grafico punti/posizione;
- tabella punti;
- badge informativo: sensore opzionale;
- indicazione di quali distributori o test usano questa calibrazione.

### 5.10 Editor MMS2218

Layout proposto:

- grafico calibrazione ADC;
- tabella punti X/Y;
- controlli range e ordinamento;
- info profilo banco associato.

### 5.11 Editor File Esterni

Scopo:

- gestire file esterni referenziati dai test anche quando non esiste ancora un
  editor specializzato.

Layout proposto:

- vista testo assistita o tabella chiave/valore, se riconoscibile;
- validazione base;
- azioni duplica, rinomina, sostituisci riferimento;
- messaggio esplicito se il formato non e ancora supportato in modo strutturato.

### 5.12 Viewer Output Banco

Scopo:

- consultare output di test senza confonderli con la configurazione.

Layout proposto:

- barra superiore con archivio output attivo e pulsante `Configura archivio output`;
- area filtri con stringa libera, data creazione da/a, data modifica da/a,
  ordinamento e refresh;
- tabella risultati con cartella, nome file CSV, data creazione, data modifica,
  dimensione e stato di lettura;
- pannello grafico con selettore asse X e selettore di uno o piu assi Y;
- banner sola lettura;
- grafico output;
- tabella dati del CSV selezionato;
- pulsanti Esporta, Confronta, Stampa, Apri cartella;
- nessun pulsante Nuovo/Salva nel flusso principale.

Requisiti UX specifici:

- il viewer non deve assumere che i file siano dentro il progetto corrente;
- il percorso archivio output deve essere parametrizzabile da impostazioni o da
  una dialog dedicata;
- nel flusso principale non va mostrato un path filesystem grezzo;
- l'utente deve poter scegliere quale colonna usare come asse X;
- l'utente deve poter scegliere uno o piu assi Y da sovrapporre sul grafico;
- il filtro principale deve coprire almeno stringa libera e data di creazione;
- se l'archivio e molto grande, l'interfaccia deve mostrare risultati parziali
  e stato dell'indicizzazione.

Requisiti di scalabilita:

- non leggere tutti i CSV all'avvio;
- indicizzare in background solo i metadati di file e cartelle;
- mantenere una cache persistente dell'indice;
- aprire e parsare il CSV solo quando l'utente seleziona un risultato;
- supportare ricerca incrementale e paginazione o virtualizzazione della lista.

---

## 6. Flussi utente principali

### 6.1 Flusso A — Creare un nuovo programma di collaudo

1. L'utente apre la dashboard.
2. Clicca Nuovo programma.
3. Compila il wizard con codice, template e sezioni.
4. Il sistema crea il distributore e le risorse minime collegate.
5. L'utente entra nell'editor distributore.
6. Per ogni sezione apre o crea la sequenza TEST.
7. Dalla sequenza TEST crea o seleziona curve, rampe, limiti e altri file.
8. Lancia la validazione complessiva.
9. Salva il programma.

### 6.2 Flusso B — Duplicare un programma esistente

1. L'utente apre Duplica programma.
2. Seleziona il distributore sorgente.
3. Indica il nuovo codice.
4. Decide quali risorse duplicare e quali riusare.
5. Visualizza l'anteprima delle nuove referenze.
6. Conferma.
7. Atterra nell'editor distributore del nuovo programma.
8. Corregge solo i parametri necessari.

### 6.3 Flusso C — Modificare una sequenza test

1. L'utente apre un distributore.
2. Clicca Apri sulla sezione desiderata.
3. Visualizza la tabella dei test.
4. Seleziona una riga.
5. Modifica i parametri nel pannello dettaglio.
6. Se un parametro punta a un file, usa Seleziona/Nuovo/Duplica/Apri/Usi.
7. Controlla errori e warning.
8. Salva.

### 6.4 Flusso D — Creare una risorsa direttamente da un parametro

1. L'utente seleziona una riga test.
2. Nel dettaglio parametro clicca Nuovo vicino al riferimento file.
3. Il sistema apre un mini-dialog con nome file e posizione suggerita.
4. Alla conferma apre subito l'editor dedicato della nuova risorsa.
5. Al salvataggio, il riferimento viene gia impostato nella riga test.

### 6.5 Flusso E — Eliminare una risorsa in sicurezza

1. L'utente seleziona una risorsa nell'albero.
2. Clicca Elimina.
3. Si apre una modale con elenco file che la referenziano.
4. Il sistema propone opzioni:
   - annulla
   - elimina comunque
   - sostituisci riferimento prima di eliminare
5. L'utente conferma l'azione desiderata.

### 6.6 Flusso F — Consultare un output di banco

1. L'utente apre la sezione Output banco.
2. Seleziona o configura l'archivio output.
3. Applica un filtro per stringa, data di creazione o altri metadati.
4. Sceglie un risultato dalla lista filtrata.
5. Seleziona la colonna asse X e uno o piu assi Y.
6. Visualizza grafico e dati in sola lettura.
7. Se necessario esporta o confronta il file.

---

## 7. Stati e feedback da rappresentare

Il mock-up deve prevedere questi stati visivi:

- file salvato / ok;
- file modificato non salvato;
- warning non bloccante;
- errore bloccante;
- risorsa opzionale non configurata;
- risorsa condivisa da piu file;
- modalita sola lettura per output banco.

Nota: il riferimento mancante ricade nello stato di errore bloccante. Lo stato
di indicizzazione archivio output in corso va rappresentato come attivita di
sistema, non come nuovo stato risorsa.

Feedback UX richiesto:

- colori sobri e chiari;
- tooltip esplicativi;
- semafori o badge stato;
- conferme prima di azioni distruttive;
- report validazione leggibile in italiano.

---

## 8. Direzione visiva per il mock-up

La UI deve avere un look tecnico-industriale, non consumer e non stile IDE.

Indicazioni visive:

- tema chiaro come default;
- palette neutra con accenti blu, verde e rosso per gli stati;
- interfaccia densa ma ordinata;
- icone semplici e funzionali;
- componenti grandi quanto basta per uso da ufficio tecnico;
- forte enfasi sulla leggibilita dei dati e sulla gerarchia dei pannelli.

Da evitare:

- estetica troppo futuristica;
- dark mode come focus principale;
- layout troppo minimal che nasconde le azioni;
- linguaggio visivo da software creativo;
- grafiche decorative non funzionali.

---

## 9. Output richiesto al LLM di design

Chiedere al modello di produrre:

- una vista generale della dashboard;
- il wizard Nuovo programma;
- l'editor Distributore;
- l'editor Sequenza TEST;
- un editor grafico tipo Curva/Rampa;
- la modale di eliminazione con impatto riferimenti;
- il viewer sola lettura per output banco.

Per ogni schermata il modello dovrebbe fornire:

- struttura del layout;
- gerarchia visiva;
- etichette principali in italiano;
- componenti chiave;
- comportamento atteso nei punti critici.

---

## 10. Prompt pronto da incollare in un LLM di design

Usa questo prompt come base:

"""
Progetta un primo mock-up UX desktop Windows per un software chiamato
AT614 Configuration Editor.

Contesto prodotto:
- il software serve a utenti non programmatori per creare e modificare programmi
  di collaudo di un banco EOL;
- la UI deve essere solo in italiano;
- il programma di collaudo e composto da piu risorse collegate: distributore,
  sequenze test, curve comando, curve limite, rampe XY, calibrazioni CE16,
  calibrazioni MMS2218 e file esterni referenziati dai test;
- CE16 e una risorsa opzionale;
- le cartelle GRAPH e RAMPE_XY_LAST contengono solo output del banco e non
  devono essere trattate come file di configurazione;
- gli output possono anche stare in archivi esterni al progetto, con struttura
  a cartelle e file CSV interni;
- gli archivi output possono contenere moltissimi file, quindi servono filtri e
  caricamento lazy.

Obiettivo UX:
- rendere semplice creare un nuovo programma da zero o duplicarne uno esistente;
- permettere apertura e modifica guidata di tutte le risorse principali;
- rendere esplicite le dipendenze tra file;
- prevenire errori prima del salvataggio.

Architettura della GUI:
- barra superiore con azioni globali: Nuovo programma, Duplica programma,
  Apri risorsa, Salva, Salva tutto, Validazione, Cerca, Configura archivio
  output;
- colonna sinistra con albero risorse raggruppato per tipologia funzionale;
- area centrale con editor della risorsa selezionata;
- pannello destro con riferimenti, validazione e azioni rapide.

Tipi di schermata da progettare:
1. dashboard iniziale con azioni principali;
2. wizard Nuovo programma di collaudo;
3. editor Distributore con sezioni, riferimenti TEST e calibrazioni CE16;
4. editor Sequenza TEST con tabella test e pannello parametri dinamico;
5. editor grafico per curve/rampe;
6. dialog di eliminazione sicura con impatto sui riferimenti;
7. viewer in sola lettura per output banco.

Dettagli richiesti per il viewer output banco:
- percorso archivio configurabile, anche esterno al progetto;
- filtro per stringa libera, data creazione e data modifica;
- lista risultati veloce anche con moltissimi file;
- scelta asse X e uno o piu assi Y prima del plot;
- indicizzazione lazy con feedback visivo dello stato.

Regole UX:
- non deve sembrare un IDE o un editor di testo generico;
- deve sembrare uno strumento tecnico-industriale guidato;
- tema chiaro, leggibile, sobrio;
- usare etichette in italiano;
- rendere molto visibili errori, warning, riferimenti mancanti e file non salvati;
- ogni parametro che punta a un file deve avere azioni: Seleziona, Nuovo,
  Duplica, Apri, Usi.
- non caricare l'intero archivio output nel tree all'avvio: usare ricerca,
  filtri e risultati progressivi.
- per componenti nominati, microcopy e stati usare rigorosamente il brief
  `LAYOUT_UX.md`.

Flussi principali da supportare:
1. creare un nuovo programma di collaudo da template;
2. duplicare un programma esistente scegliendo quali risorse duplicare o riusare;
3. modificare una sequenza test e creare al volo curve, rampe o limiti dai
   parametri;
4. eliminare una risorsa in sicurezza sapendo chi la usa;
5. consultare output banco in sola lettura.
6. consultare un archivio output esterno filtrando per codice prodotto o data e
  scegliendo quali colonne graficare come asse X e assi Y.

Output desiderato:
- descrizione delle schermate;
- wireframe testuali o mock-up strutturati;
- note di comportamento UX;
- proposta visiva coerente e non generica.
"""