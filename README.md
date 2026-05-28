# AT614 Configuration Editor

**Editor Python GUI per la configurazione dei programmi di collaudo del banco AT614**

[![Python](https://img.shields.io/badge/Python-3.11%2B-blue?logo=python)](https://www.python.org/)
[![PySide6](https://img.shields.io/badge/PySide6-6.11%2B-green?logo=qt)](https://wiki.qt.io/PySide6)
[![License](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

## 📋 Descrizione

AT614 Configuration Editor è un'applicazione desktop Python che semplifica la creazione e la modifica dei programmi di collaudo per il banco EOL AT614. Consente a tecnici di collaudo e ingegneri di test di configurare procedure complesse senza necessità di conoscenza approfondita di VB6 o delle strutture interne dei file.

L'applicazione è stata progettata per **utenti non programmatori** con conoscenza del dominio tecnico, offrendo un'interfaccia intuitiva e guidata.

## ✨ Caratteristiche Principali

### 🎯 Editor Specializzati
- **Distributore**: Gestione sezioni, calibrazioni CE16 e parametri
- **Sequenze Test**: Editor visuale con tabella test e pannello parametri dinamico
- **Curve Comando**: Editor grafico con visualizzazione punti e valori
- **Curve Limite**: Gestione limiti superiore/inferiore con confronto banda
- **Rampe XY**: Editor tempo-tensione con punti trascinabili
- **Calibrazioni**: Supporto CE16 e MMS2218
- **File Esterni**: Editor assistito per file non strutturati

### 💡 Funzionalità UX
- **Tooltip Intelligenti**: Descrizioni contestuali per ogni campo editabile, estratte dagli schemi YAML
- **File Picker**: Dialog comune per la selezione dei file con filtri specifici per tipo di risorsa
- **Percorsi Relativi**: Gestione automatica dei percorsi relativi alla project root
- **Validazione**: Controlli di errore e warning in tempo reale
- **Grafico Interattivo**: Visualizzazione dinamica delle curve con aggiornamento in tempo reale

### 🔗 Gestione Risorse
- **Albero Risorse**: Organizzazione per tipologia funzionale (distribuitori, test, curve, etc.)
- **Tracciamento Dipendenze**: Monitoraggio di chi usa quale risorsa
- **Operazioni Sicure**: Conferme prima di azioni distruttive con impatto sui riferimenti
- **Duplicazione Intelligente**: Gestione automatica delle risorse collegate

### 📁 Supporto Cartelle
- Supporto completo per tutte le cartelle definite in `settings.ini`
- Mapping automatico: FolderConfigurazioneTest, FolderFileCurveComando, FolderFileCurveLimite, FolderRampeXY, FolderConfigurazioneModuli, etc.
- Risoluzione automatica dei percorsi

## 🚀 Installazione Veloce

### Prerequisiti
- Python 3.11 o superiore
- pip

### Step 1: Clonare il repository
```bash
git clone https://github.com/emiliaguerzoni/AT614_CONFIGURATION.git
cd AT614_CONFIGURATION
```

### Step 2: Installare le dipendenze
```bash
pip install -e .
```

Opzionalmente, per sviluppo con test e documentazione:
```bash
pip install -e ".[dev]"
```

### Step 3: Lanciare l'applicazione
```bash
at614-editor
```

## 📖 Utilizzo

### Flusso Principale

#### 1. Creare un nuovo programma
1. Clicca "Nuovo programma" dalla dashboard
2. Compila il wizard:
   - Codice nuovo distributore
   - Scegli template o crea vuoto
   - Numero sezioni attive
   - Configurazione opzionale CE16
3. Il sistema crea il distributore e le risorse minime
4. Entra nell'editor distributore per configurare i dettagli

#### 2. Modificare una sequenza test
1. Seleziona un distributore dall'albero risorse
2. Clicca "Apri" sulla sezione desiderata
3. Visualizza la tabella dei test
4. Seleziona una riga per modificare i parametri
5. Usa il file picker per selezionare risorse collegate
6. I tooltip guidano il significato di ogni campo

#### 3. Selezionare file con File Picker
1. Clicca il pulsante "Seleziona" in qualsiasi campo file
2. Si apre un dialog con:
   - Filtri specifici per tipo di risorsa
   - Directory di partenza appropriata
   - Supporto alle cartelle da settings.ini
3. Seleziona il file
4. Il percorso viene automaticamente salvato come relativo

#### 4. Duplicare un programma
1. Clicca "Duplica programma"
2. Seleziona distributore sorgente
3. Indica nuovo codice
4. Per ogni risorsa: decidi se duplicare o riusare
5. Visualizza anteprima dei file che verranno creati
6. Conferma e il nuovo programma è pronto

## 🏗️ Struttura del Progetto

```
AT614_CONFIGURATION/
├── src/at614_editor/
│   ├── domain/                 # Logica applicativa
│   │   ├── models.py          # Modelli dati
│   │   ├── project.py         # Gestione progetto
│   │   ├── parsers/           # Parser per CSV, YAML, INI
│   │   ├── settings_ini.py    # Parser settings.ini
│   │   ├── refactor.py        # Operazioni di refactor
│   │   ├── program_clone.py   # Logica clonazione
│   │   └── test_schema.py     # Schemi parametri test
│   ├── ui/                     # Interfaccia utente
│   │   ├── main_window.py     # Finestra principale
│   │   ├── editors/           # Editor specializzati
│   │   ├── components/        # Componenti riutilizzabili
│   │   ├── file_picker.py     # File picker con dialog
│   │   └── tooltip_manager.py # Gestione tooltip
│   └── __main__.py            # Entry point
├── schemas/                    # Schemi parametri test
├── tests/                      # Test suite
├── docs/                       # Documentazione
├── pyproject.toml            # Configurazione progetto
└── README.md                 # Questo file
```

## 🧑‍💻 Sviluppo

### Configurazione ambiente di sviluppo

```bash
# Clonare il repository
git clone https://github.com/emiliaguerzoni/AT614_CONFIGURATION.git
cd AT614_CONFIGURATION

# Installare in modalità sviluppo
pip install -e ".[dev]"
```

### Eseguire i test
```bash
pytest
```

## 📊 Schema Dati

### Distributore (DISTRIBUTORE/*.csv)
- Sezioni test (5 colonne per i codici TEST)
- Calibrazioni CE16 (5 slot opzionali)
- Parametri aggiuntivi (coppia chiave-valore)

### Sequenze Test (TEST/*.csv)
- Nome test
- ID test (determina tipo e parametri specifici)
- Indice (numero o espressione i++/++i)
- Parametri specifici

## 🔧 Configurazione

L'applicazione legge i percorsi delle cartelle da `settings.ini`:

```ini
FolderConfigurazioneBancoCollaudo=C:\Path\To\DISTRIBUTORE
FolderConfigurazioneTest=C:\Path\To\TEST
FolderFileCurveComando=C:\Path\To\CURVE_COMANDO
FolderFileCurveLimite=C:\Path\To\CURVE_LIMITE
FolderRampeXY=C:\Path\To\RAMPE_XY
FolderConfigurazioneModuli=C:\Path\To\CONFIGURAZIONE
```

## 📝 Tooltip e Descrizioni

Ogni campo editabile ha un **tooltip** che descrive il significato e come compilarlo. I tooltip sono generati automaticamente da schemi YAML e descrizioni predefinite.

## 🔍 Debug e Logging

L'applicazione registra automaticamente **tutte le operazioni importanti** in tempo reale nel terminale. Questo è molto utile per diagnosticare i problemi.

### Visualizzare i Log nel Terminale

Quando avvii l'applicazione, vedrai i log nel terminale:

```
08:15:23 | at614_editor | INFO     | ================================================================================
08:15:23 | at614_editor | INFO     | AT614 Configuration Editor - Avvio applicazione
08:15:23 | at614_editor | INFO     | ================================================================================
08:15:24 | at614_editor.ui.editors.dist_editor | INFO     | DistEditor: inizializzazione per distributore_001.csv
08:15:24 | at614_editor.ui.editors.dist_editor | DEBUG    | DistEditor: configurazione caricata con successo
08:15:25 | at614_editor.ui.file_picker | DEBUG    | FilePicker: apertura dialog per tipo 'sequenze_test'
08:15:26 | at614_editor.ui.file_picker | INFO     | FilePicker: file selezionato: C:\path\to\TEST\test_001.csv
```

### Dove Trovare i Log Completi

I log vengono registrati anche in un file permanente:

**Windows:**
```
C:\Users\<tuonome>\.at614-editor\logs\at614-editor.log
```

**macOS/Linux:**
```
~/.at614-editor/logs/at614-editor.log
```

### Informazioni Registrate

Ogni log contiene:
- **Timestamp**: Ora dell'evento (HH:MM:SS)
- **Modulo**: Quale parte del codice ha generato il log
- **Livello**: DEBUG, INFO, WARNING, ERROR, CRITICAL
- **Messaggio**: Descrizione dell'evento

### Livelli di Log

| Livello | Significato | Quando Compare |
|---------|-------------|-----------------|
| **DEBUG** | Informazioni dettagliate per sviluppatori | Sempre durante le operazioni |
| **INFO** | Informazioni generali importanti | Avvio app, caricamento file, salvataggio |
| **WARNING** | Avvisi di situazioni insolite | Percorsi mancanti, valori anomali |
| **ERROR** | Errori che impediscono un'operazione | File corrotto, parsing fallito |
| **CRITICAL** | Errori critici dell'applicazione | Crash imminente |

### Interpretare i Log in Caso di Crash

Se l'applicazione va in crash, guarda i log recenti nel terminale:

1. **Cerca EXCEPTION o ERROR** nei log - lì troverai il motivo del crash
2. **Nota il timestamp** - aiuta a identificare quale azione ha causato il problema
3. **Copia il messaggio di errore intero** - utile per segnalare il bug

Esempio di log di errore:
```
08:20:15 | at614_editor.domain.parsers.distributore | ERROR | Exception in parse_distributore
Traceback (most recent call last):
  File ".../parsers/distributore.py", line 36, in parse
    ...
```

### Aumentare il Verbosity (Avanzato)

Se hai bisogno di più dettagli, puoi modificare il file `src/at614_editor/__main__.py` e cercare `level=logging.DEBUG` per cambiarla in `level=logging.DEBUG`. I log di DEBUG contengono dettagli delle operazioni file picker, parsing, e serializzazione.

## 🐛 Segnalazione Bug

Apri una [Issue su GitHub](https://github.com/emiliaguerzoni/AT614_CONFIGURATION/issues) con descrizione del problema e passi per riprodurlo.

## 🎯 Roadmap

- [ ] Integrazione con banco AT614 (lettura dati real-time)
- [ ] Export configurazione in formati standard
- [ ] Supporto multi-lingua
- [ ] Dark mode
- [ ] Plugin system per editor personalizzati

## 📄 Licenza

Questo progetto è distribuito sotto licenza MIT.

## 👥 Contatti

**Sviluppo:** Simone Pandini  
**Email:** simone.pandini@gmail.com  
**GitHub:** https://github.com/emiliaguerzoni/AT614_CONFIGURATION

---

**Versione:** 0.1.0 (Beta)
