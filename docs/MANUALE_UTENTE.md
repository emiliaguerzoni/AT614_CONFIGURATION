# Manuale Utente — AT614 Configuration Editor

## 1. Avvio e Utilizzo
L'applicazione non richiede installazione. È sufficiente avviare l'eseguibile `at614-editor.exe`.
Al primo avvio, l'editor cercherà una cartella di progetto valida. Puoi forzare l'apertura di un progetto specifico passandolo come argomento.

## 2. Interfaccia Principale
L'interfaccia è suddivisa in tre colonne principali:
1. **Risorse (sinistra)**: Mostra l'albero di tutti i file configurazione trovati nel progetto (Distributori, Moduli TEST, Curve, ecc.).
2. **Editor (centro)**: L'area di lavoro principale dove visualizzare e modificare la risorsa selezionata.
3. **Pannello di Controllo (destra)**: Mostra informazioni aggiuntive come i riferimenti ("Chi uso" e "Chi mi usa") e lo stato di validazione.

## 3. Operazioni di Base
- **Salvataggio**: Il sistema salva automaticamente o ti avvisa se ci sono modifiche pendenti non valide.
- **Validazione**: Durante la modifica, gli errori bloccanti vengono mostrati a schermo. Puoi forzare una validazione completa dal Pannello di Controllo.
- **Rinomina**: Per rinominare una risorsa, usa l'azione dal menu contestuale. Tutti i file che dipendono da quella risorsa verranno aggiornati automaticamente.
- **Clonazione Programmi**: Dalla dashboard principale è possibile clonare un intero programma di test, decidendo quali dipendenze duplicare e quali riutilizzare.
- **Visualizzazione Output**: Dalla tab dedicata puoi navigare gli archivi CSV dei test al banco in sola lettura, visualizzando i grafici e confrontando serie di dati.

## 4. File di Log
In caso di problemi, i file di log vengono salvati automaticamente in:
`~/.at614-editor/logs/at614-editor.log`
(la tilde `~` rappresenta la tua cartella utente di Windows, es. `C:\Users\NomeUtente`).
