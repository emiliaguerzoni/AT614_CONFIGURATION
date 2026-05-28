# Checklist di Collaudo Manuale

Prima di rilasciare una nuova versione in produzione dell'eseguibile, verificare manualmente i seguenti punti:

## 1. Avvio Standalone
- [ ] Avviare `at614-editor.exe` su un PC Windows **senza Python preinstallato**.
- [ ] L'applicazione si apre senza mostrare la console nera di debug in background.
- [ ] La cartella log `~/.at614-editor/logs/` viene creata e il file `at614-editor.log` si popola regolarmente.

## 2. Apertura Progetto
- [ ] Aprire un progetto reale (es. `AT614_BANCO_0001`).
- [ ] Verificare che l'albero delle risorse (colonna di sinistra) sia popolato con tutte le categorie (Distributori, TEST, Curve, ecc.).

## 3. Test degli Schemi
- [ ] Aprire un file `Module_TEST_*.bas`.
- [ ] Verificare che i parametri specifici del test siano decodificati e visualizzati correttamente con le loro etichette e unità di misura, a conferma che la cartella `schemas` è stata integrata nell'eseguibile.

## 4. Riferimenti e Modifiche (Safe Operations)
- [ ] Rinominare una `Curva di Comando` e verificare che tutti i moduli TEST che la utilizzano vengano aggiornati di conseguenza.
- [ ] Provare l'azione "Clona Programma" dalla dashboard e confermare che la copia generi i file attesi.

## 5. Viewer di Output
- [ ] Aprire la sezione "Viewer output banco".
- [ ] Filtrare e selezionare un file archivio `.csv`.
- [ ] Verificare che i dati vengano plottati correttamente selezionando gli assi X/Y desiderati.
