# CERCA_UDI
Software per la ricerca di codici UDI sulla banca dati EUDAMED partendo da una serie di codici REF.

[Scarica l'ultima versione](https://github.com/marchimatteo/cerca_udi/releases/latest)

Appena lanciato, il software chiederà di selezionare un file excel in cui dovranno essere presenti 3 colonne:
- PRODUTTORE: con il nome del produttore
- REF: con il codice REF del dispositivo
- DESCRIZIONE: con il nome del dispositivo

Nella release è disponibile il file di esempio *esempio_file_da_caricare.xlsx* che può essere utilizzato direttamente.

Verso EUDAMED viene inviato solo il REF. Il campo PRODUTTORE serve a filtrare i risultati ricevuti: se per quel REF almeno un risultato ha un produttore corrispondente a quello indicato, vengono mostrati solo quei risultati; altrimenti vengono restituiti tutti. La DESCRIZIONE non influisce sulla ricerca e serve solo per verificare manualmente che il dispositivo trovato sia quello desiderato.

Al termine il software produce un file Excel con gli UDI trovati, il file avrà preimpostate 2 regole di formattazione condizionale:
- riga verde in caso di REF trovato e produttore fornito uguale a quello trovato in Eudamed;
- riga rossa in caso di REF non trovato o errore.

Se si desidera modificare il codice Python poi per rigenerare l'eseguibile è sufficiente eseguire *pyinstaller CercaUDI.spec*