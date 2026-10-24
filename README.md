# CERCA_UDI
Software per la ricerca di codici UDI sulla banca dati EUDAMED partendo da una serie di codici REF.

[Scarica l'ultima versione](https://github.com/marchimatteo/cerca_udi/releases/latest)

Appena lanciato, il software chiederà di selezionare un file excel in cui dovranno essere presenti 3 colonne:
- PRODUTTORE: con il nome del produttore
- REF: con il codice REF del dispositivo
- DESCRIZIONE: con il nome del dispositivo

Nel repo è disponibile il file di esempio *esempio_file_da_caricare.xlsx* che può essere utilizzato direttamente.

Il REF è l'unico valore che effettivamente viene inviato ad EUDAMED, produttore e descrizione servono per verificare manualmente che il dispositivo trovato corrisponda effettivamente a quello desiderato.

Se si desidera modificare il codice Python poi per rigenerare l'eseguibile è sufficiente eseguire *pyinstaller CercaUDI.spec*