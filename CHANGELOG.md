# Changelog

Formato basato su [Keep a Changelog](https://keepachangelog.com/it/1.1.0/); versioni secondo [Semantic Versioning](https://semver.org/lang/it/).

## [0.1.0] - beta - in preparazione

Prima versione pubblica. Sostituisce il prototipo con server Node locale e macro VBA.

### Aggiunto
- **Installatore unico per Windows** (`PersonalAIAssistant-0.1.0-Setup.exe`, Inno Setup), per utente, che compare in *App installate*.
  - Word ed Excel: catalogo attendibile in una cartella condivisa locale in sola lettura (`\\localhost\PersonalAIAssistant$`), creata con una sola conferma da amministratore per PC; poi un clic su *Aggiungi* la prima volta.
  - Outlook classico: add-in VSTO per utente, senza amministratore e senza avvisi di fiducia (elenco di inclusione VSTO limitato al manifest firmato del progetto).
  - Controllo dei prerequisiti (.NET Framework 4.8, runtime VSTO, Outlook classico), richiesta di chiudere Office, rimozione opzionale della condivisione in disinstallazione.
- **Add-in Outlook** (C#, VSTO): gruppo *Personal AI* nella scheda Home con *Riassumi selezionate*, *Rispondi con AI* (conferma, declina, chiarimenti, personalizzata) e *Impostazioni*.
  Chiavi cifrate con la protezione dati di Windows, chiamate asincrone annullabili, importazione automatica delle impostazioni della vecchia macro.
- **PowerPoint**: riscrittura, correzione, sintesi e traduzione del testo selezionato; creazione di slide da una scaletta (layout "Titolo e contenuto" del tema, aggiunte in fondo); miglioramento di tutti i testi della slide selezionata, con anteprima e annullamento.
- **Excel: grafici.** Il modello propone tipo, titolo e intervalli; il pannello valida la proposta e la mostra prima di inserirla. Conteggi, somme e medie per categoria sono calcolati dal pannello in modo esatto, nel foglio *Dati grafici PAI*.
- Pannello: primo avvio guidato con i link *Ottieni chiave*, pulsante *Cancella* chiavi, numero di versione.
- Pagine [installazione](public/installa.html) e [informativa privacy](public/privacy.html).
- Licenza MIT.

### Modificato
- Il pannello Word/Excel chiama direttamente i provider (Gemini, OpenAI, Claude): nessun server locale, nessun certificato.
- Excel: il tipo di inserimento (formula, tabella, testo, grafico) dipende dal pulsante premuto, non dal testo del modello; la conferma di sovrascrittura vale solo per la stessa area e lo stesso contenuto.
- Formule Excel scritte nella lingua di Excel (italiano con `;`), con la destinazione che salta l'intestazione e l'estensione automatica alle righe di dati.
- Word: il risultato si inserisce con un'unica scrittura (un solo Ctrl+Z per annullare).
- Il manifest Word/Excel (versione 0.1.0.0) viene generato dal modello con l'indirizzo reale del sito.
- Le chiavi del pannello sono salvate con nomi propri dell'app (migrazione automatica).

### Corretto
- Il server del prototipo esponeva `.env` e il certificato a qualunque sito web: architettura rimossa.
- La macro Outlook troncava i riassunti con elenchi puntati e perdeva accenti e a capo.
- Chiave Gemini non valida (HTTP 400) segnalata come chiave non valida.

### Note
- Il sideload di Word/Excel tramite la chiave di registro `WEF\Developer` è stato scartato: l'add-in partiva solo dopo aver cliccato sul messaggio di errore di Office (verificato su due utenti Windows).
- Provato su Windows 11 con Office LTSC 2024 a 32 bit (utente amministratore e standard). PowerPoint e l'add-in Outlook definitivo non sono ancora stati provati dentro Office; Office a 64 bit e altre edizioni non provati.

### Sicurezza
- Chiavi del pannello non piu' ricopiate nei campi dopo il salvataggio (solo le ultime 4 cifre nel segnaposto); opzione "solo per questa sessione" che non le scrive su disco.
- Installatore: la parte eseguita come amministratore riceve codice e manifest incorporati nel comando (nessun file intermedio in cartelle scrivibili dall'utente).
- Avvertenze per proteggere la chiave API (chiave dedicata, limiti di spesa, revoca) nell'installatore, nella guida e nell'informativa.
- Ramo main protetto da cancellazione e riscrittura della storia.
- Opzione "solo sessione" corretta: la scelta vale subito in memoria (anche se il suo salvataggio fallisce), la migrazione delle vecchie chiavi la rispetta, le cancellazioni non riuscite vengono segnalate.
- Installatore: il catalogo di Word/Excel/PowerPoint viene considerato attendibile solo dopo una verifica senza privilegi di condivisione, cartella, permessi e manifest; in caso di rifiuto o errore nessuna registrazione residua. La disinstallazione rimuove solo i propri file e segnala l'esito.
- Outlook: la pulizia delle chiavi della vecchia macro dal registro viene verificata, ritentata a ogni avvio e segnalata se non riesce.

### Revisione
- Correzioni dalle revisioni di Codex e Gemini: conferma di sovrascrittura legata a foglio, indirizzo e contenuto (anche celle con formule); add-in Outlook protetto da errori del modello a oggetti dopo l'attesa; script della condivisione che rifiuta cartelle non sicure e condivisioni non sue e imposta i permessi prima di scrivere; PowerPoint che non sovrascrive modifiche fatte dopo la proposta; foglio dei grafici appena creato; manifest senza AppDomain; pulizia della cache per tutte le versioni dei comandi; build che rileva la versione di Visual Studio.
