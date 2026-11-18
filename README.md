# Personal AI Assistant

Assistente AI per **Word, Excel** (pannello laterale) e **Outlook POP/SMTP** (macro VBA), con modelli **Google Gemini, OpenAI e Anthropic Claude** a scelta.

Da novembre 2026 **non serve più nessun server locale** (niente Node, certificati o `avvia_server.bat`):
- il pannello Word/Excel è una pagina statica pubblicata su un hosting HTTPS e chiama direttamente le API dei provider;
- la macro Outlook chiama direttamente le API dei provider.

Le chiavi API restano sul tuo PC (archivio del pannello e registro utente per la macro) e vengono inviate solo al provider scelto. I testi elaborati vanno al provider cloud selezionato.

---

## Struttura

| Percorso | Contenuto |
| --- | --- |
| `public/` | Pagina del pannello: `index.html`, `style.css`, `taskpane.js`, `ai.js` (chiamate ai provider + **catalogo modelli**), icone. È l'unica cartella da pubblicare. |
| `manifest-word-excel.template.xml` | Manifest dell'add-in con segnaposto `__BASE_URL__`. |
| `installa.ps1` | Genera il manifest, lo registra per l'utente, pulisce la vecchia versione, importa le chiavi per la macro. |
| `outlook_macro.vba` | Codice del modulo VBA per Outlook. |
| `_obsoleto/` | Vecchia architettura con server Node (conservata per riferimento, si può cancellare). |

## 1. Pubblicare il pannello (una volta)

Office carica i pannelli solo da indirizzi `https://`. I file in `public/` non contengono segreti, quindi possono stare su un hosting statico gratuito.

**GitHub Pages:**
1. Crea un repository (es. `personal-ai-assistant`) e carica il **contenuto** di `public/` nella radice.
2. Vai in *Settings → Pages → Branch: main / root → Save*.
3. Dopo un minuto l'indirizzo è `https://TUO-UTENTE.github.io/personal-ai-assistant`.

In alternativa puoi usare Cloudflare Pages o Netlify: si trascina la cartella `public/` nel sito.

Per aggiornare il pannello basta ricaricare i file: Office prende la nuova versione alla riapertura del pannello.

## 2. Installare l'add-in Word/Excel

Con Word, Excel e Outlook chiusi, apri PowerShell nella cartella del progetto:

```powershell
powershell -ExecutionPolicy Bypass -File .\installa.ps1 -BaseUrl https://TUO-UTENTE.github.io/personal-ai-assistant
```

Apri Word: compare la scheda **Personal AI**. Apri il pannello, clicca **⚙️**, inserisci le chiavi, poi **Salva chiavi** e **Verifica chiavi**.

## 3. Installare la macro Outlook

1. Per copiare le chiavi del vecchio `.env` nelle impostazioni della macro:
   `powershell -ExecutionPolicy Bypass -File .\installa.ps1 -ImportaChiaviDaEnv`
   In alternativa, dopo il punto 2 esegui la macro **ImpostazioniAI**.
2. In Outlook premi `Alt+F11`. Se c'è il vecchio modulo, cancellane tutto il contenuto; altrimenti usa *Inserisci → Modulo*. Incolla `outlook_macro.vba` e salva.
3. Esegui **ImpostazioniAI** per scegliere provider e modello.
4. I pulsanti già presenti sulla barra multifunzione continuano a funzionare, perché i nomi delle macro non sono cambiati: `RiassumiMailSelezionate`, `RispondiConAI_Conferma`, `RispondiConAI_Declina`, `RispondiConAI_Personalizzata`. Puoi aggiungere anche un pulsante per `ImpostazioniAI`.

## Modelli

L'elenco è in `AI_CATALOG` in [public/ai.js](public/ai.js): la tendina viene generata da lì. Il pulsante **Verifica chiavi** segnala i modelli del catalogo che il provider non offre più.
Per Claude Opus/Sonnet 5.5 è attivo il ripiego automatico lato server: se il modello declina una richiesta, ne risponde un altro. Il nome del modello che ha risposto davvero compare nell'intestazione del risultato.

## Comportamento

- **Word:** il testo viene sostituito o inserito con una sola operazione, quindi basta un `Ctrl+Z` per annullarla. Vengono mantenuti font, dimensione, colore e stile del paragrafo della selezione; la formattazione mista all'interno della selezione non viene mantenuta. L'effetto digitazione c'è solo nel pannello.
- **Excel:** la formula va nella prima cella della selezione. Le tabelle sono separate da TAB e vengono riempite fino a formare un rettangolo. Se l'area di destinazione contiene già dati, il pannello chiede conferma (secondo clic). Un testo che inizia con `=` viene inserito come testo.
- **Outlook (macro):**
  - JSON corretto con accenti, virgolette ed emoji, con timeout espliciti.
  - Errori distinti per chiave non valida, modello inesistente, quota esaurita e rete.
  - Il riassunto ordina le email per data e dichiara nella bozza se qualcosa è stato troncato.
  - La risposta viene inserita sopra il messaggio citato, mantenendo la firma.
  - Nessun invio automatico: si apre sempre una bozza.
- Il pulsante **Annulla** interrompe la richiesta in corso. Il timeout è di 120 s; gli errori temporanei (429/5xx) vengono ritentati al massimo 2 volte.

## Test rapido senza Office

Apri `public/index.html` con un doppio clic: si apre in modalità *Browser / Test* con le schede Word, Excel e Outlook simulate. Le chiamate ai modelli sono reali.

## Disinstallazione

```powershell
powershell -ExecutionPolicy Bypass -File .\installa.ps1 -Disinstalla
```

Rimuove la registrazione dell'add-in, le impostazioni e le chiavi della macro, e la cache di Office. Il modulo VBA va eliminato a mano.
