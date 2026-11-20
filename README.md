# Personal AI Assistant

> 🇮🇹 **Italiano** | 🇬🇧 **[English](README.en.md)**

**Assistente AI integrato dentro Microsoft Word, Excel, PowerPoint e Outlook classico per Windows** — versione **beta 0.2.1**

## Cos'è e come funziona

**Personal AI Assistant** è un componente aggiuntivo open source che porta l'intelligenza artificiale **direttamente dentro la barra multifunzione di Microsoft Office** (scheda *Personal AI*), senza dover fare copia-incolla verso chat esterne e senza passare da server intermedi.

**Come funziona in pratica (in 3 passi):**
1. **Apri il documento o l'email in Office** (Word, Excel, PowerPoint o Outlook classico) e seleziona il testo, le celle, la slide o le email su cui vuoi lavorare.
2. **Scegli un'azione dal pannello *Personal AI*** (riscrivi, correggi, traduci, riassumi, genera formule/grafici, crea slide o prepara una bozza di risposta).
3. **Il tuo PC chiama direttamente il provider AI che hai scelto** (con la tua chiave API personale, oppure un modello locale come Ollama / `llama-server` / LM Studio) e ti mostra il risultato pronto da rileggere prima di applicarlo al documento o salvarlo come bozza.

> [!IMPORTANT]
> **Serve una tua chiave API personale (modello BYOK).** Personal AI Assistant non è model-free e non funziona senza almeno una chiave API di **Google Gemini, OpenAI, Anthropic Claude, Qwen, MiniMax, DeepSeek o OpenRouter** (in Outlook è supportato anche un modello locale senza chiave; nel pannello in via sperimentale).
> La chiave **non è inclusa** nell'installazione: la ottieni tu dal sito del provider ed è legata al **tuo account e al tuo contratto** con quel provider (piano gratuito o a pagamento, limiti, costi e condizioni d'uso sono quelli del provider).
> Senza chiave il pannello si apre ma non può elaborare nulla. Per iniziare gratis il modo più semplice è una chiave [Google AI Studio](https://aistudio.google.com/apikey).

## A cosa serve in dettaglio e con quali connettori lavora

Personal AI Assistant aggiunge un **assistente AI nativo direttamente dentro l'interfaccia di Microsoft Office**, così da elaborare testi, dati, slide ed email senza uscire dalle applicazioni e senza dover usare chat o agenti esterni:

- **Word**: riscrivi con il tono scelto, correggi le bozze, riassumi, traduci, genera testo e fai domande sull'intero documento.
- **Excel**: formule in italiano a partire da una descrizione, analisi dei dati selezionati (con statistiche esatte), grafici e tabelle di esempio.
- **PowerPoint**: riscrivi, correggi, accorcia o traduci il testo selezionato, crea slide da una scaletta e migliora i testi della slide selezionata.
- **Outlook classico**: riassunto di più email selezionate e bozze di risposta (conferma, declina, chiarimenti o su misura), anche con account POP e IMAP.

**Connettori utilizzati:**
- **Verso Office:** pannello web basato su **Office.js** (eseguito nel motore WebView2 di Office e caricato via manifest XML locale + pagine statiche su GitHub Pages) per **Word, Excel e PowerPoint**; add-in nativo **VSTO (.NET Framework 4.8)** integrato nella barra *Home* di **Outlook classico**. Non gestisce file PDF/Acrobat.
- **Verso l'AI:** chiamate **HTTPS REST dirette** dal tuo PC alle API di **Google Gemini, OpenAI, Anthropic Claude, Qwen (Alibaba), MiniMax, DeepSeek o OpenRouter**, oppure chiamate HTTP/HTTPS verso un **server locale compatibile OpenAI** (**Ollama, `llama-server`, LM Studio** — pienamente supportato in Outlook, sperimentale nel pannello Office). Non esiste alcun server intermedio del progetto e le chiavi restano sul tuo PC.

## Differenze rispetto a `mcp4office`

Entrambi i progetti permettono di interagire con Microsoft Office su Windows tramite AI, ma con architetture e requisiti molto diversi:

| Aspetto | Personal AI Assistant (questo progetto) | `mcp4office` |
| --- | --- | --- |
| **Dove si usa** | **Dentro Office**, come scheda *Personal AI* (pannello laterale in Word/Excel/PowerPoint e pulsanti in Outlook) | **Dall'esterno di Office**, dalla chat di un client AI compatibile con MCP (Antigravity, Claude Desktop/Code, Codex, ecc.) |
| **Modello AI e costi** | **BYOK via API dirette**: richiede una tua chiave API (a consumo o piano gratuito del provider) oppure un modello locale (Ollama / `llama-server` / LM Studio) | **Model-free**: nessun modello integrato e nessuna chiave API; sfrutta l'abbonamento del tuo client MCP |
| **Connettori verso l'AI** | **HTTPS REST diretti** verso Gemini, OpenAI, Claude, Qwen, MiniMax, DeepSeek, OpenRouter o endpoint locale OpenAI-compatibile | **MCP (JSON-RPC 2.0 su `stdio`)** verso client/agenti AI compatibili |
| **Connettori verso Office** | **Office.js** (WebView2) per Word, Excel, PowerPoint + add-in **VSTO (.NET 4.8)** per Outlook classico (nessun supporto PDF/Acrobat) | **COM locale** su istanze aperte di Word, Excel, PowerPoint, Outlook classico, Acrobat + lettura diretta **PDF locali** |
| **Tipo di utilizzo** | Azioni **rapide e guidate** sul documento o email aperta, con interfaccia a pulsanti senza uscire da Office | Flussi **agentici e cross-app** (confronto tra documenti/fogli, lettura PDF, compilazione multi-file in chat) |
| **Diritti di installazione** | Profilo utente per Outlook; richiede **1 conferma amministratore (UAC)** la prima volta per la cartella condivisa del manifest di Word/Excel/PowerPoint | 100% profilo utente, **nessun UAC / amministratore** |

In sintesi: scegli **Personal AI Assistant** se vuoi un assistente con pulsanti integrati direttamente in Office usando la tua chiave API (o un modello locale); scegli **[`mcp4office`](https://github.com/pai-office-automation/mcp4office/releases/tag/v0.5.7)** se preferisci pilotare Office e leggere PDF dal tuo client MCP (Antigravity, Claude Desktop, Codex) senza gestire chiavi API.

## Installazione

Segui la **[guida illustrata](https://pai-office-automation.github.io/personal-ai-assistant-office/installa.html)**. In breve:

1. Scarica `PersonalAIAssistant-<versione>-Setup.exe` dalla pagina [Releases](../../releases/latest).
2. L'installatore non è ancora firmato: clic destro sul file → *Proprietà* → **Annulla blocco**; se compare *«Windows ha protetto il PC»* → *Ulteriori informazioni* → **Esegui comunque**.
3. Avanti, avanti, fine. Per Word, Excel e PowerPoint Windows chiede una conferma da amministratore la prima volta (e di nuovo solo se un aggiornamento cambia il file di configurazione) (crea una cartella condivisa locale in sola lettura con il solo file di configurazione).
4. In Word, Excel e PowerPoint, **una sola volta** in ciascuno: *Home → Componenti aggiuntivi → Altri componenti aggiuntivi → Cartella condivisa → Personal AI Assistant → Aggiungi*. Poi trovi la scheda **Personal AI**.
5. Inserisci **la tua** chiave API (obbligatoria, non inclusa): il pannello te la chiede al primo avvio (*Ottieni chiave* porta al sito del provider).

Verifica dell'integrità del file scaricato:

```powershell
Get-FileHash .\PersonalAIAssistant-0.2.1-Setup.exe -Algorithm SHA256
```

Il valore deve coincidere con quello pubblicato nella Release.

**Requisiti:** Windows 10/11, Office desktop per Windows, .NET Framework 4.8 per Outlook.
**Non supportati:** nuovo Outlook per Windows, Outlook sul web, Office per Mac.

### Cosa è stato provato

Ambiente: Windows 11, Office LTSC 2024 **a 32 bit**, utente amministratore e utente standard.

| Applicazione | Stato |
| --- | --- |
| Installazione, catalogo condiviso, *Aggiungi* | Provati su utente standard |
| Word ed Excel (pannello, formule, analisi) | Provati in Office |
| Excel - grafici | Provati in Excel |
| PowerPoint | Creazione slide provata in PowerPoint 2024 |
| Outlook (add-in) | Provato in Outlook classico |
| Qwen, MiniMax, DeepSeek, OpenRouter (0.2.0) | Indirizzi e accesso dal pannello verificati; non ancora provati con una chiave |
| Outlook con modello locale (0.2.0) | Provato in Outlook classico (llama-server); chiamate provate anche con Ollama |
| Office a 64 bit, Microsoft 365, 2016-2021 | Non provati |

Le funzioni di PowerPoint *Crea slide* e *Migliora slide* richiedono le API PowerPoint 1.4 e 1.5 (Office 2024 o Microsoft 365): con versioni precedenti il pannello lo segnala.

## Privacy

Leggi l'**[informativa](https://pai-office-automation.github.io/personal-ai-assistant-office/privacy.html)**. Punti essenziali:

- i testi selezionati, i documenti (solo se lo chiedi) e le email selezionate vengono inviati al provider scelto, secondo le **sue** condizioni: alcuni piani gratuiti possono usare i dati per migliorare i servizi, verificale prima di elaborare documenti riservati;
- **Qwen (Alibaba), MiniMax e DeepSeek sono aziende con sede in Cina**: prima di salvarne la chiave il programma chiede conferma; con un modello locale i testi restano nella tua rete;
- nessuna telemetria, nessuna email inviata in automatico (Outlook apre sempre una bozza).

## Limiti noti della beta

- I risultati dell'AI possono contenere errori: rileggili sempre.
- Sui PC aziendali senza diritti di amministratore la parte Word/Excel può richiedere l'IT; Outlook si installa comunque.
- Vedi la tabella *Cosa è stato provato*: diverse combinazioni non sono ancora verificate.

## Per chi sviluppa

Architettura, build, firma e rilascio: [docs/SVILUPPO.md](docs/SVILUPPO.md). Storia del progetto: [CHANGELOG.md](CHANGELOG.md).

## Licenza

[MIT](LICENSE). Progetto indipendente, non affiliato a Microsoft, Google, OpenAI o Anthropic.
