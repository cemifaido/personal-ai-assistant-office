# Personal AI Assistant

**Assistente AI per Word, Excel, PowerPoint e Outlook per Windows** - versione beta 0.2.0

> [!IMPORTANT]
> **Serve una tua chiave API personale.** Personal AI Assistant non funziona senza almeno una chiave API di **Google Gemini, OpenAI, Anthropic Claude, Qwen, MiniMax, DeepSeek o OpenRouter** (in Outlook anche un modello locale, senza chiave).
> La chiave **non è inclusa** nell'installazione: la ottieni tu dal sito del provider ed è legata al **tuo account e al tuo contratto** con quel provider (piano gratuito o a pagamento, limiti, costi e condizioni d'uso sono quelli del provider).
> Senza chiave il pannello si apre ma non può elaborare nulla. Per iniziare gratis il modo più semplice è una chiave [Google AI Studio](https://aistudio.google.com/apikey).

Usa il modello che preferisci - **Google Gemini, OpenAI, Anthropic Claude, Qwen, MiniMax, DeepSeek o OpenRouter** - con la tua chiave API, direttamente dentro Office:

- **Word**: riscrivi con il tono scelto, correggi le bozze, riassumi, traduci, genera testo e fai domande sull'intero documento.
- **Excel**: formule in italiano a partire da una descrizione, analisi dei dati selezionati (con statistiche esatte), grafici e tabelle di esempio.
- **PowerPoint**: riscrivi, correggi, accorcia o traduci il testo selezionato, crea slide da una scaletta e migliora i testi della slide selezionata.
- **Outlook classico**: riassunto di più email selezionate e bozze di risposta (conferma, declina, chiarimenti o su misura), anche con account POP e IMAP.

I testi vanno **direttamente** dal tuo PC al provider scelto: non c'è un server del progetto. Le chiavi restano sul tuo PC.

## Installazione

Segui la **[guida illustrata](https://pai-office-automation.github.io/personal-ai-assistant-office/installa.html)**. In breve:

1. Scarica `PersonalAIAssistant-<versione>-Setup.exe` dalla pagina [Releases](../../releases/latest).
2. L'installatore non è ancora firmato: clic destro sul file → *Proprietà* → **Annulla blocco**; se compare *«Windows ha protetto il PC»* → *Ulteriori informazioni* → **Esegui comunque**.
3. Avanti, avanti, fine. Per Word, Excel e PowerPoint Windows chiede una conferma da amministratore la prima volta (e di nuovo solo se un aggiornamento cambia il file di configurazione) (crea una cartella condivisa locale in sola lettura con il solo file di configurazione).
4. In Word, Excel e PowerPoint, **una sola volta** in ciascuno: *Home → Componenti aggiuntivi → Altri componenti aggiuntivi → Cartella condivisa → Personal AI Assistant → Aggiungi*. Poi trovi la scheda **Personal AI**.
5. Inserisci **la tua** chiave API (obbligatoria, non inclusa): il pannello te la chiede al primo avvio (*Ottieni chiave* porta al sito del provider).

Verifica dell'integrità del file scaricato:

```powershell
Get-FileHash .\PersonalAIAssistant-0.2.0-Setup.exe -Algorithm SHA256
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
| Outlook con modello locale (0.2.0) | Chiamate provate con llama-server e Ollama; finestra Impostazioni da provare in Outlook |
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
