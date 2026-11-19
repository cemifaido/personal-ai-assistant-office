# Personal AI Assistant

**Assistente AI per Word, Excel e Outlook per Windows** - versione beta 0.1.0

Usa il modello che preferisci - **Google Gemini, OpenAI o Anthropic Claude** - con la tua chiave API, direttamente dentro Office:

- **Word**: riscrivi con il tono scelto, correggi le bozze, riassumi, traduci, genera testo e fai domande sull'intero documento.
- **Excel**: formule in italiano a partire da una descrizione, analisi dei dati selezionati (con statistiche esatte), grafici e tabelle di esempio.
- **Outlook classico**: riassunto di più email selezionate e bozze di risposta (conferma, declina, chiarimenti o su misura), anche con account POP e IMAP.

I testi vanno **direttamente** dal tuo PC al provider scelto: non c'è un server del progetto. Le chiavi restano sul tuo PC.

## Installazione

Segui la **[guida illustrata](https://pai-office-automation.github.io/personal-ai-assistant-office/installa.html)**. In breve:

1. Scarica `PersonalAIAssistant-<versione>-Setup.exe` dalla pagina [Releases](../../releases/latest).
2. L'installatore non è ancora firmato: clic destro sul file → *Proprietà* → **Annulla blocco**; se compare *«Windows ha protetto il PC»* → *Ulteriori informazioni* → **Esegui comunque**.
3. Avanti, avanti, fine. Per Word ed Excel Windows chiede **una volta** una conferma da amministratore (crea una cartella condivisa locale in sola lettura con il solo file di configurazione).
4. In Word e in Excel, **una sola volta**: *Home → Componenti aggiuntivi → Altri componenti aggiuntivi → Cartella condivisa → Personal AI Assistant → Aggiungi*. Poi trovi la scheda **Personal AI**.
5. Inserisci la chiave API: il pannello te la chiede al primo avvio (*Ottieni chiave* porta al sito del provider).

Verifica dell'integrità del file scaricato:

```powershell
Get-FileHash .\PersonalAIAssistant-0.1.0-Setup.exe -Algorithm SHA256
```

Il valore deve coincidere con quello pubblicato nella Release.

**Requisiti:** Windows 10/11, Office desktop per Windows (provato su LTSC 2024 a 32 bit), .NET Framework 4.8 per Outlook.
**Non supportati:** nuovo Outlook per Windows, Outlook sul web, Office per Mac.

## Privacy

Leggi l'**[informativa](https://pai-office-automation.github.io/personal-ai-assistant-office/privacy.html)**. Punti essenziali:

- i testi selezionati, i documenti (solo se lo chiedi) e le email selezionate vengono inviati al provider scelto, secondo le **sue** condizioni: alcuni piani gratuiti possono usare i dati per migliorare i servizi, verificale prima di elaborare documenti riservati;
- nessuna telemetria, nessuna email inviata in automatico (Outlook apre sempre una bozza).

## Limiti noti della beta

- I risultati dell'AI possono contenere errori: rileggili sempre.
- Sui PC aziendali senza diritti di amministratore la parte Word/Excel può richiedere l'IT; Outlook si installa comunque.
- Provato solo su Office LTSC 2024; altre edizioni non ancora verificate.

## Per chi sviluppa

Architettura, build, firma e rilascio: [docs/SVILUPPO.md](docs/SVILUPPO.md). Storia del progetto: [CHANGELOG.md](CHANGELOG.md).

## Licenza

[MIT](LICENSE). Progetto indipendente, non affiliato a Microsoft, Google, OpenAI o Anthropic.
