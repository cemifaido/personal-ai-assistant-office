# Sviluppo

## Struttura

| Percorso | Contenuto |
| --- | --- |
| `public/` | Pannello Word/Excel/PowerPoint (Office.js), pubblicato su GitHub Pages: `index.html`, `taskpane.js`, `ai.js` (chiamate ai provider e **catalogo modelli**), pagine `installa.html` e `privacy.html`. |
| `manifest-word-excel.template.xml` | Manifest Word/Excel con segnaposto `__BASE_URL__`; versione nel tag `<Version>`. |
| `outlook/PersonalAIAssistant.Outlook/` | Add-in Outlook VSTO (C#, .NET Framework 4.8). Il catalogo modelli in `ModelCatalog.cs` va tenuto allineato a `public/ai.js`. |
| `installer/` | Script Inno Setup, `condivisione.ps1` (unica parte eseguita come amministratore), testi e icona. |
| `build.ps1` | Build della release: add-in Outlook, manifest, installatore, SHA-256. **Unico punto con gli indirizzi** di sito e repository. |
| `.github/workflows/pages.yml` | Pubblica `public/` su GitHub Pages e genera `manifest.xml` con l'indirizzo reale. |
| `spike/` | Prova tecnica di installazione (non distribuita), conservata per documentare le scelte. |
| `docs/` | Piano di miglioramento (revisione di Codex), prototipo storico, macro VBA storica. |

## Architettura

```
Word / Excel ──(catalogo \\localhost\PersonalAIAssistant$ → manifest)──► pannello su GitHub Pages ──HTTPS──► provider
Outlook classico ──(HKCU\...\Outlook\Addins + VSTO)──► PersonalAIAssistant.Outlook.dll ───────────HTTPS──► provider
```

- Nessun backend: le chiamate partono dal pannello (CORS consentito da Gemini, OpenAI e Anthropic con l'header `anthropic-dangerous-direct-browser-access`) e dall'add-in .NET.
- Chiavi: `localStorage` del pannello (prefisso `paa.key.`), DPAPI per Outlook (`%APPDATA%\PersonalAIAssistant\outlook-settings.json`).
- Nessun cambio silenzioso di modello o provider; nuovi tentativi solo per 408/429/5xx; Claude Opus/Sonnet 5.5 con `fallbacks: "default"` (il modello effettivo è mostrato).

### Perché il catalogo in cartella condivisa

Verificato con lo spike (`spike/`) su utenti Windows standard e Office LTSC 2024:

| Metodo Word/Excel | Esito |
| --- | --- |
| `HKCU\...\WEF\Developer` (sideload "sviluppatore") | Errore «caricamento dei componenti aggiuntivi» all'avvio; l'add-in parte solo cliccando sull'errore. Scartato. |
| Catalogo `TrustedCatalogs` su percorso UNC | Funziona; dopo il primo *Aggiungi* la scheda resta a ogni avvio. Scelto. |
| Pre-inserimento scrivendo `Wef\AddinInfo\...\boot.json` | Ignorato da Office. Scartato (non documentato). |

Office accetta cataloghi solo da percorsi di rete: la condivisione locale in sola lettura richiede un'elevazione, eseguita solo per `installer/condivisione.ps1` (l'installatore gira senza privilegi, così le impostazioni finiscono nel profilo dell'utente giusto).

## Build della release

Requisiti: Visual Studio 2026 con il carico di lavoro **Sviluppo di Microsoft 365** (componente *Visual Studio Tools per Office*), Inno Setup 6, certificato di firma.

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1 -Versione 0.1.0
```

Produce `release\PersonalAIAssistant-0.1.0-Setup.exe` e `release\SHA256SUMS.txt`. Prima di cambiare versione aggiornare: `<Version>` nel manifest, `PAA_VERSION` in `public/ai.js`, `ApplicationVersion` nel `.csproj`, `AssemblyVersion` in `AssemblyInfo.cs`, `CHANGELOG.md`.

### Certificato di firma dei manifest VSTO

I manifest VSTO devono essere firmati. Si usa un certificato del progetto, `CN=Personal AI Assistant`, nell'archivio *CurrentUser\My* di chi compila; il suo thumbprint è nel `.csproj` e la sua chiave pubblica finisce nell'elenco di inclusione VSTO scritto dall'installatore.

- **Va conservato**: con un certificato diverso, gli utenti già installati vedrebbero di nuovo l'avviso di fiducia di VSTO.
- Esportarlo una volta in un `.pfx` protetto da password, **fuori dal repository** (i `.pfx` sono esclusi da `.gitignore`).
- Non è una firma Authenticode: l'installatore resta non firmato (avvisi SmartScreen documentati nella guida).

## Pubblicazione

1. `public/` si pubblica da solo a ogni push su `main` (workflow *Pubblica pannello*): prima il job **verifica**
   (sintassi JavaScript, `node --test tests/*.test.cjs`, manifest generato e validato), poi il job **pubblica**, che
   parte solo se la verifica è passata e pubblica esattamente il pacchetto verificato.
2. La release si crea a mano su GitHub con il Setup e lo SHA-256 da `release\`.
3. Aggiornare la data e la versione in `CHANGELOG.md`.

Il pannello mostra in basso versione e commit pubblicato (es. `v0.2.1 · 1a2b3c4d`): è il riferimento da chiedere
in caso di segnalazione.

### Versioni dei componenti

I componenti avanzano separatamente; la compatibilità è questa:

| Componente | Dove si aggiorna | Versione attuale | Note |
|---|---|---|---|
| Pannello Word/Excel/PowerPoint | Pages, a ogni push (tutti gli utenti alla prossima apertura) | `PAA_VERSION` in `public/ai.js` + commit | Funziona con qualsiasi manifest `0.1.0.0` |
| Manifest Word/Excel/PowerPoint | Setup (cartella condivisa) | `0.1.0.0` | Cambiarlo solo se serve: richiede di nuovo la conferma da amministratore |
| Add-in Outlook | Setup | uguale al programma | Indipendente dal pannello |
| Installatore | Release | uguale al programma | |

Il pannello deve restare compatibile con i manifest già installati: un'aggiunta che richieda un manifest nuovo va
pubblicata insieme a un nuovo Setup.

### Tornare a una versione precedente del pannello

Il pannello pubblicato arriva subito a tutti: se una pubblicazione introduce un problema, si ripristina così.

1. Individuare il commit difettoso (il pannello lo mostra in basso; `git log -- public/`).
2. Annullarlo con un nuovo commit, senza riscrivere la storia (il ramo `main` è protetto):
   ```
   git revert <commit>
   git push origin main
   ```
3. Il workflow rifà la verifica e ripubblica il pannello precedente (circa un minuto). Controllare online il commit
   mostrato dal pannello o il contenuto di `ai.js`.
4. Se il problema era in una Release (Outlook o installatore), segnarla come non consigliata nelle note e ripubblicare
   la Release precedente come *latest*; i Setup restano scaricabili dalla pagina delle Release.

La cronologia delle pubblicazioni del pannello è nella scheda *Actions* (workflow *Pubblica pannello*) e
nell'ambiente *github-pages* del repository.

## Prove manuali prima di una release

Prima delle prove dentro Office, eseguire anche i controlli del salvataggio Outlook con Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/outlook-settings.tests.ps1
```

Il test compila i sorgenti delle impostazioni e del catalogo in un ambiente isolato: usa una cartella temporanea
e un percorso di registro dedicato, senza leggere o modificare le impostazioni dell'utente.
Verifica sostituzione con backup, rifiuto di JSON incompleti, compatibilita' dei file precedenti e recupero.
Le impostazioni recuperate devono essere controllate e salvate prima di usarle per richieste o verifiche delle chiavi:
un backup puo' contenere chiavi, modelli o destinatari precedenti.

- Utente Windows **standard** pulito: installazione, richiesta di elevazione, *Aggiungi* in Word ed Excel, scheda presente dopo riavvio.
- Outlook: nessun avviso di fiducia, riassunto di più email con accenti ed elenchi, risposta con firma e citazione intatte.
- PowerPoint: testo selezionato (sostituisci e nuova casella), creazione di slide da scaletta, miglioramento della slide selezionata, Ctrl+Z.
- Excel: formula per riga e calcolo unico, tabella, grafico per intervalli e per conteggio, conferma di sovrascrittura.
- Disinstallazione: schede rimosse, condivisione rimossa se richiesto.
