> **Stato (18/11/2026, dopo il passaggio all'architettura senza server):**
> - §1 Gateway: superato. Il server non esiste più; restano le chiavi in localStorage e nel registro, che è una scelta accettata (punto 7 non applicato).
> - §2 Word: scrittura unica, quindi Annulla in un passo, e niente digitazione nel documento. **Da fare:** destinazione stabile e gestione dei conflitti.
> - §3 Outlook: JSON corretto, HTML con escape, troncamenti dichiarati, ordine cronologico, timeout, errori distinti, provider configurabile. **Da fare:** richieste asincrone e riassunto a blocchi per thread lunghi.
> - §4 Excel: cella iniziale normalizzata, separatore TAB, matrice rettangolare, conferma prima di sovrascrivere, testo "=" protetto. **Da fare:** dati strutturati in JSON e localizzazione delle formule.
> - §5 Provider: catalogo unico, modello effettivo mostrato, nessun cambio silenzioso, nessun nuovo tentativo su 401/403/404, timeout e Annulla.
> - §6: contenuti trattati come dati nei prompt.
> - §7: README riscritto, script di installazione idempotente.

# Piano di miglioramento — Office AI Copilot

Data: 18 novembre 2026

## Scopo e stato della revisione

Questo documento raccoglie la revisione di `README.md` e `RELAZIONE_PROGETTO.md`, confrontata con i principali flussi implementati in `server.js`, `public/taskpane.js` e `outlook_macro.vba`.

L'architettura con pannello Office, gateway locale e provider intercambiabili è adatta al progetto. La priorità è consolidare sicurezza, integrità dei documenti e gestione degli errori prima di ampliare le funzionalità.

I rilievi derivano da una lettura statica del codice. Non sono stati eseguiti test dentro Office né verifiche live della disponibilità dei modelli. Le soluzioni sotto sono proposte da implementare, non interventi già completati.

## 1. Protezione del gateway locale — P0

### Problemi osservati

- Il server cerca nella radice del progetto i file non trovati in `public/`. Questo rende potenzialmente scaricabili `.env`, certificati e sorgenti.
- `server.listen(PORT)` non limita esplicitamente l'ascolto al loopback.
- CORS permette qualsiasi origine e le chiamate AI non richiedono autenticazione locale.
- Il corpo delle richieste viene accumulato senza un limite esplicito.
- Le chiavi inserite dal pannello vengono conservate nel `localStorage`.
- Le macro Outlook ignorano gli errori del certificato TLS.

### Soluzioni da adottare

1. Eliminare il fallback alla radice. Servire soltanto gli asset autorizzati di `public/`; se necessari, esporre i manifest con rotte esplicite.
2. Verificare che ogni percorso statico risolto resti nella directory pubblica, anche con percorsi codificati, separatori Windows e collegamenti simbolici. Rifiutare file nascosti e percorsi non validi.
3. Configurare l'ascolto esclusivamente su loopback, verificando la risoluzione IPv4/IPv6 di `localhost` nei client Office.
4. Validare `Host` e le origini browser ammesse. Rimuovere CORS aperto; se il pannello lavora sulla stessa origine, non aggiungere autorizzazioni cross-origin inutili.
5. Introdurre un token locale per le operazioni API. Provisionarlo ai client autorizzati senza renderlo recuperabile da un endpoint pubblico. Per VBA, che può non inviare `Origin`, verificare comunque il token: l'assenza di `Origin` non costituisce autenticazione.
6. Applicare limiti configurabili alla dimensione delle richieste e al numero di elaborazioni concorrenti. Restituire errori espliciti per input troppo grandi o server occupato.
7. Conservare le chiavi sul backend. Per la prima versione usare `.env` con accessi locali limitati; successivamente valutare Windows Credential Manager o DPAPI. Il pannello deve mostrare soltanto lo stato della configurazione.
8. Aggiungere esclusioni per `.env`, chiavi private, PFX e log sensibili al controllo versione. Mantenere `.env.example` privo di segreti.
9. Installare correttamente il certificato locale attendibile e rimuovere il bypass TLS dalle macro. Se il certificato manca, interrompere l'avvio con un errore comprensibile.

Se il server è già stato raggiungibile da altri dispositivi, verificare l'eventuale esposizione e valutare la rotazione delle credenziali coinvolte.

### Criteri di accettazione

- Richieste a `/.env`, `/server.js`, `/certs/localhost.pfx` e tentativi di uscita dalla directory pubblica non restituiscono file privati.
- Il servizio non ascolta sulle interfacce di rete esterne.
- Chiamate senza token valido e richieste browser da origini non autorizzate vengono rifiutate.
- Richieste oltre il limite vengono interrotte senza invocare un provider.
- Chiavi e contenuti dei documenti non compaiono nei log.
- Word, Excel e VBA comunicano senza disabilitare la verifica TLS.

## 2. Integrità delle modifiche Word — P1

### Problemi osservati

La conservazione di nome del font, dimensione, colore e stile della selezione non equivale alla conservazione completa della formattazione mista. L'inserimento progressivo usa molte scritture e può lasciare testo parziale in caso di errore. La destinazione viene letta al momento dell'applicazione: spostare la selezione può cambiare il punto modificato.

### Soluzioni da adottare

1. Acquisire testo originale e riferimento alla destinazione quando parte la richiesta. Usare un meccanismo di riferimento supportato dalle API disponibili nell'host e rilasciarlo al termine.
2. Prima di applicare il risultato, verificare che la destinazione esista ancora e che il contenuto non sia cambiato. In caso di conflitto, lasciare il risultato nel pannello e chiedere una nuova selezione.
3. Mostrare anteprima e confronto tra originale e proposta, con azioni distinte: sostituire, inserire sotto, copiare.
4. Applicare il risultato con il minor numero possibile di scritture. Conservare l'originale per il recupero e verificare concretamente il comportamento di Annulla; non presumere che una singola sincronizzazione sia una transazione atomica.
5. Mantenere l'effetto digitazione nel pannello. Disattivarlo per impostazione predefinita dentro Word.
6. Definire il livello di fedeltà supportato: inizialmente testo semplice con stile di paragrafo. Segnalare i casi complessi, come tabelle, campi e formattazione mista.
7. Valutare una gestione per paragrafi o OOXML per casi avanzati. La corrispondenza tra segmenti originali e testo riscritto deve essere definita: una riscrittura libera non consente sempre di trasferire ogni stile in modo univoco.

### Criteri di accettazione

- Spostare il cursore durante la generazione non modifica una destinazione diversa da quella iniziale.
- Modificare il testo originale durante l'attesa produce un conflitto gestito.
- Un errore di generazione non modifica il documento.
- Il recupero dell'originale e Annulla vengono provati su documenti reali.
- Le limitazioni su grassetto, corsivo, elenchi, tabelle e stili misti sono documentate.

## 3. Robustezza di Outlook — P1

### Problemi osservati

- Il JSON viene costruito e letto con sostituzioni e ricerche di virgolette. Il parsing può interrompersi sulle virgolette escapate presenti nell'HTML.
- La conversione Markdown/HTML usa regex e non effettua una sanitizzazione completa.
- Le email vengono troncate a 2.500 o 3.000 caratteri senza evidenziare la perdita di contenuto.
- Le richieste HTTP sono sincrone e le macro non specificano provider e modello.

### Soluzioni da adottare

1. Usare un serializzatore e parser JSON affidabile in VBA, includendolo nel progetto e documentandone provenienza e licenza.
2. Centralizzare chiamata HTTP, autenticazione, timeout e parsing in un modulo VBA condiviso.
3. Distinguere errori di rete, TLS, autenticazione, quota e risposta non valida. Non inserire l'intero JSON di risposta nel corpo della bozza come ripiego.
4. Convertire Markdown con un parser mantenuto e sanitizzare l'HTML con una lista restrittiva di tag e attributi. Per una prima versione semplice, preferire testo correttamente escapato e pochi elementi di formattazione controllati.
5. Evitare il blocco prolungato di Outlook: valutare richieste asincrone oppure un protocollo a job con polling compatibile con VBA. Introdurre comunque timeout espliciti.
6. Dichiarare quante email e quanto contenuto vengono elaborati. Per testi lunghi, segmentare e riassumere per passaggi, mantenendo riferimenti alle email originali.
7. Ordinare esplicitamente la cronologia, contare soltanto gli elementi email validi e gestire citazioni ripetute. Conservare mittente, data e oggetto come metadati strutturati.
8. Usare una configurazione backend comune per provider e modello, con eventuale override esplicito dal client.
9. Conservare il flusso in bozza: l'utente rilegge e invia. Verificare il mantenimento della firma e della conversazione originale.

### Criteri di accettazione

- Virgolette, backslash, accenti, emoji e ritorni a capo attraversano correttamente il flusso JSON.
- HTML con attributi tra virgolette non viene troncato.
- Markup non autorizzato viene rimosso o mostrato come testo.
- Una scadenza in fondo a un'email lunga viene considerata oppure l'esclusione del contenuto viene dichiarata.
- Gli errori non creano bozze contenenti JSON grezzo e non causano invii automatici.

## 4. Inserimento affidabile in Excel — P1

### Problemi osservati

La divisione dei risultati sulle virgole può spezzare numeri decimali e testi. Il ridimensionamento parte dalla selezione corrente anziché da una cella iniziale normalizzata, con possibili dimensioni errate quando sono selezionate più celle.

### Soluzioni da adottare

1. Separare i risultati in tipi espliciti: formula, tabella, spiegazione.
2. Per le tabelle richiedere una matrice JSON rettangolare, validata sul backend e sul client. Usare output strutturati del provider quando disponibili, senza considerarli un sostituto della validazione.
3. Derivare l'intervallo di destinazione dalla cella iniziale e dal numero esatto di righe e colonne.
4. Mostrare un'anteprima dell'area di scrittura. Se contiene dati, chiedere conferma della sovrascrittura oppure permettere una nuova destinazione.
5. Definire una convenzione coerente per formule e localizzazione, verificando le proprietà Office.js usate sul target LTSC.
6. Separare testo letterale e formule: un dataset generato non deve trasformare involontariamente stringhe in formule eseguibili.
7. Gestire intervalli protetti, celle unite e dimensioni massime con errori comprensibili.

### Criteri di accettazione

- Valori come `1,25` e testi contenenti virgole restano integri.
- Una tabella di 3 righe e 2 colonne occupa sempre 3 × 2 celle, anche partendo da una selezione più grande.
- Matrici irregolari vengono rifiutate prima della scrittura.
- Celle occupate non vengono sovrascritte senza una scelta esplicita.
- Formule e testo letterale sono trattati secondo il tipo dichiarato.

## 5. Provider, fallback e gestione delle richieste — P1

### Problemi osservati

I fallback possono cambiare modello, ma la risposta riporta quello richiesto. I tentativi non distinguono adeguatamente le cause di errore. Manca una politica condivisa per timeout, annullamento e limiti. Gli elenchi dei modelli divergono tra documentazione e implementazione.

### Soluzioni da adottare

1. Centralizzare il catalogo modelli: provider, ID API, etichetta, parametri supportati e limiti configurabili. Verificare gli ID sulle fonti ufficiali al momento dell'implementazione.
2. Creare adapter separati per provider e un contratto interno uniforme, indipendente dai formati delle API esterne.
3. Restituire almeno `requestedModel`, `actualModel`, `provider`, `fallbackUsed`, `requestId` e utilizzo token quando disponibile.
4. Definire fallback espliciti e visibili. Non cambiare provider senza una scelta dell'utente, perché cambia anche il destinatario dei dati.
5. Non riprovare errori di autenticazione o richieste malformate cambiando modello. Per errori temporanei, usare tentativi limitati, attesa progressiva e rispetto di `Retry-After` quando presente.
6. Applicare timeout con annullamento della richiesta upstream. Collegare il pulsante Annulla del client al backend; chiarire che annullare non garantisce l'assenza di costi già maturati.
7. Validare provider, modello, tipi dei campi, dimensioni e parametri prima della chiamata esterna.
8. Per contenuti lunghi, stimare il budget del contesto, segmentare quando necessario e rendere visibili le parti escluse.
9. Registrare metadati operativi, durata e categorie di errore senza salvare per impostazione predefinita prompt, documenti o chiavi.

### Criteri di accettazione

- L'interfaccia mostra il modello effettivamente utilizzato.
- Errori 401/403 non attivano una catena di tentativi inutili.
- Timeout, quota esaurita e annullamento hanno messaggi distinti.
- Un provider non valido viene rifiutato invece di essere trattato implicitamente come Gemini.
- Ogni richiesta ha un identificatore utile alla diagnosi.

## 6. Trattamento dei contenuti e qualità delle risposte — P2

### Soluzioni da adottare

- Mostrare prima dell'elaborazione la fonte dei dati: selezione, intero documento, intervallo o email selezionate, insieme al provider destinatario.
- Preferire il contesto minimo necessario. Rendere esplicita l'opzione di inviare l'intero documento.
- Trattare istruzioni contenute nei documenti e nelle email come dati, non come comandi dell'utente. Separare istruzioni e fonti nei prompt; questa misura riduce il rischio, ma non garantisce da sola l'immunità alla prompt injection.
- Per riepiloghi e Q&A, associare affermazioni e scadenze a riferimenti riconoscibili: email, sezioni o paragrafi.
- Distinguere informazioni presenti nelle fonti, inferenze e informazioni mancanti.
- Aggiungere preset di attività con istruzioni mantenute in un solo punto, evitando duplicazioni tra client.

### Criteri di accettazione

- L'utente comprende quali dati vengono inviati e a chi.
- Una fonte che contiene istruzioni estranee non cambia il comportamento autorizzato del programma.
- I riepiloghi consentono di risalire alle fonti e non presentano come certe scadenze assenti dal materiale.

## 7. Documentazione e installazione — P1/P2

### Correzioni necessarie

- Il README presenta Outlook tramite manifest; la relazione descrive il percorso VBA per POP/SMTP. Documentare separatamente i due percorsi e il loro stato di verifica.
- Allineare nomi dei modelli, branding del progetto e descrizione del server multi-provider.
- Sostituire “perfetto”, “impeccabile” e “100%” con capacità precise e risultati di prove riproducibili.
- Eliminare percorsi assoluti personali e link `file:///`, usando collegamenti relativi.
- Distinguere animazione locale della risposta da streaming effettivo del provider.
- Non qualificare `.env` come deposito sicuro né presentare il gateway locale come elaborazione interamente locale: i contenuti vengono inviati al cloud.
- Verificare e datare eventuali affermazioni su prezzi, quote gratuite e compatibilità. Evitare promesse generiche sul costo per documento.

### Struttura proposta

| Documento | Contenuto |
| --- | --- |
| `README.md` | Scopo, funzioni disponibili, avvio rapido, collegamenti alle guide |
| `docs/INSTALLAZIONE.md` | Prerequisiti, certificati, sideloading, VBA, aggiornamento e rimozione |
| `docs/ARCHITETTURA.md` | Componenti, flussi, contratto API e gestione dei segreti |
| `docs/COMPATIBILITA.md` | Host, versioni, account, API richieste, prove e limiti noti |
| `docs/PRIVACY_E_SICUREZZA.md` | Destinatari dei dati, storage locale, log e protezioni del gateway |
| `docs/VERIFICHE.md` | Scenari manuali, controlli automatici e risultati |
| `RELAZIONE_PROGETTO.md` | Stato del progetto con risultati dimostrati e limiti |
| `PIANO_MIGLIORAMENTO.md` | Priorità e criteri di completamento degli interventi |

### Installazione riproducibile

Fornire una procedura idempotente che controlli versione Node supportata, porta disponibile, certificato e configurazione. La procedura deve segnalare cosa manca e indicare come rimuovere registrazioni e certificati creati per il progetto. Non deve modificare impostazioni Office globali senza necessità.

## 8. Organizzazione del codice e verifiche — P2

Evitare una riscrittura completa. Estrarre progressivamente i componenti mentre si correggono i flussi:

```text
server.js                  # Avvio e collegamento dei componenti
src/config/                # Configurazione e catalogo modelli
src/http/                  # Routing, autenticazione, limiti e file statici
src/providers/             # Adapter Gemini, OpenAI, Anthropic
src/services/              # Orchestrazione, timeout e fallback
src/rendering/             # Conversione e sanitizzazione degli output
public/hosts/              # Operazioni specifiche Word, Excel, Outlook
public/ui/                 # Stato del pannello, anteprima e messaggi
tests/                     # Verifiche delle proprietà critiche
docs/                      # Guide operative
```

Mantenere poche dipendenze è utile, ma non sostituire parser e sanitizzatori affidabili con regex fragili soltanto per dichiarare “zero dipendenze”. Se si introducono librerie, fissarne le versioni tramite lockfile e verificarne licenza e manutenzione.

Automatizzare soprattutto le verifiche che proteggono proprietà reali: accesso ai file, autenticazione, limiti, parsing, sanitizzazione e fallback. Simulare le risposte dei provider per evitare costi e dipendenze dalla rete nei test ordinari. Per integrità del documento, rendering e compatibilità, affiancare prove manuali riproducibili negli host Office reali.

## 9. Ordine operativo consigliato

| Fase | Interventi | Condizione di completamento |
| --- | --- | --- |
| 1 — Gateway | File statici, loopback, autenticazione, segreti, TLS e limiti | Nessun file privato esposto; accessi non autorizzati rifiutati |
| 2 — Affidabilità comune | Contratto API, validazione, timeout, errori e modello effettivo | Comportamento osservabile e prevedibile con provider simulati |
| 3 — Word | Destinazione stabile, conflitti, anteprima e applicazione controllata | Modifiche verificabili e recuperabili su documenti reali |
| 4 — Outlook ed Excel | JSON, HTML, testi lunghi, matrici e destinazioni | Superati gli scenari di errore e integrità descritti sopra |
| 5 — Distribuzione | Documentazione, installazione e matrice di compatibilità | Installazione ripetibile con istruzioni allineate al codice |
| 6 — Evoluzione | Streaming reale, preset, riferimenti alle fonti e UX avanzata | Nuove funzioni senza regressioni sui flussi consolidati |

La prima versione consolidata dovrebbe privilegiare Word e il gateway condiviso. Outlook ed Excel possono evolvere sulle stesse fondamenta una volta stabilizzati autenticazione, contratto delle risposte ed errori.
