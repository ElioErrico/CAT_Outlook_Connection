# CheshireCat VBA for Outlook

Automatizza le risposte alle email di Outlook usando il backend CheshireCat direttamente dall’editor di messaggi (WordEditor).
I moduli convertono automaticamente eventuali tabelle Word ⇄ Markdown, inviano il testo (con contesto del thread) all’API, e inseriscono la risposta formattata nel corpo dell’email.

## Repository:

- modCheshireCatApi – chiamate HTTP (auth, /message, clear history).

- modCheshireCatCore – logica core: costruzione payload, conversione Markdown⇄Word, inserimento risposta.

- modMarkdownHelpers – funzioni helper per Markdown, sanitizzazione testo, formattazioni inline.

- modOutlookCheshireCat – macro “pronte all’uso” per Outlook (usa il WordEditor della finestra di composizione).

- modPromptHelper – costanti per il prompt (contesto+richiesta).

## Requisiti

Windows + Outlook 2016 o superiore (32/64 bit).

Word come editor di posta (predefinito su Outlook 2016+).

Backend CheshireCat raggiungibile da Outlook (LAN/localhost o HTTPS).

Riferimenti VBA (VBE → Strumenti → Riferimenti):

✅ Visual Basic for Applications

✅ Microsoft Outlook xx.0 Object Library

✅ Microsoft Word xx.0 Object Library

✅ Microsoft Office xx.0 Object Library

✅ Microsoft XML, v6.0 (MSXML2) ← necessario per le HTTP call

Se “Microsoft XML, v6.0” non fosse disponibile, puoi sostituire nelle CreateObject con "MSXML2.XMLHTTP" (ma ServerXMLHTTP.6.0 è consigliato).

## Installazione

- Apri Outlook.

- Premi Alt+F11 per aprire l’Editor VBA (VBE).

- Menu File → Importa file… e importa i 5 moduli (.bas)

- In VBE: Strumenti → Riferimenti… spunta i riferimenti elencati sopra.

- Salva (Ctrl+S) e chiudi il VBE.

## Configurazione

Apri modCheshireCatApi e imposta i parametri dell’endpoint:


=========== CONFIG ===========
Public Const DEFAULT_URL As String = "http://localhost:1865"

Public Const DEFAULT_USERNAME As String = "admin"

Public Const DEFAULT_PASSWORD As String = "admin"

## Abilitare le macro in Outlook

In Outlook: File → Opzioni → Centro protezione → Impostazioni Centro protezione → Impostazioni Macro
Scegli una delle opzioni (per sviluppo è più semplice la prima):

Attiva notifiche per tutte le macro (consigliato in sviluppo), oppure

Consenti solo macro firmate (richiede la firma del progetto VBA).

Spunta anche, se necessario: “Considera attendibile l’accesso al modello a oggetti del progetto VBA”.

Firma del progetto e installazione certificati (se usi “solo macro firmate”)
1) Ottieni/crea un certificato di firma

Opzione A – Certificato aziendale/pubblico (consigliato in produzione)
Usa un certificato di Code Signing fornito dall’IT/PKI aziendale (o da una CA pubblica).

Opzione B – SelfCert (sviluppo/test)
Crea un certificato locale con SelfCert.exe (installato con Office):

Percorso tipico:
C:\Program Files\Microsoft Office\root\Office16\SELFCERT.EXE
(su Office 2019/2021/365 è analogo; se 32-bit su OS 64-bit: Program Files (x86))

Avvia SelfCert.exe, dai un nome (es. “VBA CheshireCat Dev”) e conferma.

2) Firma il progetto VBA

Apri Outlook → Alt+F11 (Editor VBA).

Menu Strumenti → Firma digitale….

Scegli il certificato → OK.

Salva il progetto, chiudi e riapri Outlook.

Importante: qualsiasi modifica al codice invalida la firma → dopo ogni modifica, rifirma con gli stessi passaggi.


## Aggiungere i bottoni alla Barra di Accesso Rapido

File → Opzioni → Barra di accesso rapido.

“Scegli comandi da” → Macro.

Aggiungi queste macro:

modOutlookCheshireCat.CCAT_InserisciRispostaDaSelezione_Outlook

modOutlookCheshireCat.CCAT_CancellaCronologia_Outlook

(Opzionale) Modifica per assegnare icona/nome visibile.

OK per salvare.




