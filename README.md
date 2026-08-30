# Report mensile georeferenziazione ANNCSU

Ricostruisce in automatico su GitHub Actions la tabella "Georeferenziazione
accessi ANNCSU": per ogni comune quanti accessi sono in archivio e quanti hanno
le coordinate, con il confronto rispetto al mese precedente.

Nessun file va caricato a mano: il workflow scarica i dati dal sito ufficiale a
ogni corsa.

## Da dove arrivano i dati

- **Accessi e coordinate**: i 20 [indirizzari regionali in open data](https://www.anncsu.gov.it/it/consultazione-dellarchivio/open-data/Accedi-ai-servizi-di-dowload-massivo-in-Open-data/)
  dell'ANNCSU, file ZIP con dentro un CSV rigenerato ogni mese. Ogni riga del CSV
  è un accesso e conta come georeferenziato se ha longitudine e latitudine valorizzate.
- **Denominazioni di comune, provincia e regione**: l'[elenco Istat dei comuni](https://www.istat.it/storage/codici-unita-amministrative/Elenco-comuni-italiani.xlsx),
  perché l'indirizzario contiene solo i codici. La copia in `data/lookup/comuni-istat.csv`
  fa da riserva se Istat non risponde, e conserva il codice IPA che Istat non pubblica.

L'aggancio fra le due fonti usa il **codice catastale** (`A018`) e non il codice
Istat: dal 1° gennaio 2026 il codice Istat è cambiato per tutti i comuni sardi e
le due fonti possono allinearsi in momenti diversi.

## Struttura

```
.github/workflows/report-anncsu.yml   il workflow schedulato
tools/
  anncsu_report.py                    scarica, conta, confronta, scrive
  requirements.txt                    openpyxl
data/lookup/
  comuni-istat.csv                    anagrafica di riserva + codici IPA
docs/                                 ← anche il sito GitHub Pages
  index.html                          pagina unica con la ricerca per comune
  dati/comuni.json                    7.894 comuni, un comune per riga
  .nojekyll                           evita che Pages processi i .md
  REPORT-ANNCSU.md                    questo file
report/                               ← quello che finisce nel repo
  README.md                           indice, tabella regioni, legenda colonne
  novita.csv                          i comuni cambiati questo mese, i piu' attivi in cima
  riepilogo-regioni.csv               20 righe, una per regione
  sorgenti.json                       data del dataset di ogni regione
  regioni/
    abruzzo.csv                       stato corrente, una riga per comune
    basilicata.csv
    …                                 20 file, uno per regione
  storico/
    2026-07/
      README.md                       sintesi del mese
      novita.csv                      copia congelata delle novita' di quel mese
    2026-08/…
dist/                                 ← non versionato (.gitignore)
  comuni-2026-07.csv                  dump nazionale completo
  anncsu-2026-07.xlsx                 stesso contenuto in Excel, tre fogli
```

Nel repo restano circa 650 KB di schede regionali riscritte ogni mese, più una
cartella di storico che pesa quanto le sole variazioni. Il dump nazionale
completo non viene committato: va in allegato alla **release** `report-YYYY-MM`,
scaricabile da chiunque senza login e senza far crescere la history.

## Come si consultano le modifiche

Quattro strade, a seconda di cosa serve.

**Il sito.** `https://checcoconf.github.io/anncsu-report/` — casella di
ricerca, filtro per regione, "solo chi si è mosso", "solo sotto il 50%",
ordinamento per colonna. Legge `docs/dati/comuni.json` (630 KB, 139 KB
compressi in transito) e fa tutto nel browser: nessun backend, nessun download.
Da attivare una volta sola in Settings → Pages → Deploy from a branch → `main` /
`/docs`.

**Che cosa è cambiato questo mese.** `report/novita.csv` contiene solo i comuni che
si sono mossi, ordinati dal comune che ha aggiunto più coordinate in giù. È il file
da aprire per primo, e i primi dieci finiscono anche nel README.

**Un comune preciso.** Apri `report/regioni/lombardia.csv`: GitHub lo mostra come
tabella con la casella di ricerca, quindi cerchi il comune e leggi la sua riga.
Il file sta sotto i 200 KB, ben dentro il limite oltre il quale GitHub smette di
renderizzare i CSV.

**La storia di una regione.** Nella pagina del file, **History** → l'ultimo commit:
il diff mostra riga per riga quali comuni si sono mossi e di quanto. È la ragione
per cui le schede sono divise per regione e ordinate sempre allo stesso modo, così
il diff resta corto e leggibile.

Colonne delle schede regionali, nell'ordine in cui compaiono:

| Colonna | Significato |
|---|---|
| `REGIONE` `PROVINCIA` `COMUNE` | denominazioni |
| `NUMERI_CIVICI` | numeri civici presenti nell'archivio |
| `CON_COORDINATE` | quanti hanno latitudine e longitudine |
| `COMPLETAMENTO_%` | percentuale georeferenziata |
| `COORDINATE_AGGIUNTE` `CIVICI_AGGIUNTI` | differenze rispetto al report precedente |
| `COSA_E_CAMBIATO` | `Invariato`, `Aggiunte coordinate`, `Aggiunti civici`, `Aggiunti civici e coordinate`, `In calo`, `Nuovo comune` |
| `ULTIMO_AGGIORNAMENTO` | ultimo mese in cui quel comune ha mosso i suoi numeri |
| `CON_COORDINATE_MESE_SCORSO` `NUMERI_CIVICI_MESE_SCORSO` | valori del report precedente |
| `DATI_AGGIORNATI_AL` | data di generazione del file sorgente della regione |
| `CODICE_ISTAT` `CODICE_CATASTALE` `CODICE_IPA` | identificativi, in fondo perché servono solo a chi incrocia altre fonti |

`ULTIMO_AGGIORNAMENTO` è l'unica colonna che non si ricava dai dati del mese: viene
copiata dalla scheda precedente quando il comune non cambia, quindi la sua storia
parte dal `2026-03` della baseline e si allunga da lì.

## Quello che il report non sa

Sa **quanti** civici hanno ricevuto le coordinate in un comune, non **quali**.
Per sapere che è stato Via Roma 12 e non Via Verdi 8 servirebbe confrontare i
27,5 milioni di accessi uno per uno fra un mese e l'altro, e quindi conservare
l'identità di ciascuno (`PROGRESSIVO_ACCESSO`) da una corsa alla successiva.
Il runner non ha stato fra un'esecuzione e l'altra, quindi quello stato andrebbe
scritto da qualche parte: circa 40-70 MB al mese di identificativi compressi,
troppo per il repo ma gestibile come allegato di release da riscaricare all'avvio.
È un secondo progetto, non una variante di questo.

## Come gira

Parte **ogni giorno alle 04:20 UTC**, oppure a mano da **Actions → Report ANNCSU →
Run workflow**.

Ogni corsa comincia da un controllo che costa una manciata di byte. Il nome del
csv dentro ogni zip contiene la data di generazione, e in un file zip quel nome
sta nell'intestazione del primo membro, cioè nei primissimi byte: con una
richiesta `Range: bytes=0-511` si legge senza scaricare il resto. Venti regioni,
circa 10 KB in tutto. Se nessuna data è cambiata rispetto a `sorgenti.json` il
job finisce lì, in una decina di secondi.

Solo quando almeno una regione ha una data nuova parte il lavoro vero: scarica
una regione alla volta, la legge in streaming dentro lo zip senza mai estrarla e
cancella il file prima di passare alla successiva. Il picco di disco è quello del
file regionale più grande e la memoria resta di qualche decina di MB anche sui
27 milioni di civici nazionali. Il runner GitHub (4 CPU, 16 GB di RAM, 14 GB di
disco, 6 ore di limite) è abbondante, e su un repo pubblico i minuti sono gratis.

Se il controllo di una regione fallisce, quella regione viene considerata
cambiata. Una corsa completa di troppo costa mezz'ora di runner; una mancata
costa un mese di ritardo.

Un'altra protezione: **report incompleto = job fallito**. Se anche una sola
regione non si scarica non viene pubblicato nulla, perché un report a cui manca
una regione produrrebbe il mese dopo migliaia di variazioni inesistenti. Basta
rilanciare il workflow.

Il commit contiene `[skip ci]`: non serve in questo repo, ma evita che un
eventuale altro workflow reagisca ai commit del bot.

## Prove in locale

```bash
pip install -r tools/requirements.txt

# due regioni piccole, senza toccare le schede versionate
python tools/anncsu_report.py --out /tmp/prova --dist /tmp/prova-dist \
  --regioni VALL,MOLI --parziale

# solo il controllo: dice se l'ANNCSU ha ripubblicato qualcosa
python tools/anncsu_report.py --out report --controlla

# tutto, come in CI
python tools/anncsu_report.py --out report --dist dist --sito docs
```

## Manutenzione

Il tracciato dell'indirizzario ha già un campo in più rispetto ai 18 dichiarati
nei metadati ufficiali, quindi lo script non si fida delle posizioni: riconosce
da solo la colonna del codice catastale, quella del codice Istat e la coppia di
coordinate leggendo un campione di righe. Se l'Agenzia aggiunge o sposta colonne
continua a funzionare; se cambia il separatore o la codifica, l'errore compare nel
log del job come `colonne non individuate`.

Il campionamento delle colonne legge le prime 4.000 righe più, se serve, altre
fino a incontrare 200 righe con le coordinate valorizzate: in alcune regioni i
primi comuni in ordine di codice non ne hanno nessuna, e un campione corto
farebbe fallire il riconoscimento. Se anche così non ne trova (regione senza
nemmeno una coordinata), ricade sulla posizione strutturale: longitudine e
latitudine sono sempre quart'ultima e terz'ultima colonna, prima di `QUOTA` e
`METODO`.

I valori di partenza per il confronto sono quelli dell'allegato ANNCSU al
31 marzo 2026, già caricati in `report/regioni/*.csv` con
`COSA_E_CAMBIATO = Punto di partenza`: il primo report automatico misura quindi
le variazioni rispetto a quella data.
