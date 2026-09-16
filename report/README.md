# I dati

Fotografia di **2026-09**. La lettura guidata, con i totali e la spiegazione delle colonne, è nel [README del repository](../README.md).

| | Cosa contiene |
|---|---|
| **[`novita.csv`](novita.csv)** | solo i comuni cambiati rispetto al report precedente, dal più attivo in giù — sono 529 |
| **[`regioni/`](regioni/)** | una tabella per regione, tutti i comuni, lo stato di adesso |
| **[`riepilogo-regioni.csv`](riepilogo-regioni.csv)** | venti righe, i totali di ogni regione |
| **[`storico/`](storico/)** | le variazioni mese per mese, congelate |
| `sorgenti.json` | la data del file ANNCSU di ogni regione: serve al workflow per capire se c'è qualcosa di nuovo |

Per il file unico con tutti i comuni, in CSV e in Excel, vedi le [release](../../../releases).

## Come leggere le variazioni

Apri una tabella regionale, clicca **History** in alto a destra e poi un commit: il confronto mostra in verde e in rosso le righe dei comuni che si sono mossi, con i numeri vecchi accanto ai nuovi. Le tabelle sono divise per regione e ordinate sempre allo stesso modo proprio per questo: il diff resta corto e leggibile.

## Mesi disponibili

[2026-09](storico/2026-09/) · [2026-08](storico/2026-08/)

