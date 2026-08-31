#!/usr/bin/env python3
"""
Report mensile georeferenziazione numeri civici ANNCSU.

Scarica gli indirizzari regionali in open data dall'Agenzia delle Entrate,
conta per ogni comune quanti numeri civici ci sono e quanti hanno le coordinate,
e confronta il risultato con il report del mese precedente.

Nel repo restano solo i file leggibili: una tabella per regione con lo stato
corrente, l'elenco delle novita' del mese e lo storico. Il dump nazionale
completo finisce in dist/ e va allegato alla release, non committato.

Uso:
    python tools/anncsu_report.py --out report
    python tools/anncsu_report.py --out /tmp/prova --regioni VALL,MOLI --parziale
    python tools/anncsu_report.py --out report --solo-se-aggiornato
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import shutil
import sys
import time
import zipfile
from datetime import date, datetime
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

BASE_URL = "https://anncsu.open.agenziaentrate.gov.it/age-inspire/opendata/anncsu/getds.php"
UA = "anncsu-report (+https://github.com/checcoconf/anncsu-report)"

# codice ANNCSU, nome regione, nome del file nel repo
REGIONI = [
    ("ABRU", "ABRUZZO", "abruzzo"),
    ("BASI", "BASILICATA", "basilicata"),
    ("CALA", "CALABRIA", "calabria"),
    ("CAMP", "CAMPANIA", "campania"),
    ("EMIL", "EMILIA ROMAGNA", "emilia-romagna"),
    ("FRIU", "FRIULI VENEZIA GIULIA", "friuli-venezia-giulia"),
    ("LAZI", "LAZIO", "lazio"),
    ("LIGU", "LIGURIA", "liguria"),
    ("LOMB", "LOMBARDIA", "lombardia"),
    ("MARC", "MARCHE", "marche"),
    ("MOLI", "MOLISE", "molise"),
    ("PIEM", "PIEMONTE", "piemonte"),
    ("PUGL", "PUGLIA", "puglia"),
    ("SARD", "SARDEGNA", "sardegna"),
    ("SICI", "SICILIA", "sicilia"),
    ("TOSC", "TOSCANA", "toscana"),
    ("TREN", "TRENTINO ALTO ADIGE", "trentino-alto-adige"),
    ("UMBR", "UMBRIA", "umbria"),
    ("VALL", "VALLE D'AOSTA", "valle-d-aosta"),
    ("VENE", "VENETO", "veneto"),
]
SLUG = {nome: slug for _, nome, slug in REGIONI}

ISTAT_URLS = [
    "https://www.istat.it/storage/codici-unita-amministrative/Elenco-comuni-italiani.xlsx",
    "https://www.istat.it/storage/codici-unita-amministrative/Elenco-comuni-italiani.csv",
]

# bounding box generosa dell'Italia, serve solo a riconoscere le colonne
LON_MIN, LON_MAX = 5.0, 20.0
LAT_MIN, LAT_MAX = 34.0, 48.5

# Le colonne piu' utili stanno davanti; i codici, che servono solo a chi
# incrocia i dati con altre fonti, stanno in fondo.
COLONNE_COMUNI = [
    "REGIONE", "PROVINCIA", "COMUNE",
    "NUMERI_CIVICI", "CON_COORDINATE", "COMPLETAMENTO_%",
    "COORDINATE_AGGIUNTE", "CIVICI_AGGIUNTI",
    "COSA_E_CAMBIATO", "ULTIMO_AGGIORNAMENTO",
    "CON_COORDINATE_MESE_SCORSO", "NUMERI_CIVICI_MESE_SCORSO",
    "DATI_AGGIORNATI_AL",
    "CODICE_ISTAT", "CODICE_CATASTALE", "CODICE_IPA",
]

COLONNE_REGIONI = [
    "REGIONE", "COMUNI", "NUMERI_CIVICI", "CON_COORDINATE", "COMPLETAMENTO_%",
    "COORDINATE_AGGIUNTE", "CIVICI_AGGIUNTI",
    "COMUNI_AGGIORNATI", "COMUNI_SENZA_NESSUNA_COORDINATA",
    "DATI_AGGIORNATI_AL", "SCHEDA",
]

INVARIATO = "Invariato"
NUOVO = "Nuovo comune"
PARTENZA = "Punto di partenza"


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


# --------------------------------------------------------------------------
# download
# --------------------------------------------------------------------------

def scarica(url: str, dest: Path, tentativi: int = 4, timeout: int = 900) -> Path:
    """Scarica un file in streaming, con qualche tentativo in caso di errore."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    ultimo = None
    for n in range(1, tentativi + 1):
        try:
            req = Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
            with urlopen(req, timeout=timeout) as r, open(dest, "wb") as f:
                shutil.copyfileobj(r, f, length=1024 * 1024)
            if dest.stat().st_size == 0:
                raise IOError("file vuoto")
            return dest
        except (URLError, HTTPError, IOError, TimeoutError) as e:  # noqa: PERF203
            ultimo = e
            attesa = 10 * n
            log(f"  tentativo {n}/{tentativi} fallito ({e}); riprovo tra {attesa}s")
            time.sleep(attesa)
    raise RuntimeError(f"download fallito: {url} ({ultimo})")


def data_dal_nome(nome: str) -> str:
    """INDIR_ABRU_20260731.csv -> 2026-07-31"""
    m = re.search(r"(\d{8})", Path(nome).stem)
    if not m:
        return ""
    try:
        return datetime.strptime(m.group(1), "%Y%m%d").date().isoformat()
    except ValueError:
        return m.group(1)


def sbircia_data(url: str, timeout: int = 60) -> str:
    """
    Legge la data del dataset senza scaricare lo zip.

    Il nome del csv contiene la data di generazione, e in uno zip quel nome sta
    nell'intestazione del primo membro, nei primi byte del file. Con una richiesta
    Range ne bastano 512 invece di centinaia di megabyte. Se il server ignora
    l'header Range chiudiamo comunque la connessione dopo il primo blocco.
    """
    req = Request(url, headers={"User-Agent": UA, "Accept": "*/*", "Range": "bytes=0-511"})
    with urlopen(req, timeout=timeout) as r:
        testa = r.read(512)
    if testa[:4] != b"PK\x03\x04" or len(testa) < 30:
        raise ValueError("non sembra uno zip")
    lung_nome = int.from_bytes(testa[26:28], "little")
    nome = testa[30:30 + lung_nome].decode("utf-8", "replace")
    data = data_dal_nome(nome)
    if not data:
        raise ValueError(f"nessuna data nel nome {nome!r}")
    return data


def controlla_aggiornamenti(sorgenti_prec: dict, pausa: float) -> tuple[dict, list]:
    """Confronta la data di ogni dataset regionale con quella dell'ultimo report."""
    date_ora, cambiate = {}, []
    for cod, nome, _ in REGIONI:
        try:
            data = sbircia_data(f"{BASE_URL}?INDIR_{cod}")
        except Exception as e:  # noqa: BLE001
            # nel dubbio la consideriamo cambiata: meglio una corsa in piu'
            log(f"  {nome}: controllo fallito ({e}), la tratto come cambiata")
            date_ora[cod], _ = "", cambiate.append(nome)
            continue
        date_ora[cod] = data
        prec = sorgenti_prec.get(cod)
        segno_ = "=" if prec == data else "→"
        log(f"  {nome}: {prec or 'mai visto'} {segno_} {data}")
        if prec != data:
            cambiate.append(nome)
        time.sleep(pausa)
    return date_ora, cambiate


# --------------------------------------------------------------------------
# lettura dell'indirizzario
# --------------------------------------------------------------------------

def apri_testo(zf: zipfile.ZipFile, nome: str):
    return io.TextIOWrapper(zf.open(nome, "r"), encoding="utf-8", errors="replace", newline="")


def _num(v: str):
    """Converte un numero ANNCSU (decimale con virgola) in float, o None."""
    if not v:
        return None
    v = v.strip().replace(",", ".")
    if not v:
        return None
    try:
        return float(v)
    except ValueError:
        return None


def individua_colonne(righe: list[list[str]]) -> dict:
    """
    Individua gli indici delle colonne senza dipendere dall'ordine dichiarato
    nei metadati: il tracciato ha gia' un campo in piu' rispetto ai 18 pubblicati.
    """
    if not righe:
        raise ValueError("nessuna riga di campione")

    ncol = max(len(r) for r in righe)
    idx = {"comune": None, "istat": None, "lon": None, "lat": None}
    re_cat = re.compile(r"^[A-Za-z]\d{3}$")
    re_istat = re.compile(r"^\d{6}$")
    punteggi = {k: [0] * ncol for k in idx}

    for r in righe:
        for i, v in enumerate(r):
            v = (v or "").strip()
            if not v:
                continue
            if re_cat.match(v):
                punteggi["comune"][i] += 1
            if re_istat.match(v):
                punteggi["istat"][i] += 1
            f = _num(v)
            if f is None or "." not in v.replace(",", "."):
                continue
            if LON_MIN <= f <= LON_MAX:
                punteggi["lon"][i] += 1
            if LAT_MIN <= f <= LAT_MAX:
                punteggi["lat"][i] += 1

    for k in ("comune", "istat"):
        best = max(range(ncol), key=lambda i: punteggi[k][i])
        if punteggi[k][best] > 0:
            idx[k] = best
    if idx["comune"] is None or idx["istat"] is None:
        raise ValueError("colonne dei codici non individuate")

    # lon e lat cadono nello stesso intervallo: la longitudine e' la prima delle due
    candidati = sorted(
        (i for i in range(ncol) if punteggi["lon"][i] or punteggi["lat"][i]),
        key=lambda i: -(punteggi["lon"][i] + punteggi["lat"][i]),
    )[:2]
    if len(candidati) == 2:
        idx["lon"], idx["lat"] = sorted(candidati)
    elif ncol >= 4:
        # Nessuna coordinata nel campione: capita se i primi comuni della regione
        # non ne hanno nessuna. Le due colonne restano comunque quart'ultima e
        # terz'ultima, prima di QUOTA e METODO, in tutte le versioni del tracciato.
        idx["lon"], idx["lat"] = ncol - 4, ncol - 3
    else:
        raise ValueError("colonne delle coordinate non individuate")
    return idx


def conta_regione(zip_path: Path) -> tuple[dict, str]:
    """
    Legge lo zip in streaming e restituisce
    {codice_catastale: [civici, con_coordinate, codice_istat]} e la data del dataset.
    """
    conteggi: dict[str, list] = {}
    with zipfile.ZipFile(zip_path) as zf:
        membri = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        if not membri:
            raise ValueError(f"nessun csv dentro {zip_path.name}: {zf.namelist()}")
        membro = membri[0]

        data_dataset = data_dal_nome(membro)

        # Campiono finche' non incontro un po' di coordinate: se i primi comuni
        # della regione non ne hanno, le prime mille righe sono tutte vuote.
        with apri_testo(zf, membro) as fh:
            campione, con_coord, letti = [], 0, 0
            for r in csv.reader(fh, delimiter=";"):
                letti += 1
                if len(campione) < 4000:
                    campione.append(r)
                elif con_coord >= 200 or letti > 400_000:
                    break
                if any("," in (v or "") and _num(v) is not None for v in r[-6:]):
                    con_coord += 1
                    if con_coord < 200 and len(campione) >= 4000:
                        campione.append(r)
        intestazione = bool(campione) and any(
            "CODICE_COMUNE" in (c or "").upper() for c in campione[0])
        idx = individua_colonne(campione[1:] if intestazione else campione)

        i_com, i_ist, i_lon, i_lat = idx["comune"], idx["istat"], idx["lon"], idx["lat"]
        larghezza = max(i_com, i_ist, i_lon, i_lat) + 1

        with apri_testo(zf, membro) as fh:
            lettore = csv.reader(fh, delimiter=";")
            if intestazione:
                next(lettore, None)
            for riga in lettore:
                if len(riga) < larghezza:
                    continue
                cod = riga[i_com].strip().upper()
                if not cod:
                    continue
                voce = conteggi.get(cod)
                if voce is None:
                    voce = [0, 0, riga[i_ist].strip()]
                    conteggi[cod] = voce
                voce[0] += 1
                x, y = _num(riga[i_lon]), _num(riga[i_lat])
                if x is not None and y is not None and (x or y):
                    voce[1] += 1

    return conteggi, data_dataset


# --------------------------------------------------------------------------
# anagrafica Istat
# --------------------------------------------------------------------------

def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def carica_anagrafica(cache: Path) -> dict:
    """
    {codice_catastale: {"comune":…, "provincia":…, "regione":…, "istat":…, "ipa":…}}

    Parte dalla copia versionata nel repo (che porta il codice IPA, non ricavabile
    dai dati Istat) e ci sovrappone l'elenco Istat aggiornato, cosi' fusioni e
    cambi di denominazione entrano da soli. Se Istat non risponde si va avanti
    con la sola copia locale.
    """
    base: dict[str, dict] = {}
    if cache.exists():
        with open(cache, newline="", encoding="utf-8") as f:
            base = {r["codice_catastale"]: dict(r) for r in csv.DictReader(f)}
        log(f"anagrafica locale: {len(base)} comuni")

    tmp = cache.parent / "_istat_tmp"
    for url in ISTAT_URLS:
        try:
            log(f"aggiorno l'anagrafica da Istat: {url}")
            dest = scarica(url, tmp / Path(url).name, tentativi=2, timeout=180)
            fresche = _leggi_istat(dest)
            if not fresche:
                continue
            for cat, r in fresche.items():
                voce = base.setdefault(cat, {"codice_catastale": cat, "ipa": ""})
                voce.update({k: v for k, v in r.items() if v})
            log(f"  {len(fresche)} comuni da Istat, anagrafica a {len(base)} voci")
            _scrivi_cache(cache, base)
            break
        except Exception as e:  # noqa: BLE001
            log(f"  non disponibile ({e})")
    else:
        log("Istat non raggiungibile: uso solo la copia nel repo")

    shutil.rmtree(tmp, ignore_errors=True)
    if not base:
        raise RuntimeError("anagrafica non disponibile: ne' Istat ne' cache nel repo")
    return base


def _leggi_istat(path: Path) -> dict:
    if path.suffix.lower() in (".xlsx", ".xls"):
        from openpyxl import load_workbook

        wb = load_workbook(path, read_only=True, data_only=True)
        it = wb[wb.sheetnames[0]].iter_rows(values_only=True)
        header = [str(c or "") for c in next(it)]
        righe = ([("" if c is None else str(c)) for c in r] for r in it)
        return _mappa_istat(header, righe)

    testo = ""
    for enc in ("utf-8-sig", "latin-1"):
        try:
            testo = path.read_text(encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    delim = ";" if testo.count(";") > testo.count(",") else ","
    lettore = csv.reader(io.StringIO(testo), delimiter=delim)
    return _mappa_istat(next(lettore), lettore)


def _mappa_istat(header: list[str], righe) -> dict:
    h = [_norm(c) for c in header]

    def trova(*chiavi):
        for k in chiavi:
            for i, c in enumerate(h):
                if k in c:
                    return i
        return None

    i_cat = trova("codice catastale")
    i_com = trova("denominazione in italiano", "denominazione italiana e straniera", "denominazione")
    i_pro = trova("denominazione dell unita territoriale sovracomunale", "denominazione provincia")
    i_reg = trova("denominazione regione")
    i_ist = trova("codice comune formato alfanumerico", "codice istat")
    if i_cat is None or i_com is None:
        raise ValueError(f"intestazione Istat non riconosciuta: {header[:8]}")

    def campo(r, i):
        return (r[i] or "").strip() if i is not None and len(r) > i else ""

    out = {}
    for r in righe:
        cat = campo(r, i_cat).upper()
        if not re.match(r"^[A-Z]\d{3}$", cat):
            continue
        out[cat] = {
            "codice_catastale": cat,
            "comune": campo(r, i_com),
            "provincia": campo(r, i_pro),
            "regione": campo(r, i_reg),
            "istat": campo(r, i_ist),
        }
    return out


def _scrivi_cache(cache: Path, righe: dict) -> None:
    cache.parent.mkdir(parents=True, exist_ok=True)
    campi = ["codice_catastale", "comune", "provincia", "regione", "istat", "ipa"]
    with open(cache, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=campi, lineterminator="\n")
        w.writeheader()
        for k in sorted(righe):
            w.writerow({c: righe[k].get(c, "") for c in campi})


# --------------------------------------------------------------------------
# confronto
# --------------------------------------------------------------------------

def carica_precedente(cartella: Path) -> dict:
    """Rilegge le schede regionali: sono loro lo stato salvato del mese scorso."""
    prec = {}
    for f in sorted(cartella.glob("*.csv")):
        with open(f, newline="", encoding="utf-8") as fh:
            for r in csv.DictReader(fh):
                cod = r.get("CODICE_CATASTALE")
                if cod:
                    prec[cod] = r
    return prec


def _int(v) -> int:
    try:
        return int(float(str(v).replace(",", ".")))
    except (TypeError, ValueError):
        return 0


def perc(num: int, den: int) -> float:
    return round(100.0 * num / den, 2) if den else 0.0


def cosa_e_cambiato(nuovo: bool, d_civici: int, d_coord: int) -> str:
    if nuovo:
        return NUOVO
    if d_civici == 0 and d_coord == 0:
        return INVARIATO
    if d_civici < 0 or d_coord < 0:
        return "In calo"
    if d_coord > 0 and d_civici == 0:
        return "Aggiunte coordinate"
    if d_coord == 0 and d_civici > 0:
        return "Aggiunti civici"
    return "Aggiunti civici e coordinate"


def costruisci_righe(dati: dict, anagrafica: dict, prec: dict,
                     date_dataset: dict, mese: str) -> list[dict]:
    righe = []
    for cod_reg, nome_reg, _ in REGIONI:
        conteggi = dati.get(cod_reg)
        if conteggi is None:
            continue
        for cat, (civici, coord, istat) in conteggi.items():
            info = anagrafica.get(cat, {})
            p = prec.get(cat)
            civici_p = _int(p["NUMERI_CIVICI"]) if p else 0
            coord_p = _int(p["CON_COORDINATE"]) if p else 0
            d_civici, d_coord = civici - civici_p, coord - coord_p
            # da quando i numeri di questo comune non si muovono
            if p is None or d_civici or d_coord:
                ultimo = mese
            else:
                ultimo = p.get("ULTIMO_AGGIORNAMENTO", "")
            righe.append({
                "REGIONE": nome_reg,
                "PROVINCIA": info.get("provincia", "").upper(),
                "COMUNE": info.get("comune", "").upper(),
                "NUMERI_CIVICI": civici,
                "CON_COORDINATE": coord,
                "COMPLETAMENTO_%": perc(coord, civici),
                "COORDINATE_AGGIUNTE": d_coord if p else "",
                "CIVICI_AGGIUNTI": d_civici if p else "",
                "COSA_E_CAMBIATO": cosa_e_cambiato(p is None, d_civici, d_coord),
                "ULTIMO_AGGIORNAMENTO": ultimo,
                "CON_COORDINATE_MESE_SCORSO": coord_p if p else "",
                "NUMERI_CIVICI_MESE_SCORSO": civici_p if p else "",
                "DATI_AGGIORNATI_AL": date_dataset.get(cod_reg, ""),
                "CODICE_ISTAT": info.get("istat") or istat,
                "CODICE_CATASTALE": cat,
                "CODICE_IPA": info.get("ipa") or ("c_" + cat.lower()),
            })
    righe.sort(key=lambda r: (r["REGIONE"], r["PROVINCIA"], r["COMUNE"], r["CODICE_CATASTALE"]))
    return righe


def riepiloga_regioni(righe: list[dict], date_dataset: dict) -> list[dict]:
    per_data = {nome: date_dataset.get(cod, "") for cod, nome, _ in REGIONI}
    agg: dict[str, dict] = {}
    for r in righe:
        a = agg.setdefault(r["REGIONE"], dict.fromkeys(
            ["COMUNI", "NUMERI_CIVICI", "CON_COORDINATE", "_CIV_P", "_COO_P",
             "COMUNI_AGGIORNATI", "COMUNI_SENZA_NESSUNA_COORDINATA"], 0))
        a["COMUNI"] += 1
        a["NUMERI_CIVICI"] += r["NUMERI_CIVICI"]
        a["CON_COORDINATE"] += r["CON_COORDINATE"]
        a["_CIV_P"] += _int(r["NUMERI_CIVICI_MESE_SCORSO"])
        a["_COO_P"] += _int(r["CON_COORDINATE_MESE_SCORSO"])
        if r["COSA_E_CAMBIATO"] not in (INVARIATO, NUOVO, PARTENZA):
            a["COMUNI_AGGIORNATI"] += 1
        if r["CON_COORDINATE"] == 0:
            a["COMUNI_SENZA_NESSUNA_COORDINATA"] += 1

    out = []
    for nome, a in agg.items():
        a.update({
            "REGIONE": nome,
            "COMPLETAMENTO_%": perc(a["CON_COORDINATE"], a["NUMERI_CIVICI"]),
            "COORDINATE_AGGIUNTE": a["CON_COORDINATE"] - a["_COO_P"],
            "CIVICI_AGGIUNTI": a["NUMERI_CIVICI"] - a["_CIV_P"],
            "DATI_AGGIORNATI_AL": per_data.get(nome, ""),
            "SCHEDA": f"regioni/{SLUG[nome]}.csv",
        })
        out.append({c: a.get(c, "") for c in COLONNE_REGIONI})
    out.sort(key=lambda r: r["REGIONE"])
    return out


# --------------------------------------------------------------------------
# scrittura
# --------------------------------------------------------------------------

def scrivi_csv(path: Path, colonne: list[str], righe: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=colonne, extrasaction="ignore", lineterminator="\n")
        w.writeheader()
        w.writerows(righe)


def scrivi_xlsx(path: Path, comuni: list[dict], regioni: list[dict], novita: list[dict]) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    testata = Font(name="Arial", bold=True, color="FFFFFF")
    sfondo = PatternFill("solid", fgColor="1F4E79")
    corpo = Font(name="Arial")

    fogli = [("Regioni", COLONNE_REGIONI, regioni),
             ("Novita del mese", COLONNE_COMUNI, novita),
             ("Tutti i comuni", COLONNE_COMUNI, comuni)]
    for nome, colonne, righe in fogli:
        ws = wb.create_sheet(nome)
        ws.append(colonne)
        for c in ws[1]:
            c.font, c.fill = testata, sfondo
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in righe:
            ws.append([r.get(c, "") for c in colonne])
        for col in range(1, len(colonne) + 1):
            lettera = get_column_letter(col)
            ws.column_dimensions[lettera].width = max(11, min(28, len(colonne[col - 1]) + 2))
            for cell in ws[lettera][1:]:
                cell.font = corpo
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(colonne))}{ws.max_row}"

    del wb[wb.sheetnames[0]]
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def mig(n) -> str:
    return f"{_int(n):,}".replace(",", ".")


def segno(n) -> str:
    v = _int(n)
    return f"+{mig(v)}" if v > 0 else ("0" if v == 0 else f"-{mig(-v)}")


def barra(percentuale: float, larghezza: int = 10) -> str:
    pieni = int(round(percentuale / 100 * larghezza))
    return "█" * pieni + "░" * (larghezza - pieni)


def scrivi_readme(path: Path, mese: str, comuni: list[dict], regioni: list[dict],
                  novita: list[dict], aveva_precedente: bool) -> str:
    """
    Il README del repo e' il report: cruscotto in cima, guida sotto.
    Viene riscritto per intero a ogni corsa completa.
    """
    tot_civ = sum(r["NUMERI_CIVICI"] for r in comuni)
    tot_coo = sum(r["CON_COORDINATE"] for r in comuni)
    tot_nuove = sum(_int(r["COORDINATE_AGGIUNTE"]) for r in comuni)
    pc = perc(tot_coo, tot_civ)
    completi = sum(1 for r in comuni if r["COMPLETAMENTO_%"] >= 99.5)
    a_zero = sum(1 for r in comuni if r["CON_COORDINATE"] == 0)
    date_valide = sorted({r["DATI_AGGIORNATI_AL"] for r in regioni if r["DATI_AGGIORNATI_AL"]})
    finestra = date_valide[-1] if len(date_valide) <= 1 else f"{date_valide[0]} → {date_valide[-1]}"
    reg_mosse = [r["REGIONE"] for r in regioni if _int(r["COMUNI_AGGIORNATI"]) > 0]

    t = ["# Numeri civici georeferenziati in Italia", "",
         "Ogni numero civico registrato nell'Anagrafe Nazionale dei Numeri Civici e delle "
         "Strade Urbane può avere latitudine e longitudine, oppure no: dipende da quanto lavoro "
         "ci ha messo il comune. Questo repository ricostruisce la situazione **comune per "
         "comune**, ogni volta che l'ANNCSU ripubblica i suoi dati aperti, e tiene traccia di "
         "cosa si muove.", "",
         f"### {barra(pc, 20)}  **{pc:.1f}%**", "",
         f"**{mig(tot_coo)}** civici con coordinate su **{mig(tot_civ)}** in archivio · "
         f"{mig(len(comuni))} comuni · dati ANNCSU al **{finestra}**", ""]

    # ---- cosa è cambiato ----
    t += ["## Cosa è cambiato", ""]
    if not aveva_precedente:
        t += ["Primo report: non c'è ancora un confronto.", ""]
    elif not novita:
        t += ["Nessun comune ha cambiato numeri rispetto al report precedente.", ""]
    else:
        t += [f"**{segno(tot_nuove)} coordinate** in **{mig(len(novita))} comuni**"
              + (f", in {len(reg_mosse)} regioni su {len(regioni)}." if reg_mosse else "."), "",
              "I venticinque comuni più attivi:", "",
              "| Comune | Provincia | Coordinate aggiunte | Ora è al |", "|---|---|--:|--:|"]
        for r in novita[:25]:
            t.append(f"| **{r['COMUNE']}** | {r['PROVINCIA']} | "
                     f"{segno(r['COORDINATE_AGGIUNTE'])} | {r['COMPLETAMENTO_%']:.0f}% |")
        t += ["", f"Tutti e {mig(len(novita))} in **[`report/novita.csv`](report/novita.csv)**, "
              "già ordinati dal più attivo. "
              f"La copia congelata di questo mese è in "
              f"[`report/storico/{mese}/`](report/storico/{mese}/).", ""]

    # ---- regioni ----
    t += ["## Le regioni", "",
          "Clicca una regione per aprire la sua tabella: GitHub la mostra con la casella di "
          "ricerca in alto, così cerchi il comune e leggi la sua riga.", "",
          "| Regione | Completamento | Civici | Con coordinate | Aggiunte | Comuni mossi | Dati al |",
          "|---|---|--:|--:|--:|--:|:--:|"]
    for r in regioni:
        p_ = r["COMPLETAMENTO_%"]
        t.append(f"| **[{r['REGIONE']}]({r['SCHEDA'].replace('regioni/', 'report/regioni/')})** | "
                 f"{barra(p_)} {p_:.1f}% | {mig(r['NUMERI_CIVICI'])} | {mig(r['CON_COORDINATE'])} | "
                 f"{segno(r['COORDINATE_AGGIUNTE']) if aveva_precedente else '—'} | "
                 f"{mig(r['COMUNI_AGGIORNATI']) if aveva_precedente else '—'} | "
                 f"{r['DATI_AGGIORNATI_AL'] or 'n/d'} |")
    t.append(f"| **ITALIA** | {barra(pc)} {pc:.1f}% | {mig(tot_civ)} | {mig(tot_coo)} | "
             f"{segno(tot_nuove) if aveva_precedente else '—'} | "
             f"{mig(len(novita)) if aveva_precedente else '—'} | {finestra} |")
    t += ["", f"Agli estremi: **{mig(completi)} comuni** hanno georeferenziato tutto, "
          f"**{mig(a_zero)}** non hanno nemmeno una coordinata.", ""]

    # ---- guida ----
    t += ["## Come trovare quello che ti serve", "",
          "**Voglio sapere come sta il mio comune.** Apri la sua regione nella tabella qui "
          "sopra, usa la casella di ricerca e leggi `COMPLETAMENTO_%`. Se hai fretta e sai già "
          "la regione, i file stanno in [`report/regioni/`](report/regioni/).", "",
          "**Voglio sapere cosa si è mosso.** "
          "[`report/novita.csv`](report/novita.csv) contiene solo i comuni cambiati, dal più "
          "attivo in giù. È il file da aprire per primo ogni mese.", "",
          "**Voglio vedere il prima e il dopo.** Apri il file di una regione, clicca **History** "
          "in alto a destra, poi un commit: il confronto evidenzia in verde e in rosso le righe "
          "dei comuni che si sono mossi, con i numeri vecchi accanto ai nuovi. È la ragione per "
          "cui le tabelle sono divise per regione e ordinate sempre allo stesso modo.", "",
          "**Voglio l'andamento nel tempo.** Ogni mese lascia la sua cartella in "
          "[`report/storico/`](report/storico/) con la sintesi e le variazioni congelate. "
          "La colonna `ULTIMO_AGGIORNAMENTO` di ogni comune dice l'ultima volta che ha mosso "
          "qualcosa: se è ferma da mesi, quel comune non sta lavorando all'archivio.", "",
          "**Voglio tutto in un file solo.** Il CSV nazionale e l'Excel con tutti i comuni sono "
          "allegati alla [release del mese](../../releases), scaricabili senza account.", "",
          "### Le colonne", "",
          "| Colonna | Cosa dice |", "|---|---|",
          "| `REGIONE` `PROVINCIA` `COMUNE` | dove siamo |",
          "| `NUMERI_CIVICI` | quanti numeri civici ha il comune in archivio |",
          "| `CON_COORDINATE` | quanti di questi hanno latitudine e longitudine |",
          "| `COMPLETAMENTO_%` | la percentuale: 100 = comune finito, 0 = nessuna coordinata |",
          "| `COORDINATE_AGGIUNTE` | quante coordinate sono comparse dal report precedente |",
          "| `CIVICI_AGGIUNTI` | quanti civici nuovi sono comparsi dal report precedente |",
          "| `COSA_E_CAMBIATO` | `Invariato`, `Aggiunte coordinate`, `Aggiunti civici`, "
          "`Aggiunti civici e coordinate`, `In calo`, `Nuovo comune` |",
          "| `ULTIMO_AGGIORNAMENTO` | l'ultimo mese in cui quel comune ha mosso qualcosa |",
          "| `CON_COORDINATE_MESE_SCORSO` `NUMERI_CIVICI_MESE_SCORSO` | i valori di prima, per "
          "controllare il conto |",
          "| `DATI_AGGIORNATI_AL` | data del file ANNCSU da cui arriva la riga |",
          "| `CODICE_ISTAT` `CODICE_CATASTALE` `CODICE_IPA` | i codici, in fondo perché servono "
          "solo a chi incrocia altre fonti |", "",
          "Un comune con percentuale bassa e `ULTIMO_AGGIORNAMENTO` vecchio non è un errore del "
          "report: è un comune che davvero non sta caricando le coordinate.", ""]

    # ---- fonti ----
    t += ["## Da dove arrivano i dati", "",
          "**Civici e coordinate**: i venti "
          "[indirizzari regionali in open data](https://www.anncsu.gov.it/it/consultazione-dellarchivio/open-data/Accedi-ai-servizi-di-dowload-massivo-in-Open-data/) "
          "dell'ANNCSU (Agenzia delle Entrate e Istat), file ZIP con dentro un CSV rigenerato "
          "periodicamente. Ogni riga è un accesso e conta come georeferenziato se ha "
          "longitudine e latitudine valorizzate e non entrambe a zero.", "",
          "**Nomi di comune, provincia e regione**: l'"
          "[elenco Istat dei comuni](https://www.istat.it/storage/codici-unita-amministrative/Elenco-comuni-italiani.xlsx), "
          "perché l'indirizzario contiene solo codici. La copia in "
          "[`data/lookup/comuni-istat.csv`](data/lookup/comuni-istat.csv) fa da riserva se Istat "
          "non risponde, e conserva il codice IPA che Istat non pubblica.", "",
          "L'aggancio fra le due fonti usa il **codice catastale** (`A018`) e non il codice "
          "Istat: dal 1° gennaio 2026 quest'ultimo è cambiato per tutti i comuni sardi, e le due "
          "fonti possono allinearsi in momenti diversi.", "",
          "I valori di partenza sono quelli dell'allegato ufficiale ANNCSU al **31 marzo 2026**. "
          "Quel primo confronto mette insieme due conteggi fatti con criteri diversi — il loro e "
          "il mio — quindi qualche scarto è atteso. Dal secondo report in poi il confronto è "
          "sempre fra numeri calcolati allo stesso modo.", ""]

    # ---- automazione ----
    t += ["## Come funziona", "",
          "Un workflow GitHub Actions gira **ogni notte** e comincia da un controllo che costa "
          "pochi kilobyte: il nome del CSV dentro ogni ZIP contiene la data di generazione, e in "
          "un file ZIP quel nome sta nei primissimi byte, quindi basta chiederne 512 al server "
          "per leggerlo. Se nessuna data è cambiata il job finisce lì in una decina di secondi.", "",
          "Quando invece almeno una regione ha una data nuova parte il lavoro vero: scarica una "
          "regione alla volta, la legge in streaming dentro lo ZIP senza mai estrarla e cancella "
          "il file prima di passare alla successiva. La memoria resta di qualche decina di MB "
          "anche sui ventisette milioni di civici nazionali.", "",
          "Due scelte pensate per non sporcare lo storico. Se il controllo di una regione "
          "fallisce, quella regione viene considerata cambiata: una corsa di troppo costa mezz'ora "
          "di runner, una mancata costa un mese di ritardo. E se anche una sola regione non si "
          "scarica, non viene pubblicato niente, perché un report incompleto genererebbe il mese "
          "dopo migliaia di variazioni inesistenti.", "",
          "Per lanciarlo a mano: **Actions → Report ANNCSU → Run workflow**. Il campo *regioni* "
          "accetta codici tipo `VALL,MOLI` per una prova veloce che non tocca i file versionati.", "",
          "### Quello che il report non sa", "",
          "Sa **quanti** civici hanno ricevuto le coordinate in un comune, non **quali**. Per "
          "dire che è stato Via Roma 12 e non Via Verdi 8 servirebbe confrontare i ventisette "
          "milioni di accessi uno per uno fra un mese e l'altro, conservando l'identità di "
          "ciascuno da una corsa alla successiva: sono decine di megabyte di stato al mese, "
          "fattibili ma è un progetto a parte.", "",
          "### In locale", "",
          "```bash", "pip install -r tools/requirements.txt", "",
          "# solo il controllo: dice se l'ANNCSU ha ripubblicato qualcosa",
          "python tools/anncsu_report.py --controlla", "",
          "# due regioni piccole, senza toccare i file versionati",
          "python tools/anncsu_report.py --out /tmp/prova --dist /tmp/prova-dist \\",
          "  --readme /tmp/prova/README.md --regioni VALL,MOLI --parziale", "```", "",
          "Il tracciato dell'indirizzario ha già un campo in più rispetto ai diciotto dichiarati "
          "nei metadati ufficiali, quindi lo script non si fida delle posizioni: riconosce da "
          "solo la colonna del codice catastale, quella del codice Istat e la coppia di "
          "coordinate leggendo un campione di righe. Se cambiasse il separatore o la codifica, "
          "l'errore comparirebbe nel log come `colonne non individuate`.", "",
          "---", "",
          f"*Report generato in automatico il {date.today().isoformat()}. "
          "Dati: ANNCSU, Agenzia delle Entrate e Istat, licenza aperta.*"]

    testo = "\n".join(t) + "\n"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(testo, encoding="utf-8")
    return testo


def readme_cartella_report(path: Path, mese: str, regioni: list[dict],
                           novita: list[dict], mesi: list[str]) -> None:
    """README di report/: si apre entrando nella cartella su GitHub."""
    t = ["# I dati", "",
         f"Fotografia di **{mese}**. La lettura guidata, con i totali e la spiegazione delle "
         "colonne, è nel [README del repository](../README.md).", "",
         "| | Cosa contiene |", "|---|---|",
         "| **[`novita.csv`](novita.csv)** | solo i comuni cambiati rispetto al report "
         f"precedente, dal più attivo in giù — sono {mig(len(novita))} |",
         "| **[`regioni/`](regioni/)** | una tabella per regione, tutti i comuni, lo stato di "
         "adesso |",
         "| **[`riepilogo-regioni.csv`](riepilogo-regioni.csv)** | venti righe, i totali di "
         "ogni regione |",
         "| **[`storico/`](storico/)** | le variazioni mese per mese, congelate |",
         "| `sorgenti.json` | la data del file ANNCSU di ogni regione: serve al workflow per "
         "capire se c'è qualcosa di nuovo |", "",
         "Per il file unico con tutti i comuni, in CSV e in Excel, vedi le "
         "[release](../../../releases).", "",
         "## Come leggere le variazioni", "",
         "Apri una tabella regionale, clicca **History** in alto a destra e poi un commit: il "
         "confronto mostra in verde e in rosso le righe dei comuni che si sono mossi, con i "
         "numeri vecchi accanto ai nuovi. Le tabelle sono divise per regione e ordinate sempre "
         "allo stesso modo proprio per questo: il diff resta corto e leggibile.", ""]
    if mesi:
        t += ["## Mesi disponibili", "", " · ".join(f"[{m}](storico/{m}/)" for m in mesi), ""]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(t) + "\n", encoding="utf-8")


def readme_cartella_regioni(path: Path, mese: str, regioni: list[dict],
                            aveva_precedente: bool) -> None:
    """README di report/regioni/: l'elenco delle venti tabelle."""
    t = ["# Le tabelle regionali", "",
         "Una riga per comune. GitHub apre ogni file come tabella con la casella di ricerca in "
         "alto: scrivi il nome del comune e leggi la sua riga. La spiegazione delle colonne è "
         "nel [README del repository](../../README.md).", "",
         f"Situazione a **{mese}**.", "",
         "| Regione | Comuni | Con coordinate | Completamento |" +
         (" Aggiunte |" if aveva_precedente else ""),
         "|---|--:|--:|---|" + ("--:|" if aveva_precedente else "")]
    for r in regioni:
        p_ = r["COMPLETAMENTO_%"]
        riga = (f"| **[{r['REGIONE']}]({Path(r['SCHEDA']).name})** | {mig(r['COMUNI'])} | "
                f"{mig(r['CON_COORDINATE'])} | {barra(p_)} {p_:.1f}% |")
        if aveva_precedente:
            riga += f" {segno(r['COORDINATE_AGGIUNTE'])} |"
        t.append(riga)
    t += ["", "Il file è ordinato per provincia e poi per comune, sempre allo stesso modo, così "
          "il diff fra un mese e l'altro resta leggibile.", ""]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(t) + "\n", encoding="utf-8")


def readme_cartella_storico(path: Path, mesi: list[str]) -> None:
    """README di report/storico/: l'indice dei mesi."""
    t = ["# Storico", "",
         "Ogni cartella è la fotografia di un mese: la sintesi in `README.md` e il dettaglio "
         "dei soli comuni cambiati in `novita.csv`. Le tabelle complete non vengono duplicate "
         "qui, stanno in [`../regioni/`](../regioni/) e la loro storia si legge dai commit.", ""]
    if mesi:
        t += ["| Mese | |", "|---|---|"]
        for m in sorted(mesi, reverse=True):
            t.append(f"| **[{m}]({m}/)** | [sintesi]({m}/README.md) · "
                     f"[variazioni]({m}/novita.csv) |")
        t.append("")
    else:
        t += ["Ancora nessun mese archiviato.", ""]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(t) + "\n", encoding="utf-8")


def scheda_mese(path: Path, mese: str, novita: list[dict], regioni: list[dict],
                aveva_precedente: bool, base: str = "") -> str:
    """
    base vuota = link relativi, giusti dentro il repo.
    base valorizzata = link assoluti, obbligatori nelle note della release:
    GitHub le rende senza un percorso di partenza e i link relativi si rompono.
    """
    link_novita = f"{base}novita.csv" if base else "novita.csv"
    t = [f"# Novità ANNCSU — {mese}", ""]
    if not aveva_precedente:
        t.append("Primo report: non c'è un mese precedente con cui confrontare.")
    elif not novita:
        t.append("Nessun comune ha cambiato numeri rispetto al report precedente.")
    else:
        t.append(f"**{mig(len(novita))} comuni** hanno cambiato qualcosa. Il dettaglio riga per "
                 f"riga è in [`novita.csv`]({link_novita}), ordinato dal comune che ha "
                 "aggiunto più coordinate.")

    t += ["", "## Per regione", "",
          "| Regione | Coordinate aggiunte | Civici aggiunti | Comuni aggiornati | Dati al |",
          "|---|--:|--:|--:|:--:|"]
    for r in regioni:
        t.append(f"| {r['REGIONE']} | {segno(r['COORDINATE_AGGIUNTE'])} | "
                 f"{segno(r['CIVICI_AGGIUNTI'])} | {mig(r['COMUNI_AGGIORNATI'])} | "
                 f"{r['DATI_AGGIORNATI_AL'] or 'n/d'} |")

    top = [r for r in novita if _int(r["COORDINATE_AGGIUNTE"]) > 0][:25]
    if top:
        t += ["", "## Chi ha aggiunto più coordinate", "",
              "| Comune | Provincia | Regione | Coordinate aggiunte | Ora è al |",
              "|---|---|---|--:|--:|"]
        for r in top:
            t.append(f"| {r['COMUNE']} | {r['PROVINCIA']} | {r['REGIONE']} | "
                     f"{segno(r['COORDINATE_AGGIUNTE'])} | {r['COMPLETAMENTO_%']:.0f}% |")

    cali = sorted((r for r in novita if _int(r["COORDINATE_AGGIUNTE"]) < 0),
                  key=lambda r: _int(r["COORDINATE_AGGIUNTE"]))[:15]
    if cali:
        t += ["", "## Comuni in calo", "",
              "Di solito è una ripulitura dell'archivio comunale, ma vale la pena controllarli.", "",
              "| Comune | Provincia | Civici | Coordinate |", "|---|---|--:|--:|"]
        for r in cali:
            t.append(f"| {r['COMUNE']} | {r['PROVINCIA']} | {segno(r['CIVICI_AGGIUNTI'])} | "
                     f"{segno(r['COORDINATE_AGGIUNTE'])} |")

    if base:
        t += ["", "---", "",
              f"Report completo e tabelle regionali: {base.split('/blob/')[0]}"]

    testo = "\n".join(t) + "\n"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(testo, encoding="utf-8")
    return testo


def assolutizza(testo: str, base: str) -> str:
    """
    Riscrive i link relativi del markdown in link assoluti.
    Serve alle note delle release: GitHub le rende senza un percorso di
    partenza, quindi un `(novita.csv)` non punta da nessuna parte.
    """
    import posixpath

    def sostituisci(m):
        etichetta, bersaglio = m.group(1), m.group(2).strip()
        if re.match(r"^(https?:|mailto:|#|/)", bersaglio):
            return m.group(0)
        # normalizzo i ../.. cosi' l'indirizzo resta leggibile
        finale = "/" if bersaglio.endswith("/") else ""
        assoluto = posixpath.normpath(base + bersaglio) + finale
        return f"[{etichetta}]({assoluto.replace('https:/', 'https://', 1)})"
    return re.sub(r"\[([^\]]*)\]\(([^)]+)\)", sostituisci, testo)


def note_storiche(out: Path, dist: Path, repo_url: str) -> list[str]:
    """
    Rigenera le note di ogni mese gia' archiviato, con i link assoluti.
    Le release pubblicate prima della correzione avevano link relativi e
    quindi rotti: cosi' il workflow puo' riallinearle tutte.
    """
    mesi = []
    for cartella in sorted((out / "storico").glob("*")):
        sorgente = cartella / "README.md"
        if not sorgente.is_file():
            continue
        mese = cartella.name
        base = f"{repo_url}/blob/main/report/storico/{mese}/"
        testo = assolutizza(sorgente.read_text(encoding="utf-8"), base)
        if "Report completo e tabelle regionali" not in testo:
            testo += f"\n---\n\nReport completo e tabelle regionali: {repo_url}\n"
        destinazione = dist / "note" / f"{mese}.md"
        destinazione.parent.mkdir(parents=True, exist_ok=True)
        destinazione.write_text(testo, encoding="utf-8")
        mesi.append(mese)
    return mesi


# --------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default="report", help="cartella del report nel repo")
    ap.add_argument("--dist", default="dist", help="cartella dei file da allegare alla release")
    ap.add_argument("--readme", default="README.md",
                    help="il README del repo, rigenerato a ogni report completo")
    ap.add_argument("--lavoro", default=".lavoro", help="cartella temporanea per gli zip")
    ap.add_argument("--regioni", default="", help="codici ANNCSU separati da virgola, per prove")
    ap.add_argument("--mese", default="", help="etichetta YYYY-MM (default: mese corrente)")
    ap.add_argument("--solo-se-aggiornato", action="store_true",
                    help="esce senza scrivere nulla se nessun dataset e' cambiato")
    ap.add_argument("--parziale", action="store_true",
                    help="consente un report incompleto; non aggiorna le schede regionali")
    ap.add_argument("--note-storiche", action="store_true",
                    help="rigenera le note di tutte le release gia' pubblicate e esce")
    ap.add_argument("--controlla", action="store_true",
                    help="guarda solo se i dataset sono cambiati, senza scaricarli")
    ap.add_argument("--pausa", type=float, default=3.0, help="secondi tra un download e l'altro")
    args = ap.parse_args()

    out, dist, lavoro = Path(args.out), Path(args.dist), Path(args.lavoro)
    lavoro.mkdir(parents=True, exist_ok=True)
    mese = args.mese or date.today().strftime("%Y-%m")

    regioni = REGIONI
    if args.regioni:
        voluti = {c.strip().upper() for c in args.regioni.split(",") if c.strip()}
        regioni = [r for r in REGIONI if r[0] in voluti]
        if not regioni:
            log(f"nessuna regione valida in {args.regioni!r}")
            return 2

    repo_url = "https://github.com/" + os.environ.get(
        "GITHUB_REPOSITORY", "checcoconf/anncsu-report")

    if args.note_storiche:
        mesi = note_storiche(out, dist, repo_url)
        log(f"note rigenerate per: {', '.join(mesi) if mesi else 'nessun mese'}")
        if os.environ.get("GITHUB_OUTPUT"):
            with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as f:
                f.write(f"mesi={' '.join(mesi)}\n")
        return 0

    sorgenti_path = out / "sorgenti.json"
    sorgenti_prec = {}
    if sorgenti_path.exists():
        sorgenti_prec = json.loads(sorgenti_path.read_text()).get("dataset", {})

    if args.controlla:
        log("controllo le date dei dataset senza scaricarli")
        _, cambiate = controlla_aggiornamenti(sorgenti_prec, min(args.pausa, 1.0))
        cambiato = bool(cambiate) or not sorgenti_prec
        if cambiato:
            elenco = ", ".join(cambiate) if cambiate else "primo controllo"
            log(f"da aggiornare: {elenco}")
            print(f"::notice::ANNCSU aggiornato ({elenco})")
        else:
            log("nessun dataset regionale e' cambiato")
        if os.environ.get("GITHUB_OUTPUT"):
            with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as f:
                f.write(f"cambiato={'true' if cambiato else 'false'}\n")
                f.write(f"regioni_cambiate={len(cambiate)}\n")
        return 0

    dati: dict[str, dict] = {}
    date_dataset: dict[str, str] = {}
    errori: list[str] = []

    for i, (cod, nome, _) in enumerate(regioni, 1):
        zip_path = lavoro / f"INDIR_{cod}.zip"
        log(f"[{i}/{len(regioni)}] {nome}")
        try:
            scarica(f"{BASE_URL}?INDIR_{cod}", zip_path)
            mb = zip_path.stat().st_size / 1e6
            conteggi, data_ds = conta_regione(zip_path)
            dati[cod], date_dataset[cod] = conteggi, data_ds
            tot = sum(v[0] for v in conteggi.values())
            geo = sum(v[1] for v in conteggi.values())
            log(f"  {mb:.0f} MB · {len(conteggi)} comuni · {tot:,} civici · "
                f"{geo:,} con coordinate ({perc(geo, tot)}%) · dataset {data_ds or 'n/d'}")
        except Exception as e:  # noqa: BLE001
            errori.append(f"{nome}: {e}")
            log(f"  ERRORE {e}")
        finally:
            zip_path.unlink(missing_ok=True)
            time.sleep(args.pausa)

    completo = len(dati) == len(REGIONI)
    if not dati:
        log("nessun dato scaricato, mi fermo")
        return 1
    if errori:
        log(f"attenzione: {len(errori)} regioni non elaborate: {'; '.join(errori)}")
    if not completo and not args.parziale:
        # un report a cui manca una regione produrrebbe il mese dopo migliaia di
        # variazioni fasulle: meglio fallire e rilanciare il workflow
        log("report incompleto: non lo pubblico (usa --parziale per forzare)")
        return 1

    if args.solo_se_aggiornato and sorgenti_prec:
        if all(sorgenti_prec.get(c) == date_dataset.get(c) for c in date_dataset):
            log("nessun dataset regionale e' cambiato: non scrivo nulla")
            print("::notice::Nessun aggiornamento ANNCSU rispetto al report precedente")
            return 0

    anagrafica = carica_anagrafica(Path("data/lookup/comuni-istat.csv"))
    prec = carica_precedente(out / "regioni")
    log(f"report precedente: {len(prec)} comuni")

    comuni = costruisci_righe(dati, anagrafica, prec, date_dataset, mese)
    regioni_agg = riepiloga_regioni(comuni, date_dataset)

    # le novita' vanno in cima chi ha aggiunto di piu': e' la prima cosa che si guarda
    novita = sorted((r for r in comuni if r["COSA_E_CAMBIATO"] not in (INVARIATO, PARTENZA)),
                    key=lambda r: (-_int(r["COORDINATE_AGGIUNTE"]), -_int(r["CIVICI_AGGIUNTI"])))

    if completo:
        for _, nome, slug in REGIONI:
            scrivi_csv(out / "regioni" / f"{slug}.csv", COLONNE_COMUNI,
                       [r for r in comuni if r["REGIONE"] == nome])
        scrivi_csv(out / "novita.csv", COLONNE_COMUNI, novita)
        scrivi_csv(out / "riepilogo-regioni.csv", COLONNE_REGIONI, regioni_agg)
        sorgenti_path.parent.mkdir(parents=True, exist_ok=True)
        sorgenti_path.write_text(
            json.dumps({"mese": mese, "dataset": date_dataset}, indent=2, sort_keys=True) + "\n",
            encoding="utf-8")
        scrivi_readme(Path(args.readme), mese, comuni, regioni_agg, novita, bool(prec))
        mesi = sorted(d.name for d in (out / "storico").glob("*") if d.is_dir()) if \
            (out / "storico").exists() else []
        if mese not in mesi:
            mesi.append(mese)
        readme_cartella_report(out / "README.md", mese, regioni_agg, novita, sorted(mesi, reverse=True)[:12])
        readme_cartella_regioni(out / "regioni" / "README.md", mese, regioni_agg, bool(prec))
        readme_cartella_storico(out / "storico" / "README.md", mesi)

    # lo storico e' parte del repo: lo scrive solo un report completo
    storico = (out if completo else dist) / "storico" / mese
    scrivi_csv(storico / "novita.csv", COLONNE_COMUNI, novita)
    sintesi = scheda_mese(storico / "README.md", mese, novita, regioni_agg, bool(prec))

    dist.mkdir(parents=True, exist_ok=True)
    # copia in dist: le note della release non devono dipendere da dove e'
    # finito lo storico, altrimenti il passo di pubblicazione si rompe
    # le note della release hanno bisogno di link assoluti
    scheda_mese(dist / "note-release.md", mese, novita, regioni_agg, bool(prec),
                base=f"{repo_url}/blob/main/report/storico/{mese}/")
    scrivi_csv(dist / f"comuni-{mese}.csv", COLONNE_COMUNI, comuni)
    scrivi_xlsx(dist / f"anncsu-{mese}.xlsx", comuni, regioni_agg, novita)

    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as f:
            f.write(sintesi)
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as f:
            # solo un report completo va committato e pubblicato: una corsa
            # parziale serve alle prove e non deve toccare il repo
            f.write(f"mese={mese}\nvariazioni={len(novita)}\n")
            f.write(f"pubblica={'true' if completo else 'false'}\n")

    log(f"fatto: {len(comuni)} comuni, {len(novita)} con novita'")
    return 0


if __name__ == "__main__":
    sys.exit(main())