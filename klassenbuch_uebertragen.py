#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Klassenbuch -> Berichtsheft
===========================

Liest die Themendokumentations-PDFs aus PDF_ORDNER, ordnet jede PDF ueber ihr
Datum der passenden Berichtsheft-Datei und dem passenden Wochentag zu und
traegt Lehrinhalte, Lernfeld-Nr. und Stundenzahl in die Tageszeile ein.

Voraussetzung:  pip install pdfplumber

Aufruf:  python klassenbuch_uebertragen.py
"""

import datetime as dt
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape, unescape

import pdfplumber

# ---------------------------------------------------------------------------
# EINSTELLUNGEN
# ---------------------------------------------------------------------------

# Ordner mit den Klassenbuch-PDFs (wird auch in Unterordnern durchsucht)
PDF_ORDNER = Path("Klassenbuch")

# Wurzelverzeichnis mit den Monatsordnern der Berichtshefte
ZIELORDNER = Path(".")

# Die Spalte "Std." meint volle Zeitstunden, das Klassenbuch zaehlt dagegen
# Unterrichtseinheiten (9 UE = 8 Stunden). Deshalb bleibt die vorbelegte 8
# stehen und wird NICHT aus der PDF uebernommen.
# True = Anzahl der Unterrichtseinheiten aus der PDF eintragen (nicht empfohlen)
STD_AUS_PDF = False

# Was in die Spalte "Lernfeld-Nr." geschrieben wird:
#   "beides"  -> "LF-ZQ3a Agile Scrum Foundation"
#   "titel"   -> "Agile Scrum Foundation"
#   "kuerzel" -> "LF-ZQ3a"
LERNFELD_INHALT = "beides"

# True  = Kursname (Titel) wird zusaetzlich als erste Zeile in die Tageszelle
#         geschrieben. Bei LERNFELD_INHALT "beides"/"titel" sonst doppelt.
MIT_TITEL = False

# Praefix vor jedem Lehrinhalt in der Tageszelle ("" = ohne Aufzaehlungszeichen)
PUNKT = "- "

# True  = Dozent wird oben in "Trainer/Ausbilder:" eingetragen
DOZENT_EINTRAGEN = True

# True  = "Muster, Erika" wird zu "Erika Muster" umgestellt
DOZENT_UMDREHEN = True

# True = nur anzeigen, was passieren wuerde, ohne die Dateien zu aendern
TESTLAUF = False

# ---------------------------------------------------------------------------

MONATE = {
    1: "Januar", 2: "Februar", 3: "M\u00e4rz", 4: "April", 5: "Mai", 6: "Juni",
    7: "Juli", 8: "August", 9: "September", 10: "Oktober", 11: "November",
    12: "Dezember",
}
TAGE = ["Mo", "Di", "Mi", "Do", "Fr"]


# ------------------------------- PDF lesen ---------------------------------

def pdf_auslesen(pfad: Path) -> dict:
    """Liest Datum, Titel, Lernfeld und Lehrinhalte aus einer Themendokumentation."""
    with pdfplumber.open(pfad) as pdf:
        text = "\n".join((s.extract_text() or "") for s in pdf.pages)
        tabellen = []
        for seite in pdf.pages:
            tabellen.extend(seite.extract_tables())

    # Datum: bevorzugt aus dem PDF-Inhalt, sonst aus dem Dateinamen
    datum = None
    m = re.search(r"Datum:\s*(\d{2})\.(\d{2})\.(\d{4})", text)
    if m:
        datum = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    else:
        m = re.search(r"(\d{2})(\d{2})(\d{4})", pfad.name)
        if m:
            datum = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    if datum is None:
        raise ValueError("Kein Datum gefunden")

    # Titel: nachgestellter Zeitraum wird abgeschnitten
    titel = ""
    m = re.search(r"Titel:\s*(.+)", text)
    if m:
        titel = re.sub(r"\s*\d{2}\.\d{2}\.\d{4}\s*-\s*\d{2}\.\d{2}\.\d{4}\s*$",
                       "", m.group(1)).strip()

    # Lernfeld-Kuerzel aus dem Titel, z. B. "LF-ZQ3a", "LF 3", "LF5b"
    m = re.search(r"\bLF[\s.-]*[A-Za-z]{0,3}\d[\w.-]*", titel)
    lernfeld = m.group(0).strip(" .,-") if m else ""
    if lernfeld:
        titel = (titel[:m.start()] + " " + titel[m.end():]).strip(" .,-")
        titel = " ".join(titel.split())

    # Lehrinhalte aus der Stundentabelle
    inhalte, dozenten, stunden = [], [], 0
    for tab in tabellen:
        for zeile in tab:
            if not zeile or len(zeile) < 2:
                continue
            nr = (zeile[0] or "").strip()
            if not nr.isdigit():
                continue
            stunden += 1
            inhalt = " ".join((zeile[1] or "").split())
            if inhalt and inhalt not in inhalte:
                inhalte.append(inhalt)
            if len(zeile) > 3:
                name = _name_formatieren(" ".join((zeile[3] or "").split()))
                if name and name not in dozenten:
                    dozenten.append(name)

    return {"datum": datum, "titel": titel, "lernfeld": lernfeld,
            "inhalte": inhalte, "dozenten": dozenten, "stunden": stunden,
            "quelle": pfad.name}


def _name_formatieren(name: str) -> str:
    """Macht aus "Muster, Erika" bei Bedarf "Erika Muster"."""
    if DOZENT_UMDREHEN and name.count(",") == 1:
        nach, vor = [t.strip() for t in name.split(",")]
        if nach and vor:
            return "%s %s" % (vor, nach)
    return name


# ------------------------------- DOCX schreiben -----------------------------

ZELLE_ENDE = "</w:tc>"


def _zellen_grenzen(xml: str, ab: int, anzahl: int) -> list:
    """Liefert die (start, ende)-Positionen der naechsten Zellen ab Position 'ab'."""
    grenzen, pos = [], ab
    for _ in range(anzahl):
        start = xml.index(ZELLE_ENDE, pos) + len(ZELLE_ENDE)
        ende = xml.index(ZELLE_ENDE, start)
        grenzen.append((start, ende))
        pos = ende
    return grenzen


def _absatz(text: str, fett: bool = False) -> str:
    rpr = '<w:rPr>%s<w:sz w:val="18"/></w:rPr>' % ("<w:b/>" if fett else "")
    return ('<w:p><w:pPr><w:rPr><w:sz w:val="18"/></w:rPr></w:pPr>'
            '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (rpr, escape(text)))


def _text_setzen(zelle: str, wert: str) -> str:
    """Setzt einen kurzen Wert in eine Zelle (ersetzt vorhandenen Text)."""
    treffer = re.search(r"<w:t(?:\s[^>]*)?>", zelle)
    if treffer:
        ende = zelle.index("</w:t>", treffer.end())
        return zelle[:treffer.end()] + escape(wert) + zelle[ende:]
    if "</w:p>" in zelle:
        p = zelle.index("</w:p>")
        run = ('<w:r><w:rPr><w:color w:val="auto"/></w:rPr>'
               '<w:t xml:space="preserve">%s</w:t></w:r>' % escape(wert))
        return zelle[:p] + run + zelle[p:]
    # Zelle enthaelt nur einen leeren, selbstschliessenden Absatz
    tcpr = zelle.index("</w:tcPr>") + len("</w:tcPr>")
    return zelle[:tcpr] + _absatz(wert)


def dozent_setzen(xml: str, namen: list) -> str:
    """Ergaenzt die Namen in der Zelle "Trainer/Ausbilder:" (ohne Dubletten)."""
    m = re.search(r"<w:t(?:\s[^>]*)?>Trainer</w:t>", xml)
    if m is None:
        return xml
    start = xml.rindex("<w:tc>", 0, m.start())
    ende = xml.index(ZELLE_ENDE, m.end())
    zelle = xml[start:ende]

    text = "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", zelle))
    fehlend = [n for n in namen if n not in text]
    if not fehlend:
        return xml

    vorhanden = text.split(":")[-1].strip()
    if vorhanden:
        trenner = ", "
    else:
        trenner = "" if text.endswith(" ") else " "
    run = ('<w:r><w:rPr><w:color w:val="auto"/><w:sz w:val="20"/></w:rPr>'
           '<w:t xml:space="preserve">%s%s</w:t></w:r>'
           % (trenner, escape(", ".join(fehlend))))
    p = zelle.rindex("</w:p>")
    return xml[:start] + zelle[:p] + run + zelle[p:] + xml[ende:]


def _zellentext(zelle: str) -> str:
    return unescape("\n".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", zelle)))


def tag_fuellen(xml: str, tag: str, absaetze: list, lernfeld: str,
                stunden: str) -> tuple:
    """Traegt die Inhalte in die Zeile des Wochentags ein. Gibt (xml, info) zurueck."""
    treffer = None
    for m in re.finditer(r"<w:t(?:\s[^>]*)?>\s*(Mo|Di|Mi|Do|Fr)\s*</w:t>", xml):
        if m.group(1) == tag:
            treffer = m
            break
    if treffer is None:
        raise ValueError("Zeile '%s' nicht gefunden" % tag)

    (t_start, t_ende), (l_start, l_ende), (s_start, s_ende) = \
        _zellen_grenzen(xml, treffer.end(), 3)

    themen = xml[t_start:t_ende]
    # Wichtig: nur den Text DIESER Tageszelle vergleichen. Ein Lehrinhalt wie
    # "Daily Scrum" kommt in derselben Woche an mehreren Tagen vor.
    bestand = [z.strip() for z in _zellentext(themen).split("\n") if z.strip()]
    offen = [a for a in absaetze if a.strip() not in bestand]
    if not offen:
        return xml, "schon vorhanden"

    neu_xml = "".join(_absatz(t, fett=t.endswith(":")) for t in offen)

    if bestand:
        themen_neu = themen + neu_xml          # bestehenden Text nicht antasten
        info = "angehaengt"
    else:
        tcpr = themen.index("</w:tcPr>") + len("</w:tcPr>")
        themen_neu = themen[:tcpr] + neu_xml
        info = "eingetragen"

    # von hinten nach vorn ersetzen, damit die Positionen gueltig bleiben
    if stunden:
        xml = xml[:s_start] + _text_setzen(xml[s_start:s_ende], stunden) + xml[s_ende:]
    if lernfeld:
        zelle = xml[l_start:l_ende]
        alt = " ".join(_zellentext(zelle).split())
        if alt and lernfeld not in alt:
            lernfeld = alt + ", " + lernfeld
        elif alt:
            lernfeld = alt
        xml = xml[:l_start] + _text_setzen(zelle, lernfeld) + xml[l_ende:]
    xml = xml[:t_start] + themen_neu + xml[t_ende:]
    return xml, info


def docx_bearbeiten(pfad: Path, tag: str, absaetze: list, lernfeld: str,
                    stunden: str, dozenten=()):
    with zipfile.ZipFile(pfad) as zin:
        eintraege = zin.infolist()
        inhalte = {e.filename: zin.read(e.filename) for e in eintraege}

    xml = inhalte["word/document.xml"].decode("utf-8")
    xml, info = tag_fuellen(xml, tag, absaetze, lernfeld, stunden)
    geaendert = info != "schon vorhanden"

    if DOZENT_EINTRAGEN and dozenten:
        vorher = xml
        xml = dozent_setzen(xml, list(dozenten))
        if xml != vorher:
            if not geaendert:
                info = "Dozent ergaenzt"
            geaendert = True

    if not geaendert:
        return info
    if TESTLAUF:
        return info + " (Testlauf)"

    inhalte["word/document.xml"] = xml.encode("utf-8")
    with zipfile.ZipFile(pfad, "w", zipfile.ZIP_DEFLATED) as zout:
        for e in eintraege:
            zout.writestr(e, inhalte[e.filename])
    return info


# ------------------------------- Ablauf -------------------------------------

def lernfeld_text(eintrag: dict) -> str:
    """Stellt den Wert fuer die Spalte "Lernfeld-Nr." zusammen."""
    kuerzel, titel = eintrag["lernfeld"], eintrag["titel"]
    if LERNFELD_INHALT == "kuerzel":
        if not kuerzel:
            print("   HINWEIS: kein LF-Kuerzel im Titel %r, nehme den Titel"
                  % titel)
            return titel
        return kuerzel
    if LERNFELD_INHALT == "titel":
        return titel or kuerzel
    return " ".join(t for t in (kuerzel, titel) if t)


def zieldatei(datum: dt.date) -> Path:
    montag = datum - dt.timedelta(days=datum.weekday())
    freitag = montag + dt.timedelta(days=4)
    ordner = ZIELORDNER / ("%02d_%d - %s" % (montag.month, montag.year,
                                             MONATE[montag.month]))
    return ordner / ("KW%02d - %s - %s.docx" % (montag.isocalendar()[1],
                                                montag.strftime("%d.%m."),
                                                freitag.strftime("%d.%m.")))


def main() -> None:
    pdfs = sorted(PDF_ORDNER.rglob("*.pdf"))
    if not pdfs:
        raise SystemExit("Keine PDFs in %s gefunden." % PDF_ORDNER.resolve())

    # alle PDFs einlesen und nach Tag buendeln
    proTag = {}
    for pdf in pdfs:
        try:
            d = pdf_auslesen(pdf)
        except Exception as fehler:
            print("FEHLER  %s: %s" % (pdf.name, fehler))
            continue
        if d["datum"].weekday() > 4:
            print("UEBERSPRUNGEN  %s: %s ist ein Wochenende"
                  % (pdf.name, d["datum"].strftime("%d.%m.%Y")))
            continue
        proTag.setdefault(d["datum"], []).append(d)

    for datum in sorted(proTag):
        eintraege = proTag[datum]
        ziel = zieldatei(datum)
        if not ziel.is_file():
            print("FEHLT   %s -> %s" % (datum.strftime("%d.%m.%Y"), ziel))
            continue

        absaetze, lernfelder, dozenten, stunden = [], [], [], 0
        for e in eintraege:
            if MIT_TITEL and e["titel"]:
                absaetze.append(e["titel"] + ":")
            absaetze.extend(PUNKT + i for i in e["inhalte"])
            lf = lernfeld_text(e)
            if lf and lf not in lernfelder:
                lernfelder.append(lf)
            for name in e["dozenten"]:
                if name not in dozenten:
                    dozenten.append(name)
            stunden += e["stunden"]

        info = docx_bearbeiten(
            ziel, TAGE[datum.weekday()], absaetze, ", ".join(lernfelder),
            str(stunden) if (STD_AUS_PDF and stunden) else "",
            dozenten)

        print("%s %s  %-16s %s  (%d UE, %d Zeilen)"
              % (datum.strftime("%d.%m.%Y"), TAGE[datum.weekday()], info,
                 ziel.name, stunden, len(absaetze)))


if __name__ == "__main__":
    main()
