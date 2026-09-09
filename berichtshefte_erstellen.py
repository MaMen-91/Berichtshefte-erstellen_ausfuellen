#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Berichtsheft-Generator
======================

Legt pro Monatsordner ("09_2026 - September") eine Word-Datei je Ausbildungswoche an
("KW37 - 07.09. - 11.09..docx") und traegt "Ausbildungswoche vom / bis" sowie das
Ausbildungsjahr direkt im Dokument ein.

Benoetigt nur die Python-Standardbibliothek (ab Python 3.7).

Aufruf:  python berichtshefte_erstellen.py
"""

import datetime as dt
import re
import zipfile
from pathlib import Path

# ---------------------------------------------------------------------------
# EINSTELLUNGEN
# ---------------------------------------------------------------------------

# Pfad zur Blanko-Vorlage
VORLAGE = Path("Blanko-Berichtsheft.docx")

# Wurzelverzeichnis, in dem die Monatsordner liegen bzw. angelegt werden
ZIELORDNER = Path(".")

# Erster Montag, der erzeugt werden soll.
# 31.08.2026 = Montag der KW36 (31.08.-04.09.) -> landet im August-Ordner.
# Auf date(2026, 9, 7) ändern, falls diese Woche schon existiert.
START = dt.date(2026, 8, 29)

# Letzter Ausbildungstag. Die Woche, in die dieses Datum faellt, wird noch erzeugt.
ENDE = dt.date(2026, 12, 31)

# True = vorhandene Dateien werden ueberschrieben, False = werden uebersprungen
UEBERSCHREIBEN = False

# Fortlaufende Nummer fuer "Ausbildungsnachweis Nr."
# Feste Bezugswoche: der Montag dieser Woche bekommt NR_START.
# 31.08.2026 = KW36, die auf die zuletzt vergebene Nr. 32 (KW35) folgt.
# Diese beiden Werte NICHT zum Testen aendern - sonst verschiebt sich die
# Nummerierung. Zum Testen nur START/ENDE oben anpassen.
NR_BEZUGSMONTAG = dt.date(2026, 8, 31)
NR_START = 33

# Datum ueber der linken Unterschriftslinie. "" = kein Datum eintragen.
# Es wird der letzte Tag der Woche (Freitag) verwendet.
DATUM_FORMAT = "%d.%m.%Y"

# Bilddatei mit der eingescannten Unterschrift (PNG oder JPG).
# None = keine Unterschrift einfuegen.
UNTERSCHRIFT = None          # z. B. Path("unterschrift.png")
UNTERSCHRIFT_HOEHE_CM = 1.2  # Anzeigehoehe, Breite wird passend berechnet

# ---------------------------------------------------------------------------

MONATE = {
    1: "Januar", 2: "Februar", 3: "M\u00e4rz", 4: "April", 5: "Mai", 6: "Juni",
    7: "Juli", 8: "August", 9: "September", 10: "Oktober", 11: "November",
    12: "Dezember",
}


EMU_PRO_CM = 360000


def _bildgroesse(daten: bytes) -> tuple:
    """Liest Breite und Hoehe aus PNG- oder JPEG-Daten (ohne Zusatzbibliothek)."""
    if daten[:8] == b"\x89PNG\r\n\x1a\n":
        return (int.from_bytes(daten[16:20], "big"),
                int.from_bytes(daten[20:24], "big"))
    if daten[:2] == b"\xff\xd8":
        i = 2
        while i < len(daten) - 9:
            if daten[i] != 0xFF:
                i += 1
                continue
            marker, laenge = daten[i + 1], int.from_bytes(daten[i + 2:i + 4], "big")
            if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                return (int.from_bytes(daten[i + 7:i + 9], "big"),
                        int.from_bytes(daten[i + 5:i + 7], "big"))
            i += 2 + laenge
    raise ValueError("Bildformat nicht erkannt (nur PNG und JPG)")


def _bild_run(rid: str, breite: int, hoehe: int) -> str:
    """Baut einen Word-Run mit einem eingebetteten Bild."""
    return (
        '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%d" cy="%d"/><wp:docPr id="99" name="Unterschrift"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="0" name="Unterschrift"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
        % (breite, hoehe, rid, breite, hoehe))


def unterschrift_absatz(xml: str, datum: str, bild_run: str) -> str:
    """Setzt Datum und Unterschrift direkt ueber die linke Unterschriftslinie."""
    if not datum and not bild_run:
        return xml
    marker = "Durch die nachfolgende Unterschrift"
    i = xml.find(marker)
    if i == -1:
        return xml
    einfuegen = xml.index("</w:p>", i) + len("</w:p>")

    inhalt = bild_run
    if datum:
        if bild_run:
            inhalt += ('<w:r><w:rPr><w:sz w:val="18"/></w:rPr>'
                       '<w:t xml:space="preserve">   </w:t></w:r>')
        inhalt += ('<w:r><w:rPr><w:sz w:val="18"/></w:rPr>'
                   '<w:t xml:space="preserve">%s</w:t></w:r>' % datum)

    absatz = ('<w:p><w:pPr><w:spacing w:after="0"/>'
              '<w:rPr><w:sz w:val="18"/></w:rPr></w:pPr>%s</w:p>' % inhalt)
    return xml[:einfuegen] + absatz + xml[einfuegen:]


def zelle_fuellen(xml: str, marker: str, wert: str, ab: int = 0) -> tuple:
    """Schreibt 'wert' in die Tabellenzelle rechts neben der Zelle mit dem Text 'marker'.

    Gibt (neues_xml, position_hinter_dem_marker) zurueck.
    """
    tag = "<w:t>%s</w:t>" % marker
    i = xml.find(tag, ab)
    if i == -1:
        raise ValueError("Text '%s' nicht in der Vorlage gefunden." % marker)

    # Ende der Zelle, die den Marker enthaelt -> danach beginnt die Zielzelle
    start = xml.index("</w:tc>", i) + len("</w:tc>")
    ende = xml.index("</w:tc>", start)
    zelle = xml[start:ende]

    treffer = re.search(r"<w:t(?:\s[^>]*)?>", zelle)
    if treffer:
        # Zelle enthaelt schon Text (z. B. "2026") -> Text ersetzen
        t_open_end = treffer.end()
        t_close = zelle.index("</w:t>", t_open_end)
        neu = zelle[:t_open_end] + wert + zelle[t_close:]
    else:
        # Leere Zelle -> neuen Run vor dem Absatzende einfuegen
        p_ende = zelle.index("</w:p>")
        run = ('<w:r><w:rPr><w:color w:val="auto"/></w:rPr>'
               '<w:t xml:space="preserve">%s</w:t></w:r>' % wert)
        neu = zelle[:p_ende] + run + zelle[p_ende:]

    return xml[:start] + neu + xml[ende:], i + len(tag)


def dokument_schreiben(vorlage: Path, ziel: Path, von: str, bis: str,
                       jahr: str, nummer: str, datum: str = "") -> None:
    """Kopiert die Vorlage und traegt die Werte ein."""
    with zipfile.ZipFile(vorlage) as zin:
        inhalte = {e.filename: zin.read(e.filename) for e in zin.infolist()}

    xml = inhalte["word/document.xml"].decode("utf-8")
    xml, pos = zelle_fuellen(xml, "Ausbildungswoche vom", von)
    xml, _ = zelle_fuellen(xml, "bis", bis, ab=pos)
    xml, _ = zelle_fuellen(xml, "Ausbildungsjahr", jahr)
    xml, _ = zelle_fuellen(xml, "Ausbildungsnachweis Nr.", nummer)

    bild_run = ""
    if UNTERSCHRIFT and Path(UNTERSCHRIFT).is_file():
        bild = Path(UNTERSCHRIFT).read_bytes()
        endung = Path(UNTERSCHRIFT).suffix.lower().lstrip(".").replace("jpeg", "jpg")
        px_b, px_h = _bildgroesse(bild)
        hoehe = int(UNTERSCHRIFT_HOEHE_CM * EMU_PRO_CM)
        breite = int(hoehe * px_b / px_h)

        rels = inhalte["word/_rels/document.xml.rels"].decode("utf-8")
        nummern = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels)]
        rid = "rId%d" % (max(nummern) + 1 if nummern else 1)
        inhalte["word/_rels/document.xml.rels"] = rels.replace(
            "</Relationships>",
            '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/image" Target="media/unterschrift.%s"/>'
            "</Relationships>" % (rid, endung)).encode("utf-8")

        typen = inhalte["[Content_Types].xml"].decode("utf-8")
        if 'Extension="%s"' % endung not in typen:
            art = "image/png" if endung == "png" else "image/jpeg"
            typen = typen.replace(
                "<Default", '<Default Extension="%s" ContentType="%s"/><Default'
                % (endung, art), 1)
            inhalte["[Content_Types].xml"] = typen.encode("utf-8")

        inhalte["word/media/unterschrift.%s" % endung] = bild
        bild_run = _bild_run(rid, breite, hoehe)

    xml = unterschrift_absatz(xml, datum, bild_run)
    inhalte["word/document.xml"] = xml.encode("utf-8")

    ziel.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, daten in inhalte.items():
            zout.writestr(name, daten)


def main() -> None:
    if not VORLAGE.is_file():
        raise SystemExit("Vorlage nicht gefunden: %s" % VORLAGE.resolve())

    montag = START - dt.timedelta(days=START.weekday())
    letzter_montag = ENDE - dt.timedelta(days=ENDE.weekday())

    erstellt = uebersprungen = 0

    while montag <= letzter_montag:
        freitag = montag + dt.timedelta(days=4)
        kw = montag.isocalendar()[1]

        ordner = ZIELORDNER / ("%02d_%d - %s" % (montag.month, montag.year,
                                                 MONATE[montag.month]))
        nummer = NR_START + (montag - NR_BEZUGSMONTAG).days // 7
        von = montag.strftime("%d.%m.")
        bis = freitag.strftime("%d.%m.")
        datei = ordner / ("KW%02d - %s - %s.docx" % (kw, von, bis))

        if datei.exists() and not UEBERSCHREIBEN:
            print("uebersprungen (existiert): %s" % datei)
            uebersprungen += 1
        else:
            dokument_schreiben(VORLAGE, datei, von, bis, str(montag.year),
                               str(nummer),
                               freitag.strftime(DATUM_FORMAT) if DATUM_FORMAT else "")
            print("angelegt: Nr. %d  %s" % (nummer, datei))
            erstellt += 1

        montag += dt.timedelta(days=7)

    print("\nFertig. %d Dateien angelegt, %d uebersprungen." % (erstellt, uebersprungen))


if __name__ == "__main__":
    main()
