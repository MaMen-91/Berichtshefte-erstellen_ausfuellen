# Berichtsheft-Automatisierung

Zwei Python-Skripte für den Ausbildungsnachweis: Das erste legt die
Wochendateien an, das zweite überträgt die Klassenbucheinträge des
Bildungsträgers hinein.

## Was die Skripte tun

**`berichtshefte_erstellen.py`** legt pro Monat einen Ordner
(`09_2026 - September`) und darin je Ausbildungswoche eine Word-Datei
(`KW37 - 07.09. - 11.09..docx`) an. Ausbildungswoche, Ausbildungsjahr und die
fortlaufende Nachweisnummer werden direkt im Dokument eingetragen, dazu das
Datum des Wochenendes über der Unterschriftslinie. Eine Unterschrift kann
optional als Bild eingebettet werden. Wochen, die über einen Monatswechsel
laufen, landen im Ordner des Monats, in dem der Montag liegt.

**`klassenbuch_uebertragen.py`** liest die Themendokumentations-PDFs, ordnet
jede über ihr Datum der richtigen Woche und dem richtigen Wochentag zu und
schreibt Lehrinhalte, Lernfeld und Dozent in die passende Tageszeile.
Mehrfaches Ausführen ist gefahrlos: Bereits vorhandene Einträge werden erkannt
und nicht doppelt eingetragen.

**`unterschrift_freistellen.py`** macht den weißen Hintergrund eines
Unterschriften-Screenshots transparent und schneidet ihn zu.

## Voraussetzungen

Damit die Vorlage auch in den Ordnern deinen Namen, Klasse etc. hat, muss das in der 
Blanko-Datei ausgefüllt sein.

```
python -m venv .venv
.\.venv\Scripts\Activate.ps1     # Windows, unter Linux/macOS: source .venv/bin/activate
pip install pdfplumber pillow
```

`berichtshefte_erstellen.py` läuft ohne zusätzliche Pakete, es nutzt nur die
Standardbibliothek. `pdfplumber` wird für die Klassenbuch-Übertragung
gebraucht, `pillow` nur für das Freistellen der Unterschrift.

## Verwendung

```
python berichtshefte_erstellen.py
python klassenbuch_uebertragen.py
```

Beide Skripte werden über Konstanten im Kopf der Datei gesteuert, dort stehen
Zeitraum, Nummerierung, Pfade und Formatierung. Für einen ersten gefahrlosen
Durchlauf lässt sich in `klassenbuch_uebertragen.py` `TESTLAUF = True` setzen,
dann wird nur angezeigt, was passieren würde.

## Ordnerstruktur

```
.
├── berichtshefte_erstellen.py
├── klassenbuch_uebertragen.py
├── unterschrift_freistellen.py
├── Blanko-Berichtsheft.docx     Vorlage ohne Namen
├── Klassenbuch/                 PDFs des Bildungsträgers (nicht im Repository)
└── 09_2026 - September/         erzeugte Berichtshefte (nicht im Repository)
```

## Hinweis zu personenbezogenen Daten

Die Klassenbuch-PDFs enthalten Namen von Dozentinnen und Dozenten, die
erzeugten Berichtshefte den eigenen Namen und gegebenenfalls ein
Unterschriftsbild. Beides ist über die `.gitignore` ausgeschlossen und sollte
dort auch bleiben. Die mitgelieferte Vorlage ist anonymisiert.

Die Vorlage stammt aus dem Ausbildungsbetrieb und ist nur zur
Veranschaulichung beigelegt.
