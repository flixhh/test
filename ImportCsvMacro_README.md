# CSV-Import-Makro (`ImportCsvMacro.bas`)

Ein Excel-VBA-Makro, das CSV-Dateien aus dem **aktuellsten** Timestamp-Ordner
ausliest und deren Inhalte in vorhandene Tabellenblätter kopiert.

## Funktionsweise

1. **Ordnerwahl:** Im Basisordner werden alle Unterordner betrachtet, deren Name
   wie ein Timestamp aussieht (beginnt mit ≥ 8 Ziffern, z. B. `20260629_140530`
   oder `2026-06-29_14-05-30`). Der „aktuellste" (größte Name per Textvergleich)
   wird gewählt.
2. **Tabellenblatt bestimmen:** Pro CSV-Datei wird der Dateiname **ohne** `.csv`
   als Tabellenblattname verwendet (z. B. `Umsatz.csv` → Blatt `Umsatz`).
3. **Leeren:** Der bisherige Inhalt des Tabellenblatts wird gelöscht
   (`Cells.Clear`).
4. **Einfügen:** Der CSV-Inhalt wird ab Zelle `A1` eingefügt.
5. **Schleife** über alle CSV-Dateien des Ordners.
6. **Log:** Im Tabellenblatt `Log` wird eine Zeile angehängt mit
   Ausführungszeitpunkt, Ordner-Timestamp sowie Anzahl importierter/übersprungener
   Dateien. Das Log-Blatt wird bei Bedarf automatisch erstellt.

## Einrichtung

1. Excel öffnen → `Alt + F11` (VBA-Editor).
2. Menü **Datei → Datei importieren…** → `ImportCsvMacro.bas` auswählen.
   (Alternativ: `Einfügen → Modul` und den Inhalt der `.bas`-Datei einfügen.)
3. Im Modul oben die **Einstellungen** anpassen:
   - `BASE_FOLDER` – Pfad zum Ordner mit den Timestamp-Unterordnern.
     Leer lassen (`""`), um den Ordner der Arbeitsmappe zu verwenden.
   - `CSV_DELIMITER` – Trennzeichen, Standard `;`. Auf `""` setzen für
     automatische Erkennung (`;`, `,`, Tab).
   - `LOG_SHEET_NAME` – Name des Log-Blatts (Standard `Log`).
   - `SKIP_IF_SHEET_MISSING` – `True` = Dateien ohne passendes Blatt überspringen,
     `False` = fehlendes Blatt automatisch anlegen.
4. Arbeitsmappe als **`.xlsm`** (mit Makros) speichern.

## Ausführen

- `Alt + F8` → `ImportiereCsvDateien` → **Ausführen**, oder
- Eine Schaltfläche (Formularsteuerelement) einfügen und mit dem Makro verknüpfen.

## Hinweise

- Der CSV-Parser berücksichtigt in Anführungszeichen stehende Felder mit
  Trennzeichen, Zeilenumbrüchen und doppelten Anführungszeichen (`""`).
- Das Einlesen versucht zunächst UTF-8 (inkl. BOM-Entfernung) und fällt sonst
  auf ANSI zurück.
- Die Daten werden als Text/Werte eingefügt (keine Formelauswertung).
