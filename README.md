# ANK_Datenbankanalyse

Struktur- und Qualitätsanalyse der ANK-Access-Datenbank (`.accdb`) mit R –
als Grundlage für automatisierte Prüfungen mit **testthat** und eine spätere
Migration nach **PostgreSQL**.

> **Die Datenbank selbst liegt nie im Repository.** Sie bleibt lokal; der Pfad
> steht in `config_lokal.R`, die von Git ignoriert wird. Auch `Output/` bleibt
> lokal, weil die Ergebnisse Datenbank-Metadaten enthalten.

## Projektstruktur

```
ANK_Datenbankanalyse/
├── ANK_Datenbankanalyse.Rproj   RStudio-Projekt (immer hierüber öffnen)
├── config_vorlage.R             Vorlage -> als config_lokal.R kopieren
├── 00_struktur.R                Schritt 0: Struktur auslesen
├── 01_erd.R                     Schritt 1: ER-Diagramm
├── 02_pruefung.R                Schritt 2: Prüfungen mit testthat
├── tests/testthat/              die Prüfregeln (eine Datei je Bereich)
├── R/
│   ├── funktionen.R             wiederverwendbare Hilfsfunktionen
│   └── dao_metadaten.ps1        liest Beziehungen/Primärschlüssel über DAO
└── Output/                      Ergebnisse (lokal, nicht in Git)
```

Geplant: Profiling der Inhalte (leere Felder, Wertebereiche, Datumsangaben),
HTML-Prüfbericht.

## Einrichtung (einmalig)

1. **Voraussetzungen** auf dem Windows-Rechner: R, RStudio, Git for Windows
   und der **64-Bit**-Treiber *Microsoft Access Database Engine*.
   Prüfen in R:
   ```r
   odbc::odbcListDrivers()$name   # muss "Microsoft Access Driver (*.mdb, *.accdb)" enthalten
   ```
   Fehlt er: 64-Bit „Microsoft Access Database Engine 2016 Redistributable“
   installieren (ggf. über die IT; bei 32-Bit-Office mit `/passive` installieren).

2. **Repository als RStudio-Projekt klonen:**
   *File → New Project → Version Control → Git*
   URL: `https://github.com/ELBE223/ANK_Datenbankanalyse.git`
   Zielordner z. B. `C:/Users/<name>/R/` (nicht in denselben Ordner wie die DB).

3. **Pakete:** nichts zu tun – jedes Skript installiert und lädt seine Pakete
   selbst mit `pacman::p_load()` (siehe Konventionen).

4. **Lokale Konfiguration:** Beim ersten Source legt `00_struktur.R` die Datei
   `config_lokal.R` (Kopie von `config_vorlage.R`) an und öffnet sie.
   `DB_PFAD` eintragen, speichern, erneut Source klicken.

5. **Beziehungen und Primärschlüssel:** normalerweise nichts zu tun.
   `00_struktur.R` liest sie über **DAO**, die Access-eigene Schnittstelle
   (PowerShell-Skript `R/dao_metadaten.ps1`, ohne Makros, schreibgeschützt).
   In der Konsole steht dann „Beziehungen und Primärschlüssel über DAO gelesen“.

   Nur wenn dort „DAO nicht verfügbar“ steht, gibt es zwei Ersatzwege
   (am besten an einer Kopie der DB):

   **Weg A – Leseberechtigung (braucht Makros):** Fehlt „Inhalt aktivieren“,
   ist die Datei meist als „aus dem Internet“ gesperrt: Access schließen,
   im Explorer Rechtsklick auf die `.accdb` → *Eigenschaften* → Haken bei
   **Zulassen** → OK. Dann in der gelben Leiste
   **Inhalt aktivieren** oder den DB-Ordner als vertrauenswürdig eintragen
   (*Datei → Optionen → Trust Center → Einstellungen für das Trust Center →
   Vertrauenswürdige Speicherorte → Neuen Speicherort hinzufügen*), Access neu
   öffnen → `Alt + F11` → `Strg + G` (Direktfenster) → Zeile ausführen:
   ```vb
   CurrentProject.Connection.Execute "GRANT SELECT ON MSysRelationships TO Admin"
   ```

   **Weg B – Excel-Export (ohne Makros, funktioniert immer):**
   *Erstellen → Abfrageentwurf* → *SQL-Ansicht* → Text ersetzen durch
   `SELECT * FROM MSysRelationships;` → *Ausführen* → `Strg + S`, Name
   `beziehungen` → *Externe Daten → Exportieren → Excel* → als
   **`beziehungen.xlsx` im selben Ordner wie die Datenbank** speichern.
   `00_struktur.R` liest diese Datei automatisch, wenn Weg A fehlt.

## Konventionen

- **Jedes Skript wird komplett mit Source gestartet** – Button *Source* oder
  `Strg + Umschalt + S`, alternativ in der Konsole `source("00_struktur.R")`.
  Kein zeilenweises Ausführen nötig.
- **Immer über das RStudio-Projekt arbeiten** (`ANK_Datenbankanalyse.Rproj`),
  damit der Projektordner das Arbeitsverzeichnis ist.
- **Pakete immer mit pacman:** am Anfang jedes Skripts
  ```r
  if (!require("pacman")) install.packages("pacman")
  pacman::p_load(DBI, odbc, writexl, readxl, jsonlite)
  ```
  Fehlende Pakete werden automatisch installiert, vorhandene nur geladen.

## Schritt 0: Struktur auslesen

Projekt in RStudio öffnen, `00_struktur.R` öffnen und **Source** klicken.
Access sollte dabei geschlossen sein.

Ergebnis: `Output/00_struktur.xlsx` mit drei Blättern:

| Blatt         | Inhalt |
|---------------|--------|
| `Objekte`     | Tabellen, Abfragen, verknüpfte Tabellen – mit Anzahl Datensätze und Spalten |
| `Spalten`     | jede Spalte mit Position, Access-Datentyp, Datentyp in R und Schlüssel-Hinweis |
| `Beziehungen` | PK-Tabelle/-Spalte → FK-Tabelle/-Spalte, 1:1 oder 1:n, referentielle Integrität, Kaskaden |

Dieselben Daten liegen als `Output/00_struktur.rds` für die nächsten Skripte vor.

**Schlüssel** in der Spalte `schluessel`:
`PK` = Primärschlüssel (über DAO exakt; ohne DAO nur Näherung: Spalte, auf die
eine Beziehung verweist), `FK` = Fremdschlüssel laut Beziehung,
`AutoWert` = Access-Datentyp COUNTER.

## Schritt 1: ER-Diagramm

`01_erd.R` öffnen und **Source** klicken (braucht nur `Output/00_struktur.rds`,
keinen Datenbankzugriff). Das Diagramm erscheint im Viewer-Fenster von RStudio
und wird als `Output/01_erd.html` gespeichert (im Browser öffnen, zoomen).

- **Farbe** = fachlicher Bereich, **grau** = Werteliste (`_tbd`).
  Die Bereiche stehen oben im Skript (`BEREICHE`) und können angepasst werden.
- **Linie**: Strich = 1-Seite, Krähenfuß = n-Seite; gestrichelt = ohne
  referentielle Integrität; **violett gepunktet** = in Access nicht definiert,
  nur aus dem Spaltennamen `fk_<tabelle>_id` abgeleitet.
- **Maus über eine Tabelle** zeigt alle Spalten mit Schlüsseln, über eine
  Linie die verbundenen Spalten.
- `MIT_HIST <- TRUE` zeigt zusätzlich die Historien-Tabellen (`*_hist`).

Zusätzlich: `Output/01_beziehungen_abgeleitet.csv` (fk_-Spalten ohne Beziehung
in Access) und `Output/01_erd.gv` (Diagramm als Graphviz-Text).

## Schritt 2: Prüfungen mit testthat

`02_pruefung.R` öffnen und **Source** klicken. Das Skript liest die
Schlüsselspalten aus der Datenbank, führt alle Prüfregeln aus und schreibt
`Output/02_pruefergebnis.xlsx` (jede Prüfung mit Status OK/FEHLER und Meldung,
z. B. „12 Verweise ohne Datensatz in wahlkreis_btw21, z. B. 3, 7, 9“).

| Datei in `tests/testthat/` | prüft |
|---|---|
| `test-1_struktur.R` | jede Tabelle hat einen Primärschlüssel, keine temporären Tabellen, Namensregel `fk_<tabelle>_id`, jede fk_-Spalte hat eine Beziehung, gleiche Datentypen |
| `test-2_primaerschluessel.R` | je Tabelle: Primärschlüssel eindeutig und nie leer |
| `test-3_fremdschluessel.R` | je Beziehung: jeder Verweis findet seinen Datensatz (auch abgeleitete Beziehungen) |
| `test-4_fachlich.R` | fachliche Regeln, z. B. 1:1-Beziehungen, leere `_tmp`-Spalten – **hier eigene Regeln ergänzen** |

Die Prüfungen 1–3 entstehen automatisch aus der Strukturanalyse und passen
sich an, wenn sich die Datenbank ändert. Die Tests immer über
`02_pruefung.R` starten, nicht einzeln.

## Typische Fehler

| Meldung | Ursache / Lösung |
|---------|------------------|
| `ODBC-Treiber ... nicht gefunden` | 64-Bit Access Database Engine fehlt (Einrichtung, Punkt 1) |
| `Falscher Arbeitsordner` | Projekt per Doppelklick auf `ANK_Datenbankanalyse.Rproj` öffnen |
| `config_lokal.R wurde angelegt` | `DB_PFAD` in der geöffneten Datei eintragen, speichern, erneut Source |
| `Datenbank nicht gefunden` | Pfad in `config_lokal.R` prüfen (Schrägstriche `/`, Dateiendung `.accdb`) |
| `Paket ... wurde unter R Version ... erstellt` | harmlos, kann ignoriert werden |
| `DAO nicht verfügbar: ...` | Meldung an Claude schicken; bis dahin Einrichtung Punkt 5, Weg A oder B |
| `Beziehungen fehlen` | Einrichtung, Punkt 5 – die übrigen Ergebnisse sind trotzdem vollständig |
| Access: `Die Makros in diesem Projekt sind deaktiviert` | Datei-Eigenschaften → **Zulassen** (Punkt 5, Weg A); sonst Weg B (Excel-Export) |
| Datei gesperrt / `.laccdb` vorhanden | Access schließen und erneut ausführen |
