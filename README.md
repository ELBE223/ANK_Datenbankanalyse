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
├── R/
│   ├── funktionen.R             wiederverwendbare Hilfsfunktionen
│   └── dao_metadaten.ps1        liest Beziehungen/Primärschlüssel über DAO
└── Output/                      Ergebnisse (lokal, nicht in Git)
```

Geplant: `01_erd.R` (ER-Diagramm), `02_profiling.R` (Datenqualität),
`tests/testthat/` (automatisierte Prüfungen), HTML-Prüfbericht.

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
