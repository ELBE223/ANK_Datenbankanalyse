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
│   └── funktionen.R             wiederverwendbare Hilfsfunktionen
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

3. **Pakete installieren:**
   ```r
   install.packages(c("DBI", "odbc", "writexl", "testthat"))
   ```

4. **Lokale Konfiguration:** `config_vorlage.R` kopieren, als `config_lokal.R`
   speichern und `DB_PFAD` eintragen.

5. **Beziehungen lesbar machen** (nur nötig, damit `00_struktur.R` die in Access
   definierten Beziehungen findet; am besten an einer Kopie der DB):
   Datenbank in Access öffnen → `Alt + F11` → `Strg + G` (Direktfenster) →
   folgende Zeile eingeben und mit Enter ausführen → Access schließen.
   ```vb
   CurrentProject.Connection.Execute "GRANT SELECT ON MSysRelationships TO Admin"
   ```

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

**Schlüssel-Hinweise** in der Spalte `schluessel`:
`AutoWert` = Access-Datentyp COUNTER (meist Primärschlüssel),
`PK` = wird von einer Beziehung referenziert, `FK` = verweist auf eine andere Tabelle.
Der Access-ODBC-Treiber gibt Primärschlüssel nicht direkt aus. Eine exakte
Übersicht aller Schlüssel und Indizes liefert Access selbst:
*Datenbanktools → Datenbankdokumentierer*.

## Typische Fehler

| Meldung | Ursache / Lösung |
|---------|------------------|
| `ODBC-Treiber ... nicht gefunden` | 64-Bit Access Database Engine fehlt (Einrichtung, Punkt 1) |
| `config_lokal.R fehlt` | Einrichtung, Punkt 4 |
| `MSysRelationships ist nicht lesbar` | Einrichtung, Punkt 5 – die übrigen Ergebnisse sind trotzdem vollständig |
| Datei gesperrt / `.laccdb` vorhanden | Access schließen und erneut ausführen |
