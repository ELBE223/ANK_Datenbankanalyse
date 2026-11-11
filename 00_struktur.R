# =============================================================================
# 00_struktur.R – Struktur der ANK-Access-Datenbank auslesen
# -----------------------------------------------------------------------------
# Fragen:   Welche Tabellen/Abfragen gibt es? Wie groß sind sie?
#           Welche Spalten und Datentypen? Wie hängen die Tabellen zusammen?
# Ergebnis: Output/00_struktur.xlsx  (Blätter: Objekte, Spalten, Beziehungen)
#           Output/00_struktur.rds   (dieselben Daten für die nächsten Skripte)
# Hinweis:  Die Datenbank wird nur gelesen, nicht verändert.
# Start:    in RStudio "Source" klicken (Strg + Umschalt + S)
# =============================================================================

# Pakete: pacman installiert fehlende Pakete automatisch und lädt sie ---------
if (!require("pacman")) install.packages("pacman")
pacman::p_load(DBI, odbc, writexl, readxl, jsonlite)

# Projektordner und lokale Konfiguration prüfen --------------------------------
if (!file.exists("ANK_Datenbankanalyse.Rproj")) {
  stop("Falscher Arbeitsordner: ", getwd(), "\nProjekt über ",
       "ANK_Datenbankanalyse.Rproj öffnen und erneut Source klicken.", call. = FALSE)
}
if (!file.exists("config_lokal.R")) {
  file.copy("config_vorlage.R", "config_lokal.R")
  if (interactive()) file.edit("config_lokal.R")
  stop("config_lokal.R wurde angelegt und geöffnet: DB_PFAD eintragen, ",
       "speichern und erneut Source klicken.", call. = FALSE)
}
source("config_lokal.R")   # setzt DB_PFAD
source("R/funktionen.R")

con <- db_verbinden(DB_PFAD)

# 1) Objekte: Tabellen, Abfragen, verknüpfte Tabellen + Anzahl Datensätze -----
objekte <- db_objekte(con)
objekte$datensaetze <- vapply(objekte$objekt, db_anzahl, numeric(1),
                              con = con, USE.NAMES = FALSE)

# 2) Spalten: Name, Reihenfolge, Access-Datentyp, Datentyp in R ---------------
spalten <- do.call(rbind, lapply(objekte$objekt, db_spalten, con = con))
objekte$spalten <- as.vector(table(factor(spalten$objekt, levels = objekte$objekt)))

# 3) Beziehungen und Primärschlüssel ------------------------------------------
#    über DAO (Access-Schnittstelle); Ersatz: Systemtabelle oder Excel-Export
dao <- db_dao(DB_PFAD)
beziehungen <- db_beziehungen(con, dao = dao,
                              export = file.path(dirname(DB_PFAD), "beziehungen.xlsx"))
spalten$schluessel <- schluessel_hinweis(spalten, beziehungen, pk = dao$pk)

DBI::dbDisconnect(con)

# 4) Speichern -----------------------------------------------------------------
struktur <- list(Objekte = objekte, Spalten = spalten, Beziehungen = beziehungen)
dir.create("Output", showWarnings = FALSE)
saveRDS(struktur, "Output/00_struktur.rds")
writexl::write_xlsx(struktur, "Output/00_struktur.xlsx")

# 5) Kurzer Überblick in der Konsole ------------------------------------------
cat("\nDatenbank:  ", basename(DB_PFAD),
    "\nObjekte:    ", paste(names(table(objekte$typ)), table(objekte$typ),
                            sep = " = ", collapse = ", "),
    "\nLeer:       ", sum(objekte$datensaetze == 0, na.rm = TRUE),
    "Tabellen ohne Datensätze",
    "\nSpalten:    ", nrow(spalten),
    "\nBeziehungen:", nrow(beziehungen),
    "\nPrimärschl.:", if (is.null(dao)) "nicht verfügbar (nur Hinweise)" else
                        paste(length(unique(dao$pk$tabelle)), "Tabellen mit Primärschlüssel"),
    "\nErgebnis:    Output/00_struktur.xlsx\n\n")
print(objekte[order(-objekte$datensaetze), ], row.names = FALSE)

# Zum Anschauen in RStudio:  View(objekte); View(spalten); View(beziehungen)
