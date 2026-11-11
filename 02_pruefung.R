# =============================================================================
# 02_pruefung.R – Datenqualität der ANK-Datenbank mit testthat prüfen
# -----------------------------------------------------------------------------
# Prüft:    Struktur, Primärschlüssel, Fremdschlüssel, fachliche Regeln
#           (die Regeln stehen in tests/testthat/test-*.R)
# Eingabe:  Output/00_struktur.rds (vorher 00_struktur.R) und die Datenbank
# Ergebnis: Output/02_pruefergebnis.xlsx  (jede Prüfung mit Status und Meldung)
# Start:    in RStudio "Source" klicken (Strg + Umschalt + S)
# Hinweis:  Die Datenbank wird nur gelesen, nicht verändert.
# =============================================================================

# Pakete: pacman installiert fehlende Pakete automatisch und lädt sie ---------
if (!require("pacman")) install.packages("pacman")
pacman::p_load(DBI, odbc, testthat, writexl)

if (!file.exists("ANK_Datenbankanalyse.Rproj") || !file.exists("config_lokal.R")) {
  stop("Projekt über ANK_Datenbankanalyse.Rproj öffnen und zuerst 00_struktur.R ausführen.",
       call. = FALSE)
}
source("config_lokal.R")
source("R/funktionen.R")
if (!file.exists("Output/00_struktur.rds")) stop("Zuerst 00_struktur.R ausführen.", call. = FALSE)

# 1) Was wird geprüft? Tabellen, Primärschlüssel, Beziehungen ------------------
PRUEF_UMFANG <- pruefumfang(readRDS("Output/00_struktur.rds"))

# 2) Schlüsselspalten aus der Datenbank lesen ----------------------------------
con <- db_verbinden(DB_PFAD)
PRUEF_DATEN <- schluessel_lesen(con, PRUEF_UMFANG)
DBI::dbDisconnect(con)

# 3) Prüfungen mit testthat ausführen ------------------------------------------
#    Ausgabe: je Bereich ein Punkt pro bestandener Prüfung, eine Nummer pro
#    Fehler; danach alle Fehler mit Meldung.
ergebnis <- testthat::test_dir(
  "tests/testthat",
  reporter = testthat::SummaryReporter$new(show_praise = FALSE, max_reports = Inf),
  stop_on_failure = FALSE
)

# 4) Prüfbericht speichern -------------------------------------------------------
bericht <- pruef_bericht(ergebnis)
dir.create("Output", showWarnings = FALSE)
writexl::write_xlsx(bericht, "Output/02_pruefergebnis.xlsx")

# 5) Kurzer Überblick in der Konsole ------------------------------------------
cat("\n================ Prüfbericht ================\n")
print(table(Bereich = bericht$bereich, Status = bericht$status))
cat("\n", sum(bericht$status == "FEHLER"), "von", nrow(bericht), "Prüfungen mit Fehlern.",
    "\nAlle Prüfungen mit Meldung: Output/02_pruefergebnis.xlsx\n")
