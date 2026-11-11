# =============================================================================
# 01_erd.R – ER-Diagramm der ANK-Datenbank
# -----------------------------------------------------------------------------
# Frage:    Wie hängen die Tabellen zusammen?
# Eingabe:  Output/00_struktur.rds  (vorher 00_struktur.R ausführen)
# Ergebnis: Output/01_erd.html                    interaktives Diagramm
#           Output/01_beziehungen_abgeleitet.csv  fk_-Spalten ohne Beziehung
# Start:    in RStudio "Source" klicken (Strg + Umschalt + S)
# Hinweis:  Braucht keinen Zugriff auf die Datenbank.
# =============================================================================

# Pakete: pacman installiert fehlende Pakete automatisch und lädt sie ---------
if (!require("pacman")) install.packages("pacman")
pacman::p_load(DiagrammeR, htmlwidgets)

source("R/funktionen.R")
if (!file.exists("Output/00_struktur.rds")) stop("Zuerst 00_struktur.R ausführen.", call. = FALSE)
struktur <- readRDS("Output/00_struktur.rds")

# Einstellungen -----------------------------------------------------------------
MIT_HIST <- FALSE   # TRUE = Historien-Tabellen (*_hist) mit anzeigen

# Fachliche Bereiche (Farbe): Muster der Reihe nach, der erste Treffer zählt.
# Wertelisten (*_tbd) sind immer grau. Hier bei Bedarf anpassen.
BEREICHE <- c(
  "Projekt & Förderung" = "^(projekt|foerder|verbundvorhaben|indikator|umsetzung|berichtswesen|sachstandsbericht|zwischenverwendungsnachweis)",
  "Beteiligte"          = "^(zuwendungsempfaenger|kontakt|massnahmenumsetzer)",
  "Geodaten"            = "^(geoinformation|wahlkreis|wkr|plz)"
)
FARBEN <- c(stats::setNames(c("#2a78d6", "#eb6834", "#1baf7a"), names(BEREICHE)),
            "Sonstiges" = "#8c8c8c")

# 1) Tabellen auswählen ---------------------------------------------------------
objekte <- struktur$Objekte[struktur$Objekte$typ == "Tabelle", ]
if (!MIT_HIST) objekte <- objekte[!endsWith(objekte$objekt, "_hist"), ]

# 2) Beziehungen: aus Access + aus Spaltennamen abgeleitet ---------------------
aus_access  <- struktur$Beziehungen[struktur$Beziehungen$fk_tabelle %in% objekte$objekt, ]
abgeleitet  <- beziehungen_ableiten(struktur$Spalten, aus_access, objekte$objekt)
beziehungen <- rbind(cbind(aus_access, quelle = rep("Access", nrow(aus_access))),
                     cbind(abgeleitet, quelle = rep("abgeleitet", nrow(abgeleitet))))

# 3) Rolle und Bereich je Tabelle ----------------------------------------------
objekte$rolle   <- tabellen_rolle(objekte, beziehungen)
objekte$bereich <- tabellen_bereich(objekte, BEREICHE)

# 4) Diagramm zeichnen und speichern -------------------------------------------
dot <- erd_dot(objekte, beziehungen, struktur$Spalten, FARBEN)
erd <- DiagrammeR::grViz(dot, width = "100%", height = "900px")
print(erd)   # erscheint im Viewer-Fenster von RStudio

dir.create("Output", showWarnings = FALSE)
writeLines(dot, "Output/01_erd.gv", useBytes = TRUE)
htmlwidgets::saveWidget(erd, file.path(getwd(), "Output/01_erd.html"),
                        selfcontained = nzchar(Sys.getenv("RSTUDIO_PANDOC")),
                        title = "ANK – ER-Diagramm")
write.csv2(abgeleitet[, c("fk_tabelle", "fk_spalte", "pk_tabelle", "pk_spalte")],
           "Output/01_beziehungen_abgeleitet.csv", row.names = FALSE)

# 5) Kurzer Überblick in der Konsole ------------------------------------------
cat("\nTabellen im Diagramm: ", nrow(objekte),
    "\nBeziehungen Access:   ", nrow(aus_access),
    "\nnur abgeleitet:       ", nrow(abgeleitet), "(violett gepunktet)",
    "\nRollen:               ", paste(names(table(objekte$rolle)), table(objekte$rolle),
                                      sep = " = ", collapse = ", "),
    "\nErgebnis:              Output/01_erd.html\n\n")
if (nrow(abgeleitet) > 0) {
  cat("fk_-Spalten ohne Beziehung in Access:\n")
  print(abgeleitet[, c("fk_tabelle", "fk_spalte", "pk_tabelle")], row.names = FALSE)
}
