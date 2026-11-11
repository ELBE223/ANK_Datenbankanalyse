# -----------------------------------------------------------------------------
# Hilfsfunktionen für die Prüfungen (lädt testthat automatisch vor den Tests).
# Die Daten PRUEF_UMFANG und PRUEF_DATEN stellt 02_pruefung.R bereit –
# die Tests daher immer über dieses Skript starten (Source).
# -----------------------------------------------------------------------------

if (!exists("PRUEF_DATEN") || !exists("PRUEF_UMFANG")) {
  stop("Die Prüfungen bitte über 02_pruefung.R starten (Source).", call. = FALSE)
}

# Erwartung: x ist leer. Sonst schlägt die Prüfung fehl und nennt Anzahl und
# Beispiele, z. B. "3 doppelte Werte, z. B. 17, 18, 21".
expect_keine <- function(x, was, beispiele = 5) {
  testthat::expect(
    length(x) == 0,
    sprintf("%d %s, z. B. %s", length(x), was,
            paste(utils::head(unique(x), beispiele), collapse = ", "))
  )
  invisible(x)
}

# Werte einer Spalte (NULL, wenn Tabelle oder Spalte nicht gelesen wurde)
werte <- function(tabelle, spalte) PRUEF_DATEN[[tabelle]][[spalte]]
