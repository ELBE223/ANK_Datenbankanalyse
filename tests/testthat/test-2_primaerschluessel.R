# -----------------------------------------------------------------------------
# Primärschlüssel: in jeder Tabelle eindeutig und nie leer.
# Eine Prüfung je Tabelle (aus der Strukturanalyse, Schritt 0).
# -----------------------------------------------------------------------------

PK <- split(PRUEF_UMFANG$pk$spalte, PRUEF_UMFANG$pk$objekt)

for (tabelle in names(PK)) {
  test_that(paste0(tabelle, ".", paste(PK[[tabelle]], collapse = "+")), {
    d <- PRUEF_DATEN[[tabelle]][PK[[tabelle]]]
    if (is.null(PRUEF_DATEN[[tabelle]])) skip("Tabelle nicht lesbar")
    schluessel <- do.call(paste, c(d, sep = "+"))
    expect_keine(which(!stats::complete.cases(d)), "leere Werte (Zeilennummer)")
    expect_keine(schluessel[duplicated(d)], "doppelte Werte")
  })
}
