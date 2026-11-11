# -----------------------------------------------------------------------------
# Fremdschlüssel: jeder Verweis zeigt auf einen vorhandenen Datensatz.
# Leere Werte sind erlaubt (Verweis nicht ausgefüllt).
# Eine Prüfung je Beziehung: aus Access und aus fk_-Spaltennamen abgeleitet.
# -----------------------------------------------------------------------------

B <- PRUEF_UMFANG$beziehungen

for (i in seq_len(nrow(B))) {
  b <- B[i, ]
  test_that(sprintf("%s.%s -> %s.%s%s", b$fk_tabelle, b$fk_spalte, b$pk_tabelle, b$pk_spalte,
                    if (b$quelle == "abgeleitet") "  (abgeleitet)" else ""), {
    fk <- werte(b$fk_tabelle, b$fk_spalte)
    pk <- werte(b$pk_tabelle, b$pk_spalte)
    if (is.null(fk) || is.null(pk)) skip("Spalte nicht lesbar")
    expect_keine(fk[!is.na(fk) & !fk %in% pk], paste("Verweise ohne Datensatz in", b$pk_tabelle))
  })
}
