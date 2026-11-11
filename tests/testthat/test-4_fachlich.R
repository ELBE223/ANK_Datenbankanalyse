# -----------------------------------------------------------------------------
# Fachliche Regeln der ANK-Datenbank – hier ergänzen und anpassen.
# Die Regeln sind Annahmen aus der Strukturanalyse und müssen fachlich
# bestätigt werden.
# -----------------------------------------------------------------------------

# 1:1-Beziehungen ---------------------------------------------------------------
# Jeder Datensatz links hat genau einen Partner rechts, und jeder Partner
# gehört zu genau einem Datensatz (z. B. jedes Projekt hat genau eine Finanzierung).
EINS_ZU_EINS <- data.frame(
  tabelle = c("projekt",          "projekt",              "geoinformation",                  "foerdergegenstand"),
  spalte  = c("fk_finanzen_id",   "fk_geoinformation_id", "fk_geoinformation_basisdaten_id", "fk_foerdergegenstand_finanzen_id"),
  ziel    = c("projekt_finanzen", "geoinformation",       "geoinformation_basisdaten",       "foerdergegenstand_finanzen")
)

for (i in seq_len(nrow(EINS_ZU_EINS))) {
  r <- EINS_ZU_EINS[i, ]
  test_that(sprintf("1:1 %s.%s <-> %s", r$tabelle, r$spalte, r$ziel), {
    fk   <- werte(r$tabelle, r$spalte)
    id   <- werte(r$tabelle, "id")
    ziel <- werte(r$ziel, "id")
    if (is.null(fk) || is.null(ziel)) skip("Spalte nicht lesbar")
    expect_keine(id[is.na(fk)], paste0(r$tabelle, " ohne ", r$ziel, " (id)"))
    expect_keine(fk[!is.na(fk) & duplicated(fk)], "mehrfach verwendete Verweise")
    expect_keine(setdiff(ziel, fk), paste0(r$ziel, " ohne Zuordnung (id)"))
  })
}

# Temporäre Spalten ---------------------------------------------------------------
# Spalten mit "_tmp" im Namen sollten leer sein (Reste eines Umbaus).
test_that("projekt.fk_projekt_massnahme_tbd_id_tmp ist leer", {
  x <- werte("projekt", "fk_projekt_massnahme_tbd_id_tmp")
  if (is.null(x)) skip("Spalte nicht vorhanden")
  expect_keine(werte("projekt", "id")[!is.na(x)], "Projekte mit Wert (id)")
})
