# -----------------------------------------------------------------------------
# Struktur: Regeln für den Aufbau der Datenbank (nur Metadaten, keine Daten).
# Wichtig vor allem für eine saubere Migration nach PostgreSQL.
# -----------------------------------------------------------------------------

U <- PRUEF_UMFANG
B <- U$beziehungen

test_that("Jede Tabelle hat einen Primärschlüssel", {
  expect_keine(setdiff(U$tabellen, U$pk$objekt), "Tabellen ohne Primärschlüssel", 20)
})

test_that("Keine temporären Tabellen (~TMP...)", {
  expect_keine(U$temporaer, "temporäre Tabellen", 20)
})

test_that("Fremdschlüssel heißen fk_<tabelle>_<schlüssel>", {
  b    <- B[B$quelle == "Access" & B$pk_tabelle != B$fk_tabelle, ]
  soll <- paste0("fk_", b$pk_tabelle, "_", b$pk_spalte)
  abw  <- b$fk_spalte != soll
  expect_keine(paste0(b$fk_tabelle[abw], ".", b$fk_spalte[abw], " (erwartet ", soll[abw], ")",
                      recycle0 = TRUE),
               "abweichende Namen", 20)
})

test_that("Jede fk_-Spalte hat eine Beziehung in Access", {
  sp   <- U$spalten[startsWith(U$spalten$spalte, "fk_") & !endsWith(U$spalten$objekt, "_hist"), ]
  a    <- B[B$quelle == "Access", ]
  ohne <- !paste(sp$objekt, sp$spalte) %in% paste(a$fk_tabelle, a$fk_spalte)
  expect_keine(paste0(sp$objekt[ohne], ".", sp$spalte[ohne], recycle0 = TRUE),
               "fk_-Spalten ohne Beziehung", 20)
})

test_that("Fremd- und Primärschlüssel haben denselben Datentyp", {
  typ <- function(t, s) {
    x <- U$spalten$access_typ[match(paste(t, s), paste(U$spalten$objekt, U$spalten$spalte))]
    sub("COUNTER", "INTEGER", x)   # AutoWert ist intern eine ganze Zahl
  }
  abw <- which(typ(B$fk_tabelle, B$fk_spalte) != typ(B$pk_tabelle, B$pk_spalte))
  expect_keine(paste0(B$fk_tabelle[abw], ".", B$fk_spalte[abw], recycle0 = TRUE),
               "abweichende Datentypen", 20)
})
