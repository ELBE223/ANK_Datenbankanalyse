# =============================================================================
# R/funktionen.R – Hilfsfunktionen für die Analyse der ANK-Access-Datenbank
#
# Alle Funktionen lesen nur (SELECT). An der Datenbank wird nichts verändert.
# Benötigte Pakete: DBI, odbc
# =============================================================================

ACCESS_TREIBER <- "Microsoft Access Driver (*.mdb, *.accdb)"


# Verbindung öffnen -----------------------------------------------------------
# Prüft vorher, ob Datei und ODBC-Treiber vorhanden sind.
db_verbinden <- function(pfad) {
  pfad <- normalizePath(pfad, winslash = "\\", mustWork = TRUE)
  if (!ACCESS_TREIBER %in% odbc::odbcListDrivers()$name) {
    stop("ODBC-Treiber '", ACCESS_TREIBER, "' nicht gefunden. ",
         "Benötigt wird die 64-Bit 'Microsoft Access Database Engine' ",
         "(siehe README).", call. = FALSE)
  }
  DBI::dbConnect(
    odbc::odbc(),
    .connection_string = paste0("Driver={", ACCESS_TREIBER, "};DBQ=", pfad, ";")
  )
}


# Objekte auflisten -----------------------------------------------------------
# Access meldet über ODBC:  TABLE   = Tabelle
#                           VIEW    = (Auswahl-)Abfrage
#                           SYNONYM = verknüpfte Tabelle
# Systemtabellen ("MSys...") haben einen eigenen Typ und fallen dadurch weg.
db_objekte <- function(con) {
  typen <- c(TABLE = "Tabelle", VIEW = "Abfrage", SYNONYM = "Verknuepfte Tabelle")
  liste <- lapply(names(typen), function(t) {
    namen <- DBI::dbListTables(con, table_type = t)
    if (length(namen) > 0) data.frame(objekt = namen, typ = typen[[t]])
  })
  do.call(rbind, liste)
}


# Anzahl Datensätze -----------------------------------------------------------
# NA, wenn das Objekt nicht lesbar ist (z. B. Parameterabfrage, defekte
# Verknüpfung zu einer anderen Datei).
db_anzahl <- function(objekt, con) {
  sql <- paste0("SELECT COUNT(*) AS n FROM [", objekt, "]")
  tryCatch(as.numeric(DBI::dbGetQuery(con, sql)$n), error = function(e) NA_real_)
}


# Spalten und Datentypen ------------------------------------------------------
# Spaltennamen/-reihenfolge kommen aus einer leeren Abfrage (exakt),
# der Access-Datentyp (z. B. COUNTER = AutoWert, VARCHAR, DATETIME) aus ODBC.
db_spalten <- function(objekt, con) {
  leer <- tryCatch(
    DBI::dbGetQuery(con, paste0("SELECT * FROM [", objekt, "] WHERE 1 = 0")),
    error = function(e) NULL
  )
  if (is.null(leer)) return(NULL)

  # "_" ist in ODBC-Suchmustern ein Platzhalter -> maskieren
  typen <- tryCatch(
    odbc::odbcListColumns(con, table = gsub("_", "\\_", objekt, fixed = TRUE)),
    error = function(e) data.frame(name = character(), type = character())
  )

  data.frame(
    objekt     = objekt,
    position   = seq_along(leer),
    spalte     = names(leer),
    access_typ = typen$type[match(names(leer), typen$name)],
    r_klasse   = vapply(leer, function(x) class(x)[1], character(1)),
    row.names  = NULL
  )
}


# Beziehungen -----------------------------------------------------------------
# Liest die in Access definierten Beziehungen aus der Systemtabelle
# MSysRelationships. Dafür ist einmalig eine Leseberechtigung nötig (README).
#   pk_* = "1"-Seite (Primärschlüssel), fk_* = "n"-Seite (Fremdschlüssel)
db_beziehungen <- function(con) {
  leer <- data.frame(
    beziehung = character(), pk_tabelle = character(), pk_spalte = character(),
    fk_tabelle = character(), fk_spalte = character(), art = character(),
    ref_integritaet = logical(), kaskade_aendern = logical(),
    kaskade_loeschen = logical()
  )
  sql <- "SELECT szRelationship, szReferencedObject, szReferencedColumn,
                 szObject, szColumn, grbit
          FROM MSysRelationships"
  b <- tryCatch(DBI::dbGetQuery(con, sql), error = function(e) {
    warning("MSysRelationships ist nicht lesbar - Beziehungen fehlen. ",
            "Einmalige Freigabe in Access siehe README.", call. = FALSE)
    NULL
  })
  if (is.null(b)) return(leer)

  b <- b[!grepl("^MSys", b$szObject) & !grepl("^MSys", b$szReferencedObject), ]
  if (nrow(b) == 0) return(leer)

  # grbit = Bit-Schalter der Beziehung (DAO RelationAttributeEnum)
  bit <- function(wert) bitwAnd(as.integer(b$grbit), wert) > 0
  data.frame(
    beziehung        = b$szRelationship,
    pk_tabelle       = b$szReferencedObject,
    pk_spalte        = b$szReferencedColumn,
    fk_tabelle       = b$szObject,
    fk_spalte        = b$szColumn,
    art              = ifelse(bit(1L), "1:1", "1:n"),
    ref_integritaet  = !bit(2L),
    kaskade_aendern  = bit(256L),
    kaskade_loeschen = bit(4096L),
    row.names        = NULL
  )
}


# Schlüssel-Hinweise je Spalte ------------------------------------------------
# Markiert AutoWert-Spalten sowie Spalten, die in Beziehungen als
# Primärschlüssel (PK) oder Fremdschlüssel (FK) vorkommen.
schluessel_hinweis <- function(spalten, beziehungen) {
  id <- paste(spalten$objekt, spalten$spalte, sep = "|")
  h  <- cbind(
    ifelse(spalten$access_typ %in% "COUNTER", "AutoWert", NA),
    ifelse(id %in% paste(beziehungen$pk_tabelle, beziehungen$pk_spalte, sep = "|"), "PK", NA),
    ifelse(id %in% paste(beziehungen$fk_tabelle, beziehungen$fk_spalte, sep = "|"), "FK", NA)
  )
  apply(h, 1, function(x) paste(stats::na.omit(x), collapse = ", "))
}
