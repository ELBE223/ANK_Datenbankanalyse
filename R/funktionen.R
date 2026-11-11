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
  if (!file.exists(pfad)) {
    stop("Datenbank nicht gefunden: ", pfad,
         "\nDB_PFAD in config_lokal.R prüfen.", call. = FALSE)
  }
  pfad <- normalizePath(pfad, winslash = "\\")
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
# Tabellen mit "~" am Anfang (z. B. ~TMPCLP...) sind Reste aus Kopieren/
# Einfügen in Access -> als "Temporaer" markiert, gehören nicht zur Logik.
db_objekte <- function(con) {
  typen <- c(TABLE = "Tabelle", VIEW = "Abfrage", SYNONYM = "Verknuepfte Tabelle")
  liste <- lapply(names(typen), function(t) {
    namen <- DBI::dbListTables(con, table_type = t)
    if (length(namen) > 0) data.frame(objekt = namen, typ = typen[[t]])
  })
  obj <- do.call(rbind, liste)
  obj$typ[startsWith(obj$objekt, "~")] <- "Temporaer"
  obj
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


# Beziehungen und Primärschlüssel über DAO ------------------------------------
# DAO ist die Access-eigene Schnittstelle. R startet dafür PowerShell mit dem
# Skript R/dao_metadaten.ps1 – ohne Makros, ohne Freigabe, schreibgeschützt.
# Ergebnis: list(beziehungen, pk) oder NULL, wenn DAO nicht verfügbar ist.
db_dao <- function(pfad, skript = "R/dao_metadaten.ps1") {
  ps <- paste(readLines(skript, encoding = "UTF-8"), collapse = "\n")
  ps <- gsub("{{PFAD}}", gsub("'", "''", normalizePath(pfad)), ps, fixed = TRUE)
  z  <- powershell_ausfuehren(ps)

  if (!is.null(attr(z, "status")) || any(startsWith(z, "ERR\t"))) {
    message("DAO nicht verfügbar: ", paste(sub("^ERR\t", "", head(z, 3)), collapse = " "))
    return(NULL)
  }
  felder <- function(typ, namen) {
    x <- strsplit(z[startsWith(z, paste0(typ, "\t"))], "\t", fixed = TRUE)
    m <- matrix(c(character(), unlist(lapply(x, `[`, -1))), ncol = length(namen), byrow = TRUE)
    stats::setNames(as.data.frame(m), namen)
  }
  message("Beziehungen und Primärschlüssel über DAO gelesen.")
  list(
    beziehungen = felder("REL", c("szRelationship", "szReferencedObject", "szReferencedColumn",
                                  "szObject", "szColumn", "grbit")),
    pk          = felder("PK", c("tabelle", "spalte"))
  )
}

# PowerShell-Befehl ausführen und Ausgabezeilen zurückgeben. -EncodedCommand
# umgeht Probleme mit Anführungszeichen und der Skript-Ausführungsrichtlinie.
powershell_ausfuehren <- function(ps, timeout = 120) {
  if (.Platform$OS.type != "windows") return(structure("nur unter Windows", status = 1))
  enc <- gsub("\n", "", jsonlite::base64_enc(iconv(ps, "UTF-8", "UTF-16LE", toRaw = TRUE)[[1]]))
  tryCatch(
    suppressWarnings(system2("powershell",
      c("-NoProfile", "-NonInteractive", "-EncodedCommand", enc),
      stdout = TRUE, stderr = TRUE, timeout = timeout)),
    error = function(e) structure(conditionMessage(e), status = 1)
  )
}


# Beziehungen -----------------------------------------------------------------
# Die in Access definierten Beziehungen. Quellen, in dieser Reihenfolge:
#   1. DAO (siehe db_dao) – der Normalfall
#   2. Systemtabelle MSysRelationships (braucht einmalige Leseberechtigung)
#   3. Excel-Export dieser Systemtabelle aus Access (README, Punkt 5)
#   pk_* = "1"-Seite (Primärschlüssel), fk_* = "n"-Seite (Fremdschlüssel)
db_beziehungen <- function(con, dao = NULL, export = NULL) {
  leer <- data.frame(
    beziehung = character(), pk_tabelle = character(), pk_spalte = character(),
    fk_tabelle = character(), fk_spalte = character(), art = character(),
    ref_integritaet = logical(), kaskade_aendern = logical(),
    kaskade_loeschen = logical()
  )
  sql <- "SELECT szRelationship, szReferencedObject, szReferencedColumn,
                 szObject, szColumn, grbit
          FROM MSysRelationships"
  b <- dao$beziehungen
  if (is.null(b)) b <- tryCatch(DBI::dbGetQuery(con, sql), error = function(e) NULL)
  if (is.null(b) && !is.null(export) && file.exists(export)) {
    message("Beziehungen gelesen aus Excel-Export: ", export)
    b <- as.data.frame(readxl::read_excel(export))
  }
  if (is.null(b)) {
    warning("Beziehungen fehlen: weder DAO noch MSysRelationships lesbar und ",
            "kein Excel-Export gefunden. Siehe README, Einrichtung Punkt 5.",
            call. = FALSE)
    return(leer)
  }

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
# PK = Primärschlüssel: aus DAO (exakt) oder, falls nicht verfügbar, die Spalten,
#      auf die eine Beziehung verweist (Näherung).
# FK = Fremdschlüssel laut Beziehung, AutoWert = Access-Datentyp COUNTER.
schluessel_hinweis <- function(spalten, beziehungen, pk = NULL) {
  id    <- paste(spalten$objekt, spalten$spalte, sep = "|")
  pk_id <- if (!is.null(pk)) paste(pk$tabelle, pk$spalte, sep = "|") else
    paste(beziehungen$pk_tabelle, beziehungen$pk_spalte, sep = "|")
  h  <- cbind(
    ifelse(spalten$access_typ %in% "COUNTER", "AutoWert", NA),
    ifelse(id %in% pk_id, "PK", NA),
    ifelse(id %in% paste(beziehungen$fk_tabelle, beziehungen$fk_spalte, sep = "|"), "FK", NA)
  )
  apply(h, 1, function(x) paste(stats::na.omit(x), collapse = ", "))
}
