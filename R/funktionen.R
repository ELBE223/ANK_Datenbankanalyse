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


# =============================================================================
# ER-Diagramm (01_erd.R)
# =============================================================================

# Beziehungen aus Spaltennamen ableiten ---------------------------------------
# Namenskonvention der ANK-DB: fk_<tabelle>_id verweist auf <tabelle>.id
# (Varianten: _objid statt _id, Tabellenname ohne "_tbd"). Findet Spalten, die
# wie Fremdschlüssel heißen, für die in Access aber keine Beziehung definiert
# ist. Nicht auflösbare Namen (z. B. fk_s3) werden ausgelassen.
beziehungen_ableiten <- function(spalten, beziehungen, tabellen) {
  fk <- spalten[startsWith(spalten$spalte, "fk_") & spalten$objekt %in% tabellen, ]
  fk <- fk[!paste(fk$objekt, fk$spalte) %in% paste(beziehungen$fk_tabelle, beziehungen$fk_spalte), ]
  name <- sub("_(id|objid)(_tmp)?$", "", sub("^fk_", "", fk$spalte))
  ziel <- ifelse(name %in% tabellen, name,
          ifelse(paste0(name, "_tbd") %in% tabellen, paste0(name, "_tbd"), NA))
  ok <- !is.na(ziel)
  pk <- spalten[grepl("PK", spalten$schluessel), ]
  data.frame(
    beziehung = rep("abgeleitet", sum(ok)), pk_tabelle = ziel[ok],
    pk_spalte = pk$spalte[match(ziel[ok], pk$objekt)],
    fk_tabelle = fk$objekt[ok], fk_spalte = fk$spalte[ok], art = rep("1:n", sum(ok)),
    ref_integritaet = rep(FALSE, sum(ok)), kaskade_aendern = rep(FALSE, sum(ok)),
    kaskade_loeschen = rep(FALSE, sum(ok)), row.names = NULL
  )
}


# Rolle und Bereich je Tabelle --------------------------------------------------
# Verknüpfung (n:m): höchstens 4 Spalten, mindestens 2 Fremdschlüssel und
#                    wird selbst von keiner Tabelle referenziert
# Werteliste:        Name endet auf "_tbd" (und ist keine Verknüpfung)
# Bereich:           erstes passende Muster aus 'bereiche', sonst "Sonstiges"
tabellen_rolle <- function(objekte, beziehungen) {
  zaehlen <- function(x) as.vector(table(factor(x, levels = objekte$objekt)))
  n_fk  <- zaehlen(beziehungen$fk_tabelle)
  n_ref <- zaehlen(beziehungen$pk_tabelle[beziehungen$pk_tabelle != beziehungen$fk_tabelle])
  ifelse(objekte$spalten <= 4 & n_fk >= 2 & n_ref == 0, "Verknuepfung",
  ifelse(endsWith(objekte$objekt, "_tbd"), "Werteliste", "Tabelle"))
}

tabellen_bereich <- function(objekte, bereiche) {
  b <- rep("Sonstiges", nrow(objekte))
  for (i in rev(seq_along(bereiche))) b[grepl(bereiche[[i]], objekte$objekt)] <- names(bereiche)[i]
  b
}


# DOT-Text für Graphviz/DiagrammeR erzeugen -----------------------------------
# objekte braucht die Spalten objekt, datensaetze, rolle, bereich;
# beziehungen zusätzlich die Spalte quelle ("Access" oder "abgeleitet").
# Lesart: Strich-Ende = "1"-Seite, Krähenfuß = "n"-Seite.
erd_dot <- function(objekte, beziehungen, spalten, farben) {
  q <- function(x) paste0('"', gsub('"', '\\\\"', x), '"')

  hover <- vapply(objekte$objekt, function(t) {
    s <- spalten[spalten$objekt == t, ]
    paste(paste0(s$spalte, ifelse(s$schluessel == "", "", paste0("  [", s$schluessel, "]"))),
          collapse = "\\n")
  }, character(1))
  wl    <- objekte$rolle == "Werteliste"
  farbe <- ifelse(wl, "#8c8c8c", farben[objekte$bereich])
  fill  <- ifelse(wl, "#f0f0ee", paste0(farbe, "38"))     # helle Tönung der Bereichsfarbe
  label <- paste0(objekte$objekt, ifelse(objekte$rolle == "Verknuepfung", "  (n:m)", ""),
                  "\\n", format(objekte$datensaetze, big.mark = ".", decimal.mark = ",", trim = TRUE),
                  " Datensätze")

  knoten <- sprintf("%s [label=%s, color=%s, fillcolor=%s, style=%s, penwidth=%s, tooltip=%s]",
    q(objekte$objekt), q(label), q(farbe), q(fill),
    q(paste0(ifelse(objekte$rolle == "Verknuepfung", "rounded,", ""), "filled",
             ifelse(objekte$datensaetze %in% 0, ",dashed", ""))),
    ifelse(wl, "1", "1.6"), q(hover))

  abg <- beziehungen$quelle == "abgeleitet"
  kanten <- sprintf("%s -> %s [style=%s, color=%s, tooltip=%s]",
    q(beziehungen$pk_tabelle), q(beziehungen$fk_tabelle),
    ifelse(abg, "dotted", ifelse(beziehungen$ref_integritaet, "solid", "dashed")),
    ifelse(abg, q("#6250d6"), q("#6b6b6b")),
    q(paste0(beziehungen$fk_tabelle, ".", beziehungen$fk_spalte, "  ->  ",
             beziehungen$pk_tabelle, ".", beziehungen$pk_spalte,
             ifelse(abg, "  (nicht in Access definiert)", ""))))

  legende <- c(
    "subgraph cluster_legende {", 'label="Legende"; fontsize=12; color="#bbbbbb"; style=rounded;',
    sprintf('"L_%s" [label=%s, color=%s, fillcolor=%s, style=filled, penwidth=1.6];',
            names(farben), q(names(farben)), q(farben), q(paste0(farben, "38"))),
    '"L_Werteliste" [label="Werteliste (_tbd)", color="#8c8c8c", fillcolor="#f0f0ee", style=filled];',
    sprintf('"L_Text" [shape=plaintext, style="", fontsize=10, label=%s];', q(paste(
      "Strich = 1-Seite, Krähenfuß = n-Seite",
      "durchgezogen = Beziehung mit ref. Integrität",
      "gestrichelt = Beziehung ohne ref. Integrität",
      "gepunktet (violett) = nur aus Spaltenname abgeleitet",
      "abgerundet = Verknüpfungstabelle (n:m)",
      "gestrichelter Rahmen = Tabelle ohne Datensätze",
      "Maus über Tabelle/Linie = Spalten bzw. Schlüssel\\l", sep = "\\l"))),
    "}")

  paste(c(
    "digraph ERD {",
    'graph [rankdir=LR, ranksep=1.1, nodesep=0.18, fontname="Helvetica"];',
    'node  [shape=box, fontname="Helvetica", fontsize=11, fontcolor="#1f1f1f", margin="0.15,0.06"];',
    'edge  [dir=both, arrowtail=tee, arrowhead=crow, arrowsize=0.7, penwidth=1.1];',
    knoten, kanten, legende, "}"), collapse = "\n")
}


# =============================================================================
# Prüfungen (02_pruefung.R und tests/testthat/)
# =============================================================================

# Prüfumfang festlegen ----------------------------------------------------------
# Tabellen (ohne temporäre), Primärschlüssel und Beziehungen: aus Access plus
# aus fk_-Spaltennamen abgeleitete (ohne *_hist-Tabellen, das sind Kopien).
pruefumfang <- function(struktur) {
  o  <- struktur$Objekte
  sp <- struktur$Spalten
  tabellen <- o$objekt[o$typ == "Tabelle"]
  access <- struktur$Beziehungen
  abg    <- beziehungen_ableiten(sp, access, tabellen[!endsWith(tabellen, "_hist")])
  list(
    tabellen    = tabellen,
    temporaer   = o$objekt[o$typ == "Temporaer"],
    spalten     = sp[sp$objekt %in% tabellen, ],
    pk          = sp[sp$objekt %in% tabellen & grepl("PK", sp$schluessel), c("objekt", "spalte")],
    beziehungen = rbind(cbind(access, quelle = rep("Access", nrow(access))),
                        cbind(abg, quelle = rep("abgeleitet", nrow(abg))))
  )
}


# Schlüsselspalten aus der Datenbank lesen --------------------------------------
# Je Tabelle nur Primärschlüssel, Beziehungs- und fk_-Spalten (schnell, klein).
# Ergebnis: Liste mit einem data.frame je Tabelle (NULL, wenn nicht lesbar).
schluessel_lesen <- function(con, umfang) {
  sp <- umfang$spalten
  b  <- umfang$beziehungen
  noetig <- unique(rbind(
    umfang$pk,
    data.frame(objekt = b$pk_tabelle, spalte = b$pk_spalte),
    data.frame(objekt = b$fk_tabelle, spalte = b$fk_spalte),
    sp[startsWith(sp$spalte, "fk_"), c("objekt", "spalte")]
  ))
  stats::setNames(lapply(umfang$tabellen, function(t) {
    s <- noetig$spalte[noetig$objekt == t]
    if (length(s) == 0) return(NULL)
    sql <- paste0("SELECT ", paste0("[", s, "]", collapse = ", "), " FROM [", t, "]")
    tryCatch(DBI::dbGetQuery(con, sql), error = function(e) NULL)
  }), umfang$tabellen)
}


# Prüfbericht aus dem testthat-Ergebnis -----------------------------------------
# Eine Zeile je Prüfung: Bereich (Testdatei), Prüfung, Status, Meldung.
pruef_bericht <- function(ergebnis) {
  zeilen <- lapply(ergebnis, function(t) {
    fehler <- Filter(function(e) inherits(e, c("expectation_failure", "expectation_error")),
                     t$results)
    skip <- any(vapply(t$results, inherits, logical(1), "expectation_skip"))
    data.frame(
      bereich  = sub("^test-\\d+_(.*)\\.R$", "\\1", basename(t$file)),
      pruefung = t$test,
      status   = if (length(fehler)) "FEHLER" else if (skip) "uebersprungen" else "OK",
      meldung  = gsub("\\s*\n\\s*", " ",
                      paste(vapply(fehler, conditionMessage, character(1)), collapse = " | "))
    )
  })
  do.call(rbind, zeilen)
}
