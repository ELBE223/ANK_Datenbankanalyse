# -----------------------------------------------------------------------------
# dao_metadaten.ps1 – Beziehungen und Primärschlüssel über DAO auslesen
#
# Wird von R aus aufgerufen (db_dao() in R/funktionen.R), nicht direkt starten.
# DAO ist die Access-eigene Schnittstelle: kein Makro, keine Freigabe nötig.
# Die Datenbank wird schreibgeschützt geöffnet.
# Ausgabe: eine Zeile je Eintrag, Felder durch Tabulator getrennt
#   REL  Name  PK-Tabelle  PK-Spalte  FK-Tabelle  FK-Spalte  Attribute
#   PK   Tabelle  Spalte
#   ERR  Fehlermeldung
# Den Platzhalter in OpenDatabase(...) ersetzt R durch den Pfad zur Datenbank.
# -----------------------------------------------------------------------------
try {
  [Console]::OutputEncoding = [Text.Encoding]::UTF8
  $db = (New-Object -ComObject DAO.DBEngine.120).OpenDatabase('{{PFAD}}', $false, $true)

  foreach ($r in $db.Relations) {
    foreach ($f in $r.Fields) {
      'REL', $r.Name, $r.Table, $f.Name, $r.ForeignTable, $f.ForeignName, $r.Attributes -join "`t"
    }
  }

  foreach ($t in $db.TableDefs) {
    if ($t.Name -like 'MSys*' -or $t.Name -like '~*' -or $t.Connect) { continue }
    foreach ($i in $t.Indexes) {
      if ($i.Primary) {
        foreach ($f in $i.Fields) { 'PK', $t.Name, $f.Name -join "`t" }
      }
    }
  }

  $db.Close()
} catch {
  'ERR', $_.Exception.Message -join "`t"
  exit 1
}
