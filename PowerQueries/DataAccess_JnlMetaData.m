let
    Source = Excel.CurrentWorkbook(){[Name="tbl_JnlMetaData"]}[Content],
    FilteredOutNulls = Table.SelectRows(Source, each ([Jnl Prefix] <> null))
in
    FilteredOutNulls