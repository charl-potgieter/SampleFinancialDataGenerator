let
    Source = Excel.CurrentWorkbook(){[Name="tbl_InputJnlMetaData"]}[Content],
    FilteredOutNulls = Table.SelectRows(Source, each ([Jnl Prefix] <> null))
in
    FilteredOutNulls