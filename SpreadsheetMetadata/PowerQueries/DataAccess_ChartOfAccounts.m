let
    Source = Excel.CurrentWorkbook(){[Name="tbl_InputChartOfAccounts"]}[Content],
    ChangedType = Table.TransformColumnTypes(Source,{{"Account Code and Description", type text}, {"Account Category 1", type text}, {"Account Category 2", type text}, {"Account Category 3", type text}, {"Account Category 4", type text}, {"Account Category 5", type text}, {"Sort Order Account Category 1", Int64.Type}, {"Sort Order Account Category 2", Int64.Type}, {"Sort Order Account Category 3", Int64.Type}, {"Sort Order Account Category 4", Int64.Type}, {"Sort Order Account Category 5", Int64.Type}})
in
    ChangedType