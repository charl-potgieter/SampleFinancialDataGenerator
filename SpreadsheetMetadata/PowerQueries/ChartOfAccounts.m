let
    ChartOfAccountsRaw = DataAccess_ChartOfAccounts,
    AddAccountCode = Table.AddColumn(ChartOfAccountsRaw, "Account Code", each Text.Start([Account Code and Description], param_NumberOfAccountDigits), type text),
    AddAccountDescription = Table.AddColumn(AddAccountCode, "Account Description", each Text.End(
            [Account Code and Description], 
            Text.Length([Account Code and Description]) -param_NumberOfAccountDigits -1)
        , type text),
    Reorder = Table.ReorderColumns(AddAccountDescription,{"Account Code", "Account Description", "Account Code and Description", "Account Category 1", "Account Category 2", "Account Category 3", "Account Category 4", "Account Category 5", "Sort Order Account Category 1", "Sort Order Account Category 2", "Sort Order Account Category 3", "Sort Order Account Category 4", "Sort Order Account Category 5"})
in
    Reorder