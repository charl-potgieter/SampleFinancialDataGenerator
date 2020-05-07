let

    /*********************************************************************************************************************
        Create a buffered table of opening balances and journals
    *********************************************************************************************************************/

    JournalsSelectedCols = Table.SelectColumns(Journals, {"EndOfMonth", "Account Code", "Jnl Amount"}),
    JournalsRenamedCol = Table.RenameColumns(JournalsSelectedCols,{{"Jnl Amount", "Amount"}}),

    OpeningBalancesRaw = DataAccess_OpeningBalalances,
    OpeningBalanceChangedType = Table.TransformColumnTypes(OpeningBalancesRaw,{{"Account Code and Description", type text}, {"Amount", type number}}),
    OpeningBalanceAddAccountCodeCol = Table.AddColumn(OpeningBalanceChangedType, "Account Code", each Text.Start([Account Code and Description], param_NumberOfAccountDigits), type text),
    OpeningBalanceAddDateCol = Table.AddColumn(OpeningBalanceAddAccountCodeCol, "Date", each param_DateStart, type date),
    OpeningBalanceSelectCols = Table.SelectColumns(OpeningBalanceAddDateCol, {"Date", "Account Code", "Amount"}),
    OpeningBalanceRenameCols = Table.RenameColumns(OpeningBalanceSelectCols,{{"Date", "EndOfMonth"}}),
    CombineJournalAndOpeningBalance = Table.Combine({OpeningBalanceRenameCols, JournalsRenamedCol}),
    GroupedRows = Table.Group(CombineJournalAndOpeningBalance, {"EndOfMonth", "Account Code"}, {{"Amount", each List.Sum([Amount]), type number}}),
    BufferedTransactions = Table.Buffer(GroupedRows),

    /*********************************************************************************************************************
        Generate a table of month ends with all accounts for each month
    *********************************************************************************************************************/
    
    MonthEndList = fn_std_DatesBetween(param_DateStart, param_DateEnd, "Month"),
    ConvertToTable = Table.FromList(MonthEndList, Splitter.SplitByNothing(), {"EndOfMonth"}),
    ChangedType2 = Table.TransformColumnTypes(ConvertToTable,{{"EndOfMonth", type date}}),
    AddChartOfAccountsTableCol = Table.AddColumn(ChangedType2, "ChartOfAccountsTable", each DataAccess_ChartOfAccounts, type table),
    ExpandedTable = Table.ExpandTableColumn(AddChartOfAccountsTableCol, "ChartOfAccountsTable", {"Account Code and Description"}, {"Account Code and Description"}),
    AddAccountCode = Table.AddColumn(ExpandedTable, "Account code", each Text.Start([Account Code and Description], param_NumberOfAccountDigits), type text),
    RemovedColumn = Table.RemoveColumns(AddAccountCode,{"Account Code and Description"}),
    

    /*********************************************************************************************************************
        Caclculate tb balance for each date and account
    *********************************************************************************************************************/

    //Add account type
    AddAccountType = Table.AddColumn(RemovedColumn, "Account type", each 
        if Number.From([Account code]) = Number.From(param_RetainedEarningsAccountCode) then
            "Retained Earnings"
        else if Number.From([Account code]) >= Number.From(param_StartOfPandLAccountCode) then
            "P&L account"
        else
            "Balance Sheet Account", 
        type text),


    //Add start of year column
    AddStartOfYearCol =Table.AddColumn(AddAccountType, "StartOfYear", each Date.StartOfYear([EndOfMonth]), type date),
    
    //Add amount column
    AddAmountCol = Table.AddColumn(AddStartOfYearCol, "Amount",
        each fn_TbRollUp([Account code], [EndOfMonth], [Account type], BufferedTransactions), 
        type number)
            
in
    AddAmountCol