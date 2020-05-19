let

    /*********************************************************************************************************************
        Function - return balance sheet portion of TB
    *********************************************************************************************************************/

    fn_BsPortionOfTb = 
    (tbl_Jnls as table, dte_EndOfMonth as date)=>
    let
        FilterTable = Table.SelectRows(tbl_Jnls, each [EndOfMonth] <= dte_EndOfMonth and [Account type] = "Balance Sheet Account"),
        GroupedByAccount = Table.Group(FilterTable, {"Account Code"}, {{"Amount", each List.Sum([Amount]), type number}}),
        AddMonthEnd = Table.AddColumn(GroupedByAccount, "EndOfMonth", each dte_EndOfMonth, type date),
        ReorderCols = Table.ReorderColumns(AddMonthEnd,{"EndOfMonth", "Account Code", "Amount"})
    in
        ReorderCols,


    /*********************************************************************************************************************
        Function - return P&L portion of TB
    *********************************************************************************************************************/

    fn_PandLPortionOfTb = 
    (tbl_Jnls as table, dte_EndOfMonth as date)=>    
    let
        FilterTable = Table.SelectRows(tbl_Jnls, each 
                [EndOfMonth] <= dte_EndOfMonth and 
                Date.Year([EndOfMonth]) =  Date.Year(dte_EndOfMonth) and
                [Account type] = "P&L account"),
                
        GroupedByAccount = Table.Group(FilterTable, {"Account Code"}, {{"Amount", each List.Sum([Amount]), type number}}),
        AddMonthEnd = Table.AddColumn(GroupedByAccount, "EndOfMonth", each dte_EndOfMonth, type date),
        ReorderCols = Table.ReorderColumns(AddMonthEnd,{"EndOfMonth", "Account Code", "Amount"})
    in
        ReorderCols,


    /*********************************************************************************************************************
        Function - return retained earnings TB
    *********************************************************************************************************************/

    fn_RetainedEarningsPortionOfTb = 
    (tbl_Jnls as table, dte_EndOfMonth as date)=>  
    let
        
        FilterTable = Table.SelectRows(tbl_Jnls, each 
                (
                    [EndOfMonth] <= dte_EndOfMonth and 
                    [Account type] = "Retained Earnings"
                )
                or
                (
                    [EndOfMonth] <= #date(Date.Year(dte_EndOfMonth)-1, 12, 31) and 
                    [Account type] = "P&L account"
                )
            ),

        RemovedOldMonthEndCol = Table.RemoveColumns(FilterTable,{"EndOfMonth"}),
        AddNewMonthEndCol = Table.AddColumn(RemovedOldMonthEndCol, "EndOfMonth", each dte_EndOfMonth, type date),
        GroupedByMonthEnd = Table.Group(AddNewMonthEndCol, {"EndOfMonth"}, {{"Amount", each List.Sum([Amount]), type number}}),
        AddAccountCodeCol = Table.AddColumn(GroupedByMonthEnd, "Account Code", each param_RetainedEarningsAccountCode, type text),
        ReorderCols = Table.ReorderColumns(AddAccountCodeCol,{"EndOfMonth", "Account Code", "Amount"})
    in
        ReorderCols,

    /*********************************************************************************************************************
        Create a buffered table of opening balances and journals
    *********************************************************************************************************************/

    //Note - need to pick up from the output tab otherwise the random number generator will change number versus original journals!
    JournalsSelectedCols = Table.SelectColumns(DataAccess_JnlsfromOutputTab, {"EndOfMonth", "Account Code", "Jnl Amount"}),
    JournalsChangedType = Table.TransformColumnTypes(JournalsSelectedCols,{{"EndOfMonth", type date}, {"Account Code", type text}, {"Jnl Amount", type number}}),
    JournalsRenamedCol = Table.RenameColumns(JournalsChangedType,{{"Jnl Amount", "Amount"}}),

    OpeningBalancesRaw = DataAccess_OpeningBalalances,
    OpeningBalanceChangedType = Table.TransformColumnTypes(OpeningBalancesRaw,{{"Account Code and Description", type text}, {"Amount", type number}}),
    OpeningBalanceAddAccountCodeCol = Table.AddColumn(OpeningBalanceChangedType, "Account Code", each Text.Start([Account Code and Description], param_NumberOfAccountDigits), type text),
    OpeningBalanceAddDateCol = Table.AddColumn(OpeningBalanceAddAccountCodeCol, "Date", each param_DateStart, type date),
    OpeningBalanceSelectCols = Table.SelectColumns(OpeningBalanceAddDateCol, {"Date", "Account Code", "Amount"}),
    OpeningBalanceRenameCols = Table.RenameColumns(OpeningBalanceSelectCols,{{"Date", "EndOfMonth"}}),
    CombineJournalAndOpeningBalance = Table.Combine({OpeningBalanceRenameCols, JournalsRenamedCol}),
    GroupedRows = Table.Group(CombineJournalAndOpeningBalance, {"EndOfMonth", "Account Code"}, {{"Amount", each List.Sum([Amount]), type number}}),
    AddAccountType = Table.AddColumn(GroupedRows, "Account type", each 
        if Number.From([Account Code]) = Number.From(param_RetainedEarningsAccountCode) then
            "Retained Earnings"
        else if Number.From([Account Code]) >= Number.From(param_StartOfPandLAccountCode) then
            "P&L account"
        else
            "Balance Sheet Account", 
        type text),

    BufferedTransactions = Table.Buffer(AddAccountType),

    /*********************************************************************************************************************
        Generate a table of month ends and generate tb per month
    *********************************************************************************************************************/
    
    MonthEndList = fn_std_DatesBetween(param_DateStart, param_DateEnd, "Month"),
    ConvertToTable = Table.FromList(MonthEndList, Splitter.SplitByNothing(), {"EndOfMonth"}),
    ChangedType = Table.TransformColumnTypes(ConvertToTable,{{"EndOfMonth", type date}}),
    
    
    AddTbTable = Table.AddColumn(ChangedType, "TB as table", 
            each fn_BsPortionOfTb(BufferedTransactions, [EndOfMonth]) & 
            fn_RetainedEarningsPortionOfTb(BufferedTransactions, [EndOfMonth]) &
            fn_PandLPortionOfTb(BufferedTransactions, [EndOfMonth]), 
            type table),

    ExpandedTBTable = Table.ExpandTableColumn(AddTbTable, "TB as table", {"Account Code", "Amount"}, {"Account Code", "Amount"}),
    ChangedType2 = Table.TransformColumnTypes(ExpandedTBTable,{{"Account Code", type text}, {"Amount", type number}})



in
    ChangedType2