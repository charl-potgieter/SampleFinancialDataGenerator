let

    BufferedJournals = Table.Buffer(Journals),

    fn_TB_RollUp = 
    (AccountCode as text, dte as date, AccountType as text)=>
    let
        //Uncomment for debuging purposes
        //account = "1210",
        //dte = #date(2018,3,31),

        FilteredJnls = Table.SelectRows(BufferedJournals, 
            each if AccountType = "Balance Sheet Account" then
                [EndOfMonth] <= dte and [Account Code] = AccountCode
            else
                BufferedJournals),
        CumulativeTotal = List.Sum(FilteredJnls[Jnl Amount])
    in
        CumulativeTotal,


    //Generate a table of month ends
    MonthEndList = fn_std_DatesBetween(param_DateStart, param_DateEnd, "Month"),
    ConvertToTable = Table.FromList(MonthEndList, Splitter.SplitByNothing(), {"EndOfMonth"}),
    ChangedType = Table.TransformColumnTypes(ConvertToTable,{{"EndOfMonth", type date}}),
    AddChartOfAccountsTableCol = Table.AddColumn(ChangedType, "ChartOfAccountsTable", each DataAccess_ChartOfAccounts, type table),
    ExpandedTable = Table.ExpandTableColumn(AddChartOfAccountsTableCol, "ChartOfAccountsTable", {"Account Code and Description"}, {"Account Code and Description"}),
    
    //Add account code 
    AddAccountCode = Table.AddColumn(ExpandedTable, "Account code", each Text.Start([Account Code and Description], param_NumberOfAccountDigits), type text),
    RemovedColumn = Table.RemoveColumns(AddAccountCode,{"Account Code and Description"}),
    
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
        each fn_TB_RollUp([Account code], [EndOfMonth], [Account type]), 
        type number)
            
in
    AddAmountCol