(AccountCode as text, dte as date, AccountType as text, tblTransactions as table)=>
let
    

    FilteredJnls = Table.SelectRows(tblTransactions, 
        each if AccountType = "Balance Sheet Account" then
            [EndOfMonth] <= dte and [Account Code] = AccountCode
        else if AccountType = "P&L account" then 
            false
        else if AccountType = "Retained Earnings" then
            false
        else
            false),

    CumulativeTotal = List.Sum(FilteredJnls[Amount])
in
    CumulativeTotal