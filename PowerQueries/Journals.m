let
    MonthEndList = fn_std_DatesBetween(param_OpeningBalanceDate, param_LastMonthEnd, "Month"),
    ConvertToTable = Table.FromList(MonthEndList, Splitter.SplitByNothing(), {"EndOfMonth"}),
    ChangedType = Table.TransformColumnTypes(ConvertToTable,{{"EndOfMonth", type date}}),

    AddMonthIndex = Table.AddColumn(ChangedType, "Month Index", each 
        (Date.Year([EndOfMonth]) * 12 + Date.Month([EndOfMonth])) - 
        (Date.Year(List.Min(ChangedType[EndOfMonth])) * 12 + Date.Month(List.Min(ChangedType[EndOfMonth]))),
        Int64.Type),

    JnlPrefixList = Table.AddColumn(AddMonthIndex, "Jnl Prefix", each DataAccess_JnlMetaData[Jnl Prefix], type list),
    ExpandedJnlPrefixList = Table.ExpandListColumn(JnlPrefixList, "Jnl Prefix"),
    ChangedType2 = Table.TransformColumnTypes(ExpandedJnlPrefixList,{{"Jnl Prefix", type text}}),
    AddJnlMetaDataTable = Table.NestedJoin(ChangedType2, "Jnl Prefix", DataAccess_JnlMetaData, "Jnl Prefix", "JnlMetaDataTable", JoinKind.LeftOuter),
    AddNumberOfTablesCol = Table.AddColumn(AddJnlMetaDataTable, "NumberOfTablesRequired", each [JnlMetaDataTable][Number of instances per month]{0}, Int64.Type),
    AddListOfTablesCol = Table.AddColumn(AddNumberOfTablesCol, "AddMultiTableCol", each List.Accumulate({1..[NumberOfTablesRequired]}, {}, (state, current)=> state & {[JnlMetaDataTable]}), type list),
    ExpandedListOfTables = Table.ExpandListColumn(AddListOfTablesCol, "AddMultiTableCol"),
    ExpandedTables = Table.ExpandTableColumn(ExpandedListOfTables, "AddMultiTableCol", {"Jnl Prefix", "Debit Account Number", "Debit Account Full", "Credit Account Number", "Credit Account Full", "Jnl Description", "Base Amount Total For Month", "Is recurring", "Monthly growth percentage", "Number of instances per month", "Randomise", "Random Variation around base"}, {"Jnl Prefix.1", "Debit Account Number", "Debit Account Full", "Credit Account Number", "Credit Account Full", "Jnl Description", "Base Amount Total For Month", "Is recurring", "Monthly growth percentage", "Number of instances per month", "Randomise", "Random Variation around base"}),

    AddRandomAdjCol = Table.AddColumn(ExpandedTables, "Random Number Factor", each if [Randomise] then
            1 + Number.RandomBetween(-[Random Variation around base], [Random Variation around base])
        else
            1, 
        type number),
    AddGrowthAdjCol = Table.AddColumn(AddRandomAdjCol, "Growth Factor", each Number.Power(1 + [Monthly growth percentage], [Month Index]), type number),
    AddAmountCol = Table.AddColumn(AddGrowthAdjCol, "Absolute Amount", each [Base Amount Total For Month] / [Number of instances per month] * [Random Number Factor] * [Growth Factor], type number),
    AddIndexCol = Table.AddIndexColumn(AddAmountCol, "Index", 0, 1),
    AddJnlIDCol = Table.AddColumn(AddIndexCol, "Jnl ID", each [Jnl Prefix] & "-" & Text.PadStart(Text.From([Index]), 10, "0"), type text),
    SelectCols = Table.SelectColumns(AddJnlIDCol,{"EndOfMonth", "Jnl ID", "Debit Account Number", "Credit Account Number", "Jnl Description", "Absolute Amount"}),
    UnpivotAccountNumbers = Table.UnpivotOtherColumns(SelectCols, {"EndOfMonth", "Jnl ID", "Jnl Description", "Absolute Amount"}, "Account Number Type", "Account Number"),

    AddJnlAmountCol = Table.AddColumn(UnpivotAccountNumbers, "Jnl Amount", each if [Account Number Type] = "Debit Account Number" then
            [Absolute Amount]
        else
            -[Absolute Amount],
        type number),
    RemovedCols = Table.RemoveColumns(AddJnlAmountCol,{"Account Number Type", "Absolute Amount"}),
    ReorderedCols = Table.ReorderColumns(RemovedCols,{"EndOfMonth", "Jnl ID", "Account Number", "Jnl Description", "Jnl Amount"}),
    ChangedType3 = Table.TransformColumnTypes(ReorderedCols,{{"EndOfMonth", type date}, {"Jnl ID", type text}, {"Account Number", type text}, {"Jnl Description", type text}})
in
    ChangedType3