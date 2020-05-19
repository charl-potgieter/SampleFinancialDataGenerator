let


    /*************************************************************************************************************************
            Get add hoc journals
    *************************************************************************************************************************/

    AdHocJnlsRaw = DataAccess_AdHocJnls,
    ChangedType0 = Table.TransformColumnTypes(AdHocJnlsRaw,{{"EndOfMonth", type date}, {"Jnl ID", type text}, {"Account Code and Description", type text}, {"Jnl Description", type text}, {"Jnl Amount", type number}}),
    AdHocJnls = Table.SelectRows(ChangedType0, each 
        [EndOfMonth] <> null and
        [EndOfMonth] >= param_DateStart and 
        [EndOfMonth] <= param_DateEnd),



    /*************************************************************************************************************************
            Get list of recurring journals
    *************************************************************************************************************************/

    //Get generate list of month ends and index
    MonthEndList = fn_std_DatesBetween(param_DateStart, param_DateEnd, "Month"),
    ConvertToTable = Table.FromList(MonthEndList, Splitter.SplitByNothing(), {"EndOfMonth"}),
    ChangedType = Table.TransformColumnTypes(ConvertToTable,{{"EndOfMonth", type date}}),
    AddMonthIndex = Table.AddColumn(ChangedType, "Month Index", each 
        (Date.Year([EndOfMonth]) * 12 + Date.Month([EndOfMonth])) - 
        (Date.Year(List.Min(ChangedType[EndOfMonth])) * 12 + Date.Month(List.Min(ChangedType[EndOfMonth]))),
        Int64.Type),

    //For each month end add the list of recurring journal id prefixes
    JnlPrefixList = Table.AddColumn(AddMonthIndex, "Jnl Prefix", each DataAccess_JnlMetaData[Jnl Prefix], type list),
    ExpandedJnlPrefixList = Table.ExpandListColumn(JnlPrefixList, "Jnl Prefix"),
    ChangedType2 = Table.TransformColumnTypes(ExpandedJnlPrefixList,{{"Jnl Prefix", type text}}),

    // Add the required number of journal tables and expand
    AddJnlMetaDataTable = Table.NestedJoin(ChangedType2, "Jnl Prefix", DataAccess_JnlMetaData, "Jnl Prefix", "JnlMetaDataTable", JoinKind.LeftOuter),
    AddNumberOfTablesCol = Table.AddColumn(AddJnlMetaDataTable, "NumberOfTablesRequired", each [JnlMetaDataTable][Number of instances per month]{0}, Int64.Type),
    AddListOfTablesCol = Table.AddColumn(AddNumberOfTablesCol, "AddMultiTableCol", each List.Accumulate({1..[NumberOfTablesRequired]}, {}, (state, current)=> state & {[JnlMetaDataTable]}), type list),
    ExpandedListOfTables = Table.ExpandListColumn(AddListOfTablesCol, "AddMultiTableCol"),
    ExpandedTables = Table.ExpandTableColumn(ExpandedListOfTables, "AddMultiTableCol", 
        {
            "Debit Account Code and Description", "Credit Account Code and Description", "Jnl Description", "Base Amount Total For Month", 
            "Is recurring", "Annual growth percentage", "Number of instances per month", "Randomise", "Random Variation around base"}, 
        {
            "Debit Account Code and Description", "Credit Account Code and Description", "Jnl Description", "Base Amount Total For Month", 
            "Is recurring", "Annual growth percentage", "Number of instances per month", "Randomise", "Random Variation around base"}),

    //Calculate amount by adding random number factor and percentage growth
    AddRandomAdjCol = Table.AddColumn(ExpandedTables, "Random Number Factor", each if [Randomise] then
            1 + Number.RandomBetween(-[Random Variation around base], [Random Variation around base])
        else
            1, 
        type number),

    AddMonthlyGrowthPercentage = Table.AddColumn(AddRandomAdjCol, "Monthly growth percentage", each fn_MonthlyFromAnnualisedPercentage([Annual growth percentage]),type number),
    AddGrowthAdjCol = Table.AddColumn(AddMonthlyGrowthPercentage, "Growth Factor", each Number.Power(1 + [Monthly growth percentage], [Month Index]), type number),
    AddAmountCol = Table.AddColumn(AddGrowthAdjCol, "Absolute Amount", 
            each Number.Round(
                    [Base Amount Total For Month] / [Number of instances per month] * [Random Number Factor] * [Growth Factor],
                    2),
            type number),

    //Add unique jnl ID and select columns
    AddIndexCol = Table.AddIndexColumn(AddAmountCol, "Index", 0, 1),
    AddJnlIDCol = Table.AddColumn(AddIndexCol, "Jnl ID", each [Jnl Prefix] & "-" & Text.PadStart(Text.From([Index]), 10, "0"), type text),
    SelectCols = Table.SelectColumns(AddJnlIDCol,{"EndOfMonth", "Jnl ID", "Debit Account Code and Description", "Credit Account Code and Description", "Jnl Description", "Absolute Amount"}),
    
    //Unpivot debit and credit accounts and clean up
    UnpivotAccountNumbers = Table.UnpivotOtherColumns(SelectCols, {"EndOfMonth", "Jnl ID", "Jnl Description", "Absolute Amount"}, "Account Number Type", "Account Code and Description"),
    AddJnlAmountCol = Table.AddColumn(UnpivotAccountNumbers, "Jnl Amount", each if [Account Number Type] = "Debit Account Code and Description" then
            [Absolute Amount]
        else
            -[Absolute Amount],
        type number),
    /*************************************************************************************************************************
            Combine ad-hoc and recurring journals
    *************************************************************************************************************************/

    Combined = Table.Combine({AdHocJnls, AddJnlAmountCol}),
    AddAccountCodeCol = Table.AddColumn(Combined, "Account Code", each Text.Start([Account Code and Description], param_NumberOfAccountDigits), type text),
    SelectColsAndReorder = Table.SelectColumns(AddAccountCodeCol,{"EndOfMonth", "Jnl ID", "Account Code", "Jnl Description", "Jnl Amount"}),
    ChangedType3 = Table.TransformColumnTypes(SelectColsAndReorder,{{"EndOfMonth", type date}, {"Jnl ID", type text}, {"Account Code", type text}, {"Jnl Description", type text}}),
    SortedRows = Table.Sort(ChangedType3,{{"EndOfMonth", Order.Ascending}, {"Jnl ID", Order.Ascending}})

in
    SortedRows