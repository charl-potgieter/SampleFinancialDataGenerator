let
    FileName = param_InputFilePath & "\JnlMetaData.txt",
    Source = Csv.Document(File.Contents(FileName),[Delimiter="|", Columns=11, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    FilteredOutNulls = Table.SelectRows(PromotedHeaders, each ([Jnl Prefix] <> "")),
    ChangedType = Table.TransformColumnTypes(FilteredOutNulls,{{"Jnl Prefix", type text}, {"First Month End", type date}, {"Debit Account Number", type text}, {"Debit Account Full", type text}, {"Credit Account Number", type text}, {"Credit Account Full", type text}, {"Base Amount", type number}, {"Is recurring", type logical}, {"Monthly change trend", type number}, {"Number of instances per month", type number}, {"Randomise", type logical}})
in
    ChangedType