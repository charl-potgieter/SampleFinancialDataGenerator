let
    /*
    This query exists to prevent recalling the Journals function in TB process as this will generate
    new random numbers that will not align to the original journals
    */

    Source = Excel.CurrentWorkbook(){[Name="tbl_Journals"]}[Content]
in
    Source