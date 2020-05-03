Attribute VB_Name = "m050_DataModel"
Option Explicit

Sub ExportPowerQueriesToFiles(ByVal sFolderPath As String, wkb As Workbook)

    Dim qry As WorkbookQuery
    
    For Each qry In wkb.Queries
        WriteStringToTextFile qry.Formula, sFolderPath & Application.PathSeparator & qry.Name & ".m"
    Next qry

End Sub
