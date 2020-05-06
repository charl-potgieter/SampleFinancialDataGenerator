Attribute VB_Name = "m000_EntryPoints"
Option Explicit

Sub RefreshOutputQueries()

    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    

    ThisWorkbook.Sheets("OutputJournals").ListObjects("tbl_Journals").QueryTable.Refresh BackgroundQuery:=False
    ThisWorkbook.Sheets("OutputChartOfAccounts").ListObjects("tbl_ChartOfAccounts").QueryTable.Refresh BackgroundQuery:=False

    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
        
    MsgBox ("Queries refreshed")



End Sub


Sub ExportOutputToFiles()

    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ExportTableToPipeDelimtedText ThisWorkbook.Sheets("OutputJournals").ListObjects("tbl_Journals")
    ExportTableToPipeDelimtedText ThisWorkbook.Sheets("OutputChartOfAccounts").ListObjects("tbl_ChartOfAccounts")
    
    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
        
    MsgBox ("Output files created")


End Sub
