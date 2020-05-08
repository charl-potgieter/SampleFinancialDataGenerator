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
    
    'Pause to ensure above are refreshed as TB is dependent on output journal tb
    Application.Wait Now + #12:00:05 AM#
    ThisWorkbook.Sheets("OutputTB").ListObjects("tbl_TrialBalance").QueryTable.Refresh BackgroundQuery:=False

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
    ExportTableToPipeDelimtedText ThisWorkbook.Sheets("OutputTB").ListObjects("tbl_TrialBalance")
    
    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
        
    MsgBox ("Output files created")


End Sub
