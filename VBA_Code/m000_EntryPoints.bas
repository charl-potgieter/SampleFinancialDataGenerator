Attribute VB_Name = "m000_EntryPoints"
Option Explicit

Sub RefreshOutputQueries()

    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    

    ThisWorkbook.Sheets("JournalOutput").ListObjects("tbl_JournalsOutput").QueryTable.Refresh BackgroundQuery:=False


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
    
    ExportTableToPipeDelimtedText Sheets("Journals").ListObjects("tbl_JournalsOutput")

    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
        
    MsgBox ("Output files created")


End Sub
