Attribute VB_Name = "m000_EntryPoints"
Option Explicit


Sub ExportTablesInActiveWorkbookToPipeDelimtedText()
'Saves all tables in active sheet pipe delimited text files in active workbook path subfolder PipeDelimitedTextFiles
'File name equals to table name, excl "tbl_" prefix if applicable
'If file already exists a warning is generated, existing file is not overwritten, new file is not generated
        
    Dim lo As ListObject
    
    Dim sFolderPath As String
    Dim sFolderPathAndName As String
    Dim sht As Worksheet

    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    sFolderPath = ActiveWorkbook.Path & Application.PathSeparator & "SampleDataFiles"
    
    If Not FolderExists(sFolderPath) Then
        CreateFolder sFolderPath
    End If
    

    For Each sht In ActiveWorkbook.Sheets
        For Each lo In sht.ListObjects
            If Left(lo.Name, 4) = "tbl_" Then
                sFolderPathAndName = sFolderPath & Application.PathSeparator & Right(lo.Name, Len(lo.Name) - 4) & ".txt"
            Else
                sFolderPathAndName = sFolderPath & Application.PathSeparator & lo.Name & ".txt"
            End If
            ExportListObjectToPipeDelimtedText lo, sFolderPathAndName
        Next lo
    Next sht


    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
        
    MsgBox ("Files created")

End Sub


Sub ExportActiveTableToPipeDelimtedText()
'Saves active table as pipe delimited text files in active workbook path subfolder PipeDelimitedTextFiles
'File name equals to table name, excl "tbl_" prefix if applicable
'If file already exists a warning is generated, existing file is not overwritten, new file is not generated
        
    Dim lo As ListObject
    
    Dim sFolderPath As String
    Dim sFolderPathAndName As String
    Dim sht As Worksheet

    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    sFolderPath = ActiveWorkbook.Path & Application.PathSeparator & "SampleDataFile"
    
    If Not FolderExists(sFolderPath) Then
        CreateFolder sFolderPath
    End If
    
    Set lo = ActiveCell.ListObject

    If Left(lo.Name, 4) = "tbl_" Then
        sFolderPathAndName = sFolderPath & Application.PathSeparator & Right(lo.Name, Len(lo.Name) - 4) & ".txt"
    Else
        sFolderPathAndName = sFolderPath & Application.PathSeparator & lo.Name & ".txt"
    End If
    ExportListObjectToPipeDelimtedText lo, sFolderPathAndName


    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
        
    MsgBox ("File created")

End Sub



Sub ExportPowerQueriesInActiveWorkbookToFiles()

    Dim sFolderSelected As String
    
    sFolderSelected = GetFolder
    If NumberOfFilesInFolder(sFolderSelected) <> 0 Then
        MsgBox ("Please select an empty folder...exiting")
        Exit Sub
    End If
    
    ExportPowerQueriesToFiles sFolderSelected, ActiveWorkbook
    MsgBox ("Queries Exported")

End Sub
