Attribute VB_Name = "m030_FileUtilities"
Option Explicit
Option Private Module
'-----------------------------------------------------------------------------
'   Requires reference to Microsoft Scripting runtime
'-----------------------------------------------------------------------------

Function FolderExists(ByVal sFolderPath) As Boolean
'Requires reference to Microsoft Scripting runtime
'An alternative solution exists using the DIR function but this seems to result in memory leak and folder is
'not released by VBA
    
    Dim FSO As Scripting.FileSystemObject
    Dim FolderPath As String
    
    Set FSO = New Scripting.FileSystemObject
    
    If Right(sFolderPath, 1) <> Application.PathSeparator Then
        FolderPath = FolderPath & Application.PathSeparator
    End If
    
    FolderExists = FSO.FolderExists(sFolderPath)
    Set FSO = Nothing

End Function

Function FileExists(ByVal sFilePath) As Boolean
'Requires reference to Microsoft Scripting runtime
'An alternative solution exists using the DIR function but this seems to result in memory leak and file is
'not released by VBA

    Dim FSO As Scripting.FileSystemObject
    Dim FolderPath As String
    
    Set FSO = New Scripting.FileSystemObject
    
    FileExists = FSO.FileExists(sFilePath)
    Set FSO = Nothing


End Function


Sub CreateFolder(ByVal sFolderPath As String)
'   Requires reference to Microsoft Scripting runtime

    Dim FSO As FileSystemObject

    If FolderExists(sFolderPath) Then
        MsgBox ("Folder already exists, new folder not created")
    Else
        Set FSO = New FileSystemObject
        FSO.CreateFolder sFolderPath
    End If
    
    Set FSO = Nothing

End Sub


Function NumberOfFilesInFolder(ByVal sFolderPath As String) As Integer
'Requires refence: Microsoft Scripting Runtime
'This is non-recursive


    Dim oFSO As FileSystemObject
    Dim oFolder As Folder
    
    Set oFSO = New FileSystemObject
    Set oFolder = oFSO.GetFolder(sFolderPath)
    NumberOfFilesInFolder = oFolder.Files.Count


End Function



Sub ExportListObjectToPipeDelimtedText(ByRef lo As ListObject, ByVal sFilePathAndName As String)
'Requires reference to Microsoft Scripting Runtime
'Saves sht as a pipe delimted text file
'Existing files will not be overwritten.  Warning is given.
    
    Dim dblNumberOfRows As Double
    Dim dblNumberOfCols As Double
    Dim iFileNo As Integer
    Dim i As Double
    Dim j As Double
    Dim sRowStringToWrite As String

    
    If FileExists(sFilePathAndName) Then
        MsgBox ("File " & sFilePathAndName & " already exists.  New file has not been generated")
    End If

    'Get first free file number
    iFileNo = FreeFile

    
    Open sFilePathAndName For Output As #iFileNo
    
    dblNumberOfRows = lo.Range.Rows.Count
    dblNumberOfCols = lo.Range.Columns.Count
    
    
    For j = 1 To dblNumberOfRows
        sRowStringToWrite = ""
        For i = 1 To dblNumberOfCols
            If i < dblNumberOfCols Then
                sRowStringToWrite = sRowStringToWrite & lo.Range.Cells(j, i) & "|"
            Else
                sRowStringToWrite = sRowStringToWrite & lo.Range.Cells(j, i)
            End If
        Next i
        Print #1, sRowStringToWrite
    Next j

    Close #iFileNo

End Sub


Function GetFolder() As String
'Returns the results of a user folder picker

    Dim fldr As FileDialog
    Dim sItem As String
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a folder"
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        If .Show = -1 Then
            GetFolder = .SelectedItems(1)
        End If
    End With
    
    Set fldr = Nothing


End Function




Function WriteStringToTextFile(ByVal sStr As String, ByVal sFilePath As String)
'Requires reference to Microsoft Scripting Runtime
'Writes sStr to a text file

    Dim FSO As Object
    Dim oFile As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = FSO.CreateTextFile(sFilePath)
    oFile.Write (sStr)
    oFile.Close
    Set FSO = Nothing
    Set oFile = Nothing

End Function
