Attribute VB_Name = "m020_FileUtilities"
Option Explicit
Option Private Module

'-----------------------------------------------------------------------------------------------------
'           Requires reference to Microsoft Scripting runtime
'-----------------------------------------------------------------------------------------------------


Function NumberOfFilesInFolder(ByVal sFolderPath As String) As Integer
'Requires refence: Microsoft Scripting Runtime
'This is non-recursive


    Dim oFSO As FileSystemObject
    Dim oFolder As Folder
    
    Set oFSO = New FileSystemObject
    Set oFolder = oFSO.GetFolder(sFolderPath)
    NumberOfFilesInFolder = oFolder.Files.Count


End Function

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

Sub ExportTableToPipeDelimtedText(lo As ListObject)
'Saves list object as pipe delimited text files in active workbook path subfolder OutputFiles
'File name equals to table name, excl "tbl_" prefix if applicable
'Any existing file will be overwritten
        
    
    Dim sFolderPath As String
    Dim sFolderPathAndName As String
    Dim sht As Worksheet

    
    sFolderPath = ThisWorkbook.Path & Application.PathSeparator & "OutputFiles"
    
    
    If Not FolderExists(sFolderPath) Then
        CreateFolder sFolderPath
    End If
    

    If Left(lo.Name, 4) = "tbl_" Then
        sFolderPathAndName = sFolderPath & Application.PathSeparator & Right(lo.Name, Len(lo.Name) - 4) & ".txt"
        sFolderPathAndName = GetNextAvailableFileName(sFolderPathAndName)
    Else
        sFolderPathAndName = sFolderPath & Application.PathSeparator & lo.Name & ".txt"
        sFolderPathAndName = GetNextAvailableFileName(sFolderPathAndName)
    End If
    ExportListObjectToPipeDelimtedText lo, sFolderPathAndName


End Sub


Sub ExportListObjectToPipeDelimtedText(ByRef lo As ListObject, ByVal sFilePathAndName As String)
'Requires reference to Microsoft Scripting Runtime
'Saves sht as a pipe delimted text file
'Existing files will be overwritten

    Dim dblNumberOfRows As Double
    Dim dblNumberOfCols As Double
    Dim iFileNo As Integer
    Dim i As Double
    Dim j As Double
    Dim sRowStringToWrite As String

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
        If j < dblNumberOfRows Then
            Print #iFileNo, sRowStringToWrite
        Else
            'note the semi-colon at end to avoid the newline
            Print #iFileNo, sRowStringToWrite;
        End If
    Next j

    Close #iFileNo

End Sub



Function GetNextAvailableFileName(ByVal sFilePath As String) As String
'Requires refence: Microsoft Scripting Runtime
'Returns next available file name.  Can be utilised to ensure files are not overwritten

    Dim oFSO As FileSystemObject
    Dim sFolder As String
    Dim sFileName As String
    Dim sFileExtension As String
    Dim i As Long

    Set oFSO = CreateObject("Scripting.FileSystemObject")

    With oFSO
        sFolder = .GetParentFolderName(sFilePath)
        sFileName = .GetBaseName(sFilePath)
        sFileExtension = .GetExtensionName(sFilePath)

        Do While .FileExists(sFilePath)
            i = i + 1
            sFilePath = .BuildPath(sFolder, sFileName & "(" & i & ")." & sFileExtension)
        Loop
        
    End With

    GetNextAvailableFileName = sFilePath

End Function
