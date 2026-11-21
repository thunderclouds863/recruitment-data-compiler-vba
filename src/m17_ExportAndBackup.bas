Attribute VB_Name = "m17_ExportAndBackup"
Option Explicit

Public Sub ExportOutputWorkbook(Optional ByVal isClearOutput As Boolean = False)

    ' Pastikan nama, urutan, status aktif, dan hidden archived column
    ' sudah sesuai Schema_Config sebelum workbook diekspor.
    SchemaB02PrepareAll

    Dim outputPath As String
    outputPath = GetSettingValue("OutputFilePath")
    If Trim(outputPath) = "" Then
        AddRunLog "Skipped", "OutputFilePath kosong."
        Exit Sub
    End If
    
    CreateFolderIfNotExistsFinal GetParentFolderFinal(outputPath)
    BackupOutputWorkbookIfExistsFinal
    
    Dim sheetList As Variant
    sheetList = Array("Blacklist Candidate", "Kode Dropdown", "DB Kode Posisi", "FPTK", _
                      "DB Sourcing", "Monitoring Sourcing", "Flow Map Model", "funneling", _
                      "funneling - Updated", "Run_Log", "Schema_Config")
    
    Dim exportSheets() As String
    Dim exportCount As Long
    exportCount = 0
    
    Dim nm As Variant
    For Each nm In sheetList
        If SheetExists(ThisWorkbook, CStr(nm)) Then
            exportCount = exportCount + 1
            ReDim Preserve exportSheets(1 To exportCount)
            exportSheets(exportCount) = CStr(nm)
        End If
    Next nm
    
    If exportCount = 0 Then
        AddRunLog "Error", "Export dibatalkan: tidak ada sheet yang diekspor."
        Exit Sub
    End If
    
    Dim oldDisplayAlerts As Boolean
    oldDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    On Error GoTo ExportCopyError
    ThisWorkbook.Worksheets(exportSheets).Copy
    
    Dim wbOut As Workbook
    Set wbOut = ActiveWorkbook
    
    On Error GoTo SaveTimestamp
    wbOut.SaveAs fileName:=outputPath, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    Application.DisplayAlerts = oldDisplayAlerts
    
    If isClearOutput Then
        AddRunLog "Output", "Cleared output saved: " & outputPath
    Else
        AddRunLog "Output", "Output saved: " & outputPath
    End If
    Exit Sub
    
ExportCopyError:
    Application.DisplayAlerts = oldDisplayAlerts
    AddRunLog "Error", "Export gagal saat copy sheet: " & Err.Description
    Exit Sub
    
SaveTimestamp:
    Dim archivePath As String
    archivePath = GetArchiveFolderPathFinal()
    If Right(archivePath, 1) <> "\" Then archivePath = archivePath & "\"
    CreateFolderIfNotExistsFinal archivePath
    
    Dim timestampPath As String
    timestampPath = archivePath & "OUTPUT_COMPILED_FPTK_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    
    On Error Resume Next
    wbOut.SaveAs fileName:=timestampPath, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    Application.DisplayAlerts = oldDisplayAlerts
    On Error GoTo 0
End Sub

Public Function BackupEngineWorkbook() As String
    Dim archivePath As String
    archivePath = GetArchiveFolderPathFinal()
    If Right(archivePath, 1) <> "\" Then archivePath = archivePath & "\"
    CreateFolderIfNotExistsFinal archivePath
    
    Dim backupFullPath As String
    backupFullPath = archivePath & "BACKUP_MASTER_COMPILE_FPTK_ENGINE_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs backupFullPath
    BackupEngineWorkbook = backupFullPath
End Function

Public Sub BackupOutputWorkbookIfExistsFinal()
    Dim outputPath As String
    outputPath = GetSettingValue("OutputFilePath")
    If Trim(outputPath) = "" Then Exit Sub
    If Dir(outputPath) = "" Then Exit Sub
    
    Dim archivePath As String
    archivePath = GetArchiveFolderPathFinal()
    If Right(archivePath, 1) <> "\" Then archivePath = archivePath & "\"
    CreateFolderIfNotExistsFinal archivePath
    
    Dim backupFullPath As String
    backupFullPath = archivePath & "BACKUP_OUTPUT_COMPILED_FPTK_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    
    On Error Resume Next
    FileCopy outputPath, backupFullPath
    On Error GoTo 0
End Sub

Public Function GetArchiveFolderPathFinal() As String
    Dim archivePath As String
    archivePath = GetSettingValue("ArchiveFolderPath")
    If Trim(archivePath) <> "" Then
        GetArchiveFolderPathFinal = archivePath
        Exit Function
    End If
    
    Dim parentFolder As String
    If Trim(ThisWorkbook.Path) <> "" Then
        parentFolder = GetParentFolderFinal(ThisWorkbook.Path)
    Else
        parentFolder = Environ("TEMP")
    End If
    
    GetArchiveFolderPathFinal = parentFolder & "\04_Archive\"
End Function

Public Function GetParentFolderFinal(pathText As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Trim(pathText) = "" Then
        GetParentFolderFinal = ""
    Else
        GetParentFolderFinal = fso.GetParentFolderName(pathText)
    End If
End Function

Public Sub CreateFolderIfNotExistsFinal(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim cleanPath As String
    cleanPath = folderPath
    If Right(cleanPath, 1) = "\" Then cleanPath = Left(cleanPath, Len(cleanPath) - 1)
    If cleanPath = "" Then Exit Sub
    
    If fso.FolderExists(cleanPath) Then Exit Sub
    
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(cleanPath)
    If parentPath <> "" And Not fso.FolderExists(parentPath) Then
        CreateFolderIfNotExistsFinal parentPath
    End If
    
    fso.CreateFolder cleanPath
End Sub

Public Sub CopyOutputToArchive()
    Dim outputPath As String
    outputPath = GetSettingValue("OutputFilePath")
    If Trim(outputPath) = "" Then
        AddRunLog "Warning", "OutputFilePath kosong, copy tidak dilakukan"
        Exit Sub
    End If
    
    Dim copyPath As String
    copyPath = GetParentFolderFinal(outputPath) & "\DB_FPTK_COMPILE_COPY.xlsx"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    If fso.FileExists(copyPath) Then
        fso.DeleteFile copyPath, True
    End If
    
    fso.CopyFile ThisWorkbook.FullName, copyPath, True
    On Error GoTo 0
    
    If fso.FileExists(copyPath) Then
        AddRunLog "Output", "Copy saved: " & copyPath
    Else
        AddRunLog "Error", "Gagal copy ke: " & copyPath
    End If
    
    Set fso = Nothing
End Sub



