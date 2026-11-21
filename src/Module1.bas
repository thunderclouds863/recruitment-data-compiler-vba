Attribute VB_Name = "Module1"
Option Explicit

' ============================================================
' ===== MODUL UPLOAD KE OUTPUT =====
' ===== Fungsi: Export, Backup, Copy, Upload =====
' ============================================================

' ============================================================
' ==== 1. UPLOAD/EXPORT KE OUTPUT ====
' ============================================================

' ===== UPLOAD UTAMA: Export semua sheet ke file output =====
Public Sub UploadToOutput()
    Dim startTime As Double
    startTime = Timer
    
    AddRunLog "Info", "===== START UPLOAD TO OUTPUT ====="
    
    ' Cek output path
    Dim outputPath As String
    outputPath = GetSettingValue("OutputFilePath")
    
    If Trim(outputPath) = "" Then
        MsgBox "OutputFilePath belum diisi di sheet Settings!", vbExclamation
        AddRunLog "Error", "Upload gagal: OutputFilePath kosong"
        Exit Sub
    End If
    
    ' Buat folder jika belum ada
    CreateFolderIfNotExists GetParentFolderPath(outputPath)
    
    ' Backup output lama jika ada
    BackupOutputFile
    
    ' Export ke output
    Dim exportSuccess As Boolean
    exportSuccess = ExportToOutput(outputPath)
    
    If exportSuccess Then
        ' Copy ke archive
        CopyToArchive outputPath
        
        ' Update log
        AddRunLog "Success", "Upload ke output selesai: " & outputPath & " (" & Format(Timer - startTime, "0.0") & "s)"
        
        ' Notifikasi
        MsgBox "Upload ke output berhasil!" & vbCrLf & _
               "File: " & outputPath & vbCrLf & _
               "Waktu: " & Format(Timer - startTime, "0.0") & " detik", vbInformation
    Else
        AddRunLog "Error", "Upload ke output gagal"
        MsgBox "Upload ke output gagal! Cek log untuk detail.", vbCritical
    End If
End Sub

' ===== EXPORT KE OUTPUT =====
Private Function ExportToOutput(ByVal outputPath As String) As Boolean
    Dim wbOut As Workbook
    Dim sheetNames As Variant
    Dim exportSheets() As String
    Dim exportCount As Long
    Dim nm As Variant
    Dim oldDisplayAlerts As Boolean
    Dim tempPath As String
    
    On Error GoTo ErrHandler
    
    ' Daftar sheet yang akan diexport
    sheetNames = Array( _
        "Blacklist Candidate", _
        "Kode Dropdown (tidak diubah)", _
        "DB Kode Posisi", _
        "FPTK", _
        "DB Sourcing", _
        "Monitoring Sourcing", _
        "Flow Map Model (tidak diubah)", _
        "funneling", _
        "funneling - Updated", _
        "Run_Log")
    
    ' Filter sheet yang ada
    exportCount = 0
    For Each nm In sheetNames
        If ThisWorkbookSheetExists(CStr(nm)) Then
            exportCount = exportCount + 1
            ReDim Preserve exportSheets(1 To exportCount)
            exportSheets(exportCount) = CStr(nm)
        End If
    Next nm
    
    If exportCount = 0 Then
        AddRunLog "Error", "Export dibatalkan: Tidak ada sheet yang bisa diexport"
        ExportToOutput = False
        Exit Function
    End If
    
    ' Nonaktifkan alert
    oldDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    ' Copy sheet ke workbook baru
    ThisWorkbook.Worksheets(exportSheets).Copy
    
    Set wbOut = ActiveWorkbook
    
    ' Hapus worksheet default jika ada
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wbOut.Worksheets
        If LCase(ws.Name) = "sheet1" Or LCase(ws.Name) = "sheet2" Or LCase(ws.Name) = "sheet3" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    On Error GoTo 0
    
    ' Simpan
    wbOut.SaveAs fileName:=outputPath, FileFormat:=xlOpenXMLWorkbook
    wbOut.Close SaveChanges:=False
    
    Application.DisplayAlerts = oldDisplayAlerts
    
    ExportToOutput = True
    AddRunLog "Info", "Export sukses: " & outputPath
    Exit Function
    
ErrHandler:
    Application.DisplayAlerts = oldDisplayAlerts
    AddRunLog "Error", "ExportToOutput: " & Err.Description
    ExportToOutput = False
End Function

' ===== COPY KE ARCHIVE =====
Private Sub CopyToArchive(ByVal outputPath As String)
    Dim archiveFolder As String
    Dim archivePath As String
    Dim fso As Object
    
    On Error GoTo ErrHandler
    
    archiveFolder = GetArchiveFolderPath()
    CreateFolderIfNotExists archiveFolder
    
    ' Nama file dengan timestamp
    archivePath = archiveFolder & "OUTPUT_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Hapus file lama jika ada (max 10 file)
    CleanOldArchiveFiles archiveFolder, 10
    
    ' Copy file
    If fso.FileExists(outputPath) Then
        fso.CopyFile outputPath, archivePath, True
        AddRunLog "Info", "Copy ke archive: " & archivePath
    End If
    
    Exit Sub
    
ErrHandler:
    AddRunLog "Warning", "CopyToArchive: " & Err.Description
End Sub

' ===== CLEAN OLD ARCHIVE FILES =====
Private Sub CleanOldArchiveFiles(ByVal folderPath As String, ByVal maxFiles As Long)
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim files As Object
    Dim i As Long
    Dim fileCount As Long
    Dim fileDict As Object
    Dim key As Variant
    Dim keys As Variant
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Sub
    
    Set folder = fso.GetFolder(folderPath)
    Set files = folder.files
    Set fileDict = CreateObject("Scripting.Dictionary")
    
    ' Kumpulkan file output
    For Each file In files
        If LCase(file.Name) Like "output_*.xlsx" Then
            fileDict.Add file.DateCreated & "|" & file.Name, file.Path
        End If
    Next file
    
    fileCount = fileDict.Count
    
    ' Hapus file tertua jika melebihi max
    If fileCount > maxFiles Then
        keys = fileDict.keys
        ' Sort keys (ascending = tertua dulu)
        Dim j As Long, k As Long
        Dim tempKey As Variant
        For j = 0 To UBound(keys) - 1
            For k = j + 1 To UBound(keys)
                If keys(j) > keys(k) Then
                    tempKey = keys(j)
                    keys(j) = keys(k)
                    keys(k) = tempKey
                End If
            Next k
        Next j
        
        ' Hapus file tertua (pertama)
        For i = 0 To fileCount - maxFiles - 1
            Dim filePath As String
            filePath = fileDict(keys(i))
            If fso.FileExists(filePath) Then
                fso.DeleteFile filePath, True
                AddRunLog "Info", "Hapus archive lama: " & filePath
            End If
        Next i
    End If
    
    On Error GoTo 0
End Sub


' ============================================================
' ==== 2. BACKUP ====
' ============================================================

' ===== BACKUP OUTPUT FILE =====
Public Sub BackupOutputFile()
    Dim outputPath As String
    Dim backupPath As String
    Dim fso As Object
    
    outputPath = GetSettingValue("OutputFilePath")
    If Trim(outputPath) = "" Then Exit Sub
    
    If Dir(outputPath) = "" Then Exit Sub
    
    ' Buat backup di folder yang sama
    backupPath = GetParentFolderPath(outputPath) & "\BACKUP_" & Format(Now, "yyyymmdd_hhnnss") & "_" & GetFileName(outputPath)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    fso.CopyFile outputPath, backupPath, True
    If Err.Number = 0 Then
        AddRunLog "Info", "Backup output: " & backupPath
    Else
        AddRunLog "Warning", "Gagal backup output: " & Err.Description
    End If
    On Error GoTo 0
End Sub

' ===== BACKUP ENGINE FILE =====
Public Sub BackupEngineFile()
    Dim backupPath As String
    Dim fso As Object
    
    backupPath = ThisWorkbook.Path & "\BACKUP_ENGINE_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    fso.CopyFile ThisWorkbook.FullName, backupPath, True
    If Err.Number = 0 Then
        AddRunLog "Info", "Backup engine: " & backupPath
    Else
        AddRunLog "Warning", "Gagal backup engine: " & Err.Description
    End If
    On Error GoTo 0
End Sub


' ============================================================
' ==== 3. SAVE CURRENT WORKBOOK ====
' ============================================================

' ===== SAVE ENGINE WORKBOOK =====
Public Sub SaveEngineWorkbook()
    On Error Resume Next
    ThisWorkbook.Save
    If Err.Number = 0 Then
        AddRunLog "Info", "Engine workbook saved"
    Else
        AddRunLog "Warning", "Gagal save engine: " & Err.Description
    End If
    On Error GoTo 0
End Sub

' ===== SAVE AS ENGINE WORKBOOK =====
Public Sub SaveEngineWorkbookAs()
    Dim savePath As String
    Dim oldDisplayAlerts As Boolean
    
    savePath = ThisWorkbook.Path & "\MASTER_COMPILE_ENGINE_" & Format(Now, "yyyymmdd") & ".xlsm"
    
    oldDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    On Error Resume Next
    ThisWorkbook.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    If Err.Number = 0 Then
        AddRunLog "Info", "Engine saved as: " & savePath
        MsgBox "Engine berhasil disimpan sebagai:" & vbCrLf & savePath, vbInformation
    Else
        AddRunLog "Error", "Gagal save engine as: " & Err.Description
        MsgBox "Gagal save engine: " & Err.Description, vbCritical
    End If
    On Error GoTo 0
    
    Application.DisplayAlerts = oldDisplayAlerts
End Sub


' ============================================================
' ==== 4. COMPILE DAN UPLOAD SEKALIGUS ====
' ============================================================

' ===== COMPILE + UPLOAD (ALL IN ONE) =====
Public Sub CompileAndUpload()
    ' Jalankan compile
    CompileFPTK_All
    
    ' Upload ke output
    UploadToOutput
    
    MsgBox "Compile dan Upload selesai!", vbInformation
End Sub


' ============================================================
' ==== 5. HELPER FUNCTIONS ====
' ============================================================

' ===== GET PARENT FOLDER =====
Private Function GetParentFolderPath(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    GetParentFolderPath = fso.GetParentFolderName(filePath)
    If GetParentFolderPath = "" Then
        GetParentFolderPath = Environ("TEMP")
    End If
    On Error GoTo 0
End Function

' ===== GET FILE NAME =====
Private Function GetFileName(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    GetFileName = fso.GetFileName(filePath)
    If GetFileName = "" Then
        GetFileName = "output.xlsx"
    End If
    On Error GoTo 0
End Function

' ===== GET ARCHIVE FOLDER =====
Private Function GetArchiveFolderPath() As String
    Dim archivePath As String
    Dim parentFolder As String
    
    archivePath = GetSettingValue("ArchiveFolderPath")
    If Trim(archivePath) <> "" Then
        GetArchiveFolderPath = archivePath
        Exit Function
    End If
    
    If Trim(ThisWorkbook.Path) <> "" Then
        parentFolder = GetParentFolderPath(ThisWorkbook.Path)
    Else
        parentFolder = Environ("TEMP")
    End If
    
    GetArchiveFolderPath = parentFolder & "\04_Archive\"
End Function

' ===== CREATE FOLDER IF NOT EXISTS =====
Private Sub CreateFolderIfNotExists(ByVal folderPath As String)
    Dim fso As Object
    Dim cleanPath As String
    
    If Trim(folderPath) = "" Then Exit Sub
    
    cleanPath = folderPath
    If Right(cleanPath, 1) = "\" Then cleanPath = Left(cleanPath, Len(cleanPath) - 1)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    If Not fso.FolderExists(cleanPath) Then
        fso.CreateFolder cleanPath
    End If
    On Error GoTo 0
End Sub

' ===== SHEET EXISTS =====
Private Function ThisWorkbookSheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    ThisWorkbookSheetExists = Not ws Is Nothing
End Function


' ============================================================
' ==== 6. UPLOAD DENGAN PROGRESS BAR ====
' ============================================================

' ===== UPLOAD KE OUTPUT DENGAN PROGRESS =====
Public Sub UploadToOutputWithProgress()
    Dim totalSteps As Long
    Dim currentStep As Long
    
    totalSteps = 5
    
    InitProgress totalSteps, "Upload ke Output...", True
    
    UpdateProgress "Membaca konfigurasi...", "-"
    currentStep = currentStep + 1
    
    ' Cek output path
    Dim outputPath As String
    outputPath = GetSettingValue("OutputFilePath")
    
    If Trim(outputPath) = "" Then
        FinishProgress "Error: OutputFilePath kosong"
        MsgBox "OutputFilePath belum diisi di sheet Settings!", vbExclamation
        Exit Sub
    End If
    
    UpdateProgress "Membuat folder output...", "-"
    currentStep = currentStep + 1
    CreateFolderIfNotExists GetParentFolderPath(outputPath)
    
    UpdateProgress "Backup output lama...", "-"
    currentStep = currentStep + 1
    BackupOutputFile
    
    UpdateProgress "Export ke output...", outputPath
    currentStep = currentStep + 1
    
    Dim exportSuccess As Boolean
    exportSuccess = ExportToOutput(outputPath)
    
    If exportSuccess Then
        UpdateProgress "Copy ke archive...", "-"
        currentStep = currentStep + 1
        CopyToArchive outputPath
        
        FinishProgress "Upload selesai!"
        
        AddRunLog "Success", "Upload ke output selesai: " & outputPath
        MsgBox "Upload ke output berhasil!" & vbCrLf & "File: " & outputPath, vbInformation
    Else
        FinishProgress "Upload gagal!"
        AddRunLog "Error", "Upload ke output gagal"
        MsgBox "Upload ke output gagal! Cek log untuk detail.", vbCritical
    End If
End Sub


' ============================================================
' ==== 7. EXPORT SEMUA SHEET KE FOLDER ====
' ============================================================

' ===== EXPORT SHEET SATU PER SATU KE FOLDER =====
Public Sub ExportAllSheetsToFolder()
    Dim folderPath As String
    Dim ws As Worksheet
    Dim wbOut As Workbook
    Dim fso As Object
    
    folderPath = ThisWorkbook.Path & "\Exported_Sheets\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    On Error GoTo 0
    
    For Each ws In ThisWorkbook.Worksheets
        ' Skip sheet tertentu
        If ws.Name <> "Settings" And ws.Name <> "Progress" Then
            ws.Copy
            Set wbOut = ActiveWorkbook
            
            ' Hapus sheet lain
            Dim ws2 As Worksheet
            Application.DisplayAlerts = False
            For Each ws2 In wbOut.Worksheets
                If ws2.Name <> ws.Name Then
                    ws2.Delete
                End If
            Next ws2
            Application.DisplayAlerts = True
            
            ' Simpan
            Dim savePath As String
            savePath = folderPath & ws.Name & ".xlsx"
            
            On Error Resume Next
            wbOut.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
            wbOut.Close SaveChanges:=False
            If Err.Number = 0 Then
                AddRunLog "Info", "Export sheet: " & ws.Name & " -> " & savePath
            End If
            On Error GoTo 0
        End If
    Next ws
    
    MsgBox "Semua sheet telah diexport ke:" & vbCrLf & folderPath, vbInformation
End Sub


' ============================================================
' ==== 8. VALIDASI OUTPUT ====
' ============================================================

' ===== CEK OUTPUT FILE =====
Public Function CheckOutputFile() As Boolean
    Dim outputPath As String
    Dim fso As Object
    Dim wb As Workbook
    
    outputPath = GetSettingValue("OutputFilePath")
    
    If Trim(outputPath) = "" Then
        AddRunLog "Warning", "CheckOutput: OutputFilePath kosong"
        CheckOutputFile = False
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(outputPath) Then
        AddRunLog "Warning", "CheckOutput: File tidak ditemukan: " & outputPath
        CheckOutputFile = False
        Exit Function
    End If
    
    ' Coba buka file
    On Error Resume Next
    Set wb = Workbooks.Open(outputPath, ReadOnly:=True, UpdateLinks:=0)
    If Err.Number <> 0 Then
        AddRunLog "Error", "CheckOutput: Gagal buka file: " & Err.Description
        CheckOutputFile = False
        Exit Function
    End If
    On Error GoTo 0
    
    wb.Close SaveChanges:=False
    
    AddRunLog "Info", "CheckOutput: File valid: " & outputPath
    CheckOutputFile = True
End Function

