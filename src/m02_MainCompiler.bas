Attribute VB_Name = "m02_MainCompiler"
Option Explicit

Private mCurrentCompileStage As String

Private mTraceStageStartedAt As Double

Private Sub TraceStageStart( _
    ByVal stageName As String, _
    ByVal detailText As String)

    mTraceStageStartedAt = Timer

    SetCompileStage stageName

    HeartbeatProgress _
        "RUNNING", _
        gSchemaCurrentFolder, _
        stageName, _
        detailText

    AddRunLog _
        "Trace", _
        "START | " & stageName & _
        " | " & detailText

End Sub

Private Sub TraceStageEnd( _
    ByVal stageName As String, _
    ByVal detailText As String)

    Dim elapsedSeconds As Double

    elapsedSeconds = Timer - mTraceStageStartedAt

    If elapsedSeconds < 0 Then
        elapsedSeconds = elapsedSeconds + 86400
    End If

    HeartbeatProgress _
        "DONE", _
        gSchemaCurrentFolder, _
        stageName, _
        detailText & _
        " | " & _
        Format$(elapsedSeconds, "0.0") & _
        " sec"

    AddRunLog _
        "Trace", _
        "DONE | " & stageName & _
        " | " & _
        Format$(elapsedSeconds, "0.0") & _
        " sec | " & detailText

End Sub

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================
Public Sub CompileFPTK_All()

    ' Default button = compile data saja.
    ' Report/export dijalankan terpisah agar compile jauh lebih cepat.
    CompileFPTK_Core _
        True, _
        False

End Sub

Public Sub CompileFPTK_All_Full()

    ' Full mode: compile + monitoring + funneling + export + backup.
    CompileFPTK_Core _
        True, _
        True

End Sub

Public Sub CompileFPTK_Silent()

    CompileFPTK_Core _
        False, _
        False

End Sub

' ============================================================
' MANUAL ENTRY POINT - DB KODE POSISI SAJA
' ============================================================
Public Sub CompileDBKodePosisiOnly()

    Dim rootFolder As String
    Dim compileFolders As Collection

    On Error GoTo ErrorHandler

    rootFolder = Trim$( _
        GetSettingValue( _
            "InputFolderPath"))

    If Len(rootFolder) = 0 Then

        Err.Raise _
            vbObjectError + 3191, _
            "CompileDBKodePosisiOnly", _
            "InputFolderPath di Settings masih kosong."

    End If

    rootFolder = _
        EnsureTrailingSlash( _
            rootFolder)

    Set compileFolders = _
        GetConfiguredCompileFolders( _
            rootFolder)

    If compileFolders.Count = 0 Then

        Err.Raise _
            vbObjectError + 3192, _
            "CompileDBKodePosisiOnly", _
            "Tidak ada folder input yang valid."

    End If

    Ensure_DB_Kode_Posisi_Target

    Compile_DB_Kode_Posisi_MultiFolders _
        compileFolders, _
        True

    Exit Sub

ErrorHandler:

    MsgBox _
        "Compile DB Kode Posisi gagal." & _
        vbCrLf & _
        "Error " & CStr(Err.Number) & _
        ": " & Err.Description, _
        vbCritical, _
        "DB Kode Posisi"

End Sub

Public Sub RefreshReportsAndExportOnly()

    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean

    On Error GoTo ErrorHandler

    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    InitProgress _
        6, _
        "Finalize Reports", _
        True

    FinalizeReportsAndExportTurbo

    ThisWorkbook.Save

    FinishProgress _
        "Reports and export completed"

SafeExit:

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.StatusBar = False

    Exit Sub

ErrorHandler:

    On Error Resume Next

    FinishProgress _
        "Finalize failed: " & _
        Err.Description

    MsgBox _
        "Finalize report gagal." & _
        vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Finalize Reports"

    Resume SafeExit

End Sub
Private Sub CompileFPTK_Core( _
    ByVal showMessage As Boolean, _
    ByVal runFinalize As Boolean)

    Dim folderPath As String
    Dim compileFolders As Collection
    Dim folderItem As Variant

    Dim totalFiles As Long
    Dim totalSteps As Long

    Dim oldCalc As XlCalculation
    Dim oldAskToUpdateLinks As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean

    Dim applicationStateCaptured As Boolean
    Dim errorText As String
    Dim compileStartedAt As Double

    If gSchemaCompileInProgress Then

        If showMessage Then

            MsgBox _
                "Compiler sedang berjalan." & vbCrLf & _
                "Tunggu proses sebelumnya selesai.", _
                vbExclamation, _
                "Compile Database"

        End If

        Exit Sub

    End If

    If ThisWorkbook.ReadOnly Then

        If showMessage Then

            MsgBox _
                "File engine sedang Read-Only.", _
                vbExclamation, _
                "Compile Database"

        End If

        Exit Sub

    End If

    On Error GoTo ErrHandler

    compileStartedAt = Timer

    oldCalc = Application.Calculation
    oldAskToUpdateLinks = _
        Application.AskToUpdateLinks

    oldScreenUpdating = _
        Application.ScreenUpdating

    oldDisplayAlerts = _
        Application.DisplayAlerts

    oldEnableEvents = _
        Application.EnableEvents

    applicationStateCaptured = True

    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    gSchemaCompileInProgress = True
    gSchemaLastError = ""
    gSchemaCurrentFolder = ""

    SetCompileStage _
        "Preparing Support Sheets"

    Application.StatusBar = _
        "Preparing Support Sheets..."

    EnsureSupportSheets

    ' ========================================================
    ' FPTK CORRUPTION TRACE
    ' Bersihkan row helper/corrupt lama dan mulai baseline.
    ' Tidak mengubah flow compile existing.
    ' ========================================================
    FPTKTrace_ResetAndClean _
        "Before Schema"

    ' Schema hanya disiapkan SATU KALI.
    ' Tidak ada post-compile/final schema pass.
    SetCompileStage _
        "Preparing Schema"

    Application.StatusBar = _
        "Preparing Schema..."

    SchemaB02PrepareAll

    FPTKTrace_Checkpoint _
        "After SchemaB02PrepareAll"

    SetCompileStage _
        "Scanning Input Folders"

    folderPath = Trim$( _
        GetSettingValue("InputFolderPath"))

    If Len(folderPath) = 0 Then

        Err.Raise _
            vbObjectError + 3101, _
            "CompileFPTK_Core", _
            "InputFolderPath di sheet Settings masih kosong."

    End If

    folderPath = EnsureTrailingSlash( _
        folderPath)

    If Not FolderExistsFSO(folderPath) Then

        Err.Raise _
            vbObjectError + 3102, _
            "CompileFPTK_Core", _
            "Folder tidak ditemukan:" & _
            vbCrLf & folderPath

    End If

    Set compileFolders = _
        GetConfiguredCompileFolders( _
            folderPath)

    If compileFolders.Count = 0 Then

        Err.Raise _
            vbObjectError + 3103, _
            "CompileFPTK_Core", _
            "Tidak ada folder input yang valid."

    End If

    totalFiles = CountFilesInFolders( _
        compileFolders)

    If totalFiles = 0 Then

        Err.Raise _
            vbObjectError + 3104, _
            "CompileFPTK_Core", _
            "Folder recruiter berhasil terbaca, tetapi tidak ada " & _
            "file .xlsx atau .xlsm yang ditemukan langsung di " & _
            "folder tersebut." & _
            vbCrLf & vbCrLf & _
            "Jalankan macro DiagnoseConfiguredInputFiles " & _
            "untuk melihat jumlah file per folder."

    End If

    If runFinalize Then
        totalSteps = _
            (compileFolders.Count * 2) + 10
    Else
        totalSteps = _
            (compileFolders.Count * 2) + 5
    End If

    InitProgress _
        totalSteps, _
        "Compile Database FPTK", _
        showMessage

    UpdateProgress _
        "Preserving existing compiled data...", _
        "APPEND / UPDATE"

    AddRunLog _
        "Info", _
        "TURBO MODE | Append/update | " & _
        "Finalize=" & CStr(runFinalize)

    ' DB Kode Posisi tidak bergantung pada hasil compile FPTK.
    ' Jalankan lebih awal supaya tetap dibuat walaupun stage
    ' FPTK/Blacklist berikutnya mengalami masalah.
    gSchemaCurrentFolder = _
        "MULTI FOLDER"

    UpdateProgress _
        "Preparing DB Kode Posisi target...", _
        "Ensuring static A:J sheet"

    TraceStageStart _
        "DB Kode Posisi - Ensure Target", _
        "Creating sheet and static headers A:J"

    Ensure_DB_Kode_Posisi_Target

    TraceStageEnd _
        "DB Kode Posisi - Ensure Target", _
        "Target sheet is ready"

    FPTKTrace_Checkpoint _
        "After Ensure DB Kode Posisi Target"

    For Each folderItem In compileFolders

        gSchemaCurrentFolder = _
            CStr(folderItem)

        UpdateProgress _
            "Starting compile folder...", _
            CStr(folderItem)

        DoEvents

        CompileOneFolder _
            CStr(folderItem)

    Next folderItem

    ' FPTK target sekarang sudah lengkap.
    ' Bentuk DB Kode Posisi dari sheet FPTK engine,
    ' lalu enrich dari sheet DB Kode Posisi source bila tersedia.
    gSchemaCurrentFolder = _
        "MULTI FOLDER"

    UpdateProgress _
        "Compiling DB Kode Posisi from compiled FPTK...", _
        CStr(compileFolders.Count) & _
        " folders"

    TraceStageStart _
        "DB Kode Posisi One Pass", _
        "Building static A:J from compiled FPTK"

    Compile_DB_Kode_Posisi_MultiFolders _
        compileFolders, _
        False

    TraceStageEnd _
        "DB Kode Posisi One Pass", _
        "Static A:J master completed"

    FPTKTrace_Checkpoint _
        "After DB Kode Posisi One Pass"

    ' ========================================================
    ' PENTING:
    ' FPTK per-folder selesai lebih dulu, sedangkan DB Kode
    ' Posisi baru final setelah one-pass di atas.
    '
    ' Karena itu User (Manager), Indirect User dan Kode Angka
    ' FPTK wajib direfresh SEKARANG, setelah master A:J final.
    ' ========================================================
    UpdateProgress _
        "Refreshing FPTK User from final DB Kode Posisi...", _
        "POST MASTER"

    TraceStageStart _
        "FPTK User Post Master", _
        "Refreshing User/Indirect after DB Kode Posisi"

    RefreshFPTKUsersAfterMaster

    TraceStageEnd _
        "FPTK User Post Master", _
        "User/Indirect refresh completed"

    FPTKTrace_Checkpoint _
        "After RefreshFPTKUsersAfterMaster"

    UpdateProgress _
        "Compiling DB Sourcing all folders...", _
        CStr(compileFolders.Count) & _
        " folders"

    TraceStageStart _
        "DB Sourcing One Pass", _
        "Building one index for all folders"

    Compile_DB_Sourcing_MultiFolders _
        compileFolders, _
        False

    TraceStageEnd _
        "DB Sourcing One Pass", _
        "All configured folders completed"

    FPTKTrace_Checkpoint _
        "After DB Sourcing One Pass"

    gSchemaCurrentFolder = ""

    ' FPTK fast compiler sudah menghitung derived values
    ' hanya untuk row insert/update. Jangan scan seluruh FPTK lagi.
    If runFinalize Then

        FinalizeReportsAndExportTurbo

        FPTKTrace_Checkpoint _
            "After FinalizeReportsAndExportTurbo"

    Else

        UpdateProgress _
            "Skipping reports/export for turbo mode...", _
            "Run RefreshReportsAndExportOnly when needed"

    End If

    UpdateProgress _
        "Saving engine file...", _
        "-"

    SetCompileStage _
        "Save Engine"

    SetSettingValue _
        "LastCompiledAt", _
        Format(Now, "yyyy-mm-dd hh:nn:ss")

    SetSettingValue _
        "LastCompiledBy", _
        Environ$("Username")

    SetSettingValue _
        "LastCompileMode", _
        IIf(runFinalize, "FULL TURBO", "DATA TURBO")

    AddRunLog _
        "Success", _
        "Turbo compile selesai | " & _
        TurboElapsedText(compileStartedAt)

    ThisWorkbook.Save

    SetCompileStage ""

    FinishProgress _
        "Compile completed successfully"

    CloseSourceExcelApp

    RestoreCompilerApplicationState _
        applicationStateCaptured, _
        oldCalc, _
        oldAskToUpdateLinks, _
        oldScreenUpdating, _
        oldDisplayAlerts, _
        oldEnableEvents

    gSchemaCompileInProgress = False
    gSchemaCurrentFolder = ""
    gSchemaLastError = ""

    If showMessage Then

        If runFinalize Then

            MsgBox _
                "Compile full selesai.", _
                vbInformation, _
                "Compile Database"

        Else

            MsgBox _
                "Compile data selesai."

        End If

    End If

    Exit Sub

ErrHandler:

    errorText = _
        "Error " & CStr(Err.Number)

    If Len(mCurrentCompileStage) > 0 Then

        errorText = _
            errorText & _
            " pada " & _
            mCurrentCompileStage

    End If

    errorText = _
        errorText & _
        ": " & Err.Description

    gSchemaLastError = errorText

    On Error Resume Next

    AddRunLog _
        "Error", _
        errorText

    FinishProgress errorText
    CloseSourceExcelApp

    On Error GoTo 0

    RestoreCompilerApplicationState _
        applicationStateCaptured, _
        oldCalc, _
        oldAskToUpdateLinks, _
        oldScreenUpdating, _
        oldDisplayAlerts, _
        oldEnableEvents

    gSchemaCompileInProgress = False
    gSchemaCurrentFolder = ""

    SetCompileStage ""

    If showMessage Then

        MsgBox _
            errorText, _
            vbCritical, _
            "Compile Database"

    End If

End Sub
Private Sub FinalizeReportsAndExportTurbo()

    UpdateProgress _
        "Creating Monitoring Sourcing...", _
        "-"

    SetCompileStage _
        "Monitoring Sourcing"

    On Error Resume Next

    CreateMonitoringSheet

    If Err.Number <> 0 Then

        AddRunLog _
            "Error", _
            "CreateMonitoringSheet failed: " & _
            Err.Description

        Err.Clear

    End If

    FPTKTrace_Checkpoint _
        "After Monitoring Sourcing"

    UpdateProgress _
        "Creating Funneling...", _
        "-"

    SetCompileStage _
        "Funneling"

    UpdatefunnelingSheet

    If Err.Number <> 0 Then

        AddRunLog _
            "Error", _
            "UpdatefunnelingSheet failed: " & _
            Err.Description

        Err.Clear

    End If

    On Error GoTo 0

    FPTKTrace_Checkpoint _
        "After Funneling"

    UpdateProgress _
        "Resizing report tables...", _
        "-"

    SetCompileStage _
        "Report Tables"

    EnsureReportTables

    FPTKTrace_Checkpoint _
        "After EnsureReportTables"

    UpdateProgress _
        "Exporting output workbook...", _
        "-"

    SetCompileStage _
        "Export Output"

    ExportOutputWorkbook False

    FPTKTrace_Checkpoint _
        "After ExportOutputWorkbook"

    UpdateProgress _
        "Creating backup copy...", _
        "-"

    SetCompileStage _
        "Backup Output"

    CopyOutputToArchive

End Sub

Private Function TurboElapsedText( _
    ByVal startedAt As Double) As String

    Dim elapsedSeconds As Double

    elapsedSeconds = Timer - startedAt

    If elapsedSeconds < 0 Then
        elapsedSeconds = elapsedSeconds + 86400
    End If

    If elapsedSeconds < 60 Then

        TurboElapsedText = _
            Format$(elapsedSeconds, "0.0") & _
            " detik"

    Else

        TurboElapsedText = _
            Format$(Int(elapsedSeconds / 60), "0") & _
            " menit " & _
            Format$(elapsedSeconds Mod 60, "0") & _
            " detik"

    End If

End Function

' ============================================================
' COMPILE ONE FOLDER
' ============================================================
Private Sub CompileOneFolder( _
    ByVal folderPath As String)

    folderPath = EnsureTrailingSlash( _
        folderPath)

    AddRunLog _
        "Info", _
        "Compiling folder: " & folderPath

    UpdateProgress _
        "Compiling FPTK...", _
        folderPath

    FPTKTrace_Checkpoint _
        "Before Compile_FPTK | " & _
        folderPath

    TraceStageStart _
        "FPTK", _
        "Entering Compile_FPTK"

    Compile_FPTK folderPath

    TraceStageEnd _
        "FPTK", _
        "Compile_FPTK returned"

    FPTKTrace_Checkpoint _
        "After Compile_FPTK | " & _
        folderPath

    ' DB Kode Posisi dan DB Sourcing diproses sekali setelah
    ' seluruh folder selesai menjalankan FPTK dan Blacklist.

    ' selesai menjalankan FPTK dan Blacklist.

    UpdateProgress _
        "Compiling Blacklist Candidate...", _
        folderPath

    TraceStageStart _
        "Blacklist Candidate", _
        "Entering Compile_Blacklist_Candidate"

    Compile_Blacklist_Candidate folderPath

    TraceStageEnd _
        "Blacklist Candidate", _
        "Compile_Blacklist_Candidate returned"

    FPTKTrace_Checkpoint _
        "After Blacklist | " & _
        folderPath

    AddRunLog _
        "Info", _
        "Compile folder completed: " & _
        folderPath

    SetCompileStage ""

End Sub

' ============================================================
' COMPILE FOLDER LIST
' ============================================================
Private Function GetConfiguredCompileFolders( _
    ByVal rootFolder As String) As Collection

    Dim result As New Collection
    Dim seenPaths As Object

    Dim subFoldersText As String
    Dim normalizedList As String
    Dim subFolderItem As Variant

    Dim cleanName As String
    Dim resolvedPath As String

    Set seenPaths = CreateObject( _
        "Scripting.Dictionary")

    seenPaths.CompareMode = vbTextCompare

    rootFolder = CleanConfiguredFolderPath( _
        rootFolder)

    If Len(rootFolder) = 0 Then

        Set GetConfiguredCompileFolders = result
        Exit Function

    End If

    subFoldersText = CStr( _
        GetSettingValue("SubFolders"))

    normalizedList = _
        NormalizeSubFolderSettingText( _
            subFoldersText)

    If Len(normalizedList) = 0 Then

        If FolderExistsFSO(rootFolder) Then

            result.Add _
                EnsureTrailingSlash(rootFolder)

        End If

    Else

        For Each subFolderItem In _
            Split(normalizedList, ",")

            cleanName = _
                CleanSubFolderToken( _
                    CStr(subFolderItem))

            If Len(cleanName) > 0 Then

                resolvedPath = _
                    ResolveConfiguredSubFolder( _
                        rootFolder, _
                        cleanName)

                If Len(resolvedPath) > 0 Then

                    resolvedPath = _
                        EnsureTrailingSlash( _
                            resolvedPath)

                    If Not seenPaths.Exists( _
                        resolvedPath) Then

                        result.Add resolvedPath

                        seenPaths.Add _
                            resolvedPath, _
                            True

                        AddRunLog _
                            "Info", _
                            "Subfolder terbaca: " & _
                            resolvedPath

                    End If

                Else

                    AddRunLog _
                        "Warning", _
                        "Subfolder tidak ditemukan | " & _
                        "Root: " & rootFolder & _
                        " | Nama: [" & _
                        cleanName & "]"

                End If

            End If

        Next subFolderItem

    End If

    AddRunLog _
        "Info", _
        "Total compile folders: " & _
        CStr(result.Count)

    Set GetConfiguredCompileFolders = _
        result

End Function

Private Function NormalizeSubFolderSettingText( _
    ByVal rawText As String) As String

    Dim resultText As String

    resultText = rawText

    resultText = Replace$( _
        resultText, _
        Chr$(160), _
        " ")

    resultText = Replace$( _
        resultText, _
        vbCrLf, _
        ",")

    resultText = Replace$( _
        resultText, _
        vbCr, _
        ",")

    resultText = Replace$( _
        resultText, _
        vbLf, _
        ",")

    resultText = Replace$( _
        resultText, _
        ";", _
        ",")

    resultText = Replace$( _
        resultText, _
        "|", _
        ",")

    Do While InStr( _
        1, _
        resultText, _
        ",,", _
        vbBinaryCompare) > 0

        resultText = Replace$( _
            resultText, _
            ",,", _
            ",")

    Loop

    NormalizeSubFolderSettingText = _
        Trim$(resultText)

End Function

Private Function CleanSubFolderToken( _
    ByVal rawToken As String) As String

    Dim resultText As String

    resultText = rawToken

    resultText = Replace$( _
        resultText, _
        Chr$(160), _
        " ")

    resultText = Replace$( _
        resultText, _
        """", _
        vbNullString)

    resultText = Replace$( _
        resultText, _
        "'", _
        vbNullString)

    resultText = Trim$(resultText)

    Do While InStr( _
        1, _
        resultText, _
        "  ", _
        vbBinaryCompare) > 0

        resultText = Replace$( _
            resultText, _
            "  ", _
            " ")

    Loop

    CleanSubFolderToken = resultText

End Function

Private Function CleanConfiguredFolderPath( _
    ByVal rawPath As String) As String

    Dim resultPath As String

    resultPath = rawPath

    resultPath = Replace$( _
        resultPath, _
        Chr$(160), _
        " ")

    resultPath = Replace$( _
        resultPath, _
        """", _
        vbNullString)

    resultPath = Trim$(resultPath)

    CleanConfiguredFolderPath = _
        EnsureTrailingSlash( _
            resultPath)

End Function

Private Function ResolveConfiguredSubFolder( _
    ByVal rootFolder As String, _
    ByVal folderToken As String) As String

    Dim fso As Object
    Dim rootObject As Object
    Dim childFolder As Object

    Dim candidatePath As String
    Dim requestedKey As String
    Dim childKey As String

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    ' Token boleh berupa full path.
    If InStr( _
        1, _
        folderToken, _
        ":\", _
        vbTextCompare) > 0 Or _
       Left$(folderToken, 2) = "\\" Then

        candidatePath = _
            CleanConfiguredFolderPath( _
                folderToken)

    Else

        candidatePath = _
            EnsureTrailingSlash(rootFolder) & _
            folderToken

    End If

    If fso.FolderExists(candidatePath) Then

        ResolveConfiguredSubFolder = _
            fso.GetFolder( _
                candidatePath).Path

        Exit Function

    End If

    ' Fallback: cocokkan nama child folder secara langsung.
    ' Berguna bila ada spasi tersembunyi atau beda kapitalisasi.
    If Not fso.FolderExists(rootFolder) Then
        Exit Function
    End If

    requestedKey = _
        NormalizeFolderNameForMatch( _
            folderToken)

    Set rootObject = _
        fso.GetFolder(rootFolder)

    For Each childFolder In _
        rootObject.SubFolders

        childKey = _
            NormalizeFolderNameForMatch( _
                childFolder.Name)

        If childKey = requestedKey Then

            ResolveConfiguredSubFolder = _
                childFolder.Path

            Exit Function

        End If

    Next childFolder

End Function

Private Function NormalizeFolderNameForMatch( _
    ByVal folderName As String) As String

    Dim resultText As String

    resultText = UCase$( _
        CleanSubFolderToken( _
            folderName))

    resultText = Replace$( _
        resultText, _
        " ", _
        vbNullString)

    resultText = Replace$( _
        resultText, _
        "_", _
        vbNullString)

    resultText = Replace$( _
        resultText, _
        "-", _
        vbNullString)

    NormalizeFolderNameForMatch = _
        resultText

End Function

Private Function FolderExistsFSO( _
    ByVal folderPath As String) As Boolean

    Dim fso As Object

    On Error GoTo SafeExit

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    FolderExistsFSO = _
        fso.FolderExists( _
            folderPath)

SafeExit:

    Set fso = Nothing

End Function

Public Sub TestConfiguredSubFolders()

    Dim rootFolder As String
    Dim folders As Collection
    Dim folderItem As Variant

    Dim resultText As String

    rootFolder = CStr( _
        GetSettingValue( _
            "InputFolderPath"))

    Set folders = _
        GetConfiguredCompileFolders( _
            rootFolder)

    resultText = _
        "Root:" & vbCrLf & _
        CleanConfiguredFolderPath( _
            rootFolder) & _
        vbCrLf & vbCrLf & _
        "Folder terbaca: " & _
        CStr(folders.Count)

    For Each folderItem In folders

        resultText = resultText & _
            vbCrLf & _
            "- " & _
            CStr(folderItem)

    Next folderItem

    MsgBox _
        resultText, _
        vbInformation, _
        "Test SubFolders"

End Sub


Private Function CountFilesInFolders( _
    ByVal compileFolders As Collection) As Long

    Dim folderItem As Variant
    Dim folderPath As String
    Dim folderFileCount As Long
    Dim totalFileCount As Long

    For Each folderItem In compileFolders

        folderPath = CStr(folderItem)

        ' Hitung langsung dari folder final.
        ' Jangan panggil CountInputFiles dari m16 karena versi lama
        ' dapat menambahkan setting SubFolders untuk kedua kalinya.
        folderFileCount = _
            CountExcelFilesInFinalFolder( _
                folderPath)

        totalFileCount = _
            totalFileCount + _
            folderFileCount

        AddRunLog _
            "Info", _
            "Input scan | " & _
            folderPath & _
            " | Files: " & _
            CStr(folderFileCount)

    Next folderItem

    CountFilesInFolders = _
        totalFileCount

End Function

Private Function CountExcelFilesInFinalFolder( _
    ByVal folderPath As String) As Long

    Dim fso As Object
    Dim folderObject As Object
    Dim fileObject As Object

    Dim extensionText As String
    Dim fileCount As Long

    On Error GoTo SafeExit

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        GoTo SafeExit
    End If

    Set folderObject = _
        fso.GetFolder(folderPath)

    For Each fileObject In _
        folderObject.files

        If Left$( _
            CStr(fileObject.Name), _
            2) <> "~$" Then

            extensionText = LCase$( _
                fso.GetExtensionName( _
                    CStr(fileObject.Name)))

            Select Case extensionText

                Case "xlsx", "xlsm"

                    fileCount = _
                        fileCount + 1

            End Select

        End If

    Next fileObject

SafeExit:

    CountExcelFilesInFinalFolder = _
        fileCount

    Set fileObject = Nothing
    Set folderObject = Nothing
    Set fso = Nothing

End Function

Public Sub DiagnoseConfiguredInputFiles()

    Dim rootFolder As String
    Dim folders As Collection
    Dim folderItem As Variant

    Dim folderPath As String
    Dim folderFileCount As Long
    Dim totalFileCount As Long
    Dim resultText As String

    rootFolder = CStr( _
        GetSettingValue( _
            "InputFolderPath"))

    Set folders = _
        GetConfiguredCompileFolders( _
            rootFolder)

    resultText = _
        "Root:" & vbCrLf & _
        CleanConfiguredFolderPath( _
            rootFolder) & _
        vbCrLf & vbCrLf & _
        "Folder terbaca: " & _
        CStr(folders.Count) & _
        vbCrLf & vbCrLf

    For Each folderItem In folders

        folderPath = CStr(folderItem)

        folderFileCount = _
            CountExcelFilesInFinalFolder( _
                folderPath)

        totalFileCount = _
            totalFileCount + _
            folderFileCount

        resultText = resultText & _
            CStr(folderFileCount) & _
            " file | " & _
            folderPath & _
            vbCrLf

    Next folderItem

    resultText = resultText & _
        vbCrLf & _
        "TOTAL FILE XLSX/XLSM: " & _
        CStr(totalFileCount)

    MsgBox _
        resultText, _
        IIf( _
            totalFileCount > 0, _
            vbInformation, _
            vbExclamation), _
        "Diagnose Input Files"

End Sub


Private Function EnsureTrailingSlash( _
    ByVal folderPath As String) As String

    folderPath = Trim$(folderPath)

    If Len(folderPath) = 0 Then
        Exit Function
    End If

    If Right$(folderPath, 1) <> "\" Then

        folderPath = _
            folderPath & "\"

    End If

    EnsureTrailingSlash = folderPath

End Function

' ============================================================
' COMPILE STATE
' ============================================================
Private Sub SetCompileStage( _
    ByVal stageName As String)

    mCurrentCompileStage = _
        stageName

    gSchemaCurrentStage = _
        stageName

End Sub

Private Sub RestoreCompilerApplicationState( _
    ByVal stateWasCaptured As Boolean, _
    ByVal oldCalculation As XlCalculation, _
    ByVal oldAskToUpdateLinks As Boolean, _
    ByVal oldScreenUpdating As Boolean, _
    ByVal oldDisplayAlerts As Boolean, _
    ByVal oldEnableEvents As Boolean)

    On Error Resume Next

    If stateWasCaptured Then

        Application.Calculation = _
            oldCalculation

        Application.AskToUpdateLinks = _
            oldAskToUpdateLinks

        Application.ScreenUpdating = _
            oldScreenUpdating

        Application.DisplayAlerts = _
            oldDisplayAlerts

        Application.EnableEvents = _
            oldEnableEvents

    Else

        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True

    End If

    Application.StatusBar = False

    On Error GoTo 0

End Sub













