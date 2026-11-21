Attribute VB_Name = "m11_DBSourcingCompiler"
Option Explicit

Private Const TARGET_SHEET_NAME As String = "DB Sourcing"
Private Const TARGET_HEADER_ROW As Long = 1
Private Const FIRST_TARGET_DATA_ROW As Long = 2
Private Const BATCH_SIZE As Long = 500

' Update existing diproses per blok agar cepat tetapi tetap hemat RAM.
Private Const UPDATE_CHUNK_ROWS As Long = 1000

Private Const LOG_SHEET_NAME As String = "Log DB Sourcing"
Private Const TEMP_SUBFOLDER_NAME As String = "DBSourcingCompilerTemp"



' ============================================================
' CHECK HEADER TARGET - READ ONLY
' Tidak memindah, menghapus, atau mengganti header.
' ============================================================
Public Sub CheckDBSourcingHeaderOrderSafe()

    Dim ws As Worksheet
    Dim expectedHeaders As Variant
    Dim expectedIndex As Long
    Dim actualHeader As String
    Dim issueText As String
    Dim issueCount As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets( _
        TARGET_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then

        MsgBox _
            "Sheet DB Sourcing tidak ditemukan.", _
            vbExclamation, _
            "Check Header DB Sourcing"

        Exit Sub

    End If

    expectedHeaders = _
        StandardDBSourcingHeaders_DB()

    For expectedIndex = LBound(expectedHeaders) _
        To UBound(expectedHeaders)

        actualHeader = Trim$(CStr( _
            ws.Cells( _
                TARGET_HEADER_ROW, _
                expectedIndex + 1).Value2))

        If CanonicalHeader_DB(actualHeader) <> _
           CanonicalHeader_DB( _
                expectedHeaders(expectedIndex)) Then

            issueCount = issueCount + 1

            If issueCount <= 15 Then

                issueText = issueText & _
                    vbCrLf & _
                    "Kolom " & _
                    CStr(expectedIndex + 1) & _
                    ": expected '" & _
                    CStr(expectedHeaders(expectedIndex)) & _
                    "', actual '" & _
                    actualHeader & "'"

            End If

        End If

    Next expectedIndex

    If issueCount = 0 Then

        MsgBox _
            "74 header DB Sourcing masih sesuai urutan." & _
            vbCrLf & _
            "Termasuk Created At, Last Updated At, " & _
            "Last Compile Action, dan Source File.", _
            vbInformation, _
            "Check Header DB Sourcing"

    Else

        MsgBox _
            CStr(issueCount) & _
            " posisi header utama berbeda." & _
            vbCrLf & _
            "Macro ini hanya mengecek dan tidak mengubah data." & _
            issueText, _
            vbExclamation, _
            "Check Header DB Sourcing"

    End If

End Sub

Public Sub Compile_DB_Sourcing( _
    ByVal folderPath As String, _
    Optional ByVal showMessage As Boolean = True)

    Dim folderPaths As New Collection

    folderPaths.Add folderPath

    Compile_DB_Sourcing_MultiFolders _
        folderPaths, _
        showMessage

End Sub

Public Sub Compile_DB_Sourcing_MultiFolders( _
    ByVal folderPaths As Collection, _
    Optional ByVal showMessage As Boolean = True)

    Dim wsTarget As Worksheet
    Dim sourceFiles As Collection
    Dim sourceFile As Variant
    Dim existingIndex As Object
    Dim observedSourceHeaders As Object

    Dim importedCount As Long
    Dim updatedCount As Long
    Dim failedFileCount As Long
    Dim processedFileCount As Long
    Dim missingColumnTotal As Long
    Dim emptyRowTotal As Long
    Dim instructionRowTotal As Long

    Dim existingTableName As String
    Dim existingTableStyle As String

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldAskToUpdateLinks As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldAutomationSecurity As Long

    Dim outputFilePath As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim startedAt As Double

    startedAt = Timer

    On Error GoTo FatalError

    If folderPaths Is Nothing Then

        Err.Raise _
            vbObjectError + 5001, _
            "Compile_DB_Sourcing_MultiFolders", _
            "Daftar folder DB Sourcing belum tersedia."

    End If

    If folderPaths.Count = 0 Then

        Err.Raise _
            vbObjectError + 5002, _
            "Compile_DB_Sourcing_MultiFolders", _
            "Tidak ada folder DB Sourcing yang akan diproses."

    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldAskToUpdateLinks = Application.AskToUpdateLinks
    oldCalculation = Application.Calculation
    oldAutomationSecurity = Application.AutomationSecurity

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3

    HeartbeatProgress _
        "RUNNING", _
        CStr(folderPaths.Count) & " folders", _
        "DB Sourcing - Prepare", _
        "Preparing target once for all folders"

    Set wsTarget = GetOrCreateWorksheet_DB( _
        ThisWorkbook, _
        TARGET_SHEET_NAME)

    EnsureDBSourcingHeaderRow1Layout wsTarget

    EnsureCoreTargetHeaders_DB wsTarget
    RemoveDuplicateHeaderRows_DB wsTarget

    CaptureExistingTableInfo_DB _
        wsTarget, _
        existingTableName, _
        existingTableStyle

    HeartbeatProgress _
        "RUNNING", _
        CStr(folderPaths.Count) & " folders", _
        "DB Sourcing - Index Existing", _
        "Building key index once"

    Set existingIndex = _
        BuildExistingRecordIndex_DB( _
            wsTarget)

    outputFilePath = ReadSettingValue_DB( _
        "OutputFilePath")

    PrepareDBCompileLog_DB

    Set observedSourceHeaders = _
        BuildStandardObservedHeaders_DB()

    Set sourceFiles = _
        CollectSourceFilesFromFolders_DB( _
            folderPaths, _
            outputFilePath)

    If sourceFiles.Count = 0 Then

        Err.Raise _
            vbObjectError + 5004, _
            "Compile_DB_Sourcing_MultiFolders", _
            "Tidak ada file Excel DB Sourcing pada seluruh " & _
            "subfolder yang dikonfigurasi."

    End If

    HeartbeatProgress _
        "RUNNING", _
        CStr(folderPaths.Count) & " folders", _
        "DB Sourcing - Source Files", _
        CStr(sourceFiles.Count) & _
        " unique files found"

    For Each sourceFile In sourceFiles

        HeartbeatProgress _
            "RUNNING", _
            FileNameOnly_DB(CStr(sourceFile)), _
            "DB Sourcing - Source File", _
            CStr(processedFileCount + 1) & _
            "/" & _
            CStr(sourceFiles.Count)

        ProcessOneSourceFile_DB _
            CStr(sourceFile), _
            wsTarget, _
            existingIndex, _
            observedSourceHeaders, _
            importedCount, _
            updatedCount, _
            failedFileCount, _
            missingColumnTotal, _
            emptyRowTotal, _
            instructionRowTotal

        processedFileCount = _
            processedFileCount + 1

        DoEvents

    Next sourceFile

    HeartbeatProgress _
        "RUNNING", _
        "DB Sourcing", _
        "DB Sourcing - Clean Headers", _
        "Removing legacy headers not found in current sources"

    CleanTargetHeadersAgainstObserved_DB _
        wsTarget, _
        observedSourceHeaders

    HeartbeatProgress _
        "RUNNING", _
        "DB Sourcing", _
        "DB Sourcing - Finalize", _
        "Renumbering once"

    RenumberTargetRows_DB wsTarget

    HeartbeatProgress _
        "RUNNING", _
        "DB Sourcing", _
        "DB Sourcing - Finalize", _
        "Applying date formats once"

    ApplyDynamicDateFormats_DB wsTarget

    HeartbeatProgress _
        "RUNNING", _
        "DB Sourcing", _
        "DB Sourcing - Finalize", _
        "Resizing table once"

    CreateOrResizeTargetTable_DB _
        wsTarget, _
        GetLastTargetDataRow_DB(wsTarget), _
        LastUsedColumnInRow_DB( _
            wsTarget, _
            TARGET_HEADER_ROW), _
        existingTableName, _
        existingTableStyle

    HeartbeatProgress _
        "RUNNING", _
        "DB Sourcing", _
        "DB Sourcing - Finalize", _
        "Formatting once"

    FormatTargetSheetDynamic_DB wsTarget

    GoTo ShowResult

SafeExit:

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    Application.Calculation = oldCalculation
    Application.AutomationSecurity = oldAutomationSecurity
    Application.StatusBar = False
    Application.CutCopyMode = False

    Set observedSourceHeaders = Nothing
    Set existingIndex = Nothing
    Set sourceFiles = Nothing
    Set wsTarget = Nothing

    Exit Sub

ShowResult:

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    Application.Calculation = oldCalculation
    Application.AutomationSecurity = oldAutomationSecurity
    Application.StatusBar = False
    Application.CutCopyMode = False

    AddRunLog _
        IIf(failedFileCount > 0, "Warning", "Summary"), _
        "DB Sourcing ONE PASS | " & _
        "Folders: " & CStr(folderPaths.Count) & _
        " | Files: " & CStr(processedFileCount) & _
        " | New: " & CStr(importedCount) & _
        " | Update: " & CStr(updatedCount) & _
        " | Failed: " & CStr(failedFileCount) & _
        " | Duration: " & FormatElapsed_DB(startedAt)

    If showMessage Then

        MsgBox _
            "DB Sourcing selesai." & vbCrLf & vbCrLf & _
            "Folder diproses: " & _
                Format$(folderPaths.Count, "#,##0") & vbCrLf & _
            "File diproses: " & _
                Format$(processedFileCount, "#,##0") & vbCrLf & _
            "Baris baru: " & _
                Format$(importedCount, "#,##0") & vbCrLf & _
            "Baris diupdate: " & _
                Format$(updatedCount, "#,##0") & vbCrLf & _
            "File gagal: " & _
                Format$(failedFileCount, "#,##0") & vbCrLf & _
            "Durasi: " & _
                FormatElapsed_DB(startedAt), _
            IIf(failedFileCount > 0, vbExclamation, vbInformation), _
            "Compile DB Sourcing"

    End If

    Exit Sub

FatalError:

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    Application.Calculation = oldCalculation
    Application.AutomationSecurity = oldAutomationSecurity
    Application.StatusBar = False
    Application.CutCopyMode = False

    AddRunLog _
        "Error", _
        "Compile DB Sourcing ONE PASS gagal | " & _
        "Error " & CStr(errorNumber) & _
        ": " & errorDescription

    On Error GoTo 0

    If showMessage Then

        MsgBox _
            "Compile DB Sourcing gagal." & vbCrLf & vbCrLf & _
            "Error " & CStr(errorNumber) & _
            ": " & errorDescription, _
            vbCritical, _
            "Compile DB Sourcing"

    End If

End Sub


' ============================================================
' PROCESS SATU FILE
' ============================================================
Private Sub ProcessOneSourceFile_DB( _
    ByVal fullPath As String, _
    ByVal wsTarget As Worksheet, _
    ByVal existingIndex As Object, _
    ByVal observedSourceHeaders As Object, _
    ByRef importedCount As Long, _
    ByRef updatedCount As Long, _
    ByRef failedFileCount As Long, _
    ByRef missingColumnTotal As Long, _
    ByRef emptyRowTotal As Long, _
    ByRef instructionRowTotal As Long)

    Dim wbSource As Workbook
    Dim wsSource As Worksheet

    Dim sourceHeaderMap As Object
    Dim targetHeaderMap As Object
    Dim pendingIndex As Object
    Dim queuedExistingUpdates As Object

    Dim sourceCanonicalByColumn As Variant
    Dim targetCanonicalByColumn As Variant

    Dim stagedPath As String
    Dim sourceWasAlreadyOpen As Boolean
    Dim openStage As String

    Dim headerRow As Long
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim sourceLastColumn As Long
    Dim targetLastColumn As Long

    Dim sourceData As Variant
    Dim sourceRowIndex As Long
    Dim targetRow As Long
    Dim recordKey As String
    Dim directOpenError As String

    Dim importedBefore As Long
    Dim updatedBefore As Long
    Dim emptyBefore As Long
    Dim instructionBefore As Long

    Dim currentErrorNumber As Long
    Dim currentErrorDescription As String

    Dim insertBuffer() As Variant
    Dim bufferKeys() As String
    Dim bufferCount As Long
    Dim nextWriteRow As Long
    Dim nextNumber As Long

    Dim missingColumns As Long

    importedBefore = importedCount
    updatedBefore = updatedCount
    emptyBefore = emptyRowTotal
    instructionBefore = instructionRowTotal

    On Error GoTo FileError

    openStage = _
        "mencari workbook yang sudah terbuka"

    Set wbSource = GetOpenWorkbookByFullPath_DB( _
        fullPath)

    If Not wbSource Is Nothing Then

        sourceWasAlreadyOpen = True

        openStage = _
            "menggunakan workbook yang sudah terbuka"

    Else

        ' Coba buka file asli dahulu agar tidak selalu menyalin
        ' workbook besar ke TEMP.
        openStage = "membuka file langsung"

        HeartbeatProgress _
            "RUNNING", _
            FileNameOnly_DB(fullPath), _
            "DB Sourcing - Open File", _
            openStage

        Set wbSource = _
            TryOpenSourceWorkbookDirect_DB( _
                fullPath, _
                directOpenError)

        If wbSource Is Nothing Then

            openStage = "menyalin file ke TEMP"

            HeartbeatProgress _
                "RUNNING", _
                FileNameOnly_DB(fullPath), _
                "DB Sourcing - Open File", _
                openStage & _
                IIf( _
                    Len(directOpenError) > 0, _
                    " | Direct open: " & _
                    directOpenError, _
                    vbNullString)

            stagedPath = _
                StageSourceFileToTemp_DB( _
                    fullPath)

            openStage = "membuka salinan TEMP"

            HeartbeatProgress _
                "RUNNING", _
                FileNameOnly_DB(fullPath), _
                "DB Sourcing - Open File", _
                openStage

            Set wbSource = Application.Workbooks.Open( _
                fileName:=stagedPath, _
                UpdateLinks:=0, _
                ReadOnly:=True, _
                AddToMru:=False, _
                IgnoreReadOnlyRecommended:=True, _
                Notify:=False)

        End If

    End If

    openStage = "mencari sheet sumber"

    Set wsSource = GetSourceSheet_DB(wbSource)

    If wsSource Is Nothing Then

        Err.Raise _
            vbObjectError + 5101, _
            "ProcessOneSourceFile_DB", _
            "Sheet DB Sourcing tidak ditemukan."

    End If

    headerRow = FindBestHeaderRow_DB(wsSource)

    If headerRow = 0 Then

        Err.Raise _
            vbObjectError + 5102, _
            "ProcessOneSourceFile_DB", _
            "Header DB Sourcing tidak ditemukan."

    End If

    sourceLastColumn = LastUsedColumnInRow_DB( _
        wsSource, _
        headerRow)

    Set sourceHeaderMap = BuildHeaderMap_DB( _
        wsSource, _
        headerRow, _
        sourceLastColumn)

    RegisterObservedSourceHeaders_DB _
        observedSourceHeaders, _
        sourceHeaderMap

    EnsureTargetHeadersFromSource_DB _
        wsTarget, _
        wsSource, _
        headerRow, _
        sourceLastColumn

    Set targetHeaderMap = BuildHeaderMap_DB( _
        wsTarget, _
        TARGET_HEADER_ROW, _
        LastUsedColumnInRow_DB( _
            wsTarget, _
            TARGET_HEADER_ROW))

    sourceCanonicalByColumn = _
        BuildCanonicalArray_DB( _
            wsSource, _
            headerRow, _
            sourceLastColumn)

    targetLastColumn = LastUsedColumnInRow_DB( _
        wsTarget, _
        TARGET_HEADER_ROW)

    targetCanonicalByColumn = _
        BuildCanonicalArray_DB( _
            wsTarget, _
            TARGET_HEADER_ROW, _
            targetLastColumn)

    firstDataRow = headerRow + 1

    lastDataRow = GetLastSourceDataRow_DB( _
        wsSource, _
        headerRow, _
        sourceHeaderMap)

    If lastDataRow < firstDataRow Then
        GoTo SafeClose
    End If

    HeartbeatProgress _
        "RUNNING", _
        FileNameOnly_DB(fullPath), _
        "DB Sourcing - Read Source", _
        "Rows " & _
        CStr(firstDataRow) & "-" & _
        CStr(lastDataRow) & _
        " | Columns " & _
        CStr(sourceLastColumn)

    sourceData = wsSource.Range( _
        wsSource.Cells( _
            firstDataRow, _
            1), _
        wsSource.Cells( _
            lastDataRow, _
            sourceLastColumn)).Value2

    If Not IsArray(sourceData) Then
        GoTo SafeClose
    End If

    missingColumns = CountMissingColumns_DB( _
        targetHeaderMap, _
        sourceHeaderMap)

    missingColumnTotal = _
        missingColumnTotal + missingColumns

    nextWriteRow = GetNextTargetRow_DB( _
        wsTarget)

    nextNumber = GetNextSequenceNumber_DB( _
        wsTarget)

    ReDim insertBuffer( _
        1 To BATCH_SIZE, _
        1 To targetLastColumn)

    ReDim bufferKeys( _
        1 To BATCH_SIZE)

    Set pendingIndex = CreateObject( _
        "Scripting.Dictionary")

    pendingIndex.CompareMode = _
        vbTextCompare

    Set queuedExistingUpdates = CreateObject( _
        "Scripting.Dictionary")

    queuedExistingUpdates.CompareMode = _
        vbTextCompare

    For sourceRowIndex = 1 _
        To UBound(sourceData, 1)

        If sourceRowIndex = 1 Or _
           sourceRowIndex Mod 1000 = 0 Or _
           sourceRowIndex = _
                UBound(sourceData, 1) Then

            HeartbeatProgress _
                "RUNNING", _
                FileNameOnly_DB(fullPath), _
                "DB Sourcing - Process Rows", _
                CStr(sourceRowIndex) & "/" & _
                CStr(UBound(sourceData, 1)) & _
                " | New " & _
                CStr(importedCount - importedBefore) & _
                " | Update " & _
                CStr(updatedCount - updatedBefore) & _
                " | Queued " & _
                CStr(queuedExistingUpdates.Count)

        End If

        If ArrayRowIsEmpty_DB( _
            sourceData, _
            sourceRowIndex) Then

            emptyRowTotal = _
                emptyRowTotal + 1

        ElseIf ArrayRowIsInstruction_DB( _
            sourceData, _
            sourceRowIndex, _
            sourceCanonicalByColumn) Then

            instructionRowTotal = _
                instructionRowTotal + 1

        Else

            recordKey = BuildRecordKeyFromArray_DB( _
                sourceData, _
                sourceRowIndex, _
                sourceHeaderMap)

            If Len(recordKey) = 0 Then

                instructionRowTotal = _
                    instructionRowTotal + 1

            ElseIf pendingIndex.Exists( _
                recordKey) Then

                UpdateBufferRowFromSource_DB _
                    insertBuffer, _
                    CLng(pendingIndex(recordKey)), _
                    sourceData, _
                    sourceRowIndex, _
                    sourceHeaderMap, _
                    targetCanonicalByColumn, _
                    targetLastColumn, _
                    FileNameOnly_DB(fullPath)

                updatedCount = _
                    updatedCount + 1

            ElseIf existingIndex.Exists( _
                recordKey) Then

                targetRow = CLng( _
                    existingIndex(recordKey))

                ' Simpan hanya target row -> source row.
                ' Last occurrence wins bila key muncul berulang.
                queuedExistingUpdates( _
                    CStr(targetRow)) = _
                        sourceRowIndex

                updatedCount = _
                    updatedCount + 1

            Else

                bufferCount = _
                    bufferCount + 1

                FillNewBufferRow_DB _
                    insertBuffer, _
                    bufferCount, _
                    sourceData, _
                    sourceRowIndex, _
                    sourceHeaderMap, _
                    targetCanonicalByColumn, _
                    targetLastColumn, _
                    nextNumber, _
                    FileNameOnly_DB(fullPath)

                bufferKeys(bufferCount) = _
                    recordKey

                pendingIndex.Add _
                    recordKey, _
                    bufferCount

                nextNumber = _
                    nextNumber + 1

                importedCount = _
                    importedCount + 1

                If bufferCount = BATCH_SIZE Then

                    FlushInsertBuffer_DB _
                        wsTarget, _
                        nextWriteRow, _
                        insertBuffer, _
                        bufferKeys, _
                        bufferCount, _
                        targetLastColumn, _
                        existingIndex

                    nextWriteRow = _
                        nextWriteRow + _
                        bufferCount

                    bufferCount = 0

                    Set pendingIndex = _
                        CreateObject( _
                            "Scripting.Dictionary")

                    pendingIndex.CompareMode = _
                        vbTextCompare

                    ReDim insertBuffer( _
                        1 To BATCH_SIZE, _
                        1 To targetLastColumn)

                    ReDim bufferKeys( _
                        1 To BATCH_SIZE)

                End If

            End If

        End If

    Next sourceRowIndex

    If queuedExistingUpdates.Count > 0 Then

        ApplyQueuedExistingUpdatesChunked_DB _
            wsTarget, _
            queuedExistingUpdates, _
            sourceData, _
            sourceCanonicalByColumn, _
            targetHeaderMap, _
            targetLastColumn, _
            FileNameOnly_DB(fullPath)

    End If

    If bufferCount > 0 Then

        FlushInsertBuffer_DB _
            wsTarget, _
            nextWriteRow, _
            insertBuffer, _
            bufferKeys, _
            bufferCount, _
            targetLastColumn, _
            existingIndex

    End If

SafeClose:

    sourceData = Empty
    sourceCanonicalByColumn = Empty
    targetCanonicalByColumn = Empty

    On Error Resume Next

    Erase insertBuffer
    Erase bufferKeys

    Set sourceHeaderMap = Nothing
    Set targetHeaderMap = Nothing
    Set pendingIndex = Nothing
    Set queuedExistingUpdates = Nothing
    Set wsSource = Nothing

    On Error GoTo 0

    AppendDBCompileLog_DB _
        FileNameOnly_DB(fullPath), _
        "OK", _
        importedCount - importedBefore, _
        updatedCount - updatedBefore, _
        emptyRowTotal - emptyBefore, _
        instructionRowTotal - instructionBefore, _
        missingColumns, _
        vbNullString

    On Error Resume Next

    If Not wbSource Is Nothing Then

        If Not sourceWasAlreadyOpen Then

            wbSource.Close _
                SaveChanges:=False

        End If

    End If

    Set wbSource = Nothing

    If Len(stagedPath) > 0 Then
        DeleteTempFile_DB stagedPath
    End If

    On Error GoTo 0

    Exit Sub

FileError:

    currentErrorNumber = Err.Number
    currentErrorDescription = Err.Description

    sourceData = Empty
    sourceCanonicalByColumn = Empty
    targetCanonicalByColumn = Empty

    On Error Resume Next

    Erase insertBuffer
    Erase bufferKeys

    Set sourceHeaderMap = Nothing
    Set targetHeaderMap = Nothing
    Set pendingIndex = Nothing
    Set queuedExistingUpdates = Nothing
    Set wsSource = Nothing

    On Error GoTo 0

    failedFileCount = _
        failedFileCount + 1

    AppendDBCompileLog_DB _
        FileNameOnly_DB(fullPath), _
        "FAILED", _
        importedCount - importedBefore, _
        updatedCount - updatedBefore, _
        emptyRowTotal - emptyBefore, _
        instructionRowTotal - instructionBefore, _
        missingColumns, _
        "Tahap: " & openStage & _
        " | Error " & _
        CStr(currentErrorNumber) & _
        ": " & _
        currentErrorDescription

    On Error Resume Next

    If Not wbSource Is Nothing Then

        If Not sourceWasAlreadyOpen Then

            wbSource.Close _
                SaveChanges:=False

        End If

    End If

    Set wbSource = Nothing

    If Len(stagedPath) > 0 Then
        DeleteTempFile_DB stagedPath
    End If

    Err.Clear

    On Error GoTo 0

End Sub


Private Function TryOpenSourceWorkbookDirect_DB( _
    ByVal fullPath As String, _
    ByRef errorText As String) As Workbook

    Dim sourceWorkbook As Workbook

    On Error Resume Next

    Err.Clear

    Set sourceWorkbook = _
        Application.Workbooks.Open( _
            fileName:=fullPath, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            AddToMru:=False, _
            IgnoreReadOnlyRecommended:=True, _
            Notify:=False)

    If Err.Number <> 0 Then

        errorText = _
            Err.Description

        Err.Clear

    End If

    On Error GoTo 0

    Set TryOpenSourceWorkbookDirect_DB = _
        sourceWorkbook

End Function

Private Sub ApplyQueuedExistingUpdatesChunked_DB( _
    ByVal wsTarget As Worksheet, _
    ByVal queuedUpdates As Object, _
    ByRef sourceData As Variant, _
    ByVal sourceCanonicalByColumn As Variant, _
    ByVal targetHeaderMap As Object, _
    ByVal targetLastColumn As Long, _
    ByVal sourceFileName As String)

    Dim chunkBuckets As Object
    Dim rowBucket As Object

    Dim targetKey As Variant
    Dim chunkKey As Variant
    Dim rowKey As Variant

    Dim targetRow As Long
    Dim sourceRow As Long

    Dim chunkStartRow As Long
    Dim chunkEndRow As Long
    Dim lastTargetRow As Long
    Dim chunkRowCount As Long
    Dim bufferRow As Long

    Dim chunkData As Variant
    Dim chunkKeys As Variant
    Dim chunkIndex As Long

    If queuedUpdates Is Nothing Then Exit Sub
    If queuedUpdates.Count = 0 Then Exit Sub
    If targetLastColumn <= 0 Then Exit Sub

    Set chunkBuckets = CreateObject( _
        "Scripting.Dictionary")

    chunkBuckets.CompareMode = _
        vbTextCompare

    For Each targetKey In _
        queuedUpdates.keys

        targetRow = CLng(targetKey)

        chunkStartRow = _
            FIRST_TARGET_DATA_ROW + _
            ((targetRow - FIRST_TARGET_DATA_ROW) \ _
                UPDATE_CHUNK_ROWS) * _
            UPDATE_CHUNK_ROWS

        chunkKey = CStr(chunkStartRow)

        If Not chunkBuckets.Exists( _
            chunkKey) Then

            Set rowBucket = CreateObject( _
                "Scripting.Dictionary")

            rowBucket.CompareMode = _
                vbTextCompare

            chunkBuckets.Add _
                chunkKey, _
                rowBucket

        End If

        Set rowBucket = _
            chunkBuckets(chunkKey)

        rowBucket(CStr(targetRow)) = _
            CLng(queuedUpdates(targetKey))

    Next targetKey

    lastTargetRow = _
        GetLastTargetDataRow_DB( _
            wsTarget)

    chunkKeys = _
        chunkBuckets.keys

    SortNumericVariantArray_DB _
        chunkKeys

    For chunkIndex = LBound(chunkKeys) _
        To UBound(chunkKeys)

        chunkStartRow = CLng( _
            chunkKeys(chunkIndex))

        chunkEndRow = _
            chunkStartRow + _
            UPDATE_CHUNK_ROWS - 1

        If chunkEndRow > lastTargetRow Then

            chunkEndRow = _
                lastTargetRow

        End If

        chunkRowCount = _
            chunkEndRow - _
            chunkStartRow + 1

        HeartbeatProgress _
            "RUNNING", _
            sourceFileName, _
            "DB Sourcing - Batch Update Existing", _
            "Chunk " & _
            CStr(chunkIndex + 1) & "/" & _
            CStr(UBound(chunkKeys) + 1) & _
            " | Rows " & _
            CStr(chunkStartRow) & "-" & _
            CStr(chunkEndRow)

        chunkData = wsTarget.Cells( _
            chunkStartRow, _
            1).Resize( _
                chunkRowCount, _
                targetLastColumn).Value2

        Set rowBucket = _
            chunkBuckets( _
                CStr(chunkStartRow))

        For Each rowKey In _
            rowBucket.keys

            targetRow = CLng(rowKey)
            sourceRow = CLng( _
                rowBucket(rowKey))

            bufferRow = _
                targetRow - _
                chunkStartRow + 1

            UpdateChunkArrayRowFromSource_DB _
                chunkData, _
                bufferRow, _
                sourceData, _
                sourceRow, _
                sourceCanonicalByColumn, _
                targetHeaderMap, _
                targetLastColumn, _
                sourceFileName

        Next rowKey

        wsTarget.Cells( _
            chunkStartRow, _
            1).Resize( _
                chunkRowCount, _
                targetLastColumn).Value2 = _
                    chunkData

        chunkData = Empty
        Set rowBucket = Nothing

        DoEvents

    Next chunkIndex

    Set chunkBuckets = Nothing

End Sub

Private Sub UpdateChunkArrayRowFromSource_DB( _
    ByRef chunkData As Variant, _
    ByVal bufferRow As Long, _
    ByRef sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceCanonicalByColumn As Variant, _
    ByVal targetHeaderMap As Object, _
    ByVal targetLastColumn As Long, _
    ByVal sourceFileName As String)

    Dim sourceColumn As Long
    Dim canonicalHeader As String
    Dim targetColumn As Long
    Dim stampTime As Date

    stampTime = Now

    For sourceColumn = LBound( _
        sourceCanonicalByColumn) _
        To UBound(sourceCanonicalByColumn)

        canonicalHeader = CStr( _
            sourceCanonicalByColumn(sourceColumn))

        If Len(canonicalHeader) > 0 And _
           canonicalHeader <> "NO" And _
           canonicalHeader <> "CREATED AT" And _
           canonicalHeader <> "LAST UPDATED AT" And _
           canonicalHeader <> "LAST COMPILE ACTION" And _
           canonicalHeader <> "SOURCE FILE" Then

            If targetHeaderMap.Exists( _
                canonicalHeader) Then

                targetColumn = CLng( _
                    targetHeaderMap(canonicalHeader))

                If targetColumn <= _
                   targetLastColumn Then

                    chunkData( _
                        bufferRow, _
                        targetColumn) = _
                            SanitizeOutputValue_DB( _
                                sourceData( _
                                    sourceRow, _
                                    sourceColumn))

                End If

            End If

        End If

    Next sourceColumn

    DB_StampChunkAudit_DB _
        chunkData, _
        bufferRow, _
        targetHeaderMap, _
        targetLastColumn, _
        stampTime, _
        "UPDATE", _
        sourceFileName

End Sub

Private Sub DB_StampChunkAudit_DB( _
    ByRef rowData As Variant, _
    ByVal bufferRow As Long, _
    ByVal targetHeaderMap As Object, _
    ByVal targetLastColumn As Long, _
    ByVal stampTime As Date, _
    ByVal compileAction As String, _
    ByVal sourceFileName As String)

    Dim targetColumn As Long

    If targetHeaderMap.Exists( _
        "LAST UPDATED AT") Then

        targetColumn = CLng( _
            targetHeaderMap( _
                "LAST UPDATED AT"))

        If targetColumn <= targetLastColumn Then

            rowData( _
                bufferRow, _
                targetColumn) = _
                    stampTime

        End If

    End If

    If targetHeaderMap.Exists( _
        "LAST COMPILE ACTION") Then

        targetColumn = CLng( _
            targetHeaderMap( _
                "LAST COMPILE ACTION"))

        If targetColumn <= targetLastColumn Then

            rowData( _
                bufferRow, _
                targetColumn) = _
                    compileAction

        End If

    End If

    If targetHeaderMap.Exists( _
        "SOURCE FILE") Then

        targetColumn = CLng( _
            targetHeaderMap( _
                "SOURCE FILE"))

        If targetColumn <= targetLastColumn Then

            rowData( _
                bufferRow, _
                targetColumn) = _
                    sourceFileName

        End If

    End If

End Sub


Private Sub SortNumericVariantArray_DB( _
    ByRef values As Variant)

    Dim outerIndex As Long
    Dim innerIndex As Long
    Dim temporaryValue As Variant

    If Not IsArray(values) Then Exit Sub

    For outerIndex = LBound(values) _
        To UBound(values) - 1

        For innerIndex = outerIndex + 1 _
            To UBound(values)

            If CLng(values(innerIndex)) < _
               CLng(values(outerIndex)) Then

                temporaryValue = _
                    values(outerIndex)

                values(outerIndex) = _
                    values(innerIndex)

                values(innerIndex) = _
                    temporaryValue

            End If

        Next innerIndex

    Next outerIndex

End Sub

' ============================================================
' EXISTING INDEX
' ============================================================
Private Function BuildExistingRecordIndex_DB( _
    ByVal ws As Worksheet) As Object

    Dim result As Object
    Dim headerMap As Object

    Dim lastRow As Long
    Dim lastColumn As Long
    Dim rowCount As Long
    Dim rowIndex As Long

    Dim colKodeUnik As Long
    Dim colNama As Long
    Dim colEmail As Long
    Dim colNomorHP As Long
    Dim colSourcingDate As Long
    Dim colTanggalSourcing As Long
    Dim colTanggalSourcingFL As Long

    Dim arrKodeUnik As Variant
    Dim arrNama As Variant
    Dim arrEmail As Variant
    Dim arrNomorHP As Variant
    Dim arrSourcingDate As Variant
    Dim arrTanggalSourcing As Variant
    Dim arrTanggalSourcingFL As Variant

    Dim recordKey As String

    Set result = CreateObject( _
        "Scripting.Dictionary")

    result.CompareMode = vbTextCompare

    lastColumn = LastUsedColumnInRow_DB( _
        ws, _
        TARGET_HEADER_ROW)

    Set headerMap = BuildHeaderMap_DB( _
        ws, _
        TARGET_HEADER_ROW, _
        lastColumn)

    lastRow = GetLastTargetDataRow_DB(ws)

    If lastRow < FIRST_TARGET_DATA_ROW Or _
       lastColumn <= 0 Then

        Set BuildExistingRecordIndex_DB = result
        Exit Function

    End If

    rowCount = _
        lastRow - FIRST_TARGET_DATA_ROW + 1

    colKodeUnik = HeaderMapColumn_DB( _
        headerMap, _
        "KODE UNIK")

    colNama = HeaderMapColumn_DB( _
        headerMap, _
        "NAMA")

    colEmail = HeaderMapColumn_DB( _
        headerMap, _
        "EMAIL")

    colNomorHP = HeaderMapColumn_DB( _
        headerMap, _
        "NOMOR HP")

    colSourcingDate = HeaderMapColumn_DB( _
        headerMap, _
        "SOURCING DATE")

    colTanggalSourcing = HeaderMapColumn_DB( _
        headerMap, _
        "TANGGAL SOURCING")

    colTanggalSourcingFL = HeaderMapColumn_DB( _
        headerMap, _
        "TANGGAL SOURCING FREELANCE")

    arrKodeUnik = ReadSingleTargetColumn_DB( _
        ws, _
        colKodeUnik, _
        rowCount)

    arrNama = ReadSingleTargetColumn_DB( _
        ws, _
        colNama, _
        rowCount)

    arrEmail = ReadSingleTargetColumn_DB( _
        ws, _
        colEmail, _
        rowCount)

    arrNomorHP = ReadSingleTargetColumn_DB( _
        ws, _
        colNomorHP, _
        rowCount)

    arrSourcingDate = ReadSingleTargetColumn_DB( _
        ws, _
        colSourcingDate, _
        rowCount)

    arrTanggalSourcing = ReadSingleTargetColumn_DB( _
        ws, _
        colTanggalSourcing, _
        rowCount)

    arrTanggalSourcingFL = ReadSingleTargetColumn_DB( _
        ws, _
        colTanggalSourcingFL, _
        rowCount)

    For rowIndex = 1 To rowCount

        If rowIndex = 1 Or _
           rowIndex Mod 5000 = 0 Or _
           rowIndex = rowCount Then

            HeartbeatProgress _
                "RUNNING", _
                "DB Sourcing", _
                "DB Sourcing - Index Existing", _
                Format$(rowIndex, "#,##0") & _
                "/" & _
                Format$(rowCount, "#,##0") & _
                " | Key columns only"

        End If

        recordKey = _
            BuildRecordKeyFromKeyColumns_DB( _
                KeyColumnValue_DB( _
                    arrKodeUnik, _
                    rowIndex), _
                KeyColumnValue_DB( _
                    arrNama, _
                    rowIndex), _
                KeyColumnValue_DB( _
                    arrEmail, _
                    rowIndex), _
                KeyColumnValue_DB( _
                    arrNomorHP, _
                    rowIndex), _
                KeyColumnValue_DB( _
                    arrSourcingDate, _
                    rowIndex), _
                KeyColumnValue_DB( _
                    arrTanggalSourcing, _
                    rowIndex), _
                KeyColumnValue_DB( _
                    arrTanggalSourcingFL, _
                    rowIndex))

        If Len(recordKey) > 0 Then

            If Not result.Exists(recordKey) Then

                result.Add _
                    recordKey, _
                    FIRST_TARGET_DATA_ROW + _
                    rowIndex - 1

            End If

        End If

    Next rowIndex

    arrKodeUnik = Empty
    arrNama = Empty
    arrEmail = Empty
    arrNomorHP = Empty
    arrSourcingDate = Empty
    arrTanggalSourcing = Empty
    arrTanggalSourcingFL = Empty

    Set BuildExistingRecordIndex_DB = result

End Function

Private Function HeaderMapColumn_DB( _
    ByVal headerMap As Object, _
    ByVal canonicalHeader As String) As Long

    If headerMap Is Nothing Then Exit Function

    If headerMap.Exists( _
        canonicalHeader) Then

        HeaderMapColumn_DB = CLng( _
            headerMap(canonicalHeader))

    End If

End Function

Private Function ReadSingleTargetColumn_DB( _
    ByVal ws As Worksheet, _
    ByVal columnIndex As Long, _
    ByVal rowCount As Long) As Variant

    Dim resultData() As Variant

    If columnIndex <= 0 Or _
       rowCount <= 0 Then

        ReDim resultData( _
            1 To IIf(rowCount > 0, rowCount, 1), _
            1 To 1)

        ReadSingleTargetColumn_DB = _
            resultData

        Exit Function

    End If

    If rowCount = 1 Then

        ReDim resultData(1 To 1, 1 To 1)

        resultData(1, 1) = _
            ws.Cells( _
                FIRST_TARGET_DATA_ROW, _
                columnIndex).Value2

        ReadSingleTargetColumn_DB = _
            resultData

    Else

        ReadSingleTargetColumn_DB = _
            ws.Cells( _
                FIRST_TARGET_DATA_ROW, _
                columnIndex).Resize( _
                    rowCount, _
                    1).Value2

    End If

End Function

Private Function KeyColumnValue_DB( _
    ByRef columnData As Variant, _
    ByVal rowIndex As Long) As Variant

    On Error GoTo SafeExit

    If IsArray(columnData) Then

        KeyColumnValue_DB = _
            columnData( _
                rowIndex, _
                1)

    End If

SafeExit:

End Function

Private Function BuildRecordKeyFromKeyColumns_DB( _
    ByVal kodeUnikValue As Variant, _
    ByVal candidateNameValue As Variant, _
    ByVal emailValue As Variant, _
    ByVal phoneValue As Variant, _
    ByVal sourcingDateValue As Variant, _
    ByVal tanggalSourcingValue As Variant, _
    ByVal tanggalSourcingFLValue As Variant) As String

    Dim kodeUnik As String
    Dim candidateName As String
    Dim emailAddress As String
    Dim phoneNumber As String

    Dim sourcingDate As Variant
    Dim dateKey As String
    Dim identityKey As String

    kodeUnik = NormalizeKey_DB( _
        kodeUnikValue)

    candidateName = NormalizeKey_DB( _
        candidateNameValue)

    emailAddress = NormalizeEmailKey_DB( _
        emailValue)

    phoneNumber = NormalizePhoneKey_DB( _
        phoneValue)

    sourcingDate = sourcingDateValue

    If Not HasMeaningfulValue_DB( _
        sourcingDate) Then

        sourcingDate = _
            tanggalSourcingValue

    End If

    If Not HasMeaningfulValue_DB( _
        sourcingDate) Then

        sourcingDate = _
            tanggalSourcingFLValue

    End If

    dateKey = NormalizeDateKey_DB( _
        sourcingDate)

    If Len(kodeUnik) = 0 Or _
       Len(candidateName) = 0 Then

        Exit Function

    End If

    If Len(emailAddress) > 0 Then

        identityKey = _
            "EMAIL|" & emailAddress

    ElseIf Len(phoneNumber) > 0 Then

        identityKey = _
            "HP|" & phoneNumber

    ElseIf Len(dateKey) > 0 Then

        identityKey = _
            "DATE|" & dateKey

    Else

        identityKey = _
            "NAME|" & candidateName

    End If

    BuildRecordKeyFromKeyColumns_DB = _
        kodeUnik & "||" & _
        candidateName & "||" & _
        identityKey

End Function


Private Function BuildRecordKeyFromTargetArray_DB( _
    ByRef targetData As Variant, _
    ByVal rowIndex As Long, _
    ByVal headerMap As Object) As String

    Dim kodeUnik As String
    Dim candidateName As String
    Dim emailAddress As String
    Dim phoneNumber As String

    Dim sourcingDate As Variant
    Dim dateKey As String
    Dim identityKey As String

    kodeUnik = NormalizeKey_DB( _
        TargetArrayValue_DB( _
            targetData, _
            rowIndex, _
            headerMap, _
            "KODE UNIK"))

    candidateName = NormalizeKey_DB( _
        TargetArrayValue_DB( _
            targetData, _
            rowIndex, _
            headerMap, _
            "NAMA"))

    emailAddress = NormalizeEmailKey_DB( _
        TargetArrayValue_DB( _
            targetData, _
            rowIndex, _
            headerMap, _
            "EMAIL"))

    phoneNumber = NormalizePhoneKey_DB( _
        TargetArrayValue_DB( _
            targetData, _
            rowIndex, _
            headerMap, _
            "NOMOR HP"))

    sourcingDate = TargetArrayValue_DB( _
        targetData, _
        rowIndex, _
        headerMap, _
        "SOURCING DATE")

    If Not HasMeaningfulValue_DB( _
        sourcingDate) Then

        sourcingDate = TargetArrayValue_DB( _
            targetData, _
            rowIndex, _
            headerMap, _
            "TANGGAL SOURCING")

    End If

    If Not HasMeaningfulValue_DB( _
        sourcingDate) Then

        sourcingDate = TargetArrayValue_DB( _
            targetData, _
            rowIndex, _
            headerMap, _
            "TANGGAL SOURCING FREELANCE")

    End If

    dateKey = NormalizeDateKey_DB( _
        sourcingDate)

    If Len(kodeUnik) = 0 Or _
       Len(candidateName) = 0 Then

        Exit Function

    End If

    If Len(emailAddress) > 0 Then

        identityKey = _
            "EMAIL|" & emailAddress

    ElseIf Len(phoneNumber) > 0 Then

        identityKey = _
            "HP|" & phoneNumber

    ElseIf Len(dateKey) > 0 Then

        identityKey = _
            "DATE|" & dateKey

    Else

        identityKey = _
            "NAME|" & candidateName

    End If

    BuildRecordKeyFromTargetArray_DB = _
        kodeUnik & "||" & _
        candidateName & "||" & _
        identityKey

End Function

Private Function TargetArrayValue_DB( _
    ByRef targetData As Variant, _
    ByVal rowIndex As Long, _
    ByVal headerMap As Object, _
    ByVal canonicalHeader As String) As Variant

    Dim columnIndex As Long

    If headerMap Is Nothing Then Exit Function

    If headerMap.Exists( _
        canonicalHeader) Then

        columnIndex = CLng( _
            headerMap(canonicalHeader))

        If columnIndex >= LBound( _
            targetData, 2) And _
           columnIndex <= UBound( _
            targetData, 2) Then

            TargetArrayValue_DB = _
                targetData( _
                    rowIndex, _
                    columnIndex)

        End If

    End If

End Function


' ============================================================
' RECORD KEY
' ============================================================
Private Function BuildRecordKeyFromArray_DB( _
    ByRef sourceData As Variant, _
    ByVal rowIndex As Long, _
    ByVal sourceHeaderMap As Object) As String

    Dim kodeUnik As String
    Dim candidateName As String
    Dim emailAddress As String
    Dim phoneNumber As String
    Dim sourcingDate As Variant
    Dim dateKey As String
    Dim identityKey As String

    kodeUnik = NormalizeKey_DB( _
        GetMappedValue_DB( _
            sourceData, rowIndex, _
            sourceHeaderMap, "KODE UNIK"))

    candidateName = NormalizeKey_DB( _
        GetMappedValue_DB( _
            sourceData, rowIndex, _
            sourceHeaderMap, "NAMA"))

    emailAddress = NormalizeEmailKey_DB( _
        GetMappedValue_DB( _
            sourceData, rowIndex, _
            sourceHeaderMap, "EMAIL"))

    phoneNumber = NormalizePhoneKey_DB( _
        GetMappedValue_DB( _
            sourceData, rowIndex, _
            sourceHeaderMap, "NOMOR HP"))

    sourcingDate = GetPrimarySourcingDateFromArray_DB( _
        sourceData, rowIndex, sourceHeaderMap)

    dateKey = NormalizeDateKey_DB(sourcingDate)

    If Len(kodeUnik) = 0 Or _
       Len(candidateName) = 0 Then

        Exit Function

    End If

    If Len(emailAddress) > 0 Then

        identityKey = "EMAIL|" & emailAddress

    ElseIf Len(phoneNumber) > 0 Then

        identityKey = "HP|" & phoneNumber

    ElseIf Len(dateKey) > 0 Then

        identityKey = "DATE|" & dateKey

    Else

        identityKey = "NAME|" & candidateName

    End If

    BuildRecordKeyFromArray_DB = _
        kodeUnik & "||" & _
        candidateName & "||" & _
        identityKey

End Function

Private Function BuildRecordKeyFromSheet_DB( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal headerMap As Object) As String

    Dim kodeUnik As String
    Dim candidateName As String
    Dim emailAddress As String
    Dim phoneNumber As String
    Dim sourcingDate As Variant
    Dim dateKey As String
    Dim identityKey As String

    kodeUnik = NormalizeKey_DB( _
        SheetValueByCanonical_DB( _
            ws, rowIndex, _
            headerMap, "KODE UNIK"))

    candidateName = NormalizeKey_DB( _
        SheetValueByCanonical_DB( _
            ws, rowIndex, _
            headerMap, "NAMA"))

    emailAddress = NormalizeEmailKey_DB( _
        SheetValueByCanonical_DB( _
            ws, rowIndex, _
            headerMap, "EMAIL"))

    phoneNumber = NormalizePhoneKey_DB( _
        SheetValueByCanonical_DB( _
            ws, rowIndex, _
            headerMap, "NOMOR HP"))

    sourcingDate = GetPrimarySourcingDateFromSheet_DB( _
        ws, rowIndex, headerMap)

    dateKey = NormalizeDateKey_DB(sourcingDate)

    If Len(kodeUnik) = 0 Or _
       Len(candidateName) = 0 Then

        Exit Function

    End If

    If Len(emailAddress) > 0 Then

        identityKey = "EMAIL|" & emailAddress

    ElseIf Len(phoneNumber) > 0 Then

        identityKey = "HP|" & phoneNumber

    ElseIf Len(dateKey) > 0 Then

        identityKey = "DATE|" & dateKey

    Else

        identityKey = "NAME|" & candidateName

    End If

    BuildRecordKeyFromSheet_DB = _
        kodeUnik & "||" & _
        candidateName & "||" & _
        identityKey

End Function

Private Function GetPrimarySourcingDateFromArray_DB( _
    ByRef sourceData As Variant, _
    ByVal rowIndex As Long, _
    ByVal sourceHeaderMap As Object) As Variant

    Dim valueFound As Variant

    valueFound = GetMappedValue_DB( _
        sourceData, rowIndex, _
        sourceHeaderMap, "SOURCING DATE")

    If HasMeaningfulValue_DB(valueFound) Then

        GetPrimarySourcingDateFromArray_DB = _
            valueFound

        Exit Function

    End If

    valueFound = GetMappedValue_DB( _
        sourceData, rowIndex, _
        sourceHeaderMap, _
        "TANGGAL SOURCING")

    If HasMeaningfulValue_DB(valueFound) Then

        GetPrimarySourcingDateFromArray_DB = _
            valueFound

        Exit Function

    End If

    valueFound = GetMappedValue_DB( _
        sourceData, rowIndex, _
        sourceHeaderMap, _
        "TANGGAL SOURCING FREELANCE")

    GetPrimarySourcingDateFromArray_DB = valueFound

End Function

Private Function GetPrimarySourcingDateFromSheet_DB( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal headerMap As Object) As Variant

    Dim valueFound As Variant

    valueFound = SheetValueByCanonical_DB( _
        ws, rowIndex, _
        headerMap, "SOURCING DATE")

    If HasMeaningfulValue_DB(valueFound) Then

        GetPrimarySourcingDateFromSheet_DB = _
            valueFound

        Exit Function

    End If

    valueFound = SheetValueByCanonical_DB( _
        ws, rowIndex, _
        headerMap, _
        "TANGGAL SOURCING")

    If HasMeaningfulValue_DB(valueFound) Then

        GetPrimarySourcingDateFromSheet_DB = _
            valueFound

        Exit Function

    End If

    valueFound = SheetValueByCanonical_DB( _
        ws, rowIndex, _
        headerMap, _
        "TANGGAL SOURCING FREELANCE")

    GetPrimarySourcingDateFromSheet_DB = valueFound

End Function

' ============================================================
' INSERT BUFFER
' ============================================================
Private Sub FillNewBufferRow_DB( _
    ByRef insertBuffer As Variant, _
    ByVal bufferRow As Long, _
    ByRef sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceHeaderMap As Object, _
    ByVal targetCanonicalByColumn As Variant, _
    ByVal targetLastColumn As Long, _
    ByVal sequenceNumber As Long, _
    ByVal sourceFileName As String)

    Dim targetColumn As Long
    Dim canonicalHeader As String
    Dim stampTime As Date

    stampTime = Now

    For targetColumn = 1 To targetLastColumn

        canonicalHeader = CStr( _
            targetCanonicalByColumn(targetColumn))

        Select Case canonicalHeader

            Case "NO"

                insertBuffer( _
                    bufferRow, targetColumn) = _
                        sequenceNumber

            Case "CREATED AT"

                insertBuffer( _
                    bufferRow, targetColumn) = _
                        stampTime

            Case "LAST UPDATED AT"

                insertBuffer( _
                    bufferRow, targetColumn) = _
                        Empty

            Case "LAST COMPILE ACTION"

                insertBuffer( _
                    bufferRow, targetColumn) = _
                        "INSERT"

            Case "SOURCE FILE"

                insertBuffer( _
                    bufferRow, targetColumn) = _
                        sourceFileName

            Case Else

                If sourceHeaderMap.Exists( _
                    canonicalHeader) Then

                    insertBuffer( _
                        bufferRow, targetColumn) = _
                            SanitizeOutputValue_DB( _
                                sourceData( _
                                    sourceRow, _
                                    CLng(sourceHeaderMap( _
                                        canonicalHeader))))

                Else

                    insertBuffer( _
                        bufferRow, targetColumn) = _
                            vbNullString

                End If

        End Select

    Next targetColumn

End Sub


Private Sub UpdateBufferRowFromSource_DB( _
    ByRef insertBuffer As Variant, _
    ByVal bufferRow As Long, _
    ByRef sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceHeaderMap As Object, _
    ByVal targetCanonicalByColumn As Variant, _
    ByVal targetLastColumn As Long, _
    ByVal sourceFileName As String)

    Dim targetColumn As Long
    Dim canonicalHeader As String
    Dim stampTime As Date

    stampTime = Now

    For targetColumn = 1 To targetLastColumn

        canonicalHeader = CStr( _
            targetCanonicalByColumn(targetColumn))

        Select Case canonicalHeader

            Case "NO", "CREATED AT"
                ' Preserve sequence and original Created At.

            Case "LAST UPDATED AT"

                insertBuffer( _
                    bufferRow, targetColumn) = _
                        stampTime

            Case "LAST COMPILE ACTION"

                insertBuffer( _
                    bufferRow, targetColumn) = _
                        "UPDATE"

            Case "SOURCE FILE"

                insertBuffer( _
                    bufferRow, targetColumn) = _
                        sourceFileName

            Case Else

                If sourceHeaderMap.Exists( _
                    canonicalHeader) Then

                    insertBuffer( _
                        bufferRow, targetColumn) = _
                            SanitizeOutputValue_DB( _
                                sourceData( _
                                    sourceRow, _
                                    CLng(sourceHeaderMap( _
                                        canonicalHeader))))

                End If

        End Select

    Next targetColumn

End Sub


Private Sub FlushInsertBuffer_DB( _
    ByVal wsTarget As Worksheet, _
    ByVal firstWriteRow As Long, _
    ByRef insertBuffer As Variant, _
    ByRef bufferKeys() As String, _
    ByVal bufferCount As Long, _
    ByVal targetLastColumn As Long, _
    ByVal existingIndex As Object)

    Dim finalData() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    If bufferCount <= 0 Then Exit Sub

    ReDim finalData( _
        1 To bufferCount, _
        1 To targetLastColumn)

    For rowIndex = 1 To bufferCount

        For columnIndex = 1 To targetLastColumn

            finalData( _
                rowIndex, columnIndex) = _
                insertBuffer( _
                    rowIndex, columnIndex)

        Next columnIndex

    Next rowIndex

    wsTarget.Cells( _
        firstWriteRow, 1).Resize( _
            bufferCount, _
            targetLastColumn).Value = finalData

    For rowIndex = 1 To bufferCount

        If Len(bufferKeys(rowIndex)) > 0 Then

            If Not existingIndex.Exists( _
                bufferKeys(rowIndex)) Then

                existingIndex.Add _
                    bufferKeys(rowIndex), _
                    firstWriteRow + rowIndex - 1

            End If

        End If

    Next rowIndex

End Sub

' ============================================================
' UPDATE EXISTING ROW
' ============================================================
Private Sub UpdateWorksheetRowFromSource_DB( _
    ByVal wsTarget As Worksheet, _
    ByVal targetRow As Long, _
    ByRef sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceCanonicalByColumn As Variant, _
    ByVal targetHeaderMap As Object)

    Dim sourceColumn As Long
    Dim canonicalHeader As String
    Dim targetColumn As Long
    Dim targetLastColumn As Long

    Dim rowBuffer As Variant

    targetLastColumn = LastUsedColumnInRow_DB( _
        wsTarget, _
        TARGET_HEADER_ROW)

    If targetLastColumn <= 0 Then Exit Sub

    rowBuffer = wsTarget.Cells( _
        targetRow, _
        1).Resize( _
            1, _
            targetLastColumn).Value2

    For sourceColumn = LBound( _
        sourceCanonicalByColumn) _
        To UBound(sourceCanonicalByColumn)

        canonicalHeader = CStr( _
            sourceCanonicalByColumn(sourceColumn))

        If Len(canonicalHeader) > 0 And _
           canonicalHeader <> "NO" Then

            If targetHeaderMap.Exists( _
                canonicalHeader) Then

                targetColumn = CLng( _
                    targetHeaderMap(canonicalHeader))

                If targetColumn <= targetLastColumn Then

                    rowBuffer(1, targetColumn) = _
                        SanitizeOutputValue_DB( _
                            sourceData( _
                                sourceRow, _
                                sourceColumn))

                End If

            End If

        End If

    Next sourceColumn

    wsTarget.Cells( _
        targetRow, _
        1).Resize( _
            1, _
            targetLastColumn).Value2 = _
                rowBuffer

End Sub


Private Sub EnsureDBSourcingHeaderRow1Layout( _
    ByVal ws As Worksheet)

    Dim row1IsHeader As Boolean
    Dim row2IsHeader As Boolean

    If ws Is Nothing Then Exit Sub

    row1IsHeader = _
        DBSourcing_RowLooksLikeHeader_DB( _
            ws, _
            1)

    If row1IsHeader Then
        Exit Sub
    End If

    row2IsHeader = _
        DBSourcing_RowLooksLikeHeader_DB( _
            ws, _
            2)

    If row2IsHeader Then

        ' Layout sebelumnya:
        ' row 1 = area kosong/judul
        ' row 2 = header
        ' row 3 = data pertama
        '
        ' Hapus row 1 agar:
        ' header pindah ke row 1
        ' data pertama pindah ke row 2
        ws.rows(1).Delete _
            Shift:=xlUp

        AddRunLog _
            "Info", _
            "DB Sourcing layout migrated: header row 2 -> row 1."

    End If

End Sub


Private Function DBSourcing_RowLooksLikeHeader_DB( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long) As Boolean

    Dim headerMap As Object
    Dim lastColumn As Long

    If ws Is Nothing Then Exit Function
    If rowIndex <= 0 Then Exit Function

    lastColumn = LastUsedColumnInRow_DB( _
        ws, _
        rowIndex)

    If lastColumn <= 0 Then Exit Function

    Set headerMap = BuildHeaderMap_DB( _
        ws, _
        rowIndex, _
        lastColumn)

    If headerMap Is Nothing Then
        Exit Function
    End If

    DBSourcing_RowLooksLikeHeader_DB = _
        headerMap.Exists("KODE UNIK") And _
        headerMap.Exists("NAMA")

End Function

' ============================================================
' TARGET HEADER DYNAMIC
' ============================================================
Private Sub EnsureCoreTargetHeaders_DB( _
    ByVal ws As Worksheet)

    Dim protectedHeaders As Variant
    Dim headerMap As Object
    Dim headerItem As Variant

    Dim headerText As String
    Dim canonicalHeader As String

    Dim existingColumn As Long
    Dim newColumn As Long

    If ws Is Nothing Then Exit Sub

    protectedHeaders = _
        StandardDBSourcingHeaders_DB()

    If LastNonEmptyHeaderColumn_DB( _
        ws) = 0 Then

        ws.Cells( _
            TARGET_HEADER_ROW, _
            1).Resize( _
                1, _
                UBound(protectedHeaders) - _
                LBound(protectedHeaders) + 1).Value = _
            OneDimensionalToRowArray_DB( _
                protectedHeaders)

        Exit Sub

    End If

    Set headerMap = _
        BuildHeaderMap_DB( _
            ws, _
            TARGET_HEADER_ROW, _
            LastUsedColumnInRow_DB( _
                ws, _
                TARGET_HEADER_ROW))

    For Each headerItem In protectedHeaders

        headerText = _
            Trim$( _
                CStr(headerItem))

        canonicalHeader = _
            CanonicalHeader_DB( _
                headerText)

        If Len(canonicalHeader) > 0 Then

            If Not headerMap.Exists( _
                canonicalHeader) Then

                existingColumn = _
                    FindEquivalentTargetColumn_DB( _
                        ws, _
                        headerText, _
                        canonicalHeader)

                If existingColumn > 0 Then

                    headerMap.Add _
                        canonicalHeader, _
                        existingColumn

                Else

                    newColumn = _
                        LastUsedColumnInRow_DB( _
                            ws, _
                            TARGET_HEADER_ROW) + 1

                    ws.Cells( _
                        TARGET_HEADER_ROW, _
                        newColumn).Value = _
                            headerText

                    headerMap.Add _
                        canonicalHeader, _
                        newColumn

                End If

            End If

        End If

    Next headerItem

End Sub
Private Sub EnsureTargetHeadersFromSource_DB( _
    ByVal wsTarget As Worksheet, _
    ByVal wsSource As Worksheet, _
    ByVal sourceHeaderRow As Long, _
    ByVal sourceLastColumn As Long)

    Dim targetHeaderMap As Object

    Dim sourceColumn As Long
    Dim sourceHeaderText As String
    Dim canonicalHeader As String

    Dim existingColumn As Long
    Dim newColumn As Long

    Set targetHeaderMap = _
        BuildHeaderMap_DB( _
            wsTarget, _
            TARGET_HEADER_ROW, _
            LastUsedColumnInRow_DB( _
                wsTarget, _
                TARGET_HEADER_ROW))

    For sourceColumn = 1 _
        To sourceLastColumn

        sourceHeaderText = _
            DB_SafeHeaderText( _
                wsSource.Cells( _
                    sourceHeaderRow, _
                    sourceColumn).Value2)

        If Len(sourceHeaderText) > 0 Then

            If Not HeaderTextLooksInvalid_DB( _
                sourceHeaderText) Then

                canonicalHeader = _
                    CanonicalHeader_DB( _
                        sourceHeaderText)

                If Len(canonicalHeader) > 0 Then

                    If Not targetHeaderMap.Exists( _
                        canonicalHeader) Then

                        existingColumn = _
                            FindEquivalentTargetColumn_DB( _
                                wsTarget, _
                                sourceHeaderText, _
                                canonicalHeader)

                        If existingColumn > 0 Then

                            targetHeaderMap.Add _
                                canonicalHeader, _
                                existingColumn

                        Else

                            newColumn = _
                                LastUsedColumnInRow_DB( _
                                    wsTarget, _
                                    TARGET_HEADER_ROW) + 1

                            wsTarget.Cells( _
                                TARGET_HEADER_ROW, _
                                newColumn).Value = _
                                    sourceHeaderText

                            targetHeaderMap.Add _
                                canonicalHeader, _
                                newColumn

                        End If

                    End If

                End If

            End If

        End If

    Next sourceColumn

End Sub


Private Function BuildStandardObservedHeaders_DB() As Object

    Dim result As Object
    Dim standardHeaders As Variant
    Dim headerItem As Variant
    Dim canonicalHeader As String

    Set result = _
        CreateObject( _
            "Scripting.Dictionary")

    result.CompareMode = _
        vbTextCompare

    standardHeaders = _
        StandardDBSourcingHeaders_DB()

    For Each headerItem In standardHeaders

        canonicalHeader = _
            CanonicalHeader_DB( _
                headerItem)

        If Len(canonicalHeader) > 0 Then

            If Not result.Exists( _
                canonicalHeader) Then

                result.Add _
                    canonicalHeader, _
                    True

            End If

        End If

    Next headerItem

    Set BuildStandardObservedHeaders_DB = _
        result

End Function

Private Sub RegisterObservedSourceHeaders_DB( _
    ByVal observedHeaders As Object, _
    ByVal sourceHeaderMap As Object)

    Dim keyItem As Variant

    If observedHeaders Is Nothing Then Exit Sub
    If sourceHeaderMap Is Nothing Then Exit Sub

    For Each keyItem In sourceHeaderMap.keys

        If Len(CStr(keyItem)) > 0 Then

            If Not observedHeaders.Exists( _
                CStr(keyItem)) Then

                observedHeaders.Add _
                    CStr(keyItem), _
                    True

            End If

        End If

    Next keyItem

End Sub

Private Sub CleanTargetHeadersAgainstObserved_DB( _
    ByVal wsTarget As Worksheet, _
    ByVal observedHeaders As Object)

    Dim earliestColumn As Object

    Dim lastColumn As Long
    Dim lastRow As Long
    Dim columnIndex As Long

    Dim headerText As String
    Dim canonicalHeader As String
    Dim keepColumn As Long

    Dim removedCount As Long
    Dim duplicateCount As Long
    Dim legacyCount As Long

    If wsTarget Is Nothing Then Exit Sub
    If observedHeaders Is Nothing Then Exit Sub

    Set earliestColumn = _
        CreateObject( _
            "Scripting.Dictionary")

    earliestColumn.CompareMode = _
        vbTextCompare

    lastColumn = _
        LastUsedColumnInRow_DB( _
            wsTarget, _
            TARGET_HEADER_ROW)

    lastRow = _
        GetLastTargetDataRow_DB( _
            wsTarget)

    If lastRow < TARGET_HEADER_ROW Then
        lastRow = TARGET_HEADER_ROW
    End If

    For columnIndex = 1 _
        To lastColumn

        headerText = _
            DB_SafeHeaderText( _
                wsTarget.Cells( _
                    TARGET_HEADER_ROW, _
                    columnIndex).Value2)

        canonicalHeader = _
            CanonicalHeader_DB( _
                headerText)

        If Len(canonicalHeader) > 0 Then

            If Not earliestColumn.Exists( _
                canonicalHeader) Then

                earliestColumn.Add _
                    canonicalHeader, _
                    columnIndex

            End If

        End If

    Next columnIndex

    For columnIndex = lastColumn _
        To 1 Step -1

        headerText = _
            DB_SafeHeaderText( _
                wsTarget.Cells( _
                    TARGET_HEADER_ROW, _
                    columnIndex).Value2)

        canonicalHeader = _
            CanonicalHeader_DB( _
                headerText)

        If Len(canonicalHeader) = 0 Then

            If columnIndex > 74 Then

                BackupTargetColumnBeforeDelete_DB _
                    wsTarget, _
                    columnIndex, _
                    lastRow, _
                    "Blank/invalid legacy header"

                wsTarget.Columns( _
                    columnIndex).Delete

                removedCount = removedCount + 1
                legacyCount = legacyCount + 1

            End If

        Else

            keepColumn = _
                CLng( _
                    earliestColumn( _
                        canonicalHeader))

            If columnIndex <> keepColumn Then

                MergeTargetColumnValues_DB _
                    wsTarget, _
                    columnIndex, _
                    keepColumn, _
                    lastRow

                BackupTargetColumnBeforeDelete_DB _
                    wsTarget, _
                    columnIndex, _
                    lastRow, _
                    "Duplicate alias of column " & _
                    CStr(keepColumn)

                wsTarget.Columns( _
                    columnIndex).Delete

                removedCount = removedCount + 1
                duplicateCount = duplicateCount + 1

            ElseIf Not observedHeaders.Exists( _
                canonicalHeader) Then

                BackupTargetColumnBeforeDelete_DB _
                    wsTarget, _
                    columnIndex, _
                    lastRow, _
                    "Not found in current DB Sourcing sources"

                wsTarget.Columns( _
                    columnIndex).Delete

                removedCount = removedCount + 1
                legacyCount = legacyCount + 1

            End If

        End If

    Next columnIndex

    If removedCount > 0 Then

        AddRunLog _
            "Info", _
            "DB Sourcing header cleanup | Removed: " & _
            CStr(removedCount) & _
            " | Duplicate aliases: " & _
            CStr(duplicateCount) & _
            " | Legacy not in source: " & _
            CStr(legacyCount) & _
            " | Backup sheet: Backup DB Sourcing"

    End If

    Set earliestColumn = Nothing

End Sub

Private Sub MergeTargetColumnValues_DB( _
    ByVal wsTarget As Worksheet, _
    ByVal sourceColumn As Long, _
    ByVal targetColumn As Long, _
    ByVal lastRow As Long)

    Dim sourceData As Variant
    Dim targetData As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim changed As Boolean

    If sourceColumn <= 0 Then Exit Sub
    If targetColumn <= 0 Then Exit Sub
    If sourceColumn = targetColumn Then Exit Sub
    If lastRow < FIRST_TARGET_DATA_ROW Then Exit Sub

    rowCount = _
        lastRow - _
        FIRST_TARGET_DATA_ROW + 1

    sourceData = _
        wsTarget.Cells( _
            FIRST_TARGET_DATA_ROW, _
            sourceColumn).Resize( _
                rowCount, _
                1).Value2

    targetData = _
        wsTarget.Cells( _
            FIRST_TARGET_DATA_ROW, _
            targetColumn).Resize( _
                rowCount, _
                1).Value2

    For rowIndex = 1 _
        To rowCount

        If Not HasMeaningfulValue_DB( _
            targetData( _
                rowIndex, _
                1)) Then

            If HasMeaningfulValue_DB( _
                sourceData( _
                    rowIndex, _
                    1)) Then

                targetData( _
                    rowIndex, _
                    1) = _
                    sourceData( _
                        rowIndex, _
                        1)

                changed = True

            End If

        End If

    Next rowIndex

    If changed Then

        wsTarget.Cells( _
            FIRST_TARGET_DATA_ROW, _
            targetColumn).Resize( _
                rowCount, _
                1).Value2 = _
                    targetData

    End If

End Sub

Private Sub BackupTargetColumnBeforeDelete_DB( _
    ByVal wsTarget As Worksheet, _
    ByVal sourceColumn As Long, _
    ByVal lastRow As Long, _
    ByVal reasonText As String)

    Dim wsBackup As Worksheet
    Dim nextBackupColumn As Long
    Dim copyRowCount As Long
    Dim columnData As Variant

    If wsTarget Is Nothing Then Exit Sub
    If sourceColumn <= 0 Then Exit Sub

    Set wsBackup = _
        GetOrCreateWorksheet_DB( _
            ThisWorkbook, _
            "Backup DB Sourcing")

    nextBackupColumn = _
        LastNonEmptyHeaderColumn_DB( _
            wsBackup) + 1

    If nextBackupColumn < 1 Then
        nextBackupColumn = 1
    End If

    wsBackup.Cells( _
        1, _
        nextBackupColumn).Value = _
            Now

    wsBackup.Cells( _
        2, _
        nextBackupColumn).Value = _
            reasonText

    wsBackup.Cells( _
        3, _
        nextBackupColumn).Value = _
            "Original column " & _
            CStr(sourceColumn)

    If lastRow < TARGET_HEADER_ROW Then
        lastRow = TARGET_HEADER_ROW
    End If

    copyRowCount = lastRow

    columnData = _
        wsTarget.Cells( _
            TARGET_HEADER_ROW, _
            sourceColumn).Resize( _
                copyRowCount, _
                1).Value2

    wsBackup.Cells( _
        4, _
        nextBackupColumn).Resize( _
            copyRowCount, _
            1).Value2 = _
                columnData

    wsBackup.Columns( _
        nextBackupColumn).columnWidth = 24

End Sub

Private Function FindEquivalentTargetColumn_DB( _
    ByVal wsTarget As Worksheet, _
    ByVal sourceHeaderText As String, _
    ByVal sourceCanonical As String) As Long

    Dim lastColumn As Long
    Dim columnIndex As Long

    Dim targetHeaderText As String
    Dim targetCanonical As String

    Dim sourceBaseKey As String
    Dim targetBaseKey As String

    If wsTarget Is Nothing Then Exit Function

    lastColumn = _
        LastUsedColumnInRow_DB( _
            wsTarget, _
            TARGET_HEADER_ROW)

    sourceBaseKey = _
        DB_HeaderInstructionBaseKey( _
            sourceHeaderText)

    For columnIndex = 1 _
        To lastColumn

        targetHeaderText = _
            DB_SafeHeaderText( _
                wsTarget.Cells( _
                    TARGET_HEADER_ROW, _
                    columnIndex).Value2)

        If Len(targetHeaderText) > 0 Then

            targetCanonical = _
                CanonicalHeader_DB( _
                    targetHeaderText)

            If StrComp( _
                targetCanonical, _
                sourceCanonical, _
                vbTextCompare) = 0 Then

                FindEquivalentTargetColumn_DB = _
                    columnIndex

                Exit Function

            End If

            targetBaseKey = _
                DB_HeaderInstructionBaseKey( _
                    targetHeaderText)

            If Len(sourceBaseKey) > 0 Then

                If sourceBaseKey = _
                   targetBaseKey Then

                    FindEquivalentTargetColumn_DB = _
                        columnIndex

                    Exit Function

                End If

            End If

        End If

    Next columnIndex

End Function

Private Function DB_IsModelHeader_DB( _
    ByVal normalizedHeader As String) As Boolean

    Dim compactText As String

    compactText = _
        Replace$( _
            UCase$( _
                Trim$( _
                    normalizedHeader)), _
            " ", _
            vbNullString)

    If compactText = "MODEL" Then

        DB_IsModelHeader_DB = True
        Exit Function

    End If

    If compactText Like _
       "MODELREKRUTMEN*" Then

        DB_IsModelHeader_DB = True
        Exit Function

    End If

    If compactText Like _
       "RECRUITMENTMODEL*" Then

        DB_IsModelHeader_DB = True
        Exit Function

    End If

    If compactText Like _
       "MODEL*DROPDOWN*" Then

        DB_IsModelHeader_DB = True
        Exit Function

    End If

    If compactText Like _
       "MODEL*LIHAT*" Then

        DB_IsModelHeader_DB = True

    End If

End Function

Private Function DB_IsInstructionVariant_DB( _
    ByVal rawHeader As String) As Boolean

    Dim normalizedText As String

    normalizedText = _
        UCase$( _
            Replace$( _
                rawHeader, _
                Chr$(160), _
                " "))

    If InStr(1, normalizedText, _
        "LIHAT DROPDOWN", _
        vbTextCompare) > 0 Then

        DB_IsInstructionVariant_DB = True
        Exit Function

    End If

    If InStr(1, normalizedText, _
        "LIHAT DI SHEET", _
        vbTextCompare) > 0 Then

        DB_IsInstructionVariant_DB = True
        Exit Function

    End If

    If InStr(1, normalizedText, _
        "COPY VALUE DARI", _
        vbTextCompare) > 0 Then

        DB_IsInstructionVariant_DB = True
        Exit Function

    End If

    If InStr(1, normalizedText, _
        "ISI SESUAI", _
        vbTextCompare) > 0 Then

        DB_IsInstructionVariant_DB = True
        Exit Function

    End If

    If InStr(1, normalizedText, _
        "(DIISI", _
        vbTextCompare) > 0 Then

        DB_IsInstructionVariant_DB = True
        Exit Function

    End If

    If InStr(1, normalizedText, _
        "(BERDASARKAN", _
        vbTextCompare) > 0 Then

        DB_IsInstructionVariant_DB = True

    End If

End Function

Private Function DB_HeaderInstructionBaseKey( _
    ByVal rawHeader As String) As String

    DB_HeaderInstructionBaseKey = _
        Replace$( _
            NormalizeHeaderText_DB( _
                DB_HeaderInstructionBaseText( _
                    rawHeader)), _
            " ", _
            vbNullString)

End Function

Private Function DB_HeaderInstructionBaseText( _
    ByVal rawHeader As String) As String

    Dim resultText As String
    Dim markers As Variant
    Dim markerItem As Variant

    Dim markerPosition As Long
    Dim earliestPosition As Long

    resultText = _
        UCase$( _
            Trim$( _
                rawHeader))

    resultText = _
        Replace$( _
            resultText, _
            Chr$(160), _
            " ")

    markers = Array( _
        " (LIHAT DROPDOWN", _
        " [LIHAT DROPDOWN", _
        " - LIHAT DROPDOWN", _
        " LIHAT DROPDOWN", _
        " (LIHAT DI SHEET", _
        " [LIHAT DI SHEET", _
        " - LIHAT DI SHEET", _
        " LIHAT DI SHEET", _
        " (COPY VALUE DARI", _
        " [COPY VALUE DARI", _
        " - COPY VALUE DARI", _
        " COPY VALUE DARI", _
        " (ISI SESUAI", _
        " [ISI SESUAI", _
        " - ISI SESUAI", _
        " (DIISI", _
        " [DIISI", _
        " - DIISI", _
        " (BERDASARKAN", _
        " [BERDASARKAN", _
        " - BERDASARKAN")

    For Each markerItem In markers

        markerPosition = _
            InStr( _
                1, _
                resultText, _
                CStr(markerItem), _
                vbTextCompare)

        If markerPosition > 1 Then

            If earliestPosition = 0 Then

                earliestPosition = _
                    markerPosition

            ElseIf markerPosition < _
                   earliestPosition Then

                earliestPosition = _
                    markerPosition

            End If

        End If

    Next markerItem

    If earliestPosition > 1 Then

        resultText = _
            Left$( _
                resultText, _
                earliestPosition - 1)

    End If

    DB_HeaderInstructionBaseText = _
        Trim$( _
            resultText)

End Function

Private Function DB_SafeHeaderText( _
    ByVal rawValue As Variant) As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    DB_SafeHeaderText = _
        Trim$( _
            CStr(rawValue))

End Function

Private Function NormalizeDBSourcingSheetName_DB( _
    ByVal rawName As String) As String

    Dim sourceText As String
    Dim resultText As String
    Dim characterIndex As Long
    Dim currentCharacter As String

    sourceText = _
        UCase$( _
            Trim$( _
                rawName))

    sourceText = _
        Replace$( _
            sourceText, _
            Chr$(160), _
            " ")

    For characterIndex = 1 _
        To Len(sourceText)

        currentCharacter = _
            Mid$( _
                sourceText, _
                characterIndex, _
                1)

        If currentCharacter Like _
           "[A-Z0-9]" Then

            resultText = _
                resultText & _
                currentCharacter

        End If

    Next characterIndex

    NormalizeDBSourcingSheetName_DB = _
        resultText

End Function

Private Function HeaderTextLooksInvalid_DB( _
    ByVal headerText As String) As Boolean

    Dim normalizedText As String

    normalizedText = _
        NormalizeHeaderText_DB( _
            headerText)

    If Len(normalizedText) = 0 Then

        HeaderTextLooksInvalid_DB = True

    ElseIf InStr( _
        1, _
        normalizedText, _
        "RUMUS", _
        vbTextCompare) > 0 Then

        HeaderTextLooksInvalid_DB = True

    ElseIf InStr( _
        1, _
        normalizedText, _
        "ISI SESUAI", _
        vbTextCompare) > 0 Then

        HeaderTextLooksInvalid_DB = True

    ElseIf normalizedText = "DROPDOWN" Then

        HeaderTextLooksInvalid_DB = True

    End If

End Function

Private Function StandardDBSourcingHeaders_DB() As Variant

    Dim headers(0 To 73) As Variant

    headers(0) = "No"
    headers(1) = "Sourcing Date"
    headers(2) = "Kode Unik"
    headers(3) = "Posisi"
    headers(4) = "Model Rekrutmen"
    headers(5) = "Rekruter"
    headers(6) = "Sumber Sourcing"
    headers(7) = "Nama"
    headers(8) = "Nama Universitas/Sekolah (TOP 10)"
    headers(9) = "Nama Universitas/Sekolah Lainnya"
    headers(10) = "Jenjang Pendidikan"
    headers(11) = "Jurusan"
    headers(12) = "Tahun Lulus"
    headers(13) = "IPK"
    headers(14) = "Skor Bahasa Inggris"
    headers(15) = "University Tier"
    headers(16) = "IPK Tier"
    headers(17) = "Nomor HP"
    headers(18) = "Email"
    headers(19) = "Domisili"
    headers(20) = "Last Position"
    headers(21) = "Last Tenure"
    headers(22) = "Last Company"
    headers(23) = "Total Tenure"
    headers(24) = "Pernah di FMCG?"
    headers(25) = "Sourcing Freelance"
    headers(26) = "Tanggal Sourcing Freelance"
    headers(27) = "Sourcing HR"
    headers(28) = "Detail Keterangan Sourcing HR"
    headers(29) = "Tanggal Sourcing"
    headers(30) = "Shortlist CV"
    headers(31) = "Detail Keterangan Shortlist CV"
    headers(32) = "Tanggal Shortlist CV"
    headers(33) = "Psikotes"
    headers(34) = "Kode Psikotes"
    headers(35) = "Detail Keterangan Psikotes"
    headers(36) = "Tanggal Psikotes / Cek psikotes"
    headers(37) = "Nilai Logika"
    headers(38) = "Nilai IQ"
    headers(39) = "Nilai Daya Tangkap"
    headers(40) = "Nilai RA"
    headers(41) = "DISC"
    headers(42) = "HR Interview"
    headers(43) = "Detail Keterangan HR Interview"
    headers(44) = "Tanggal HR Interview"
    headers(45) = "Technical Test/ Case Study"
    headers(46) = "Detail Keterangan Technical Test/ Case Study"
    headers(47) = "Tanggal Technical Test/ Case Study"
    headers(48) = "Market Visit"
    headers(49) = "Detail Market Visit"
    headers(50) = "Tanggal Market Visit"
    headers(51) = "User Interview"
    headers(52) = "Detail Keterangan User Interview"
    headers(53) = "Tanggal User Interview"
    headers(54) = "Panel Interview"
    headers(55) = "Detail Keterangan Panel Interview"
    headers(56) = "Tanggal Panel Interview"
    headers(57) = "Reference Check"
    headers(58) = "Detail Keterangan Reference Check"
    headers(59) = "Tanggal Reference Check"
    headers(60) = "MCU"
    headers(61) = "Detail Keterangan MCU"
    headers(62) = "Tanggal MCU"
    headers(63) = "Offering"
    headers(64) = "Detail Keterangan Offering"
    headers(65) = "Tanggal Offering"
    headers(66) = "Notes"
    headers(67) = "Day 1"
    headers(68) = "Detail Keterangan Day 1"
    headers(69) = "Tanggal Day 1"
    headers(70) = "Created At"
    headers(71) = "Last Updated At"
    headers(72) = "Last Compile Action"
    headers(73) = "Source File"

    StandardDBSourcingHeaders_DB = headers

End Function

Private Function OneDimensionalToRowArray_DB( _
    ByVal sourceArray As Variant) As Variant

    Dim resultArray() As Variant
    Dim itemIndex As Long

    ReDim resultArray( _
        1 To 1, _
        1 To UBound(sourceArray) - _
             LBound(sourceArray) + 1)

    For itemIndex = LBound(sourceArray) _
        To UBound(sourceArray)

        resultArray( _
            1, itemIndex - _
               LBound(sourceArray) + 1) = _
            sourceArray(itemIndex)

    Next itemIndex

    OneDimensionalToRowArray_DB = resultArray

End Function

' ============================================================
' HEADER MAP
' ============================================================
Private Function BuildHeaderMap_DB( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal lastColumn As Long) As Object

    Dim headerMap As Object
    Dim columnIndex As Long
    Dim canonicalHeader As String

    Set headerMap = CreateObject( _
        "Scripting.Dictionary")

    headerMap.CompareMode = vbTextCompare

    For columnIndex = 1 To lastColumn

        canonicalHeader = CanonicalHeader_DB( _
            ws.Cells( _
                headerRow, columnIndex).Value2)

        If Len(canonicalHeader) > 0 Then

            If Not headerMap.Exists( _
                canonicalHeader) Then

                headerMap.Add _
                    canonicalHeader, columnIndex

            End If

        End If

    Next columnIndex

    Set BuildHeaderMap_DB = headerMap

End Function

Private Function BuildCanonicalArray_DB( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal lastColumn As Long) As Variant

    Dim resultArray() As String
    Dim columnIndex As Long

    ReDim resultArray(1 To lastColumn)

    For columnIndex = 1 To lastColumn

        resultArray(columnIndex) = _
            CanonicalHeader_DB( _
                ws.Cells( _
                    headerRow, columnIndex).Value2)

    Next columnIndex

    BuildCanonicalArray_DB = resultArray

End Function

' ============================================================
' HEADER CANONICAL / ALIAS
' ============================================================
Private Function CanonicalHeader_DB( _
    ByVal rawHeader As Variant) As String

    Dim normalizedHeader As String
    Dim schemaKey As String

    schemaKey = SchemaCanonicalHeaderKey( _
        TARGET_SHEET_NAME, rawHeader)

    If Len(schemaKey) > 0 Then

        normalizedHeader = _
            NormalizeHeaderText_DB( _
                schemaKey)

    Else

        normalizedHeader = _
            NormalizeHeaderText_DB( _
                rawHeader)

    End If

    If Len(normalizedHeader) = 0 Then
        Exit Function
    End If

    If DB_IsInstructionVariant_DB( _
        CStr(rawHeader)) Then

        normalizedHeader = _
            NormalizeHeaderText_DB( _
                DB_HeaderInstructionBaseText( _
                    CStr(rawHeader)))

    End If

    Select Case True

        Case normalizedHeader = "NO" Or _
             normalizedHeader = "NOMOR" Or _
             normalizedHeader = "NUMBER"

            CanonicalHeader_DB = "NO"

        Case InStr( _
            1, normalizedHeader, _
            "KODE UNIK", vbTextCompare) = 1

            CanonicalHeader_DB = "KODE UNIK"

        Case normalizedHeader = "POSISI" Or _
             normalizedHeader = "POSITION" Or _
             InStr( _
                1, normalizedHeader, _
                "POSISI KEBUTUHAN TA", _
                vbTextCompare) > 0

            CanonicalHeader_DB = "POSISI"

        Case normalizedHeader = "NAMA" Or _
             normalizedHeader = "NAMA KANDIDAT" Or _
             normalizedHeader = "CANDIDATE NAME"

            CanonicalHeader_DB = "NAMA"

        Case DB_IsModelHeader_DB( _
            normalizedHeader)

            CanonicalHeader_DB = _
                "MODEL REKRUTMEN"

        Case normalizedHeader = "REKRUTER" Or _
             normalizedHeader = "RECRUITER" Or _
             normalizedHeader = "PIC RECRUITER" Or _
             normalizedHeader = "PIC REKRUTER"

            CanonicalHeader_DB = "REKRUTER"

        Case normalizedHeader = "SUMBER SOURCING" Or _
             normalizedHeader = "SOURCING SOURCE" Or _
             normalizedHeader = "SOURCE"

            CanonicalHeader_DB = "SUMBER SOURCING"

        Case normalizedHeader = "SOURCING DATE"

            CanonicalHeader_DB = "SOURCING DATE"

        Case normalizedHeader = "TANGGAL SOURCING"

            CanonicalHeader_DB = "TANGGAL SOURCING"

        Case normalizedHeader = _
                "TANGGAL SOURCING FREELANCE" Or _
             normalizedHeader = _
                "FREELANCE SOURCING DATE"

            CanonicalHeader_DB = _
                "TANGGAL SOURCING FREELANCE"

        Case normalizedHeader = "SOURCING FREELANCE" Or _
             normalizedHeader = "SOURCIN FREELANCE" Or _
             normalizedHeader Like "SOURCING FREELANCE#*"

            CanonicalHeader_DB = "SOURCING FREELANCE"

        Case normalizedHeader = "SOURCING HR"
            CanonicalHeader_DB = "SOURCING HR"

        Case normalizedHeader = _
                "DETAIL KETERANGAN SOURCING HR" Or _
             normalizedHeader = _
                "DETAIL SOURCING HR"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN SOURCING HR"

        Case InStr( _
            1, normalizedHeader, _
            "NAMA UNIVERSITAS SEKOLAH TOP 10", _
            vbTextCompare) = 1 Or _
             normalizedHeader = "UNIVERSITAS TOP 10"

            CanonicalHeader_DB = _
                "UNIVERSITAS TOP 10"

        Case InStr( _
            1, normalizedHeader, _
            "NAMA UNIVERSITAS SEKOLAH LAINNYA", _
            vbTextCompare) = 1 Or _
             normalizedHeader = "UNIVERSITAS LAINNYA"

            CanonicalHeader_DB = _
                "UNIVERSITAS LAINNYA"

        Case normalizedHeader = _
                "JENJANG PENDIDIKAN" Or _
             normalizedHeader = "EDUCATION LEVEL"

            CanonicalHeader_DB = _
                "JENJANG PENDIDIKAN"

        Case normalizedHeader = "JURUSAN" Or _
             normalizedHeader = "MAJOR"

            CanonicalHeader_DB = "JURUSAN"

        Case normalizedHeader = "TAHUN LULUS" Or _
             normalizedHeader = "GRADUATION YEAR"

            CanonicalHeader_DB = "TAHUN LULUS"

        Case normalizedHeader = "IPK" Or _
             normalizedHeader = "GPA"

            CanonicalHeader_DB = "IPK"

        Case normalizedHeader = _
                "SKOR BAHASA INGGRIS" Or _
             normalizedHeader = "ENGLISH SCORE"

            CanonicalHeader_DB = _
                "SKOR BAHASA INGGRIS"

        Case normalizedHeader = "UNIVERSITY TIER" Or _
             normalizedHeader = "UNIV TIER"

            CanonicalHeader_DB = "UNIVERSITY TIER"

        Case normalizedHeader = "IPK TIER" Or _
             normalizedHeader = "GPA TIER"

            CanonicalHeader_DB = "IPK TIER"

        Case normalizedHeader = "NOMOR HP" Or _
             normalizedHeader = "NO HP" Or _
             normalizedHeader = "PHONE" Or _
             normalizedHeader = "PHONE NUMBER"

            CanonicalHeader_DB = "NOMOR HP"

        Case normalizedHeader = "EMAIL" Or _
             normalizedHeader = "EMAIL ADDRESS"

            CanonicalHeader_DB = "EMAIL"

        Case normalizedHeader = "DOMISILI" Or _
             normalizedHeader = "DOMICILE"

            CanonicalHeader_DB = "DOMISILI"

        Case normalizedHeader = "LAST POSITION" Or _
             normalizedHeader = "POSISI TERAKHIR" Or _
             InStr( _
                1, _
                normalizedHeader, _
                "LAST POSITION", _
                vbTextCompare) = 1

            CanonicalHeader_DB = "LAST POSITION"

        Case normalizedHeader = "LAST TENURE" Or _
             normalizedHeader = _
                "TENURE IN LAST POSITION" Or _
             InStr( _
                1, _
                normalizedHeader, _
                "LAST TENURE", _
                vbTextCompare) = 1

            CanonicalHeader_DB = "LAST TENURE"

        Case normalizedHeader = "LAST COMPANY" Or _
             normalizedHeader = "COMPANY TERAKHIR" Or _
             InStr( _
                1, _
                normalizedHeader, _
                "LAST COMPANY", _
                vbTextCompare) = 1

            CanonicalHeader_DB = "LAST COMPANY"

        Case normalizedHeader = "TOTAL TENURE" Or _
             normalizedHeader = _
                "LENGTH OF EXPERIENCE" Or _
             InStr( _
                1, _
                normalizedHeader, _
                "TOTAL TENURE", _
                vbTextCompare) = 1

            CanonicalHeader_DB = "TOTAL TENURE"

        Case normalizedHeader = "PERNAH DI FMCG" Or _
             normalizedHeader = "FMCG" Or _
             normalizedHeader = "FMCG EXPERIENCE" Or _
             InStr( _
                1, _
                normalizedHeader, _
                "BERPENGALAMAN DI INDUSTRI FMCG", _
                vbTextCompare) = 1

            CanonicalHeader_DB = "FMCG"

        Case normalizedHeader = "SHORTLIST CV"
            CanonicalHeader_DB = "SHORTLIST CV"

        Case normalizedHeader = _
                "DETAIL KETERANGAN SHORTLIST CV"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN SHORTLIST CV"

        Case normalizedHeader = _
                "TANGGAL SHORTLIST CV"

            CanonicalHeader_DB = _
                "TANGGAL SHORTLIST CV"

        Case normalizedHeader = "PSIKOTES"
            CanonicalHeader_DB = "PSIKOTES"

        Case normalizedHeader = "KODE PSIKOTES"
            CanonicalHeader_DB = "KODE PSIKOTES"

        Case normalizedHeader = _
                "DETAIL KETERANGAN PSIKOTES"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN PSIKOTES"

        Case InStr( _
            1, normalizedHeader, _
            "TANGGAL PSIKOTES", _
            vbTextCompare) = 1

            CanonicalHeader_DB = "TANGGAL PSIKOTES"

        Case normalizedHeader = "NILAI LOGIKA"
            CanonicalHeader_DB = "NILAI LOGIKA"

        Case normalizedHeader = "NILAI IQ"
            CanonicalHeader_DB = "NILAI IQ"

        Case normalizedHeader = "NILAI DAYA TANGKAP"
            CanonicalHeader_DB = "NILAI DAYA TANGKAP"

        Case normalizedHeader = "NILAI RA"
            CanonicalHeader_DB = "NILAI RA"

        Case normalizedHeader = "DISC"
            CanonicalHeader_DB = "DISC"

        Case normalizedHeader = "HR INTERVIEW"
            CanonicalHeader_DB = "HR INTERVIEW"

        Case normalizedHeader = _
                "DETAIL KETERANGAN HR INTERVIEW"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN HR INTERVIEW"

        Case normalizedHeader = _
                "TANGGAL HR INTERVIEW"

            CanonicalHeader_DB = _
                "TANGGAL HR INTERVIEW"

        Case normalizedHeader = _
                "TECHNICAL TEST CASE STUDY"

            CanonicalHeader_DB = _
                "TECHNICAL TEST CASE STUDY"

        Case normalizedHeader = _
                "DETAIL KETERANGAN TECHNICAL TEST CASE STUDY"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN TECHNICAL TEST CASE STUDY"

        Case normalizedHeader = _
                "TANGGAL TECHNICAL TEST CASE STUDY"

            CanonicalHeader_DB = _
                "TANGGAL TECHNICAL TEST CASE STUDY"

        Case normalizedHeader = "MARKET VISIT"
            CanonicalHeader_DB = "MARKET VISIT"

        Case normalizedHeader = "DETAIL MARKET VISIT"
            CanonicalHeader_DB = "DETAIL MARKET VISIT"

        Case normalizedHeader = "TANGGAL MARKET VISIT"
            CanonicalHeader_DB = "TANGGAL MARKET VISIT"

        Case normalizedHeader = "USER INTERVIEW"
            CanonicalHeader_DB = "USER INTERVIEW"

        Case normalizedHeader = _
                "DETAIL KETERANGAN USER INTERVIEW"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN USER INTERVIEW"

        Case normalizedHeader = _
                "TANGGAL USER INTERVIEW"

            CanonicalHeader_DB = _
                "TANGGAL USER INTERVIEW"

        Case normalizedHeader = "PANEL INTERVIEW"
            CanonicalHeader_DB = "PANEL INTERVIEW"

        Case normalizedHeader = _
                "DETAIL KETERANGAN PANEL INTERVIEW"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN PANEL INTERVIEW"

        Case normalizedHeader = _
                "TANGGAL PANEL INTERVIEW"

            CanonicalHeader_DB = _
                "TANGGAL PANEL INTERVIEW"

        Case normalizedHeader = "REFERENCE CHECK"
            CanonicalHeader_DB = "REFERENCE CHECK"

        Case normalizedHeader = _
                "DETAIL KETERANGAN REFERENCE CHECK"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN REFERENCE CHECK"

        Case normalizedHeader = _
                "TANGGAL REFERENCE CHECK"

            CanonicalHeader_DB = _
                "TANGGAL REFERENCE CHECK"

        Case normalizedHeader = "MCU"
            CanonicalHeader_DB = "MCU"

        Case normalizedHeader = _
                "DETAIL KETERANGAN MCU"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN MCU"

        Case normalizedHeader = "TANGGAL MCU"
            CanonicalHeader_DB = "TANGGAL MCU"

        Case normalizedHeader = "OFFERING"
            CanonicalHeader_DB = "OFFERING"

        Case normalizedHeader = _
                "DETAIL KETERANGAN OFFERING"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN OFFERING"

        Case normalizedHeader = "TANGGAL OFFERING"
            CanonicalHeader_DB = "TANGGAL OFFERING"

        Case normalizedHeader = "NOTES" Or _
             normalizedHeader = "CATATAN"

            CanonicalHeader_DB = "NOTES"

        Case normalizedHeader = "DAY 1" Or _
             normalizedHeader = "DAY ONE"

            CanonicalHeader_DB = "DAY 1"

        Case normalizedHeader = _
                "DETAIL KETERANGAN DAY 1"

            CanonicalHeader_DB = _
                "DETAIL KETERANGAN DAY 1"

        Case normalizedHeader = "TANGGAL DAY 1"
            CanonicalHeader_DB = "TANGGAL DAY 1"

        Case normalizedHeader = "CREATED AT" Or _
             normalizedHeader = "CREATED DATE" Or _
             normalizedHeader = "CREATED DATETIME"

            CanonicalHeader_DB = "CREATED AT"

        Case normalizedHeader = "LAST UPDATED AT" Or _
             normalizedHeader = "LAST UPDATED" Or _
             normalizedHeader = "UPDATED AT"

            CanonicalHeader_DB = "LAST UPDATED AT"

        Case normalizedHeader = "LAST COMPILE ACTION" Or _
             normalizedHeader = "COMPILE ACTION" Or _
             normalizedHeader = "ACTION"

            CanonicalHeader_DB = "LAST COMPILE ACTION"

        Case normalizedHeader = "SOURCE FILE" Or _
             normalizedHeader = "SOURCE FILENAME" Or _
             normalizedHeader = "LAST SOURCE FILE"

            CanonicalHeader_DB = "SOURCE FILE"

        Case Else
            CanonicalHeader_DB = normalizedHeader

    End Select

End Function

Private Function NormalizeHeaderText_DB( _
    ByVal rawHeader As Variant) As String

    Dim sourceText As String
    Dim resultText As String
    Dim characterIndex As Long
    Dim currentCharacter As String
    Dim characterCode As Long

    If IsError(rawHeader) Or _
       IsNull(rawHeader) Or _
       IsEmpty(rawHeader) Then

        Exit Function

    End If

    sourceText = UCase$(CStr(rawHeader))
    sourceText = Replace$( _
        sourceText, Chr$(160), " ")

    For characterIndex = 1 To Len(sourceText)

        currentCharacter = Mid$( _
            sourceText, characterIndex, 1)

        characterCode = AscW(currentCharacter)

        If characterCode >= 48 And _
           characterCode <= 57 Then

            resultText = resultText & _
                currentCharacter

        ElseIf characterCode >= 65 And _
               characterCode <= 90 Then

            resultText = resultText & _
                currentCharacter

        Else

            resultText = resultText & " "

        End If

    Next characterIndex

    NormalizeHeaderText_DB = _
        Application.Trim(resultText)

End Function

' ============================================================
' FIND SOURCE SHEET / HEADER
' ============================================================
Private Function GetSourceSheet_DB( _
    ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet
    Dim preferredSheet As Worksheet
    Dim candidateHeaderRow As Long
    Dim normalizedSheetName As String

    If wb Is Nothing Then
        Exit Function
    End If

    On Error Resume Next

    Set preferredSheet = _
        wb.Worksheets( _
            TARGET_SHEET_NAME)

    On Error GoTo 0

    If Not preferredSheet Is Nothing Then

        candidateHeaderRow = _
            FindBestHeaderRow_DB( _
                preferredSheet)

        If candidateHeaderRow > 0 Then

            Set GetSourceSheet_DB = _
                preferredSheet

            Exit Function

        End If

    End If

    ' Jangan fallback ke sembarang sheet.
    ' Sheet FPTK juga memiliki Kode Unik/Nama/Posisi dan pada
    ' versi lama dapat salah dianggap sebagai DB Sourcing.
    For Each ws In wb.Worksheets

        normalizedSheetName = _
            NormalizeDBSourcingSheetName_DB( _
                CStr(ws.Name))

        If normalizedSheetName = _
           "DBSOURCING" Or _
           normalizedSheetName = _
           "SOURCINGDATABASE" Or _
           normalizedSheetName = _
           "DATABASESOURCING" Then

            candidateHeaderRow = _
                FindBestHeaderRow_DB( _
                    ws)

            If candidateHeaderRow > 0 Then

                Set GetSourceSheet_DB = ws
                Exit Function

            End If

        End If

    Next ws

End Function

Private Function FindBestHeaderRow_DB( _
    ByVal ws As Worksheet) As Long

    Dim rowIndex As Long
    Dim currentScore As Long
    Dim bestScore As Long

    For rowIndex = 1 To 30

        currentScore = HeaderRowScore_DB( _
            ws, rowIndex)

        If currentScore > bestScore Then

            bestScore = currentScore
            FindBestHeaderRow_DB = rowIndex

        End If

    Next rowIndex

    If bestScore < 9 Then
        FindBestHeaderRow_DB = 0
    End If

End Function

Private Function HeaderRowScore_DB( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long) As Long

    Dim columnIndex As Long
    Dim lastColumn As Long
    Dim canonicalHeader As String
    Dim rowHeaders As Object

    lastColumn = LastUsedColumnInRow_DB( _
        ws, rowIndex)

    If lastColumn < 1 Then Exit Function

    Set rowHeaders = CreateObject( _
        "Scripting.Dictionary")

    rowHeaders.CompareMode = vbTextCompare

    For columnIndex = 1 To lastColumn

        canonicalHeader = CanonicalHeader_DB( _
            ws.Cells( _
                rowIndex, columnIndex).Value2)

        If Len(canonicalHeader) > 0 Then

            If Not rowHeaders.Exists( _
                canonicalHeader) Then

                rowHeaders.Add _
                    canonicalHeader, True

            End If

        End If

    Next columnIndex

    If rowHeaders.Exists("KODE UNIK") Then

        HeaderRowScore_DB = _
            HeaderRowScore_DB + 5

    Else

        Exit Function

    End If

    If rowHeaders.Exists("NAMA") Then
        HeaderRowScore_DB = _
            HeaderRowScore_DB + 4
    End If

    If rowHeaders.Exists("POSISI") Then
        HeaderRowScore_DB = _
            HeaderRowScore_DB + 3
    End If

    If rowHeaders.Exists("REKRUTER") Then
        HeaderRowScore_DB = _
            HeaderRowScore_DB + 2
    End If

    If rowHeaders.Exists("SOURCING DATE") Then
        HeaderRowScore_DB = _
            HeaderRowScore_DB + 2
    End If

    If rowHeaders.Exists("EMAIL") Then
        HeaderRowScore_DB = _
            HeaderRowScore_DB + 1
    End If

    If rowHeaders.Exists("NOMOR HP") Then
        HeaderRowScore_DB = _
            HeaderRowScore_DB + 1
    End If

End Function

' ============================================================
' SOURCE / TARGET LAST ROW
' ============================================================
Private Function GetLastSourceDataRow_DB( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal headerMap As Object) As Long

    Dim preferredHeaders As Variant
    Dim headerItem As Variant
    Dim columnIndex As Long
    Dim candidateRow As Long
    Dim lastRow As Long

    preferredHeaders = Array( _
        "KODE UNIK", _
        "NAMA", _
        "POSISI", _
        "EMAIL", _
        "NOMOR HP", _
        "SOURCING DATE")

    lastRow = headerRow

    For Each headerItem In preferredHeaders

        If headerMap.Exists( _
            CStr(headerItem)) Then

            columnIndex = CLng( _
                headerMap(CStr(headerItem)))

            candidateRow = ws.Cells( _
                ws.rows.Count, _
                columnIndex).End(xlUp).Row

            If candidateRow > lastRow Then
                lastRow = candidateRow
            End If

        End If

    Next headerItem

    GetLastSourceDataRow_DB = lastRow

End Function

Private Function GetLastTargetDataRow_DB( _
    ByVal ws As Worksheet) As Long

    Dim headerMap As Object
    Dim preferredHeaders As Variant
    Dim headerItem As Variant
    Dim columnIndex As Long
    Dim candidateRow As Long
    Dim lastRow As Long

    Set headerMap = BuildHeaderMap_DB( _
        ws, _
        TARGET_HEADER_ROW, _
        LastUsedColumnInRow_DB( _
            ws, TARGET_HEADER_ROW))

    preferredHeaders = Array( _
        "KODE UNIK", _
        "NAMA", _
        "POSISI", _
        "EMAIL", _
        "NOMOR HP")

    lastRow = 1

    For Each headerItem In preferredHeaders

        If headerMap.Exists( _
            CStr(headerItem)) Then

            columnIndex = CLng( _
                headerMap(CStr(headerItem)))

            candidateRow = ws.Cells( _
                ws.rows.Count, _
                columnIndex).End(xlUp).Row

            If candidateRow > lastRow Then
                lastRow = candidateRow
            End If

        End If

    Next headerItem

    GetLastTargetDataRow_DB = lastRow

End Function

Private Function GetNextTargetRow_DB( _
    ByVal ws As Worksheet) As Long

    Dim lastRow As Long

    lastRow = GetLastTargetDataRow_DB(ws)

    If lastRow < FIRST_TARGET_DATA_ROW Then

        GetNextTargetRow_DB = _
            FIRST_TARGET_DATA_ROW

    Else

        GetNextTargetRow_DB = _
            lastRow + 1

    End If

End Function

' ============================================================
' ROW FILTER
' ============================================================
Private Function ArrayRowIsEmpty_DB( _
    ByRef sourceData As Variant, _
    ByVal rowIndex As Long) As Boolean

    Dim columnIndex As Long

    ArrayRowIsEmpty_DB = True

    For columnIndex = 1 _
        To UBound(sourceData, 2)

        If HasMeaningfulValue_DB( _
            sourceData( _
                rowIndex, columnIndex)) Then

            ArrayRowIsEmpty_DB = False
            Exit Function

        End If

    Next columnIndex

End Function

Private Function ArrayRowIsInstruction_DB( _
    ByRef sourceData As Variant, _
    ByVal rowIndex As Long, _
    ByVal sourceCanonicalByColumn As Variant) As Boolean

    Dim checkText As String
    Dim columnIndex As Long
    Dim canonicalHeader As String
    Dim cellText As String

    For columnIndex = 1 _
        To UBound(sourceData, 2)

        canonicalHeader = CStr( _
            sourceCanonicalByColumn(columnIndex))

        If canonicalHeader = "KODE UNIK" Or _
           canonicalHeader = "NAMA" Or _
           canonicalHeader = "POSISI" Then

            cellText = UCase$( _
                Trim$(CStr( _
                    sourceData( _
                        rowIndex, columnIndex))))

            checkText = checkText & _
                " " & cellText

        End If

    Next columnIndex

    checkText = Application.Trim(checkText)

    If InStr( _
        1, checkText, _
        "RUMUS", vbTextCompare) > 0 Then

        ArrayRowIsInstruction_DB = True

    ElseIf InStr( _
        1, checkText, _
        "DROPDOWN", vbTextCompare) > 0 Then

        ArrayRowIsInstruction_DB = True

    ElseIf InStr( _
        1, checkText, _
        "ISI SESUAI", vbTextCompare) > 0 Then

        ArrayRowIsInstruction_DB = True

    ElseIf InStr( _
        1, checkText, _
        "KODE UNIK", vbTextCompare) > 0 And _
           InStr( _
        1, checkText, _
        "NAMA", vbTextCompare) > 0 Then

        ArrayRowIsInstruction_DB = True

    End If

End Function

Private Sub RemoveDuplicateHeaderRows_DB( _
    ByVal ws As Worksheet)

    Dim attemptNumber As Long

    For attemptNumber = 1 To 5

        If RowLooksLikeDBHeader_DB( _
            ws, FIRST_TARGET_DATA_ROW) Then

            ws.rows( _
                FIRST_TARGET_DATA_ROW).Delete _
                Shift:=xlUp

        Else

            Exit For

        End If

    Next attemptNumber

End Sub

Private Function RowLooksLikeDBHeader_DB( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long) As Boolean

    Dim headerMap As Object
    Dim score As Long

    Set headerMap = BuildHeaderMap_DB( _
        ws, _
        rowIndex, _
        LastUsedColumnInRow_DB( _
            ws, rowIndex))

    If headerMap.Exists("KODE UNIK") Then
        score = score + 4
    End If

    If headerMap.Exists("NAMA") Then
        score = score + 3
    End If

    If headerMap.Exists("POSISI") Then
        score = score + 2
    End If

    If headerMap.Exists("REKRUTER") Then
        score = score + 1
    End If

    RowLooksLikeDBHeader_DB = (score >= 7)

End Function

' ============================================================
' VALUE READERS
' ============================================================
Private Function GetMappedValue_DB( _
    ByRef sourceData As Variant, _
    ByVal rowIndex As Long, _
    ByVal sourceHeaderMap As Object, _
    ByVal canonicalHeader As String) As Variant

    Dim sourceColumn As Long

    If Not sourceHeaderMap.Exists( _
        canonicalHeader) Then

        GetMappedValue_DB = Empty
        Exit Function

    End If

    sourceColumn = CLng( _
        sourceHeaderMap(canonicalHeader))

    If sourceColumn < 1 Or _
       sourceColumn > UBound(sourceData, 2) Then

        GetMappedValue_DB = Empty
        Exit Function

    End If

    GetMappedValue_DB = _
        sourceData(rowIndex, sourceColumn)

End Function

Private Function SheetValueByCanonical_DB( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal headerMap As Object, _
    ByVal canonicalHeader As String) As Variant

    Dim columnIndex As Long

    If Not headerMap.Exists( _
        canonicalHeader) Then

        SheetValueByCanonical_DB = Empty
        Exit Function

    End If

    columnIndex = CLng( _
        headerMap(canonicalHeader))

    SheetValueByCanonical_DB = _
        ws.Cells( _
            rowIndex, columnIndex).Value2

End Function

Private Function HasMeaningfulValue_DB( _
    ByVal rawValue As Variant) As Boolean

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    HasMeaningfulValue_DB = _
        (Len(Trim$(CStr(rawValue))) > 0)

End Function

' ============================================================
' NORMALIZE KEY
' ============================================================
Private Function NormalizeKey_DB( _
    ByVal rawValue As Variant) As String

    Dim textValue As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    textValue = UCase$( _
        Trim$(CStr(rawValue)))

    textValue = Replace$( _
        textValue, Chr$(160), " ")

    textValue = Replace$( _
        textValue, vbTab, " ")

    textValue = Replace$( _
        textValue, vbCr, " ")

    textValue = Replace$( _
        textValue, vbLf, " ")

    Do While InStr( _
        1, textValue, "  ", _
        vbBinaryCompare) > 0

        textValue = Replace$( _
            textValue, "  ", " ")

    Loop

    NormalizeKey_DB = textValue

End Function

Private Function NormalizeEmailKey_DB( _
    ByVal rawValue As Variant) As String

    NormalizeEmailKey_DB = LCase$( _
        Trim$(CStr(rawValue)))

End Function

Private Function NormalizePhoneKey_DB( _
    ByVal rawValue As Variant) As String

    Dim sourceText As String
    Dim resultText As String
    Dim characterIndex As Long
    Dim currentCharacter As String

    sourceText = CStr(rawValue)

    For characterIndex = 1 To Len(sourceText)

        currentCharacter = Mid$( _
            sourceText, characterIndex, 1)

        If currentCharacter Like "#" Then

            resultText = resultText & _
                currentCharacter

        End If

    Next characterIndex

    NormalizePhoneKey_DB = resultText

End Function

Private Function NormalizeDateKey_DB( _
    ByVal rawValue As Variant) As String

    Dim parsedDate As Date

    parsedDate = ParseDateFlexible_DB(rawValue)

    If parsedDate > 0 Then

        NormalizeDateKey_DB = _
            Format$(parsedDate, "yyyymmdd")

    End If

End Function

Private Function ParseDateFlexible_DB( _
    ByVal rawValue As Variant) As Date

    Dim textValue As String
    Dim dateParts() As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long

    On Error GoTo InvalidDate

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        GoTo InvalidDate

    End If

    If IsNumeric(rawValue) Then

        If CDbl(rawValue) > 0 Then

            ParseDateFlexible_DB = _
                dateValue(CDate(CDbl(rawValue)))

            Exit Function

        End If

    End If

    textValue = Trim$(CStr(rawValue))

    If Len(textValue) = 0 Then
        GoTo InvalidDate
    End If

    textValue = Replace$(textValue, "-", "/")
    textValue = Replace$(textValue, ".", "/")
    textValue = Replace$(textValue, "\", "/")

    dateParts = Split(textValue, "/")

    If UBound(dateParts) = 2 Then

        If IsNumeric(dateParts(0)) And _
           IsNumeric(dateParts(1)) And _
           IsNumeric(dateParts(2)) Then

            dayValue = CLng(dateParts(0))
            monthValue = CLng(dateParts(1))
            yearValue = CLng(dateParts(2))

            If yearValue < 100 Then
                yearValue = yearValue + 2000
            End If

            ParseDateFlexible_DB = DateSerial( _
                yearValue, monthValue, dayValue)

            Exit Function

        End If

    End If

    If IsDate(textValue) Then

        ParseDateFlexible_DB = _
            dateValue(CDate(textValue))

        Exit Function

    End If

InvalidDate:

    ParseDateFlexible_DB = 0

End Function

' ============================================================
' SANITIZE OUTPUT
' ============================================================
Private Function SanitizeOutputValue_DB( _
    ByVal rawValue As Variant) As Variant

    Dim textValue As String
    Dim firstCharacter As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Then

        SanitizeOutputValue_DB = vbNullString
        Exit Function

    End If

    If VarType(rawValue) = vbString Then

        textValue = CStr(rawValue)

        If Len(textValue) > 0 Then

            firstCharacter = Left$(textValue, 1)

            If firstCharacter = "=" Or _
               firstCharacter = "+" Or _
               firstCharacter = "-" Or _
               firstCharacter = "@" Then

                SanitizeOutputValue_DB = _
                    "'" & textValue

                Exit Function

            End If

        End If

    End If

    SanitizeOutputValue_DB = rawValue

End Function

' ============================================================
' MISSING COLUMN COUNT
' ============================================================
Private Function CountMissingColumns_DB( _
    ByVal targetHeaderMap As Object, _
    ByVal sourceHeaderMap As Object) As Long

    Dim keyItem As Variant

    For Each keyItem In targetHeaderMap.keys

        If CStr(keyItem) <> "NO" Then

            If Not sourceHeaderMap.Exists( _
                CStr(keyItem)) Then

                CountMissingColumns_DB = _
                    CountMissingColumns_DB + 1

            End If

        End If

    Next keyItem

End Function

' ============================================================
' NUMBERING
' ============================================================
Private Function GetNextSequenceNumber_DB( _
    ByVal ws As Worksheet) As Long

    Dim headerMap As Object
    Dim noColumn As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim currentValue As Variant
    Dim maxNumber As Long

    Set headerMap = BuildHeaderMap_DB( _
        ws, _
        TARGET_HEADER_ROW, _
        LastUsedColumnInRow_DB( _
            ws, TARGET_HEADER_ROW))

    If Not headerMap.Exists("NO") Then

        GetNextSequenceNumber_DB = 1
        Exit Function

    End If

    noColumn = CLng(headerMap("NO"))
    lastRow = GetLastTargetDataRow_DB(ws)

    For rowIndex = FIRST_TARGET_DATA_ROW _
        To lastRow

        currentValue = ws.Cells( _
            rowIndex, noColumn).Value2

        If IsNumeric(currentValue) Then

            If CLng(currentValue) > maxNumber Then
                maxNumber = CLng(currentValue)
            End If

        End If

    Next rowIndex

    GetNextSequenceNumber_DB = maxNumber + 1

End Function

Private Sub RenumberTargetRows_DB( _
    ByVal ws As Worksheet)

    Dim headerMap As Object
    Dim noColumn As Long
    Dim lastRow As Long
    Dim rowCount As Long
    Dim numberArray() As Variant
    Dim rowIndex As Long

    Set headerMap = BuildHeaderMap_DB( _
        ws, _
        TARGET_HEADER_ROW, _
        LastUsedColumnInRow_DB( _
            ws, TARGET_HEADER_ROW))

    If Not headerMap.Exists("NO") Then Exit Sub

    noColumn = CLng(headerMap("NO"))
    lastRow = GetLastTargetDataRow_DB(ws)

    If lastRow < FIRST_TARGET_DATA_ROW Then
        Exit Sub
    End If

    rowCount = _
        lastRow - FIRST_TARGET_DATA_ROW + 1

    ReDim numberArray(1 To rowCount, 1 To 1)

    For rowIndex = 1 To rowCount

        numberArray(rowIndex, 1) = rowIndex

    Next rowIndex

    ws.Cells( _
        FIRST_TARGET_DATA_ROW, _
        noColumn).Resize( _
            rowCount, 1).Value = numberArray

End Sub

' ============================================================
' DATE FORMAT DYNAMIC
' ============================================================
Private Sub ApplyDynamicDateFormats_DB( _
    ByVal ws As Worksheet)

    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim canonicalHeader As String

    lastColumn = LastUsedColumnInRow_DB( _
        ws, TARGET_HEADER_ROW)

    For columnIndex = 1 To lastColumn

        canonicalHeader = CanonicalHeader_DB( _
            ws.Cells( _
                TARGET_HEADER_ROW, _
                columnIndex).Value2)

        If canonicalHeader = "CREATED AT" Or _
           canonicalHeader = "LAST UPDATED AT" Then

            ws.Columns( _
                columnIndex).numberFormat = _
                "dd/mm/yyyy hh:mm:ss"

        ElseIf canonicalHeader = "SOURCING DATE" Or _
               canonicalHeader = "TANGGAL SOURCING" Or _
               canonicalHeader = _
                    "TANGGAL SOURCING FREELANCE" Or _
               Left$(canonicalHeader, 8) = "TANGGAL " Or _
               Right$(canonicalHeader, 5) = " DATE" Then

            ws.Columns( _
                columnIndex).numberFormat = _
                "dd/mm/yyyy"

        End If

    Next columnIndex

End Sub

' ============================================================
' TABLE
' ============================================================
Private Sub CaptureExistingTableInfo_DB( _
    ByVal ws As Worksheet, _
    ByRef tableName As String, _
    ByRef tableStyle As String)

    If ws.ListObjects.Count = 0 Then Exit Sub

    tableName = ws.ListObjects(1).Name
    tableStyle = ws.ListObjects(1).tableStyle

End Sub

Private Sub CreateOrResizeTargetTable_DB( _
    ByVal ws As Worksheet, _
    ByVal lastDataRow As Long, _
    ByVal lastColumn As Long, _
    ByVal preferredTableName As String, _
    ByVal preferredTableStyle As String)

    Dim lo As ListObject
    Dim tableRange As Range
    Dim selectedTableName As String

    If ws Is Nothing Then Exit Sub

    If lastDataRow < FIRST_TARGET_DATA_ROW Then
        Exit Sub
    End If

    If lastColumn < 1 Then Exit Sub

    Set tableRange = ws.Range( _
        ws.Cells(TARGET_HEADER_ROW, 1), _
        ws.Cells(lastDataRow, lastColumn))

    If ws.ListObjects.Count > 0 Then

        Set lo = ws.ListObjects(1)

        ' Resize tidak mengubah urutan/nama header existing.
        lo.Resize tableRange

    Else

        Set lo = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=tableRange, _
            XlListObjectHasHeaders:=xlYes)

        selectedTableName = preferredTableName

        If Len(selectedTableName) = 0 Then
            selectedTableName = "tblDBSourcing"
        End If

        On Error Resume Next

        lo.Name = selectedTableName

        If Err.Number <> 0 Then

            Err.Clear

            lo.Name = _
                "tblDBSourcing_" & _
                Format$(Now, "hhnnss")

        End If

        On Error GoTo 0

    End If

    If Len(preferredTableStyle) > 0 Then

        lo.tableStyle = preferredTableStyle

    ElseIf Len(lo.tableStyle) = 0 Then

        lo.tableStyle = "TableStyleMedium2"

    End If

End Sub

' ============================================================
' FORMAT TARGET BERDASARKAN HEADER
' ============================================================
Private Sub FormatTargetSheetDynamic_DB( _
    ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim lastColumn As Long
    Dim headerMap As Object

    lastRow = GetLastTargetDataRow_DB(ws)

    lastColumn = LastUsedColumnInRow_DB( _
        ws, TARGET_HEADER_ROW)

    With ws.Range( _
        ws.Cells(TARGET_HEADER_ROW, 1), _
        ws.Cells(TARGET_HEADER_ROW, lastColumn))

        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlCenter

    End With

    ws.rows(TARGET_HEADER_ROW).RowHeight = 36

    Set headerMap = BuildHeaderMap_DB( _
        ws, TARGET_HEADER_ROW, lastColumn)

    SetColumnWidthByHeader_DB _
        ws, headerMap, "NO", 8

    SetColumnWidthByHeader_DB _
        ws, headerMap, "KODE UNIK", 23

    SetColumnWidthByHeader_DB _
        ws, headerMap, "POSISI", 34

    SetColumnWidthByHeader_DB _
        ws, headerMap, "REKRUTER", 15

    SetColumnWidthByHeader_DB _
        ws, headerMap, "NAMA", 28

    SetColumnWidthByHeader_DB _
        ws, headerMap, "NOMOR HP", 18

    SetColumnWidthByHeader_DB _
        ws, headerMap, "EMAIL", 28

    If lastRow >= FIRST_TARGET_DATA_ROW Then

        ws.Range( _
            ws.Cells(FIRST_TARGET_DATA_ROW, 1), _
            ws.Cells(lastRow, lastColumn)). _
            VerticalAlignment = xlCenter

    End If

End Sub

Private Sub SetColumnWidthByHeader_DB( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object, _
    ByVal canonicalHeader As String, _
    ByVal columnWidth As Double)

    If headerMap.Exists( _
        canonicalHeader) Then

        ws.Columns( _
            CLng(headerMap( _
                canonicalHeader))).columnWidth = _
            columnWidth

    End If

End Sub

' ============================================================
' WORKSHEET / LAST CELL
' ============================================================
Private Function GetOrCreateWorksheet_DB( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet

    On Error Resume Next

    Set GetOrCreateWorksheet_DB = _
        wb.Worksheets(sheetName)

    On Error GoTo 0

    If GetOrCreateWorksheet_DB Is Nothing Then

        Set GetOrCreateWorksheet_DB = _
            wb.Worksheets.Add( _
                After:=wb.Worksheets( _
                    wb.Worksheets.Count))

        GetOrCreateWorksheet_DB.Name = sheetName

    End If

End Function

Private Function LastUsedColumnInRow_DB( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long) As Long

    Dim lastColumn As Long

    lastColumn = ws.Cells( _
        rowNumber, _
        ws.Columns.Count).End(xlToLeft).Column

    If lastColumn < 1 Then
        lastColumn = 1
    End If

    LastUsedColumnInRow_DB = lastColumn

End Function

Private Function LastNonEmptyHeaderColumn_DB( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    On Error Resume Next

    Set foundCell = ws.rows( _
        TARGET_HEADER_ROW).Find( _
            What:="*", _
            After:=ws.Cells( _
                TARGET_HEADER_ROW, 1), _
            LookIn:=xlValues, _
            LookAt:=xlPart, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False)

    On Error GoTo 0

    If Not foundCell Is Nothing Then
        LastNonEmptyHeaderColumn_DB = foundCell.Column
    End If

End Function

' ============================================================
' FOLDER / FILE COLLECTION
' ============================================================
Private Function NormalizeFolderPath_DB( _
    ByVal folderPath As String) As String

    Dim result As String

    result = Trim$(folderPath)

    If Len(result) = 0 Then Exit Function

    result = Replace$(result, "/", "\")

    If Right$(result, 1) <> "\" Then
        result = result & "\"
    End If

    NormalizeFolderPath_DB = result

End Function

Private Function CollectSourceFilesFromFolders_DB( _
    ByVal folderPaths As Collection, _
    ByVal outputFilePath As String) As Collection

    Dim result As New Collection
    Dim seenFiles As Object
    Dim folderItem As Variant
    Dim folderPath As String

    Dim folderFiles As Collection
    Dim fileItem As Variant
    Dim normalizedFileKey As String

    Set seenFiles = CreateObject( _
        "Scripting.Dictionary")

    seenFiles.CompareMode = vbTextCompare

    For Each folderItem In folderPaths

        folderPath = _
            NormalizeFolderPath_DB( _
                CStr(folderItem))

        If FolderExistsForDBSourcing_DB( _
            folderPath) Then

            HeartbeatProgress _
                "RUNNING", _
                folderPath, _
                "DB Sourcing - Scan Folder", _
                "Collecting Excel files"

            Set folderFiles = _
                CollectSourceFiles_DB( _
                    folderPath, _
                    outputFilePath)

            For Each fileItem In folderFiles

                normalizedFileKey = _
                    LCase$(CStr(fileItem))

                If Not seenFiles.Exists( _
                    normalizedFileKey) Then

                    seenFiles.Add _
                        normalizedFileKey, _
                        True

                    result.Add CStr(fileItem)

                End If

            Next fileItem

        Else

            AddRunLog _
                "Warning", _
                "DB Sourcing folder tidak ditemukan: " & _
                folderPath

        End If

        Set folderFiles = Nothing

    Next folderItem

    Set CollectSourceFilesFromFolders_DB = _
        result

End Function

Private Function FolderExistsForDBSourcing_DB( _
    ByVal folderPath As String) As Boolean

    Dim fso As Object

    On Error GoTo SafeExit

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    FolderExistsForDBSourcing_DB = _
        fso.FolderExists( _
            folderPath)

SafeExit:

    Set fso = Nothing

End Function

Private Function CollectSourceFiles_DB( _
    ByVal folderPath As String, _
    ByVal outputFilePath As String) As Collection

    Dim result As New Collection
    Dim extensions As Variant
    Dim extensionItem As Variant
    Dim fileName As String
    Dim fullPath As String

    extensions = Array( _
        "*.xlsx", "*.xlsm", _
        "*.xlsb", "*.xls")

    For Each extensionItem In extensions

        fileName = Dir$( _
            folderPath & CStr(extensionItem))

        Do While Len(fileName) > 0

            If Left$(fileName, 2) <> "~$" Then

                fullPath = folderPath & fileName

                If Not SamePath_DB( _
                    fullPath, _
                    ThisWorkbook.FullName) Then

                    If Len(outputFilePath) = 0 Or _
                       Not SamePath_DB( _
                            fullPath, _
                            outputFilePath) Then

                        result.Add fullPath

                    End If

                End If

            End If

            fileName = Dir$()

        Loop

    Next extensionItem

    Set CollectSourceFiles_DB = result

End Function

' ============================================================
' SAFE OPEN / TEMP
' ============================================================
Private Function GetOpenWorkbookByFullPath_DB( _
    ByVal fullPath As String) As Workbook

    Dim wb As Workbook
    Dim currentFullName As String

    For Each wb In Application.Workbooks

        currentFullName = vbNullString

        On Error Resume Next
        currentFullName = wb.FullName
        On Error GoTo 0

        If Len(currentFullName) > 0 Then

            If SamePath_DB( _
                currentFullName, fullPath) Then

                Set GetOpenWorkbookByFullPath_DB = wb
                Exit Function

            End If

        End If

    Next wb

End Function

Private Function StageSourceFileToTemp_DB( _
    ByVal sourcePath As String) As String

    Dim fso As Object
    Dim tempFolder As String
    Dim tempFileName As String
    Dim extensionName As String
    Dim attemptNumber As Long
    Dim destinationPath As String

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    If Not fso.FileExists(sourcePath) Then

        Err.Raise _
            vbObjectError + 5201, _
            "StageSourceFileToTemp_DB", _
            "File sumber tidak ditemukan: " & _
            sourcePath

    End If

    tempFolder = Environ$("TEMP") & _
        Application.PathSeparator & _
        TEMP_SUBFOLDER_NAME

    If Not fso.FolderExists(tempFolder) Then
        fso.CreateFolder tempFolder
    End If

    extensionName = _
        fso.GetExtensionName(sourcePath)

    tempFileName = _
        "SRC_" & _
        Format$(Now, "yyyymmdd_hhnnss") & _
        "_" & _
        Format$(CLng(Timer * 1000), "00000000") & _
        "." & extensionName

    destinationPath = tempFolder & _
        Application.PathSeparator & _
        tempFileName

    For attemptNumber = 1 To 3

        On Error Resume Next
        Err.Clear

        fso.CopyFile _
            Source:=sourcePath, _
            Destination:=destinationPath, _
            OverWriteFiles:=True

        If Err.Number = 0 Then

            On Error GoTo 0

            StageSourceFileToTemp_DB = _
                destinationPath

            Exit Function

        End If

        On Error GoTo 0

        Application.StatusBar = _
            "Menunggu file OneDrive siap, percobaan " & _
            CStr(attemptNumber) & "/3..."

        DoEvents

        Application.Wait _
            Now + TimeSerial(0, 0, 1)

    Next attemptNumber

    Err.Raise _
        vbObjectError + 5202, _
        "StageSourceFileToTemp_DB", _
        "File tidak dapat disalin ke TEMP: " & _
        sourcePath

End Function

Private Sub DeleteTempFile_DB( _
    ByVal tempPath As String)

    Dim fso As Object

    If Len(tempPath) = 0 Then Exit Sub

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    On Error Resume Next

    If fso.FileExists(tempPath) Then
        fso.DeleteFile tempPath, True
    End If

    On Error GoTo 0

End Sub

' ============================================================
' PATH HELPERS
' ============================================================
Private Function SamePath_DB( _
    ByVal firstPath As String, _
    ByVal secondPath As String) As Boolean

    Dim fso As Object
    Dim normalizedFirst As String
    Dim normalizedSecond As String

    If Len(Trim$(firstPath)) = 0 Or _
       Len(Trim$(secondPath)) = 0 Then

        Exit Function

    End If

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    On Error Resume Next

    normalizedFirst = _
        fso.GetAbsolutePathName(firstPath)

    normalizedSecond = _
        fso.GetAbsolutePathName(secondPath)

    On Error GoTo 0

    SamePath_DB = _
        (StrComp( _
            normalizedFirst, _
            normalizedSecond, _
            vbTextCompare) = 0)

End Function

Private Function FileNameOnly_DB( _
    ByVal fullPath As String) As String

    Dim fso As Object

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    FileNameOnly_DB = _
        fso.GetFileName(fullPath)

End Function

' ============================================================
' SETTINGS
' ============================================================
Private Function ReadSettingValue_DB( _
    ByVal settingName As String) As String

    Dim wsSettings As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    On Error Resume Next

    Set wsSettings = _
        ThisWorkbook.Worksheets("Settings")

    On Error GoTo 0

    If wsSettings Is Nothing Then Exit Function

    lastRow = wsSettings.Cells( _
        wsSettings.rows.Count, "A").End(xlUp).Row

    For rowIndex = 1 To lastRow

        If StrComp( _
            Trim$(CStr( _
                wsSettings.Cells( _
                    rowIndex, "A").Value2)), _
            settingName, _
            vbTextCompare) = 0 Then

            ReadSettingValue_DB = Trim$(CStr( _
                wsSettings.Cells( _
                    rowIndex, "B").Value2))

            Exit Function

        End If

    Next rowIndex

End Function

' ============================================================
' LOG
' ============================================================
Private Sub PrepareDBCompileLog_DB()

    Dim wsLog As Worksheet
    Dim logHeaders As Variant

    Set wsLog = GetOrCreateWorksheet_DB( _
        ThisWorkbook, LOG_SHEET_NAME)

    wsLog.Cells.Clear

    logHeaders = Array( _
        "Waktu", _
        "File", _
        "Status", _
        "Baris Masuk", _
        "Baris Update", _
        "Baris Kosong", _
        "Instruction Row", _
        "Missing Columns / Error")

    wsLog.Range("A1").Resize( _
        1, UBound(logHeaders) + 1).Value = _
        OneDimensionalToRowArray_DB(logHeaders)

    With wsLog.Range("A1:H1")

        .Font.Bold = True
        .Interior.Color = RGB(44, 62, 80)
        .Font.Color = vbWhite

    End With

    wsLog.Columns("A").columnWidth = 20
    wsLog.Columns("B").columnWidth = 45
    wsLog.Columns("C").columnWidth = 12
    wsLog.Columns("D:G").columnWidth = 14
    wsLog.Columns("H").columnWidth = 70
    wsLog.Columns("H").WrapText = True

End Sub

Private Sub AppendDBCompileLog_DB( _
    ByVal fileName As String, _
    ByVal statusText As String, _
    ByVal importedRows As Long, _
    ByVal updatedRows As Long, _
    ByVal emptyRows As Long, _
    ByVal instructionRows As Long, _
    ByVal missingColumns As Long, _
    ByVal errorText As String)

    Dim wsLog As Worksheet
    Dim NextRow As Long
    Dim detailText As String

    Set wsLog = GetOrCreateWorksheet_DB( _
        ThisWorkbook, LOG_SHEET_NAME)

    NextRow = wsLog.Cells( _
        wsLog.rows.Count, "A").End(xlUp).Row + 1

    If missingColumns > 0 Then

        detailText = CStr(missingColumns) & _
            " kolom target tidak tersedia di source"

    End If

    If Len(errorText) > 0 Then

        If Len(detailText) > 0 Then
            detailText = detailText & " | "
        End If

        detailText = detailText & errorText

    End If

    wsLog.Cells(NextRow, "A").Value = Now
    wsLog.Cells(NextRow, "B").Value = fileName
    wsLog.Cells(NextRow, "C").Value = statusText
    wsLog.Cells(NextRow, "D").Value = importedRows
    wsLog.Cells(NextRow, "E").Value = updatedRows
    wsLog.Cells(NextRow, "F").Value = emptyRows
    wsLog.Cells(NextRow, "G").Value = instructionRows
    wsLog.Cells(NextRow, "H").Value = detailText

    wsLog.Cells( _
        NextRow, "A").numberFormat = _
        "dd/mm/yyyy hh:mm:ss"

    If UCase$(statusText) = "FAILED" Then

        wsLog.Range( _
            "A" & CStr(NextRow) & _
            ":H" & CStr(NextRow)).Interior.Color = _
            RGB(255, 199, 206)

    End If

End Sub

' ============================================================
' FORMAT DURASI
' ============================================================
Private Function FormatElapsed_DB( _
    ByVal startedAt As Double) As String

    Dim elapsedSeconds As Double

    elapsedSeconds = Timer - startedAt

    If elapsedSeconds < 0 Then
        elapsedSeconds = elapsedSeconds + 86400
    End If

    If elapsedSeconds < 60 Then

        FormatElapsed_DB = _
            Format$(elapsedSeconds, "0.0") & _
            " detik"

    Else

        FormatElapsed_DB = _
            Format$(Int(elapsedSeconds / 60), "0") & _
            " menit " & _
            Format$(elapsedSeconds Mod 60, "0") & _
            " detik"

    End If

End Function









