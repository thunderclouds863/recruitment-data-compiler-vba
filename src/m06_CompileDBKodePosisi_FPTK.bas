Attribute VB_Name = "m06_CompileDBKodePosisi_FPTK"
Option Explicit

Private Const DBK_TARGET_SHEET As String = "DB Kode Posisi"
Private Const DBK_FPTK_SHEET As String = "FPTK"

Private Const DBK_HEADER_ROW As Long = 1
Private Const DBK_FIRST_DATA_ROW As Long = 2
Private Const DBK_STATIC_COLS As Long = 10

' ============================================================
' ENSURE TARGET - BISA DIPANGGIL DARI MAIN ENGINE
'
' Menjamin sheet "DB Kode Posisi" sudah tersedia sejak awal.
' Hanya membuat/menyiapkan header A:J.
' Tidak membaca source dan tidak menghapus data existing A:J.
' ============================================================
Public Sub Ensure_DB_Kode_Posisi_Target()

    Dim wsTarget As Worksheet

    Set wsTarget = _
        DBK_GetOrCreateWorksheet( _
            DBK_TARGET_SHEET)

    DBK_PrepareStaticTarget _
        wsTarget

    Set wsTarget = Nothing

End Sub

' ============================================================
' COMPATIBILITY WRAPPER ï¿½ SATU FOLDER
' ============================================================
Public Sub Compile_DB_Kode_Posisi( _
    ByVal folderPath As String, _
    Optional ByVal showMessage As Boolean = True)

    Dim folders As New Collection

    folders.Add folderPath

    Compile_DB_Kode_Posisi_MultiFolders _
        folders, _
        showMessage

End Sub

' ============================================================
' ONE PASS ï¿½ SEMUA SUBFOLDER
'
' PERBAIKAN PENTING:
' - Dedup bukan berdasarkan KODE.
' - Dedup berdasarkan POSITION + LOCATION + BUSINESS UNIT.
' - Row source tetap masuk walaupun KODE masih kosong.
' - KODE disinkronkan dari sheet FPTK setelah seluruh file dibaca.
' ============================================================
Public Sub Compile_DB_Kode_Posisi_MultiFolders( _
    ByVal folderPaths As Collection, _
    Optional ByVal showMessage As Boolean = True)

    Dim wsTarget As Worksheet

    Dim rowByKey As Object
    Dim orderedKeys As Collection

    Dim fptkExactMap As Object
    Dim fptkPositionBUMap As Object
    Dim fptkPositionMap As Object

    Dim sourceFiles As Collection
    Dim fileItem As Variant

    Dim existingCount As Long
    Dim importedCount As Long
    Dim updatedCount As Long

    Dim filesFound As Long
    Dim filesOpened As Long
    Dim dbKodeSheetsFound As Long
    Dim fptkSheetsFound As Long

    Dim sourceRowsRead As Long
    Dim blankKeyRows As Long
    Dim skippedFileCount As Long

    Dim fptkMappingCount As Long
    Dim fptkMasterRowsRead As Long
    Dim kodeSyncedCount As Long
    Dim kodeNotFoundCount As Long
    Dim mappingConflictCount As Long

    Dim duplicateFileCount As Long

    Dim startedAt As Double
    Dim errorNumber As Long
    Dim errorText As String
    Dim summaryText As String

    On Error GoTo ErrorHandler

    startedAt = Timer

    If folderPaths Is Nothing Then

        Err.Raise _
            vbObjectError + 6901, _
            "Compile_DB_Kode_Posisi_MultiFolders", _
            "Daftar folder DB Kode Posisi belum tersedia."

    End If

    If folderPaths.Count = 0 Then

        Err.Raise _
            vbObjectError + 6902, _
            "Compile_DB_Kode_Posisi_MultiFolders", _
            "Tidak ada folder DB Kode Posisi yang akan diproses."

    End If

    Set wsTarget = _
        DBK_GetOrCreateWorksheet( _
            DBK_TARGET_SHEET)

    DBK_PrepareStaticTarget wsTarget

    Set rowByKey = _
        DBK_NewDictionary()

    Set orderedKeys = _
        New Collection

    Set fptkExactMap = _
        DBK_NewDictionary()

    Set fptkPositionBUMap = _
        DBK_NewDictionary()

    Set fptkPositionMap = _
        DBK_NewDictionary()

    DBK_LoadExistingTargetRows _
        wsTarget, _
        rowByKey, _
        orderedKeys, _
        existingCount

    ' ========================================================
    ' SUMBER UTAMA MASTER:
    ' sheet FPTK hasil compile di engine ini.
    '
    ' Versi sebelumnya hanya mengambil mapping KODE dari FPTK
    ' dan baru membuat row bila file source memiliki sheet
    ' "DB Kode Posisi". Akibatnya target kosong bila sheet
    ' tersebut tidak ada di file recruiter.
    ' ========================================================
    If DBK_SheetExists( _
        ThisWorkbook, _
        DBK_FPTK_SHEET) Then

        HeartbeatProgress _
            "RUNNING", _
            DBK_FPTK_SHEET, _
            "DB Kode Posisi - Build From FPTK", _
            "Reading compiled FPTK rows"

        DBK_LoadFPTKMappings _
            ThisWorkbook.Worksheets( _
                DBK_FPTK_SHEET), _
            fptkExactMap, _
            fptkPositionBUMap, _
            fptkPositionMap, _
            fptkMappingCount, _
            mappingConflictCount

        DBK_LoadRowsFromFPTK _
            ThisWorkbook.Worksheets( _
                DBK_FPTK_SHEET), _
            rowByKey, _
            orderedKeys, _
            importedCount, _
            updatedCount, _
            sourceRowsRead, _
            blankKeyRows, _
            fptkMasterRowsRead

    Else

        AddRunLog _
            "Warning", _
            "DB Kode Posisi | Sheet FPTK engine tidak ditemukan"

    End If

    HeartbeatProgress _
        "RUNNING", _
        CStr(folderPaths.Count) & " folders", _
        "DB Kode Posisi - Scan Files", _
        "Collecting XLSX/XLSM files"

    Set sourceFiles = _
        DBK_CollectSourceFiles( _
            folderPaths, _
            duplicateFileCount)

    filesFound = sourceFiles.Count

    For Each fileItem In sourceFiles

        HeartbeatProgress _
            "RUNNING", _
            DBK_FileNameOnly( _
                CStr(fileItem)), _
            "DB Kode Posisi - Read Source", _
            CStr(filesOpened + 1) & _
            "/" & _
            CStr(filesFound)

        DBK_ProcessOneWorkbook _
            CStr(fileItem), _
            rowByKey, _
            orderedKeys, _
            fptkExactMap, _
            fptkPositionBUMap, _
            fptkPositionMap, _
            importedCount, _
            updatedCount, _
            filesOpened, _
            dbKodeSheetsFound, _
            fptkSheetsFound, _
            sourceRowsRead, _
            blankKeyRows, _
            skippedFileCount, _
            fptkMappingCount, _
            mappingConflictCount

        DoEvents

    Next fileItem

    ' ========================================================
    ' Setelah master source terbaca, collapse placeholder FPTK
    ' yang memiliki KODE + POSITION + BUSINESS UNIT sama.
    '
    ' Source DB Kode Posisi diproses setelah row dari FPTK,
    ' sehingga row source yang lebih lengkap menjadi incoming
    ' row dan akan overwrite field placeholder yang kosong.
    ' ========================================================
    HeartbeatProgress _
        "RUNNING", _
        DBK_TARGET_SHEET, _
        "DB Kode Posisi - Collapse Duplicate Master", _
        CStr(orderedKeys.Count) & _
        " rows before collapse"

    DBK_CollapseByKodePositionBU _
        rowByKey, _
        orderedKeys

    HeartbeatProgress _
        "RUNNING", _
        DBK_TARGET_SHEET, _
        "DB Kode Posisi - Sync KODE", _
        CStr(orderedKeys.Count) & _
        " master rows | " & _
        CStr(fptkMappingCount) & _
        " FPTK mappings"

    DBK_SynchronizeKode _
        rowByKey, _
        orderedKeys, _
        fptkExactMap, _
        fptkPositionBUMap, _
        fptkPositionMap, _
        kodeSyncedCount, _
        kodeNotFoundCount

    ' ========================================================
    ' FINAL RECONCILIATION
    '
    ' Source DB Kode Posisi bisa memiliki LOCATION kosong,
    ' sedangkan row hasil FPTK memiliki LOCATION terisi.
    ' Keduanya menjadi record key berbeda.
    '
    ' Karena itu sebelum write output, isi ulang field master
    ' yang masih kosong berdasarkan Position + BU:
    ' - USER (MANAGER)
    ' - INDIRECT USER
    ' - DIRECTORATE
    ' - YEAR
    '
    ' User/Indirect juga fallback ke Mapping User.
    ' YEAR juga fallback ke FPTK compiled.
    ' ========================================================
    HeartbeatProgress _
        "RUNNING", _
        DBK_TARGET_SHEET, _
        "DB Kode Posisi - Final Reconcile", _
        CStr(orderedKeys.Count) & _
        " master rows"

    DBK_FinalReconcileMasterRows _
        rowByKey, _
        orderedKeys

    HeartbeatProgress _
        "RUNNING", _
        DBK_TARGET_SHEET, _
        "DB Kode Posisi - Write Output", _
        CStr(orderedKeys.Count) & _
        " unique POSITION/LOCATION/BU rows"

    DBK_WriteAllRows _
        wsTarget, _
        rowByKey, _
        orderedKeys

    DBK_PrepareStaticTarget wsTarget

    summaryText = _
        "DB Kode Posisi KEYSYNC | " & _
        "Folders: " & CStr(folderPaths.Count) & _
        " | Files found: " & CStr(filesFound) & _
        " | Files opened: " & CStr(filesOpened) & _
        " | DBKode sheets: " & CStr(dbKodeSheetsFound) & _
        " | FPTK sheets: " & CStr(fptkSheetsFound) & _
        " | FPTK master rows: " & CStr(fptkMasterRowsRead) & _
        " | Existing: " & CStr(existingCount) & _
        " | New: " & CStr(importedCount) & _
        " | Updated: " & CStr(updatedCount) & _
        " | KODE synced: " & CStr(kodeSyncedCount) & _
        " | KODE not found: " & CStr(kodeNotFoundCount) & _
        " | Source rows: " & CStr(sourceRowsRead) & _
        " | Blank keys: " & CStr(blankKeyRows) & _
        " | Skipped files: " & CStr(skippedFileCount) & _
        " | Mapping conflicts: " & CStr(mappingConflictCount) & _
        " | Total target: " & CStr(orderedKeys.Count) & _
        " | Duration: " & DBK_ElapsedText(startedAt)

    AddRunLog _
        IIf( _
            importedCount > 0 Or _
            updatedCount > 0, _
            "Summary", _
            "Warning"), _
        summaryText

    If showMessage Then

        MsgBox _
            DBK_BuildResultMessage( _
                folderPaths.Count, _
                filesFound, _
                filesOpened, _
                dbKodeSheetsFound, _
                fptkSheetsFound, _
                existingCount, _
                importedCount, _
                updatedCount, _
                kodeSyncedCount, _
                kodeNotFoundCount, _
                sourceRowsRead, _
                blankKeyRows, _
                skippedFileCount, _
                orderedKeys.Count, _
                startedAt), _
            IIf( _
                importedCount > 0 Or _
                updatedCount > 0, _
                vbInformation, _
                vbExclamation), _
            "Compile DB Kode Posisi"

    End If

    Exit Sub

ErrorHandler:

    errorNumber = Err.Number
    errorText = Err.Description

    AddRunLog _
        "Error", _
        "DB Kode Posisi KEYSYNC gagal | Error " & _
        CStr(errorNumber) & _
        ": " & errorText

    If showMessage Then

        MsgBox _
            "Compile DB Kode Posisi gagal." & _
            vbCrLf & _
            "Error " & CStr(errorNumber) & _
            ": " & errorText, _
            vbCritical, _
            "DB Kode Posisi"

    Else

        Err.Raise _
            errorNumber, _
            "Compile_DB_Kode_Posisi_MultiFolders", _
            errorText

    End If

End Sub

' ============================================================
' PROCESS ONE WORKBOOK
'
' Workbook dibuka satu kali:
' 1. baca mapping KODE dari FPTK;
' 2. baca master DB Kode Posisi.
' ============================================================
Private Sub DBK_ProcessOneWorkbook( _
    ByVal fullPath As String, _
    ByVal rowByKey As Object, _
    ByVal orderedKeys As Collection, _
    ByVal fptkExactMap As Object, _
    ByVal fptkPositionBUMap As Object, _
    ByVal fptkPositionMap As Object, _
    ByRef importedCount As Long, _
    ByRef updatedCount As Long, _
    ByRef filesOpened As Long, _
    ByRef dbKodeSheetsFound As Long, _
    ByRef fptkSheetsFound As Long, _
    ByRef sourceRowsRead As Long, _
    ByRef blankKeyRows As Long, _
    ByRef skippedFileCount As Long, _
    ByRef fptkMappingCount As Long, _
    ByRef mappingConflictCount As Long)

    Dim wbSource As Object
    Dim openError As String
    Dim workbookWasOpened As Boolean

    On Error GoTo FileError

    Set wbSource = _
        OpenSourceWorkbook( _
            fullPath, _
            openError)

    If wbSource Is Nothing Then

        skippedFileCount = _
            skippedFileCount + 1

        AddRunLog _
            "Error", _
            "DB Kode Posisi | Gagal buka | " & _
            DBK_FileNameOnly(fullPath) & _
            " | " & openError

        Exit Sub

    End If

    workbookWasOpened = True
    filesOpened = filesOpened + 1

    If DBK_SheetExists( _
        wbSource, _
        DBK_FPTK_SHEET) Then

        fptkSheetsFound = _
            fptkSheetsFound + 1

        DBK_LoadFPTKMappings _
            wbSource.Worksheets( _
                DBK_FPTK_SHEET), _
            fptkExactMap, _
            fptkPositionBUMap, _
            fptkPositionMap, _
            fptkMappingCount, _
            mappingConflictCount

    End If

    If DBK_SheetExists( _
        wbSource, _
        DBK_TARGET_SHEET) Then

        dbKodeSheetsFound = _
            dbKodeSheetsFound + 1

        DBK_LoadDBKodeRows _
            wbSource.Worksheets( _
                DBK_TARGET_SHEET), _
            rowByKey, _
            orderedKeys, _
            importedCount, _
            updatedCount, _
            sourceRowsRead, _
            blankKeyRows, _
            DBK_FileNameOnly(fullPath)

    Else

        AddRunLog _
            "Info", _
            "DB Kode Posisi | Sheet tidak ada | " & _
            DBK_FileNameOnly(fullPath)

    End If

SafeClose:

    If workbookWasOpened Then
        CloseSourceWorkbook wbSource
    End If

    Set wbSource = Nothing

    Exit Sub

FileError:

    skippedFileCount = _
        skippedFileCount + 1

    AddRunLog _
        "Error", _
        "DB Kode Posisi | Error file | " & _
        DBK_FileNameOnly(fullPath) & _
        " | " & Err.Description

    On Error Resume Next

    If workbookWasOpened Then
        CloseSourceWorkbook wbSource
    End If

    Set wbSource = Nothing

    On Error GoTo 0

End Sub

' ============================================================
' LOAD DB KODE POSISI SOURCE ROWS
'
' Kunci = POSITION + LOCATION + BUSINESS UNIT.
' KODE boleh kosong dan tidak menentukan apakah row masuk.
' ============================================================
Private Sub DBK_LoadDBKodeRows( _
    ByVal wsSource As Object, _
    ByVal rowByKey As Object, _
    ByVal orderedKeys As Collection, _
    ByRef importedCount As Long, _
    ByRef updatedCount As Long, _
    ByRef sourceRowsRead As Long, _
    ByRef blankKeyRows As Long, _
    ByVal sourceFileName As String)

    Dim headerRow As Long
    Dim lastColumn As Long
    Dim lastRow As Long

    Dim sourceMap As Object
    Dim sourceData As Variant

    Dim sourceRow As Long
    Dim rowData As Variant
    Dim existingRow As Variant
    Dim recordKey As String

    headerRow = _
        DBK_FindHeaderRow( _
            wsSource, _
            True)

    If headerRow = 0 Then

        AddRunLog _
            "Warning", _
            "DB Kode Posisi | Header POSITION tidak ditemukan | " & _
            sourceFileName

        Exit Sub

    End If

    lastColumn = _
        DBK_LastHeaderColumn( _
            wsSource, _
            headerRow)

    Set sourceMap = _
        DBK_BuildSourceMap( _
            wsSource, _
            headerRow, _
            lastColumn)

    If Not sourceMap.Exists( _
        "POSITION") Then

        AddRunLog _
            "Warning", _
            "DB Kode Posisi | Kolom POSITION tidak ditemukan | " & _
            sourceFileName

        Exit Sub

    End If

    lastRow = _
        DBK_LastSourceRow( _
            wsSource, _
            sourceMap)

    If lastRow < headerRow + 1 Then
        Exit Sub
    End If

    sourceData = wsSource.Range( _
        wsSource.Cells( _
            headerRow + 1, _
            1), _
        wsSource.Cells( _
            lastRow, _
            lastColumn)).Value2

    If Not IsArray(sourceData) Then
        Exit Sub
    End If

    For sourceRow = 1 _
        To UBound(sourceData, 1)

        rowData = _
            DBK_ReadStaticRow( _
                sourceData, _
                sourceRow, _
                sourceMap)

        recordKey = _
            DBK_RecordKeyFromRow( _
                rowData)

        If Len(recordKey) = 0 Then

            blankKeyRows = _
                blankKeyRows + 1

        Else

            sourceRowsRead = _
                sourceRowsRead + 1

            If rowByKey.Exists( _
                recordKey) Then

                existingRow = _
                    rowByKey(recordKey)

                rowByKey(recordKey) = _
                    DBK_MergeRows( _
                        existingRow, _
                        rowData)

                updatedCount = _
                    updatedCount + 1

            Else

                rowByKey.Add _
                    recordKey, _
                    rowData

                orderedKeys.Add _
                    recordKey

                importedCount = _
                    importedCount + 1

            End If

        End If

    Next sourceRow

End Sub


' ============================================================
' LOAD MASTER ROWS DARI FPTK ENGINE
'
' Membentuk target statis A:J langsung dari sheet FPTK:
' KODE              <- Kode Angka
' POSITION          <- Posisi
' LOCATION          <- Lokasi Kerja
' BUSINESS UNIT     <- Business Unit
' DIVISION CHRIS    <- Divisi / Divisi (Sesuai SO)
' DEPARTMENT CHRIS  <- Department
' USER (MANAGER)    <- User (Manager)
' INDIRECT USER     <- Indirect User
' DIRECTORATE       <- Direktorat
' YEAR              <- Year FPTK atau tahun dari FPTK Date
'
' Tidak menghapus target existing.
' Dedup tetap POSITION + LOCATION + BUSINESS UNIT.
' ============================================================
Private Sub DBK_EnrichUserFromExistingMaster( _
    ByRef incomingRow As Variant, _
    ByVal rowByKey As Object, _
    ByVal orderedKeys As Collection)

    Dim recordKey As Variant
    Dim candidateRow As Variant

    Dim targetPosition As String
    Dim targetBU As String
    Dim targetKode As String

    Dim candidatePosition As String
    Dim candidateBU As String
    Dim candidateKode As String

    Dim candidateUser As String
    Dim candidateIndirect As String

    Dim score As Long
    Dim bestScore As Long
    Dim bestRow As Variant
    Dim foundBest As Boolean

    If rowByKey Is Nothing Then Exit Sub
    If orderedKeys Is Nothing Then Exit Sub

    targetPosition = _
        DBK_NormalizeKey( _
            incomingRow(2))

    targetBU = _
        DBK_NormalizeKey( _
            incomingRow(4))

    targetKode = _
        DBK_NormalizeKey( _
            incomingRow(1))

    If Len(targetPosition) = 0 Then
        Exit Sub
    End If

    bestScore = -1

    For Each recordKey In orderedKeys

        candidateRow = _
            rowByKey( _
                CStr(recordKey))

        candidateUser = _
            Trim$( _
                CStr( _
                    candidateRow(7)))

        candidateIndirect = _
            Trim$( _
                CStr( _
                    candidateRow(8)))

        If Len(candidateUser) > 0 Or _
           Len(candidateIndirect) > 0 Then

            candidatePosition = _
                DBK_NormalizeKey( _
                    candidateRow(2))

            candidateBU = _
                DBK_NormalizeKey( _
                    candidateRow(4))

            candidateKode = _
                DBK_NormalizeKey( _
                    candidateRow(1))

            score = -1

            If candidatePosition = _
               targetPosition Then

                score = 1000

                If Len(targetBU) > 0 And _
                   candidateBU = targetBU Then

                    score = score + 300

                End If

                If Len(targetKode) > 0 And _
                   candidateKode = targetKode Then

                    score = score + 200

                End If

                If Len(candidateUser) > 0 And _
                   Len(candidateIndirect) > 0 Then

                    score = score + 50

                End If

            ElseIf Len(targetKode) > 0 And _
                   candidateKode = targetKode Then

                ' KODE fallback lebih rendah karena ada
                ' kemungkinan legacy duplicate/corrupt KODE.
                score = 100

            End If

            If score > bestScore Then

                bestScore = score
                bestRow = candidateRow
                foundBest = True

            End If

        End If

    Next recordKey

    If foundBest Then

        If Not DBK_HasValue( _
            incomingRow(7)) Then

            incomingRow(7) = _
                bestRow(7)

        End If

        If Not DBK_HasValue( _
            incomingRow(8)) Then

            incomingRow(8) = _
                bestRow(8)

        End If

        If Not DBK_HasValue( _
            incomingRow(9)) Then

            incomingRow(9) = _
                bestRow(9)

        End If

    End If

End Sub

Private Sub DBK_LoadRowsFromFPTK( _
    ByVal wsFPTK As Object, _
    ByVal rowByKey As Object, _
    ByVal orderedKeys As Collection, _
    ByRef importedCount As Long, _
    ByRef updatedCount As Long, _
    ByRef sourceRowsRead As Long, _
    ByRef blankKeyRows As Long, _
    ByRef fptkMasterRowsRead As Long)

    Dim headerRow As Long
    Dim lastColumn As Long
    Dim lastRow As Long

    Dim sourceMap As Object
    Dim sourceData As Variant

    Dim rowIndex As Long
    Dim rowData As Variant
    Dim existingRow As Variant
    Dim recordKey As String

    Dim uniqueCode As String
    Dim positionText As String
    Dim dateValue As Variant

    headerRow = _
        DBK_FindHeaderRow( _
            wsFPTK, _
            False)

    If headerRow = 0 Then

        AddRunLog _
            "Warning", _
            "DB Kode Posisi | Header FPTK engine tidak ditemukan"

        Exit Sub

    End If

    lastColumn = _
        DBK_LastHeaderColumn( _
            wsFPTK, _
            headerRow)

    Set sourceMap = _
        DBK_BuildSourceMap( _
            wsFPTK, _
            headerRow, _
            lastColumn)

    If Not sourceMap.Exists( _
        "POSITION") Then

        AddRunLog _
            "Warning", _
            "DB Kode Posisi | Kolom Posisi pada FPTK tidak ditemukan"

        Exit Sub

    End If

    lastRow = _
        DBK_LastSourceRow( _
            wsFPTK, _
            sourceMap)

    If lastRow < headerRow + 1 Then
        Exit Sub
    End If

    sourceData = wsFPTK.Range( _
        wsFPTK.Cells( _
            headerRow + 1, _
            1), _
        wsFPTK.Cells( _
            lastRow, _
            lastColumn)).Value2

    If Not IsArray(sourceData) Then
        Exit Sub
    End If

    For rowIndex = 1 _
        To UBound(sourceData, 1)

        positionText = _
            DBK_SourceText( _
                sourceData, _
                rowIndex, _
                sourceMap, _
                "POSITION")

        uniqueCode = _
            DBK_SourceText( _
                sourceData, _
                rowIndex, _
                sourceMap, _
                "KODE UNIK")

        If DBK_IsValidFPTKMasterRow( _
            positionText, _
            uniqueCode, _
            sourceData, _
            rowIndex, _
            sourceMap) Then

            rowData = _
                DBK_ReadStaticRow( _
                    sourceData, _
                    rowIndex, _
                    sourceMap)

            ' ==================================================
            ' Preserve USER / INDIRECT dari master existing.
            '
            ' FPTK bisa memakai LOCATION "CMD Sentul" sedangkan
            ' DB Kode Posisi existing memakai "Sentul".
            ' Karena record key berbeda, tanpa langkah ini
            ' row baru akan dibuat dengan User/Indirect kosong.
            ' ==================================================
            DBK_EnrichUserFromExistingMaster _
                rowData, _
                rowByKey, _
                orderedKeys

            ' Bila YEAR belum tersedia, ambil tahun dari FPTK Date.
            If Not DBK_HasValue( _
                rowData(10)) Then

                dateValue = _
                    DBK_SourceValue( _
                        sourceData, _
                        rowIndex, _
                        sourceMap, _
                        "FPTK DATE")

                rowData(10) = _
                    DBK_YearFromValue( _
                        dateValue)

            End If

            recordKey = _
                DBK_RecordKeyFromRow( _
                    rowData)

            If Len(recordKey) = 0 Then

                blankKeyRows = _
                    blankKeyRows + 1

            Else

                sourceRowsRead = _
                    sourceRowsRead + 1

                fptkMasterRowsRead = _
                    fptkMasterRowsRead + 1

                If rowByKey.Exists( _
                    recordKey) Then

                    existingRow = _
                        rowByKey(recordKey)

                    rowByKey(recordKey) = _
                        DBK_MergeRows( _
                            existingRow, _
                            rowData)

                    updatedCount = _
                        updatedCount + 1

                Else

                    rowByKey.Add _
                        recordKey, _
                        rowData

                    orderedKeys.Add _
                        recordKey

                    importedCount = _
                        importedCount + 1

                End If

            End If

        End If

    Next rowIndex

    AddRunLog _
        "Info", _
        "DB Kode Posisi | FPTK master rows loaded: " & _
        CStr(fptkMasterRowsRead)

End Sub

Private Function DBK_IsValidFPTKMasterRow( _
    ByVal positionText As String, _
    ByVal uniqueCode As String, _
    ByRef sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceMap As Object) As Boolean

    Dim kodeText As String
    Dim businessUnitText As String

    If Len(Trim$(positionText)) = 0 Then
        Exit Function
    End If

    ' Bila Kode Unik tersedia, harus berupa kombinasi huruf+angka.
    ' Ini menolak baris helper seperti MPRec20/CORPRec09.
    If sourceMap.Exists( _
        "KODE UNIK") Then

        If Len(uniqueCode) = 0 Then
            Exit Function
        End If

        If Not DBK_HasLetterAndDigit( _
            uniqueCode) Then

            Exit Function

        End If

    Else

        ' Fallback untuk format lama yang belum punya Kode Unik.
        kodeText = _
            DBK_SourceText( _
                sourceData, _
                sourceRow, _
                sourceMap, _
                "KODE")

        businessUnitText = _
            DBK_SourceText( _
                sourceData, _
                sourceRow, _
                sourceMap, _
                "BUSINESS UNIT")

        If Len(kodeText) = 0 And _
           Len(businessUnitText) = 0 Then

            Exit Function

        End If

    End If

    DBK_IsValidFPTKMasterRow = True

End Function

Private Function DBK_HasLetterAndDigit( _
    ByVal rawText As String) As Boolean

    Dim textValue As String
    Dim oneCharacter As String
    Dim characterIndex As Long

    Dim hasLetter As Boolean
    Dim hasDigit As Boolean

    textValue = UCase$( _
        Trim$(rawText))

    For characterIndex = 1 _
        To Len(textValue)

        oneCharacter = Mid$( _
            textValue, _
            characterIndex, _
            1)

        If oneCharacter >= "A" And _
           oneCharacter <= "Z" Then

            hasLetter = True

        ElseIf oneCharacter >= "0" And _
               oneCharacter <= "9" Then

            hasDigit = True

        End If

        If hasLetter And hasDigit Then

            DBK_HasLetterAndDigit = True
            Exit Function

        End If

    Next characterIndex

End Function

Private Function DBK_YearFromValue( _
    ByVal rawValue As Variant) As Variant

    Dim dateValue As Date
    Dim textValue As String
    Dim parts() As String
    Dim partIndex As Long
    Dim numericValue As Long

    On Error GoTo SafeExit

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    If IsDate(rawValue) Then

        dateValue = CDate(rawValue)
        DBK_YearFromValue = Year(dateValue)
        Exit Function

    End If

    If IsNumeric(rawValue) Then

        If CDbl(rawValue) > 20000 And _
           CDbl(rawValue) < 80000 Then

            dateValue = CDate(CDbl(rawValue))
            DBK_YearFromValue = Year(dateValue)

        End If

        Exit Function

    End If

    textValue = Trim$( _
        CStr(rawValue))

    textValue = Replace$( _
        textValue, _
        "-", _
        "/")

    textValue = Replace$( _
        textValue, _
        ".", _
        "/")

    parts = Split( _
        textValue, _
        "/")

    For partIndex = LBound(parts) _
        To UBound(parts)

        If IsNumeric(parts(partIndex)) Then

            numericValue = CLng( _
                parts(partIndex))

            If numericValue >= 1900 And _
               numericValue <= 2200 Then

                DBK_YearFromValue = _
                    numericValue

                Exit Function

            End If

        End If

    Next partIndex

SafeExit:

End Function

' ============================================================
' LOAD FPTK KODE MAPPINGS
'
' Prioritas saat sync:
' 1. POSITION + LOCATION + BUSINESS UNIT
' 2. POSITION + BUSINESS UNIT
' 3. POSITION saja
' ============================================================
Private Sub DBK_LoadFPTKMappings( _
    ByVal wsFPTK As Object, _
    ByVal exactMap As Object, _
    ByVal positionBUMap As Object, _
    ByVal positionMap As Object, _
    ByRef mappingCount As Long, _
    ByRef conflictCount As Long)

    Dim headerRow As Long
    Dim lastColumn As Long
    Dim lastRow As Long

    Dim sourceMap As Object
    Dim sourceData As Variant

    Dim rowIndex As Long
    Dim kodeText As String
    Dim positionText As String
    Dim locationText As String
    Dim businessUnitText As String

    Dim exactKey As String
    Dim positionBUKey As String
    Dim positionKey As String

    headerRow = _
        DBK_FindHeaderRow( _
            wsFPTK, _
            False)

    If headerRow = 0 Then
        Exit Sub
    End If

    lastColumn = _
        DBK_LastHeaderColumn( _
            wsFPTK, _
            headerRow)

    Set sourceMap = _
        DBK_BuildSourceMap( _
            wsFPTK, _
            headerRow, _
            lastColumn)

    If Not sourceMap.Exists("KODE") Or _
       Not sourceMap.Exists("POSITION") Then

        Exit Sub

    End If

    lastRow = _
        DBK_LastSourceRow( _
            wsFPTK, _
            sourceMap)

    If lastRow < headerRow + 1 Then
        Exit Sub
    End If

    sourceData = wsFPTK.Range( _
        wsFPTK.Cells( _
            headerRow + 1, _
            1), _
        wsFPTK.Cells( _
            lastRow, _
            lastColumn)).Value2

    If Not IsArray(sourceData) Then
        Exit Sub
    End If

    For rowIndex = 1 _
        To UBound(sourceData, 1)

        kodeText = DBK_SourceText( _
            sourceData, _
            rowIndex, _
            sourceMap, _
            "KODE")

        positionText = DBK_SourceText( _
            sourceData, _
            rowIndex, _
            sourceMap, _
            "POSITION")

        locationText = DBK_SourceText( _
            sourceData, _
            rowIndex, _
            sourceMap, _
            "LOCATION")

        businessUnitText = DBK_SourceText( _
            sourceData, _
            rowIndex, _
            sourceMap, _
            "BUSINESS UNIT")

        If Len(kodeText) > 0 And _
           Len(positionText) > 0 Then

            exactKey = _
                DBK_RecordKey( _
                    positionText, _
                    locationText, _
                    businessUnitText)

            positionBUKey = _
                DBK_PositionBUKey( _
                    positionText, _
                    businessUnitText)

            positionKey = _
                DBK_NormalizeKey( _
                    positionText)

            DBK_AddMappingSafe _
                exactMap, _
                exactKey, _
                kodeText, _
                mappingCount, _
                conflictCount

            DBK_AddMappingSafe _
                positionBUMap, _
                positionBUKey, _
                kodeText, _
                mappingCount, _
                conflictCount

            DBK_AddMappingSafe _
                positionMap, _
                positionKey, _
                kodeText, _
                mappingCount, _
                conflictCount

        End If

    Next rowIndex

End Sub

Private Sub DBK_AddMappingSafe( _
    ByVal mappingDictionary As Object, _
    ByVal mappingKey As String, _
    ByVal kodeValue As String, _
    ByRef mappingCount As Long, _
    ByRef conflictCount As Long)

    If Len(mappingKey) = 0 Then
        Exit Sub
    End If

    If mappingDictionary.Exists( _
        mappingKey) Then

        If StrComp( _
            CStr(mappingDictionary( _
                mappingKey)), _
            kodeValue, _
            vbTextCompare) <> 0 Then

            conflictCount = _
                conflictCount + 1

        End If

    Else

        mappingDictionary.Add _
            mappingKey, _
            kodeValue

        mappingCount = _
            mappingCount + 1

    End If

End Sub

' ============================================================
' SYNCHRONIZE KODE
' ============================================================
Private Sub DBK_FinalReconcileMasterRows( _
    ByVal rowByKey As Object, _
    ByVal orderedKeys As Collection)

    Dim mapUser As Object
    Dim mapYearExact As Object
    Dim mapYearPositionBU As Object
    Dim mapYearPosition As Object

    Dim recordKey As Variant
    Dim candidateKey As Variant

    Dim rowData As Variant
    Dim candidateRow As Variant
    Dim mapPair As Variant

    Dim positionKey As String
    Dim buKey As String
    Dim kodeKey As String
    Dim exactYearKey As String
    Dim positionBUKey As String

    Dim candidatePosition As String
    Dim candidateBU As String

    Dim bestUserScore As Long
    Dim userScore As Long

    Dim bestUser As String
    Dim bestIndirect As String
    Dim bestDirectorate As String
    Dim bestYear As Variant

    Dim candidateUser As String
    Dim candidateIndirect As String
    Dim candidateDirectorate As String
    Dim candidateYear As Variant

    Dim changedCount As Long
    Dim userFilledCount As Long
    Dim indirectFilledCount As Long
    Dim yearFilledCount As Long
    Dim directorateFilledCount As Long

    If rowByKey Is Nothing Then Exit Sub
    If orderedKeys Is Nothing Then Exit Sub

    Set mapUser = _
        DBK_BuildMappingUserFallback()

    Set mapYearExact = _
        DBK_NewDictionary()

    Set mapYearPositionBU = _
        DBK_NewDictionary()

    Set mapYearPosition = _
        DBK_NewDictionary()

    DBK_BuildFPTKYearFallbacks _
        mapYearExact, _
        mapYearPositionBU, _
        mapYearPosition

    For Each recordKey In orderedKeys

        rowData = _
            rowByKey( _
                CStr(recordKey))

        positionKey = _
            DBK_NormalizeKey( _
                rowData(2))

        buKey = _
            DBK_NormalizeKey( _
                rowData(4))

        kodeKey = _
            DBK_NormalizeKey( _
                rowData(1))

        If Len(positionKey) = 0 Then
            GoTo NextRecord
        End If

        positionBUKey = _
            positionKey & _
            "|" & _
            buKey

        exactYearKey = _
            kodeKey & _
            "|" & _
            positionKey & _
            "|" & _
            buKey

        bestUserScore = -1
        bestUser = ""
        bestIndirect = ""
        bestDirectorate = ""
        bestYear = Empty

        ' ----------------------------------------------------
        ' 1. Cari sesama final master row dengan Position sama.
        '    BU sama diberi prioritas tinggi.
        '    LOCATION sengaja tidak dipakai.
        ' ----------------------------------------------------
        For Each candidateKey In orderedKeys

            If CStr(candidateKey) <> _
               CStr(recordKey) Then

                candidateRow = _
                    rowByKey( _
                        CStr(candidateKey))

                candidatePosition = _
                    DBK_NormalizeKey( _
                        candidateRow(2))

                If candidatePosition = _
                   positionKey Then

                    candidateBU = _
                        DBK_NormalizeKey( _
                            candidateRow(4))

                    userScore = 1000

                    If Len(buKey) > 0 And _
                       candidateBU = buKey Then

                        userScore = _
                            userScore + 500

                    End If

                    candidateUser = _
                        Trim$( _
                            CStr( _
                                candidateRow(7)))

                    candidateIndirect = _
                        Trim$( _
                            CStr( _
                                candidateRow(8)))

                    candidateDirectorate = _
                        Trim$( _
                            CStr( _
                                candidateRow(9)))

                    candidateYear = _
                        candidateRow(10)

                    If Len(candidateUser) > 0 And _
                       Len(candidateIndirect) > 0 Then

                        userScore = _
                            userScore + 200

                    ElseIf Len(candidateUser) > 0 Or _
                           Len(candidateIndirect) > 0 Then

                        userScore = _
                            userScore + 100

                    End If

                    If DBK_HasValue( _
                        candidateYear) Then

                        userScore = _
                            userScore + 50

                    End If

                    If userScore > _
                       bestUserScore Then

                        bestUserScore = _
                            userScore

                        bestUser = _
                            candidateUser

                        bestIndirect = _
                            candidateIndirect

                        bestDirectorate = _
                            candidateDirectorate

                        bestYear = _
                            candidateYear

                    End If

                End If

            End If

        Next candidateKey

        ' ----------------------------------------------------
        ' 2. Isi dari sesama master row.
        ' ----------------------------------------------------
        If Not DBK_HasValue( _
            rowData(7)) And _
           Len(bestUser) > 0 Then

            rowData(7) = _
                bestUser

            userFilledCount = _
                userFilledCount + 1

        End If

        If Not DBK_HasValue( _
            rowData(8)) And _
           Len(bestIndirect) > 0 Then

            rowData(8) = _
                bestIndirect

            indirectFilledCount = _
                indirectFilledCount + 1

        End If

        If Not DBK_HasValue( _
            rowData(9)) And _
           Len(bestDirectorate) > 0 Then

            rowData(9) = _
                bestDirectorate

            directorateFilledCount = _
                directorateFilledCount + 1

        End If

        If Not DBK_HasValue( _
            rowData(10)) And _
           DBK_HasValue( _
                bestYear) Then

            rowData(10) = _
                bestYear

            yearFilledCount = _
                yearFilledCount + 1

        End If

        ' ----------------------------------------------------
        ' 3. USER / INDIRECT fallback Mapping User.
        '    Key = Position exact.
        ' ----------------------------------------------------
        If Not mapUser Is Nothing Then

            If mapUser.Exists( _
                positionKey) Then

                mapPair = _
                    mapUser( _
                        positionKey)

                If Not DBK_HasValue( _
                    rowData(7)) Then

                    If Len(Trim$(CStr( _
                        mapPair(0)))) > 0 Then

                        rowData(7) = _
                            mapPair(0)

                        userFilledCount = _
                            userFilledCount + 1

                    End If

                End If

                If Not DBK_HasValue( _
                    rowData(8)) Then

                    If Len(Trim$(CStr( _
                        mapPair(1)))) > 0 Then

                        rowData(8) = _
                            mapPair(1)

                        indirectFilledCount = _
                            indirectFilledCount + 1

                    End If

                End If

                If Not DBK_HasValue( _
                    rowData(9)) Then

                    If Len(Trim$(CStr( _
                        mapPair(2)))) > 0 Then

                        rowData(9) = _
                            mapPair(2)

                        directorateFilledCount = _
                            directorateFilledCount + 1

                    End If

                End If

            End If

        End If

        ' ----------------------------------------------------
        ' 4. YEAR fallback dari FPTK compiled.
        '
        ' Priority:
        ' KODE + POSITION + BU
        ' POSITION + BU
        ' POSITION
        ' ----------------------------------------------------
        If Not DBK_HasValue( _
            rowData(10)) Then

            If mapYearExact.Exists( _
                exactYearKey) Then

                rowData(10) = _
                    mapYearExact( _
                        exactYearKey)

                yearFilledCount = _
                    yearFilledCount + 1

            ElseIf mapYearPositionBU.Exists( _
                positionBUKey) Then

                rowData(10) = _
                    mapYearPositionBU( _
                        positionBUKey)

                yearFilledCount = _
                    yearFilledCount + 1

            ElseIf mapYearPosition.Exists( _
                positionKey) Then

                rowData(10) = _
                    mapYearPosition( _
                        positionKey)

                yearFilledCount = _
                    yearFilledCount + 1

            End If

        End If

        rowByKey( _
            CStr(recordKey)) = _
                rowData

        changedCount = _
            changedCount + 1

NextRecord:

    Next recordKey

    AddRunLog _
        "Info", _
        "DB Kode Posisi Final Reconcile | Rows: " & _
        CStr(changedCount) & _
        " | User filled: " & _
        CStr(userFilledCount) & _
        " | Indirect filled: " & _
        CStr(indirectFilledCount) & _
        " | Directorate filled: " & _
        CStr(directorateFilledCount) & _
        " | Year filled: " & _
        CStr(yearFilledCount)

    Set mapUser = Nothing
    Set mapYearExact = Nothing
    Set mapYearPositionBU = Nothing
    Set mapYearPosition = Nothing

End Sub

Private Function DBK_BuildMappingUserFallback() As Object

    Dim result As Object
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim rowIndex As Long

    Dim positionKey As String
    Dim userText As String
    Dim indirectText As String
    Dim directorateText As String

    Dim existingPair As Variant
    Dim existingScore As Long
    Dim newScore As Long

    Set result = _
        DBK_NewDictionary()

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            "Mapping User")

    On Error GoTo 0

    If ws Is Nothing Then

        Set DBK_BuildMappingUserFallback = _
            result

        Exit Function

    End If

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            1).End( _
                xlUp).Row

    For rowIndex = 2 To lastRow

        positionKey = _
            DBK_NormalizeKey( _
                ws.Cells( _
                    rowIndex, _
                    1).Value2)

        userText = _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        rowIndex, _
                        2).Value2))

        indirectText = _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        rowIndex, _
                        3).Value2))

        directorateText = _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        rowIndex, _
                        4).Value2))

        If Len(positionKey) > 0 Then

            newScore = 0

            If Len(userText) > 0 Then
                newScore = newScore + 100
            End If

            If Len(indirectText) > 0 Then
                newScore = newScore + 100
            End If

            If Len(directorateText) > 0 Then
                newScore = newScore + 25
            End If

            If result.Exists( _
                positionKey) Then

                existingPair = _
                    result( _
                        positionKey)

                existingScore = 0

                If Len(Trim$(CStr( _
                    existingPair(0)))) > 0 Then

                    existingScore = _
                        existingScore + 100

                End If

                If Len(Trim$(CStr( _
                    existingPair(1)))) > 0 Then

                    existingScore = _
                        existingScore + 100

                End If

                If Len(Trim$(CStr( _
                    existingPair(2)))) > 0 Then

                    existingScore = _
                        existingScore + 25

                End If

                If newScore > existingScore Then

                    result( _
                        positionKey) = _
                            Array( _
                                userText, _
                                indirectText, _
                                directorateText)

                End If

            Else

                result.Add _
                    positionKey, _
                    Array( _
                        userText, _
                        indirectText, _
                        directorateText)

            End If

        End If

    Next rowIndex

    Set DBK_BuildMappingUserFallback = _
        result

    Set ws = Nothing

End Function

Private Sub DBK_BuildFPTKYearFallbacks( _
    ByVal exactMap As Object, _
    ByVal positionBUMap As Object, _
    ByVal positionMap As Object)

    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastColumn As Long
    Dim lastRow As Long

    Dim sourceMap As Object
    Dim sourceData As Variant

    Dim rowIndex As Long

    Dim kodeKey As String
    Dim positionKey As String
    Dim buKey As String

    Dim exactKey As String
    Dim positionBUKey As String

    Dim dateValue As Variant
    Dim yearValue As Variant

    If exactMap Is Nothing Then Exit Sub
    If positionBUMap Is Nothing Then Exit Sub
    If positionMap Is Nothing Then Exit Sub

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            DBK_FPTK_SHEET)

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    headerRow = _
        DBK_FindHeaderRow( _
            ws, _
            False)

    If headerRow = 0 Then
        Exit Sub
    End If

    lastColumn = _
        DBK_LastHeaderColumn( _
            ws, _
            headerRow)

    Set sourceMap = _
        DBK_BuildSourceMap( _
            ws, _
            headerRow, _
            lastColumn)

    If Not sourceMap.Exists( _
        "POSITION") Then

        GoTo SafeExit

    End If

    lastRow = _
        DBK_LastSourceRow( _
            ws, _
            sourceMap)

    If lastRow < headerRow + 1 Then
        GoTo SafeExit
    End If

    sourceData = _
        ws.Range( _
            ws.Cells( _
                headerRow + 1, _
                1), _
            ws.Cells( _
                lastRow, _
                lastColumn)).Value2

    If Not IsArray(sourceData) Then
        GoTo SafeExit
    End If

    For rowIndex = 1 To UBound( _
        sourceData, _
        1)

        positionKey = _
            DBK_NormalizeKey( _
                DBK_SourceValue( _
                    sourceData, _
                    rowIndex, _
                    sourceMap, _
                    "POSITION"))

        If Len(positionKey) > 0 Then

            kodeKey = _
                DBK_NormalizeKey( _
                    DBK_SourceValue( _
                        sourceData, _
                        rowIndex, _
                        sourceMap, _
                        "KODE"))

            buKey = _
                DBK_NormalizeKey( _
                    DBK_SourceValue( _
                        sourceData, _
                        rowIndex, _
                        sourceMap, _
                        "BUSINESS UNIT"))

            dateValue = _
                DBK_SourceValue( _
                    sourceData, _
                    rowIndex, _
                    sourceMap, _
                    "FPTK DATE")

            yearValue = _
                DBK_YearFromValue( _
                    dateValue)

            If DBK_HasValue( _
                yearValue) Then

                exactKey = _
                    kodeKey & _
                    "|" & _
                    positionKey & _
                    "|" & _
                    buKey

                positionBUKey = _
                    positionKey & _
                    "|" & _
                    buKey

                DBK_StoreLatestYear _
                    exactMap, _
                    exactKey, _
                    yearValue

                DBK_StoreLatestYear _
                    positionBUMap, _
                    positionBUKey, _
                    yearValue

                DBK_StoreLatestYear _
                    positionMap, _
                    positionKey, _
                    yearValue

            End If

        End If

    Next rowIndex

SafeExit:

    Set sourceMap = Nothing
    Set ws = Nothing

End Sub

Private Sub DBK_StoreLatestYear( _
    ByVal targetMap As Object, _
    ByVal mapKey As String, _
    ByVal yearValue As Variant)

    Dim numericYear As Long
    Dim existingYear As Long

    If targetMap Is Nothing Then Exit Sub
    If Len(mapKey) = 0 Then Exit Sub
    If Not IsNumeric(yearValue) Then Exit Sub

    numericYear = _
        CLng( _
            yearValue)

    If numericYear < 1900 Or _
       numericYear > 2200 Then

        Exit Sub

    End If

    If targetMap.Exists( _
        mapKey) Then

        existingYear = _
            CLng( _
                targetMap( _
                    mapKey))

        If numericYear > _
           existingYear Then

            targetMap( _
                mapKey) = _
                    numericYear

        End If

    Else

        targetMap.Add _
            mapKey, _
            numericYear

    End If

End Sub
Private Sub DBK_CollapseByKodePositionBU( _
    ByVal rowByKey As Object, _
    ByVal orderedKeys As Collection)

    Dim identityMap As Object

    Dim rowIndex As Long
    Dim currentKey As String
    Dim masterKey As String
    Dim identityKey As String

    Dim currentRow As Variant
    Dim masterRow As Variant

    Dim kodeKey As String
    Dim positionKey As String
    Dim buKey As String

    Dim collapsedCount As Long

    If rowByKey Is Nothing Then Exit Sub
    If orderedKeys Is Nothing Then Exit Sub

    Set identityMap = _
        DBK_NewDictionary()

    rowIndex = 1

    Do While rowIndex <= _
        orderedKeys.Count

        currentKey = _
            CStr( _
                orderedKeys( _
                    rowIndex))

        currentRow = _
            rowByKey( _
                currentKey)

        kodeKey = _
            DBK_NormalizeKey( _
                DBK_Text( _
                    currentRow(1)))

        positionKey = _
            DBK_NormalizeKey( _
                DBK_Text( _
                    currentRow(2)))

        buKey = _
            DBK_NormalizeKey( _
                DBK_Text( _
                    currentRow(4)))

        If Len(kodeKey) > 0 And _
           Len(positionKey) > 0 Then

            identityKey = _
                kodeKey & _
                "|" & _
                positionKey & _
                "|" & _
                buKey

            If identityMap.Exists( _
                identityKey) Then

                masterKey = _
                    CStr( _
                        identityMap( _
                            identityKey))

                masterRow = _
                    rowByKey( _
                        masterKey)

                ' currentRow muncul lebih belakangan.
                ' Karena source DB Kode Posisi dibaca setelah FPTK,
                ' nilai source yang nonblank harus menang.
                masterRow = _
                    DBK_MergeRows( _
                        masterRow, _
                        currentRow)

                rowByKey( _
                    masterKey) = _
                        masterRow

                If rowByKey.Exists( _
                    currentKey) Then

                    rowByKey.Remove _
                        currentKey

                End If

                orderedKeys.Remove _
                    rowIndex

                collapsedCount = _
                    collapsedCount + 1

                ' Jangan increment rowIndex karena Collection
                ' sudah bergeser ke item berikutnya.

            Else

                identityMap.Add _
                    identityKey, _
                    currentKey

                rowIndex = _
                    rowIndex + 1

            End If

        Else

            rowIndex = _
                rowIndex + 1

        End If

    Loop

    AddRunLog _
        "Info", _
        "DB Kode Posisi collapse KODE+POSITION+BU | " & _
        "Collapsed: " & _
        CStr(collapsedCount) & _
        " | Remaining: " & _
        CStr(orderedKeys.Count)

    Set identityMap = Nothing

End Sub


Private Sub DBK_SynchronizeKode( _
    ByVal rowByKey As Object, _
    ByVal orderedKeys As Collection, _
    ByVal exactMap As Object, _
    ByVal positionBUMap As Object, _
    ByVal positionMap As Object, _
    ByRef syncedCount As Long, _
    ByRef notFoundCount As Long)

    Dim recordKey As Variant
    Dim rowData As Variant

    Dim positionText As String
    Dim locationText As String
    Dim businessUnitText As String

    Dim exactKey As String
    Dim positionBUKey As String
    Dim positionKey As String
    Dim kodeValue As String

    For Each recordKey In orderedKeys

        rowData = _
            rowByKey( _
                CStr(recordKey))

        positionText = _
            DBK_Text(rowData(2))

        locationText = _
            DBK_Text(rowData(3))

        businessUnitText = _
            DBK_Text(rowData(4))

        exactKey = _
            DBK_RecordKey( _
                positionText, _
                locationText, _
                businessUnitText)

        positionBUKey = _
            DBK_PositionBUKey( _
                positionText, _
                businessUnitText)

        positionKey = _
            DBK_NormalizeKey( _
                positionText)

        kodeValue = vbNullString

        If exactMap.Exists( _
            exactKey) Then

            kodeValue = CStr( _
                exactMap(exactKey))

        ElseIf positionBUMap.Exists( _
            positionBUKey) Then

            kodeValue = CStr( _
                positionBUMap( _
                    positionBUKey))

        ElseIf positionMap.Exists( _
            positionKey) Then

            kodeValue = CStr( _
                positionMap( _
                    positionKey))

        End If

        If Len(kodeValue) > 0 Then

            rowData(1) = kodeValue

            rowByKey( _
                CStr(recordKey)) = _
                    rowData

            syncedCount = _
                syncedCount + 1

        Else

            notFoundCount = _
                notFoundCount + 1

        End If

    Next recordKey

End Sub

' ============================================================
' TARGET EXISTING DATA
' ============================================================
Private Sub DBK_LoadExistingTargetRows( _
    ByVal ws As Worksheet, _
    ByVal rowByKey As Object, _
    ByVal orderedKeys As Collection, _
    ByRef existingCount As Long)

    Dim lastRow As Long
    Dim existingData As Variant

    Dim rowIndex As Long
    Dim rowData As Variant
    Dim recordKey As String

    lastRow = _
        DBK_LastDataRow( _
            ws)

    If lastRow < DBK_FIRST_DATA_ROW Then
        Exit Sub
    End If

    existingData = ws.Range( _
        ws.Cells( _
            DBK_FIRST_DATA_ROW, _
            1), _
        ws.Cells( _
            lastRow, _
            DBK_STATIC_COLS)).Value2

    If Not IsArray(existingData) Then
        Exit Sub
    End If

    For rowIndex = 1 _
        To UBound(existingData, 1)

        rowData = _
            DBK_RowFromArray( _
                existingData, _
                rowIndex)

        recordKey = _
            DBK_RecordKeyFromRow( _
                rowData)

        If Len(recordKey) > 0 Then

            If Not rowByKey.Exists( _
                recordKey) Then

                rowByKey.Add _
                    recordKey, _
                    rowData

                orderedKeys.Add _
                    recordKey

                existingCount = _
                    existingCount + 1

            Else

                rowByKey(recordKey) = _
                    DBK_MergeRows( _
                        rowByKey(recordKey), _
                        rowData)

            End If

        End If

    Next rowIndex

End Sub

' ============================================================
' WRITE TARGET
' ============================================================
Private Sub DBK_WriteAllRows( _
    ByVal ws As Worksheet, _
    ByVal rowByKey As Object, _
    ByVal orderedKeys As Collection)

    Dim oldLastRow As Long
    Dim outputData() As Variant

    Dim recordKey As Variant
    Dim rowData As Variant

    Dim rowIndex As Long
    Dim columnIndex As Long

    oldLastRow = _
        DBK_LastDataRow( _
            ws)

    If oldLastRow >= _
       DBK_FIRST_DATA_ROW Then

        ws.Range( _
            ws.Cells( _
                DBK_FIRST_DATA_ROW, _
                1), _
            ws.Cells( _
                oldLastRow, _
                DBK_STATIC_COLS)).ClearContents

    End If

    If orderedKeys.Count = 0 Then
        Exit Sub
    End If

    ReDim outputData( _
        1 To orderedKeys.Count, _
        1 To DBK_STATIC_COLS)

    For Each recordKey In orderedKeys

        rowIndex = rowIndex + 1

        rowData = rowByKey( _
            CStr(recordKey))

        For columnIndex = 1 _
            To DBK_STATIC_COLS

            outputData( _
                rowIndex, _
                columnIndex) = _
                    rowData(columnIndex)

        Next columnIndex

    Next recordKey

    ws.Cells( _
        DBK_FIRST_DATA_ROW, _
        1).Resize( _
            orderedKeys.Count, _
            DBK_STATIC_COLS).Value2 = _
                outputData

End Sub

' ============================================================
' SOURCE FILE COLLECTION
' ============================================================
Private Function DBK_CollectSourceFiles( _
    ByVal folderPaths As Collection, _
    ByRef duplicateFileCount As Long) As Collection

    Dim result As New Collection
    Dim seenFiles As Object
    Dim fso As Object

    Dim folderItem As Variant
    Dim folderObject As Object
    Dim fileObject As Object

    Dim folderPath As String
    Dim extensionText As String
    Dim fullPath As String
    Dim normalizedPath As String

    Set seenFiles = _
        DBK_NewDictionary()

    Set fso = _
        CreateObject( _
            "Scripting.FileSystemObject")

    For Each folderItem In folderPaths

        folderPath = _
            CStr(folderItem)

        If fso.FolderExists( _
            folderPath) Then

            Set folderObject = _
                fso.GetFolder( _
                    folderPath)

            For Each fileObject In _
                folderObject.files

                If Left$( _
                    CStr(fileObject.Name), _
                    2) <> "~$" Then

                    extensionText = _
                        LCase$( _
                            fso.GetExtensionName( _
                                CStr(fileObject.Name)))

                    If extensionText = "xlsx" Or _
                       extensionText = "xlsm" Then

                        fullPath = _
                            CStr(fileObject.Path)

                        normalizedPath = _
                            LCase$(fullPath)

                        If seenFiles.Exists( _
                            normalizedPath) Then

                            duplicateFileCount = _
                                duplicateFileCount + 1

                        Else

                            seenFiles.Add _
                                normalizedPath, _
                                True

                            result.Add _
                                fullPath

                        End If

                    End If

                End If

            Next fileObject

        Else

            AddRunLog _
                "Warning", _
                "DB Kode Posisi folder tidak ditemukan: " & _
                folderPath

        End If

        Set folderObject = Nothing

    Next folderItem

    Set DBK_CollectSourceFiles = _
        result

End Function

' ============================================================
' HEADER MAPPING
' ============================================================
Private Function DBK_FindHeaderRow( _
    ByVal ws As Object, _
    ByVal isDBKodeSheet As Boolean) As Long

    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim lastColumn As Long

    Dim canonicalHeader As String
    Dim hasPosition As Boolean
    Dim hasKode As Boolean

    For rowIndex = 1 To 15

        lastColumn = _
            DBK_LastHeaderColumn( _
                ws, _
                rowIndex)

        If lastColumn > 150 Then
            lastColumn = 150
        End If

        hasPosition = False
        hasKode = False

        For columnIndex = 1 To lastColumn

            canonicalHeader = _
                DBK_CanonicalHeader( _
                    ws.Cells( _
                        rowIndex, _
                        columnIndex).Value2)

            If canonicalHeader = _
               "POSITION" Then

                hasPosition = True

            ElseIf canonicalHeader = _
                   "KODE" Then

                hasKode = True

            End If

        Next columnIndex

        If isDBKodeSheet Then

            If hasPosition Then

                DBK_FindHeaderRow = _
                    rowIndex

                Exit Function

            End If

        Else

            If hasPosition And hasKode Then

                DBK_FindHeaderRow = _
                    rowIndex

                Exit Function

            End If

        End If

    Next rowIndex

End Function

Private Function DBK_BuildSourceMap( _
    ByVal ws As Object, _
    ByVal headerRow As Long, _
    ByVal lastColumn As Long) As Object

    Dim result As Object
    Dim columnIndex As Long
    Dim canonicalHeader As String

    Set result = _
        DBK_NewDictionary()

    For columnIndex = 1 To lastColumn

        canonicalHeader = _
            DBK_CanonicalHeader( _
                ws.Cells( _
                    headerRow, _
                    columnIndex).Value2)

        If Len(canonicalHeader) > 0 Then

            If Not result.Exists( _
                canonicalHeader) Then

                result.Add _
                    canonicalHeader, _
                    columnIndex

            End If

        End If

    Next columnIndex

    Set DBK_BuildSourceMap = _
        result

End Function
Private Function DBK_CanonicalHeader( _
    ByVal rawHeader As Variant) As String

    Dim normalizedHeader As String

    normalizedHeader = _
        DBK_NormalizeHeader( _
            rawHeader)

    If Len(normalizedHeader) = 0 Then
        Exit Function
    End If

    ' ========================================================
    ' IMPORTANT:
    ' Source DB Kode Posisi memakai header panjang, contoh:
    '
    ' Kode Angka
    ' (tidak pakai tanda petik ')
    '
    ' POSITION (Sesuai FPTK, khusus sales diberikan
    ' tambahan "- lokasi")
    '
    ' Karena itu jangan hanya exact-match.
    ' Header utama dikenali berdasarkan prefix/kata awal.
    ' ========================================================

    ' KODE / POSITION CODE harus dicek sebelum POSITION.
    If normalizedHeader = "KODE" Or _
       normalizedHeader Like "KODE ANGKA*" Or _
       normalizedHeader Like "KODE POSISI*" Or _
       normalizedHeader Like "POSITION CODE*" Then

        DBK_CanonicalHeader = _
            "KODE"

        Exit Function

    End If

    If normalizedHeader = "KODE UNIK" Or _
       normalizedHeader Like "UNIQUE CODE*" Or _
       normalizedHeader Like "UNIQUE KEY*" Then

        DBK_CanonicalHeader = _
            "KODE UNIK"

        Exit Function

    End If

    If normalizedHeader = "POSITION" Or _
       normalizedHeader = "POSISI" Or _
       normalizedHeader Like "POSITION *" Or _
       normalizedHeader Like "POSISI *" Or _
       normalizedHeader Like "NAMA POSISI*" Then

        DBK_CanonicalHeader = _
            "POSITION"

        Exit Function

    End If

    If normalizedHeader = "LOCATION" Or _
       normalizedHeader = "LOKASI" Or _
       normalizedHeader Like "LOCATION *" Or _
       normalizedHeader Like "LOKASI KERJA*" Then

        DBK_CanonicalHeader = _
            "LOCATION"

        Exit Function

    End If

    If normalizedHeader = "BUSINESS UNIT" Or _
       normalizedHeader = "BUSINESSUNIT" Or _
       normalizedHeader = "BU" Or _
       normalizedHeader Like "BUSINESS UNIT *" Then

        DBK_CanonicalHeader = _
            "BUSINESS UNIT"

        Exit Function

    End If

    If normalizedHeader = "DIVISION CHRIS" Or _
       normalizedHeader = "DIVISION" Or _
       normalizedHeader = "DIVISI CHRIS" Or _
       normalizedHeader = "DIVISI" Or _
       normalizedHeader Like "DIVISION CHRIS *" Or _
       normalizedHeader Like "DIVISI CHRIS *" Then

        DBK_CanonicalHeader = _
            "DIVISION CHRIS"

        Exit Function

    End If

    If normalizedHeader = "DEPARTMENT CHRIS" Or _
       normalizedHeader = "DEPARTMENT" Or _
       normalizedHeader = "DEPARTEMEN CHRIS" Or _
       normalizedHeader = "DEPARTEMEN" Or _
       normalizedHeader Like "DEPARTMENT CHRIS *" Or _
       normalizedHeader Like "DEPARTEMEN CHRIS *" Then

        DBK_CanonicalHeader = _
            "DEPARTMENT CHRIS"

        Exit Function

    End If

    If normalizedHeader = "USER MANAGER" Or _
       normalizedHeader = "USER (MANAGER)" Or _
       normalizedHeader = "USER" Or _
       normalizedHeader = "MANAGER" Or _
       normalizedHeader Like "USER (MANAGER) *" Or _
       normalizedHeader Like "USER MANAGER *" Then

        DBK_CanonicalHeader = _
            "USER (MANAGER)"

        Exit Function

    End If

    If normalizedHeader = "INDIRECT USER" Or _
       normalizedHeader = "INDIRECTUSER" Or _
       normalizedHeader Like "INDIRECT USER *" Then

        DBK_CanonicalHeader = _
            "INDIRECT USER"

        Exit Function

    End If

    If normalizedHeader = "DIRECTORATE" Or _
       normalizedHeader = "DIREKTORAT" Or _
       normalizedHeader Like "DIRECTORATE *" Or _
       normalizedHeader Like "DIREKTORAT *" Then

        DBK_CanonicalHeader = _
            "DIRECTORATE"

        Exit Function

    End If

    If normalizedHeader = "YEAR" Or _
       normalizedHeader = "TAHUN" Or _
       normalizedHeader Like "YEAR FPTK*" Or _
       normalizedHeader Like "TAHUN FPTK*" Or _
       normalizedHeader Like "YEAR *" Then

        DBK_CanonicalHeader = _
            "YEAR"

        Exit Function

    End If

    If normalizedHeader = "FPTK DATE REAL" Or _
       normalizedHeader = "FPTK DATE KODE" Or _
       normalizedHeader = "FPTK DATE CODE" Or _
       normalizedHeader = "FPTK DATE PAKE INI" Or _
       normalizedHeader = "FPTK DATE" Or _
       normalizedHeader Like "FPTK DATE *" Then

        DBK_CanonicalHeader = _
            "FPTK DATE"

        Exit Function

    End If

End Function

Private Function DBK_NormalizeHeader( _
    ByVal rawHeader As Variant) As String

    Dim resultText As String

    If IsError(rawHeader) Or _
       IsNull(rawHeader) Or _
       IsEmpty(rawHeader) Then

        Exit Function

    End If

    resultText = _
        UCase$( _
            Trim$( _
                CStr(rawHeader)))

    resultText = _
        Replace$( _
            resultText, _
            Chr$(160), _
            " ")

    resultText = _
        Replace$( _
            resultText, _
            vbCr, _
            " ")

    resultText = _
        Replace$( _
            resultText, _
            vbLf, _
            " ")

    resultText = _
        Replace$( _
            resultText, _
            "_", _
            " ")

    resultText = _
        Replace$( _
            resultText, _
            "-", _
            " ")

    Do While InStr( _
        1, _
        resultText, _
        "  ", _
        vbBinaryCompare) > 0

        resultText = _
            Replace$( _
                resultText, _
                "  ", _
                " ")

    Loop

    DBK_NormalizeHeader = _
        Trim$(resultText)

End Function

' ============================================================
' ROW HELPERS
' ============================================================
Private Function DBK_ReadStaticRow( _
    ByRef sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceMap As Object) As Variant

    Dim rowData(1 To 10) As Variant
    Dim columnIndex As Long

    For columnIndex = 1 _
        To DBK_STATIC_COLS

        rowData(columnIndex) = _
            DBK_SourceValue( _
                sourceData, _
                sourceRow, _
                sourceMap, _
                DBK_TargetHeader( _
                    columnIndex))

    Next columnIndex

    DBK_ReadStaticRow = _
        rowData

End Function

Private Function DBK_RowFromArray( _
    ByRef sourceData As Variant, _
    ByVal sourceRow As Long) As Variant

    Dim rowData(1 To 10) As Variant
    Dim columnIndex As Long

    For columnIndex = 1 _
        To DBK_STATIC_COLS

        rowData(columnIndex) = _
            sourceData( _
                sourceRow, _
                columnIndex)

    Next columnIndex

    DBK_RowFromArray = _
        rowData

End Function

Private Function DBK_MergeRows( _
    ByVal existingRow As Variant, _
    ByVal incomingRow As Variant) As Variant

    Dim resultRow(1 To 10) As Variant
    Dim columnIndex As Long

    For columnIndex = 1 _
        To DBK_STATIC_COLS

        If DBK_HasValue( _
            incomingRow(columnIndex)) Then

            resultRow(columnIndex) = _
                incomingRow(columnIndex)

        Else

            resultRow(columnIndex) = _
                existingRow(columnIndex)

        End If

    Next columnIndex

    DBK_MergeRows = _
        resultRow

End Function

Private Function DBK_RecordKeyFromRow( _
    ByVal rowData As Variant) As String

    DBK_RecordKeyFromRow = _
        DBK_RecordKey( _
            DBK_Text(rowData(2)), _
            DBK_Text(rowData(3)), _
            DBK_Text(rowData(4)))

End Function

Private Function DBK_RecordKey( _
    ByVal positionText As String, _
    ByVal locationText As String, _
    ByVal businessUnitText As String) As String

    Dim positionKey As String
    Dim locationKey As String
    Dim businessUnitKey As String

    positionKey = _
        DBK_NormalizeKey( _
            positionText)

    locationKey = _
        DBK_NormalizeKey( _
            locationText)

    businessUnitKey = _
        DBK_NormalizeKey( _
            businessUnitText)

    If Len(positionKey) = 0 Then
        Exit Function
    End If

    DBK_RecordKey = _
        positionKey & "|" & _
        locationKey & "|" & _
        businessUnitKey

End Function

Private Function DBK_PositionBUKey( _
    ByVal positionText As String, _
    ByVal businessUnitText As String) As String

    Dim positionKey As String
    Dim businessUnitKey As String

    positionKey = _
        DBK_NormalizeKey( _
            positionText)

    businessUnitKey = _
        DBK_NormalizeKey( _
            businessUnitText)

    If Len(positionKey) = 0 Then
        Exit Function
    End If

    DBK_PositionBUKey = _
        positionKey & "|" & _
        businessUnitKey

End Function

Private Function DBK_NormalizeKey( _
    ByVal rawText As String) As String

    Dim resultText As String

    resultText = _
        UCase$( _
            Trim$(rawText))

    resultText = _
        Replace$( _
            resultText, _
            Chr$(160), _
            " ")

    resultText = _
        Replace$( _
            resultText, _
            " ", _
            vbNullString)

    resultText = _
        Replace$( _
            resultText, _
            "-", _
            vbNullString)

    resultText = _
        Replace$( _
            resultText, _
            "_", _
            vbNullString)

    resultText = _
        Replace$( _
            resultText, _
            ".", _
            vbNullString)

    resultText = _
        Replace$( _
            resultText, _
            "/", _
            vbNullString)

    DBK_NormalizeKey = _
        resultText

End Function

Private Function DBK_SourceValue( _
    ByRef sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceMap As Object, _
    ByVal canonicalHeader As String) As Variant

    Dim sourceColumn As Long

    If sourceMap.Exists( _
        canonicalHeader) Then

        sourceColumn = _
            CLng( _
                sourceMap( _
                    canonicalHeader))

        DBK_SourceValue = _
            sourceData( _
                sourceRow, _
                sourceColumn)

    Else

        DBK_SourceValue = _
            vbNullString

    End If

End Function

Private Function DBK_SourceText( _
    ByRef sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceMap As Object, _
    ByVal canonicalHeader As String) As String

    DBK_SourceText = _
        DBK_Text( _
            DBK_SourceValue( _
                sourceData, _
                sourceRow, _
                sourceMap, _
                canonicalHeader))

End Function

Private Function DBK_TargetHeader( _
    ByVal columnIndex As Long) As String

    Select Case columnIndex

        Case 1
            DBK_TargetHeader = "KODE"

        Case 2
            DBK_TargetHeader = "POSITION"

        Case 3
            DBK_TargetHeader = "LOCATION"

        Case 4
            DBK_TargetHeader = "BUSINESS UNIT"

        Case 5
            DBK_TargetHeader = "DIVISION CHRIS"

        Case 6
            DBK_TargetHeader = "DEPARTMENT CHRIS"

        Case 7
            DBK_TargetHeader = "USER (MANAGER)"

        Case 8
            DBK_TargetHeader = "INDIRECT USER"

        Case 9
            DBK_TargetHeader = "DIRECTORATE"

        Case 10
            DBK_TargetHeader = "YEAR"

    End Select

End Function

' ============================================================
' TARGET LAYOUT
' ============================================================
Private Sub DBK_PrepareStaticTarget( _
    ByVal ws As Worksheet)

    Dim headers(1 To 1, 1 To 10) As Variant
    Dim tableIndex As Long
    Dim lastRow As Long
    Dim lastColumn As Long

    If ws Is Nothing Then
        Exit Sub
    End If

    headers(1, 1) = "KODE"
    headers(1, 2) = "POSITION"
    headers(1, 3) = "LOCATION"
    headers(1, 4) = "BUSINESS UNIT"
    headers(1, 5) = "DIVISION CHRIS"
    headers(1, 6) = "DEPARTMENT CHRIS"
    headers(1, 7) = "USER (MANAGER)"
    headers(1, 8) = "INDIRECT USER"
    headers(1, 9) = "DIRECTORATE"
    headers(1, 10) = "YEAR"

    For tableIndex = _
        ws.ListObjects.Count _
        To 1 Step -1

        On Error Resume Next

        ws.ListObjects( _
            tableIndex).Unlist

        On Error GoTo 0

    Next tableIndex

    lastRow = _
        DBK_LastUsedRow( _
            ws)

    lastColumn = _
        DBK_LastUsedColumn( _
            ws)

    If lastColumn > _
       DBK_STATIC_COLS Then

        ws.Range( _
            ws.Cells( _
                1, _
                DBK_STATIC_COLS + 1), _
            ws.Cells( _
                IIf(lastRow > 1, lastRow, 1), _
                lastColumn)).Clear

    End If

    ws.Range("A1:J1").Value2 = _
        headers

    With ws.Range("A1:J1")

        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True

    End With

    ws.Columns("A:J").AutoFit

End Sub

' ============================================================
' GENERIC HELPERS
' ============================================================
Private Function DBK_GetOrCreateWorksheet( _
    ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            sheetName)

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = _
            ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        ws.Name = _
            sheetName

    End If

    Set DBK_GetOrCreateWorksheet = _
        ws

End Function

Private Function DBK_NewDictionary() As Object

    Dim result As Object

    Set result = _
        CreateObject( _
            "Scripting.Dictionary")

    result.CompareMode = _
        vbTextCompare

    Set DBK_NewDictionary = _
        result

End Function

Private Function DBK_SheetExists( _
    ByVal wb As Object, _
    ByVal sheetName As String) As Boolean

    Dim ws As Object

    On Error Resume Next

    Set ws = _
        wb.Worksheets( _
            sheetName)

    DBK_SheetExists = _
        Not ws Is Nothing

    Set ws = Nothing

    On Error GoTo 0

End Function

Private Function DBK_LastHeaderColumn( _
    ByVal ws As Object, _
    ByVal rowIndex As Long) As Long

    Dim lastColumn As Long

    lastColumn = _
        ws.Cells( _
            rowIndex, _
            ws.Columns.Count). _
        End(xlToLeft).Column

    If lastColumn < 1 Then
        lastColumn = 1
    End If

    DBK_LastHeaderColumn = _
        lastColumn

End Function

Private Function DBK_LastSourceRow( _
    ByVal ws As Object, _
    ByVal sourceMap As Object) As Long

    Dim canonicalHeader As Variant
    Dim sourceColumn As Long
    Dim candidateRow As Long
    Dim lastRow As Long

    For Each canonicalHeader In _
        sourceMap.keys

        sourceColumn = _
            CLng( _
                sourceMap( _
                    canonicalHeader))

        candidateRow = _
            ws.Cells( _
                ws.rows.Count, _
                sourceColumn). _
            End(xlUp).Row

        If candidateRow > lastRow Then

            lastRow = _
                candidateRow

        End If

    Next canonicalHeader

    DBK_LastSourceRow = _
        lastRow

End Function

Private Function DBK_LastDataRow( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    On Error Resume Next

    Set foundCell = _
        ws.Range("A:J").Find( _
            What:="*", _
            After:=ws.Cells(1, 1), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False)

    On Error GoTo 0

    If foundCell Is Nothing Then

        DBK_LastDataRow = _
            DBK_HEADER_ROW

    Else

        DBK_LastDataRow = _
            foundCell.Row

    End If

End Function

Private Function DBK_LastUsedRow( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    On Error Resume Next

    Set foundCell = _
        ws.Cells.Find( _
            What:="*", _
            After:=ws.Cells(1, 1), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False)

    On Error GoTo 0

    If foundCell Is Nothing Then

        DBK_LastUsedRow = 1

    Else

        DBK_LastUsedRow = _
            foundCell.Row

    End If

End Function

Private Function DBK_LastUsedColumn( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    On Error Resume Next

    Set foundCell = _
        ws.Cells.Find( _
            What:="*", _
            After:=ws.Cells(1, 1), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False)

    On Error GoTo 0

    If foundCell Is Nothing Then

        DBK_LastUsedColumn = _
            DBK_STATIC_COLS

    Else

        DBK_LastUsedColumn = _
            foundCell.Column

    End If

End Function

Private Function DBK_Text( _
    ByVal rawValue As Variant) As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    DBK_Text = _
        Trim$( _
            CStr(rawValue))

End Function

Private Function DBK_HasValue( _
    ByVal rawValue As Variant) As Boolean

    DBK_HasValue = _
        (Len(DBK_Text(rawValue)) > 0)

End Function

Private Function DBK_FileNameOnly( _
    ByVal fullPath As String) As String

    Dim separatorPosition As Long

    separatorPosition = _
        InStrRev( _
            fullPath, _
            "\")

    If separatorPosition > 0 Then

        DBK_FileNameOnly = _
            Mid$( _
                fullPath, _
                separatorPosition + 1)

    Else

        DBK_FileNameOnly = _
            fullPath

    End If

End Function

Private Function DBK_ElapsedText( _
    ByVal startedAt As Double) As String

    Dim elapsedSeconds As Double

    elapsedSeconds = _
        Timer - startedAt

    If elapsedSeconds < 0 Then

        elapsedSeconds = _
            elapsedSeconds + 86400

    End If

    If elapsedSeconds < 60 Then

        DBK_ElapsedText = _
            Format$( _
                elapsedSeconds, _
                "0.0") & _
            " detik"

    Else

        DBK_ElapsedText = _
            Format$( _
                Int(elapsedSeconds / 60), _
                "0") & _
            " menit " & _
            Format$( _
                elapsedSeconds Mod 60, _
                "0") & _
            " detik"

    End If

End Function

Private Function DBK_BuildResultMessage( _
    ByVal folderCount As Long, _
    ByVal filesFound As Long, _
    ByVal filesOpened As Long, _
    ByVal dbKodeSheetsFound As Long, _
    ByVal fptkSheetsFound As Long, _
    ByVal existingCount As Long, _
    ByVal importedCount As Long, _
    ByVal updatedCount As Long, _
    ByVal kodeSyncedCount As Long, _
    ByVal kodeNotFoundCount As Long, _
    ByVal sourceRowsRead As Long, _
    ByVal blankKeyRows As Long, _
    ByVal skippedFileCount As Long, _
    ByVal totalTarget As Long, _
    ByVal startedAt As Double) As String

    Dim resultText As String

    resultText = _
        "Compile DB Kode Posisi selesai." & _
        vbCrLf & vbCrLf

    resultText = resultText & _
        "Folder: " & _
        CStr(folderCount) & _
        vbCrLf

    resultText = resultText & _
        "File ditemukan/dibuka: " & _
        CStr(filesFound) & _
        "/" & _
        CStr(filesOpened) & _
        vbCrLf

    resultText = resultText & _
        "Sheet DB Kode Posisi: " & _
        CStr(dbKodeSheetsFound) & _
        vbCrLf

    resultText = resultText & _
        "Sheet FPTK: " & _
        CStr(fptkSheetsFound) & _
        vbCrLf

    resultText = resultText & _
        "Data existing: " & _
        CStr(existingCount) & _
        vbCrLf

    resultText = resultText & _
        "Data baru: " & _
        CStr(importedCount) & _
        vbCrLf

    resultText = resultText & _
        "Data diperbarui: " & _
        CStr(updatedCount) & _
        vbCrLf

    resultText = resultText & _
        "KODE synced/not found: " & _
        CStr(kodeSyncedCount) & _
        "/" & _
        CStr(kodeNotFoundCount) & _
        vbCrLf

    resultText = resultText & _
        "Source rows/blank key: " & _
        CStr(sourceRowsRead) & _
        "/" & _
        CStr(blankKeyRows) & _
        vbCrLf

    resultText = resultText & _
        "File gagal: " & _
        CStr(skippedFileCount) & _
        vbCrLf

    resultText = resultText & _
        "Total target: " & _
        CStr(totalTarget) & _
        vbCrLf

    resultText = resultText & _
        "Durasi: " & _
        DBK_ElapsedText(startedAt)

    DBK_BuildResultMessage = _
        resultText

End Function






