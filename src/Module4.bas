Attribute VB_Name = "Module4"
Option Explicit

Private Const STATIC_DBKODE_COLS As Long = 10
' Layout final sheet FPTK:
' row 1 = header
' row 2 = data pertama

Private Const FPTK_HEADER_ROW As Long = 1
Private Const FPTK_FIRST_DATA_ROW As Long = 2

' ============================================================

' ============================================================
Private gFPTKDateSettingsLoaded As Boolean
Private gFPTKDefaultDateOrder As String
Private gFPTKDateFormatOverrides As String

Private gFPTKAmbiguousDateCount As Long
Private gFPTKResolvedBySourceMonthCount As Long
Private gFPTKResolvedBySourceWeekCount As Long
Private gFPTKResolvedByMonthWeekCount As Long
Private gFPTKSourceDateConflictCount As Long
Private gFPTKResolvedByFallbackCount As Long

' ============================================================
' SIAPKAN TARGET STATIC
' ============================================================
Private Sub PrepareStaticDBKodeTarget( _
    ByVal ws As Worksheet)

    Dim headerRow As Long
    Dim tableIndex As Long
    Dim headers(1 To 1, 1 To 10) As Variant

    If ws Is Nothing Then Exit Sub

    For tableIndex = ws.ListObjects.Count _
        To 1 Step -1

        On Error Resume Next
        ws.ListObjects(tableIndex).Unlist
        On Error GoTo 0

    Next tableIndex

    headerRow = FindStaticDBKodeHeaderRow(ws)

    If headerRow > 1 Then

        ws.rows( _
            "1:" & CStr(headerRow - 1)).Delete _
            Shift:=xlUp

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

    ws.Range("A1:J1").Value = headers

    CleanStaticDBKodeGenericHeaders ws

    With ws.Range("A1:J1")

        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True

    End With

End Sub

' ============================================================
' HAPUS HEADER GENERIK COLUMN2, COLUMN3, DST
'
' Hanya header kosong di luar A:J yang dibersihkan.
' Data yang sudah terisi tidak dihapus.
' ============================================================
Private Sub CleanStaticDBKodeGenericHeaders( _
    ByVal ws As Worksheet)

    Dim lastColumn As Long
    Dim lastRow As Long
    Dim columnIndex As Long
    Dim headerText As String
    Dim dataCount As Long

    lastColumn = ws.Cells( _
        1, ws.Columns.Count).End(xlToLeft).Column

    If lastColumn <= STATIC_DBKODE_COLS Then
        Exit Sub
    End If

    lastRow = ws.Cells( _
        ws.rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        lastRow = 2
    End If

    For columnIndex = lastColumn _
        To STATIC_DBKODE_COLS + 1 Step -1

        headerText = UCase$( _
            Trim$(CStr( _
                ws.Cells( _
                    1, columnIndex).Value2)))

        If IsStaticGenericColumnHeader( _
            headerText) Then

            dataCount = Application.CountA( _
                ws.Range( _
                    ws.Cells(2, columnIndex), _
                    ws.Cells(lastRow, columnIndex)))

            If dataCount = 0 Then

                ws.Columns(columnIndex).Delete

            End If

        End If

    Next columnIndex

End Sub

Private Function IsStaticGenericColumnHeader( _
    ByVal headerText As String) As Boolean

    Dim compactText As String
    Dim numberText As String
    Dim characterIndex As Long
    Dim currentCharacter As String

    compactText = Replace$( _
        UCase$(Trim$(headerText)), _
        " ", vbNullString)

    If Left$(compactText, 6) <> "COLUMN" Then
        Exit Function
    End If

    numberText = Mid$(compactText, 7)

    If Len(numberText) = 0 Then Exit Function

    For characterIndex = 1 To Len(numberText)

        currentCharacter = Mid$( _
            numberText, characterIndex, 1)

        If currentCharacter < "0" Or _
           currentCharacter > "9" Then

            Exit Function

        End If

    Next characterIndex

    IsStaticGenericColumnHeader = True

End Function

' ============================================================
' CARI HEADER SOURCE STATIC
' ============================================================
Private Function FindStaticDBKodeHeaderRow( _
    ByVal ws As Object) As Long

    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim headerText As String
    Dim hasKode As Boolean
    Dim hasPosition As Boolean

    For rowIndex = 1 To 10

        hasKode = False
        hasPosition = False

        For columnIndex = 1 _
            To STATIC_DBKODE_COLS

            headerText = UCase$( _
                Trim$(CStr( _
                    ws.Cells( _
                        rowIndex, _
                        columnIndex).Value2)))

            If headerText = "KODE" Then
                hasKode = True
            End If

            If headerText = "POSITION" Or _
               headerText = "POSISI" Then

                hasPosition = True

            End If

        Next columnIndex

        If hasKode And hasPosition Then

            FindStaticDBKodeHeaderRow = rowIndex
            Exit Function

        End If

    Next rowIndex

End Function

' ============================================================
' LOAD KEY EXISTING DARI KOLOM A
' ============================================================
Private Sub LoadExistingStaticDBKodeKeys( _
    ByVal ws As Worksheet, _
    ByRef existingKeys As Object)

    Dim lastRow As Long
    Dim existingData As Variant
    Dim rowIndex As Long
    Dim keyValue As String

    lastRow = ws.Cells( _
        ws.rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    existingData = ws.Range( _
        "A2:A" & CStr(lastRow)).Value2

    If IsArray(existingData) Then

        For rowIndex = 1 _
            To UBound(existingData, 1)

            keyValue = NormalizeDataKey( _
                existingData(rowIndex, 1))

            If Len(keyValue) > 0 Then

                If Not existingKeys.Exists( _
                    keyValue) Then

                    existingKeys.Add _
                        keyValue, True

                End If

            End If

        Next rowIndex

    End If

End Sub

' ============================================================
' NEXT ROW BERDASARKAN KOLOM A
' ============================================================
Private Function GetNextStaticDBKodeRow( _
    ByVal ws As Worksheet) As Long

    Dim lastRow As Long

    lastRow = ws.Cells( _
        ws.rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then

        GetNextStaticDBKodeRow = 2

    Else

        GetNextStaticDBKodeRow = _
            lastRow + 1

    End If

End Function

' ============================================================
' POTONG ARRAY KE JUMLAH ROW AKTUAL
' ============================================================
Private Function CopyFirstRowsStaticDBKode( _
    ByVal sourceData As Variant, _
    ByVal rowCount As Long, _
    ByVal columnCount As Long) As Variant

    Dim resultData() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    ReDim resultData( _
        1 To rowCount, _
        1 To columnCount)

    For rowIndex = 1 To rowCount

        For columnIndex = 1 To columnCount

            resultData( _
                rowIndex, columnIndex) = _
                sourceData( _
                    rowIndex, columnIndex)

        Next columnIndex

    Next rowIndex

    CopyFirstRowsStaticDBKode = resultData

End Function

' ============================================================
' FPTK DI BAWAH INI TETAP DYNAMIC DAN FLEXIBLE
' ============================================================
Public Sub Compile_FPTK(ByVal folderPath As String)

    Dim wsTarget As Worksheet
    Dim targetMap As Object
    Dim existingIndex As Object
    Dim sourceFiles As Collection

    Dim targetData As Variant
    Dim insertedRows() As Variant

    Dim targetRowCount As Long
    Dim targetColumnCount As Long
    Dim originalTargetRowCount As Long

    Dim insertedCount As Long
    Dim insertedCapacity As Long

    Dim fileIndex As Long
    Dim fileName As String
    Dim fullPath As String
    Dim detectedSTOSourcePath As String

    Dim wbSource As Object
    Dim wsSource As Object
    Dim openError As String

    Dim headerRow As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim sourceData As Variant
    Dim sourceKeys As Variant

    Dim reasonSourceColumn As Long
    Dim reasonTargetColumn As Long

    Dim sourceRow As Long
    Dim sourceRowCount As Long
    Dim recordPointer As Long
    Dim rowKey As String
    Dim rowData As Variant
    Dim rowChanged As Boolean

    Dim totalImported As Long
    Dim totalUpdated As Long
    Dim totalUnchanged As Long
    Dim totalSkipped As Long

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldStatusBar As Variant

    Dim errorNumber As Long
    Dim errorText As String
    Dim startedAt As Double
    Dim summaryText As String

    On Error GoTo ErrorHandler

    startedAt = Timer

    FPTK_ResetDateResolverCache

    folderPath = EnsureTrailingSeparator( _
        folderPath)

    Set sourceFiles = CollectFPTKSourceFiles( _
        folderPath)

    If sourceFiles.Count = 0 Then

        AddRunLog _
            "Warning", _
            "FPTK | Tidak ada file .xlsx/.xlsm di folder"

        Exit Sub

    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    oldStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wsTarget = GetOrCreateWorksheet( _
        "FPTK")

    HeartbeatProgress _
        "RUNNING", _
        "FPTK", _
        "FPTK - Prepare Existing", _
        "Checking header row 1"

    Application.StatusBar = _
        "FPTK | Preparing existing data..."

    DoEvents

    ' Layout Row 1:
    ' row 1 = header
    ' row 2 = data pertama
    EnsureFPTKHeaderRow1Layout _
        wsTarget

    ' Header hanya dijaga sekali di awal.
    ' Schema tidak dijalankan saat compile agar lebih cepat
    ' dan tidak mengubah visibility/urutan header.
    ProtectFPTKHeader

    ' Pemindahan kolom dilakukan setelah output selesai ditulis.
    ' Hanya range header-data yang dipindahkan, sehingga tombol
    ' header row 1 tetap dipertahankan saat kolom dipindahkan.
    FPTK_EnsureSpecialColumnOrder _
        wsTarget

    FPTK_ApplyAvailabilityValidation _
        wsTarget

    StyleHeaderDynamic _
        wsTarget

    Set targetMap = BuildHeaderMap( _
        wsTarget, _
        FPTK_HEADER_ROW)

    targetColumnCount = LastHeaderColumn( _
        wsTarget, _
        FPTK_HEADER_ROW)

    targetRowCount = FPTK_FastTargetRowCount( _
        wsTarget, _
        targetMap)

    originalTargetRowCount = targetRowCount

    HeartbeatProgress _
        "RUNNING", _
        "FPTK", _
        "FPTK - Read Existing", _
        Format$(targetRowCount, "#,##0") & " rows"

    Application.StatusBar = _
        "FPTK | Reading " & _
        Format$(targetRowCount, "#,##0") & _
        " existing rows..."

    DoEvents

    targetData = FPTK_LoadTargetDataArray( _
        wsTarget, _
        targetRowCount, _
        targetColumnCount)

    HeartbeatProgress _
        "RUNNING", _
        "FPTK", _
        "FPTK - Index Existing", _
        Format$(targetRowCount, "#,##0") & " rows"

    Application.StatusBar = _
        "FPTK | Building existing key index..."

    DoEvents

    Set existingIndex = _
        FPTK_BuildExistingIndexFromArray( _
            targetData, _
            targetRowCount, _
            targetMap)

    insertedCapacity = 500
    ReDim insertedRows( _
        1 To insertedCapacity) As Variant

    For fileIndex = 1 To sourceFiles.Count

        fullPath = CStr( _
            sourceFiles(fileIndex))

        fileName = FPTKFileNameFromPath( _
            fullPath)

        ' ====================================================
        ' FAST STO DETECTION BY FILE NAME.
        '
        ' Tidak lagi bergantung ke STO_ResolveSourcePath.
        ' Kalau user compile STO.xlsm dari folder/copy lain,
        ' file tetap diproses sebagai STO khusus.
        ' ====================================================
        If UCase$(Trim$(fileName)) = "STO.XLSM" Or _
           UCase$(Trim$(fileName)) = "STO.XLSX" Then

            detectedSTOSourcePath = _
                fullPath

            AddRunLog _
                "Info", _
                "FPTK | STO detected by filename | " & _
                fullPath

            GoTo NextFile

        End If

        Application.StatusBar = _
            "FPTK | File " & _
            CStr(fileIndex) & "/" & _
            CStr(sourceFiles.Count) & _
            " | Membuka " & fileName

        HeartbeatProgress _
            "RUNNING", _
            fileName, _
            "FPTK - Open Source File", _
            CStr(fileIndex) & "/" & _
            CStr(sourceFiles.Count)

        openError = vbNullString
        Set wbSource = Nothing
        Set wsSource = Nothing

        Set wbSource = OpenSourceWorkbook( _
            fullPath, _
            openError)

        If wbSource Is Nothing Then

            AddRunLog _
                "Error", _
                "FPTK | " & fileName & _
                " | " & openError

            GoTo NextFile

        End If

        Set wsSource = GetSheetByPossibleNames( _
            wbSource, _
            "FPTK")

        If wsSource Is Nothing Then

            AddRunLog _
                "Warning", _
                "FPTK | " & fileName & _
                " | Sheet FPTK tidak ditemukan"

            GoTo NextFile

        End If

        headerRow = FindHeaderRowDynamic( _
            wsSource)

        If headerRow = 0 Then

            AddRunLog _
                "Warning", _
                "FPTK | " & fileName & _
                " | Header tidak ditemukan"

            GoTo NextFile

        End If

        ' ====================================================
        ' STRUCTURAL STO DETECTION.
        '
        ' Source dianggap STO control bila sheet FPTK memiliki
        ' FPTK Availability dengan value V/X.
        '
        ' Ini adalah guard utama. Jadi walaupun file:
        ' - namanya bukan STO.xlsm,
        ' - dipindah/copy ke folder lain,
        ' - setting path lama berbeda,
        ' V/X tetap TIDAK pernah masuk generic compiler.
        ' ====================================================
        ' Jangan mendeteksi STO hanya dari Status FPTK /
        ' FPTK Availability karena file reguler juga bisa
        ' memakai V/X, Ya/Tidak, atau Y/N.
        If FPTK_IsDedicatedSTOSourceFile( _
            fullPath) Then

            detectedSTOSourcePath = fullPath

            AddRunLog _
                "Info", _
                "FPTK | Dedicated STO source detected | " & _
                fullPath

            GoTo NextFile

        End If

        ' ====================================================
        ' FILTER / HIDDEN ROWS MUST NOT AFFECT COMPILE
        '
        ' Source workbook dibuka read-only dan ditutup tanpa save.
        ' Jadi filter/hide hanya dibuka sementara di memory agar
        ' SEMUA row source ikut dibaca ke sourceData.
        ' ====================================================
        FPTK_ExposeAllSourceRows _
            wsSource, _
            headerRow

        lastColumn = LastHeaderColumn( _
            wsSource, _
            headerRow)

        If lastColumn <= 0 Then
            GoTo NextFile
        End If

        ' Header baru hanya ditambahkan ke kanan.
        EnsureTargetHeadersFromSource _
            wsTarget, _
            targetMap, _
            wsSource, _
            headerRow, _
            lastColumn

        If LastHeaderColumn( _
            wsTarget, _
            FPTK_HEADER_ROW) > targetColumnCount Then

            targetColumnCount = LastHeaderColumn( _
                wsTarget, _
                FPTK_HEADER_ROW)

            FPTK_ExpandTargetDataColumns _
                targetData, _
                targetRowCount, _
                targetColumnCount

            FPTK_ExpandInsertedRowsColumns _
                insertedRows, _
                insertedCount, _
                targetColumnCount

        End If

        sourceKeys = BuildColumnKeyArray( _
            wsSource, _
            headerRow, _
            lastColumn)

        ' ====================================================
        ' DIRECT SOURCE REASON MAPPING
        '
        ' Jangan bergantung ke schema/alias untuk Category.
        ' Cari kolom raw source langsung:
        ' Category / Alasan / Alasan FPTK /
        ' Alasan Permintaan FPTK.
        ' ====================================================
        reasonSourceColumn = _
            FPTK_FindReasonSourceColumn( _
                wsSource, _
                headerRow, _
                lastColumn)

        reasonTargetColumn = _
            FPTK_DirectFindHeader( _
                wsTarget, _
                FPTK_HEADER_ROW, _
                "Alasan Permintaan FPTK")

        lastRow = LastSourceDataRow( _
            wsSource, _
            headerRow, _
            sourceKeys)

        If lastRow < headerRow + 1 Then
            GoTo NextFile
        End If

        sourceData = wsSource.Range( _
            wsSource.Cells(headerRow + 1, 1), _
            wsSource.Cells(lastRow, lastColumn)).Value2

        If Not IsArray(sourceData) Then
            GoTo NextFile
        End If

        sourceRowCount = UBound( _
            sourceData, _
            1)

        For sourceRow = 1 To sourceRowCount

            If sourceRow = 1 Or _
               sourceRow Mod 1000 = 0 Or _
               sourceRow = sourceRowCount Then

                HeartbeatProgress _
                    "RUNNING", _
                    fileName, _
                    "FPTK - Process Rows", _
                    "File " & _
                    CStr(fileIndex) & "/" & _
                    CStr(sourceFiles.Count) & _
                    " | Row " & _
                    CStr(sourceRow) & "/" & _
                    CStr(sourceRowCount) & _
                    " | Insert " & _
                    CStr(totalImported) & _
                    " | Update " & _
                    CStr(totalUpdated)

                DoEvents

            End If

            If FPTK_IsValidSourceRecord( _
                sourceData, _
                sourceRow, _
                sourceKeys) Then

                rowKey = BuildFPTKKeyFromArray( _
                    sourceData, _
                    sourceRow, _
                    sourceKeys)

                If Len(rowKey) = 0 Then

                    totalSkipped = totalSkipped + 1

                ElseIf existingIndex.Exists(rowKey) Then

                    recordPointer = CLng( _
                        existingIndex(rowKey))

                    If recordPointer > 0 Then

                        rowData = FPTK_ExtractTargetRow( _
                            targetData, _
                            recordPointer, _
                            targetColumnCount)

                    Else

                        rowData = insertedRows( _
                            Abs(recordPointer))

                    End If

                    rowChanged = _
                        FPTK_MapSourceIntoRowArray( _
                            rowData, _
                            targetMap, _
                            sourceData, _
                            sourceRow, _
                            sourceKeys, _
                            fullPath, _
                            headerRow + sourceRow)

                    If FPTK_MapReasonDirect( _
                        rowData, _
                        reasonTargetColumn, _
                        sourceData, _
                        sourceRow, _
                        reasonSourceColumn) Then

                        rowChanged = True

                    End If

                    If rowChanged Then

                        FPTK_StampAuditRowArray _
                            rowData, _
                            targetMap, _
                            "UPDATE", _
                            fileName, _
                            False

                        FPTK_RefreshDerivedRowArray _
                            rowData, _
                            targetMap

                        If recordPointer > 0 Then

                            FPTK_WriteTargetRowBack _
                                targetData, _
                                recordPointer, _
                                rowData, _
                                targetColumnCount

                        Else

                            insertedRows( _
                                Abs(recordPointer)) = _
                                    rowData

                        End If

                        totalUpdated = totalUpdated + 1

                    Else

                        totalUnchanged = _
                            totalUnchanged + 1

                    End If

                Else

                    rowData = FPTK_CreateBlankRow( _
                        targetColumnCount)

                    Call FPTK_MapSourceIntoRowArray( _
                        rowData, _
                        targetMap, _
                        sourceData, _
                        sourceRow, _
                        sourceKeys, _
                        fullPath, _
                        headerRow + sourceRow)

                    Call FPTK_MapReasonDirect( _
                        rowData, _
                        reasonTargetColumn, _
                        sourceData, _
                        sourceRow, _
                        reasonSourceColumn)

                    FPTK_StampAuditRowArray _
                        rowData, _
                        targetMap, _
                        "INSERT", _
                        fileName, _
                        True

                    FPTK_RefreshDerivedRowArray _
                        rowData, _
                        targetMap

                    insertedCount = insertedCount + 1

                    If insertedCount > insertedCapacity Then

                        insertedCapacity = _
                            insertedCapacity * 2

                        ReDim Preserve insertedRows( _
                            1 To insertedCapacity) As Variant

                    End If

                    insertedRows(insertedCount) = rowData

                    ' Pointer negatif berarti row baru di buffer.
                    existingIndex.Add _
                        rowKey, _
                        -insertedCount

                    totalImported = totalImported + 1

                End If

            End If

        Next sourceRow

NextFile:

        If Not wbSource Is Nothing Then
            CloseSourceWorkbook wbSource
        End If

        Set wsSource = Nothing
        Set wbSource = Nothing

    Next fileIndex

    HeartbeatProgress _
        "RUNNING", _
        "FPTK", _
        "FPTK - Write Output", _
        "Writing existing and inserted rows"

    DoEvents

    FPTK_WriteTargetDataBlock _
        wsTarget, _
        targetData, _
        targetRowCount, _
        targetColumnCount

    FPTK_WriteInsertedRowsBlock _
        wsTarget, _
        insertedRows, _
        insertedCount, _
        targetRowCount, _
        targetColumnCount

    FPTK_ApplyAvailabilityValidation _
        wsTarget

    ProtectFPTKHeader

    Set targetMap = BuildHeaderMap( _
        wsTarget, _
        FPTK_HEADER_ROW)

    ' ========================================================
    ' STO CURRENT-RUN PROCESSING
    '
    ' MATCH KEY STO = POSISI + FPTK DATE (REAL).
    '
    ' Rule:
    ' - Match existing -> HANYA Filter Kategorisasi FPTK = STO.
    ' - Tidak match -> INSERT row baru dari source STO.
    ' - Tidak ada delete legacy otomatis.
    ' ========================================================
    If Len(detectedSTOSourcePath) > 0 Then

        STO_SaveSettingValue _
            "STOCompiledFilePath", _
            detectedSTOSourcePath

        AddRunLog _
            "Info", _
            "STO Current Run | Match: Posisi + FPTK Date (Real) | Source: " & _
            detectedSTOSourcePath

        Set targetMap = BuildHeaderMap( _
            wsTarget, _
            FPTK_HEADER_ROW)

        STO_SyncAllAssignmentsByPositionRealDate _
            wsTarget, _
            True, _
            detectedSTOSourcePath

        Set targetMap = BuildHeaderMap( _
            wsTarget, _
            FPTK_HEADER_ROW)

    End If

    ' ========================================================
    ' REFRESH FINAL WEEK / MONTH VALUES
    '
    ' Source Week/Month hanya dipakai untuk menentukan apakah
    ' tanggal ambigu memakai DMY atau MDY.
    '
    ' Setelah tanggal final tersimpan:
    ' - Week FPTK Date (CODE) dihitung dari FPTK Date (Kode)
    ' - Week Cancel Date dihitung dari FPTK Cancel Date
    ' - Week Offering Date dihitung dari Offering Date
    ' - seluruh Month terkait juga dihitung ulang
    '
    ' Refresh dilakukan untuk seluruh row agar row existing
    ' yang dianggap unchanged tetap terisi.
    ' ========================================================
    ' Kode BU selalu diturunkan dari Kode PIC.
    ' Nilai Kode BU dari source tidak dipakai sebagai hasil final.
    RefreshFPTKKodeBUValues _
        wsTarget, _
        targetMap

    ' User dan Indirect User selalu ditarik dari
    ' DB Kode Posisi: Kode Angka lebih dulu,
    ' lalu fallback berdasarkan nama Posisi.
    RefreshFPTKUsersFromDBKodePosisi _
        wsTarget, _
        targetMap

    ' Source boleh memakai header Category sebagai Alasan.
    ' Value hanya dipertahankan bila ada di Kode Dropdown kolom C.
    RefreshFPTKReasonValues _
        wsTarget, _
        targetMap

    ' Status hanya boleh OP, Closed, atau Cancel.
    ' STATUS FPTK adalah alias FPTK Availability, bukan Status.
    RefreshFPTKStatusValues _
        wsTarget, _
        targetMap

    RefreshFPTKWeekMonthValues _
        wsTarget, _
        targetMap

    ' Recalculate SLA untuk seluruh row valid, termasuk row
    ' UNCHANGED dan kolom yang sebelumnya kosong.
    RefreshFPTKSLADerivedValues _
        wsTarget, _
        targetMap

    ' ========================================================
    ' FILTER KATEGORISASI STO
    '
    ' Override kategori dasar bila pasangan Posisi + Nama
    ' Kandidat ditemukan di Tulang Punggung Sales Cimory Group.
    ' ========================================================
    If Len(detectedSTOSourcePath) > 0 Then

        RefreshFPTKSTOCategorization _
            wsTarget, _
            targetMap, _
            True, _
            detectedSTOSourcePath

    End If

    ApplyFPTKDateFormats _
        wsTarget, _
        targetMap

    ApplyFPTKAuditFormats _
        wsTarget, _
        targetMap

    summaryText = _
        "FPTK FAST ARRAY | Files: " & _
        CStr(sourceFiles.Count) & _
        " | Existing rows: " & _
        CStr(originalTargetRowCount) & _
        " | Imported: " & _
        CStr(totalImported) & _
        " | Updated: " & _
        CStr(totalUpdated)

    summaryText = summaryText & _
        " | Unchanged: " & _
        CStr(totalUnchanged) & _
        " | Skipped: " & _
        CStr(totalSkipped) & _
        " | Date by Month: " & _
        CStr(gFPTKResolvedBySourceMonthCount) & _
        " | Date by Week: " & _
        CStr(gFPTKResolvedBySourceWeekCount)

    summaryText = summaryText & _
        " | Date by Month+Week: " & _
        CStr(gFPTKResolvedByMonthWeekCount) & _
        " | Date fallback: " & _
        CStr(gFPTKResolvedByFallbackCount) & _
        " | Date conflict: " & _
        CStr(gFPTKSourceDateConflictCount)

    summaryText = summaryText & _
        " | Ambiguous: " & _
        CStr(gFPTKAmbiguousDateCount) & _
        " | Duration: " & _
        FPTK_FastElapsedText(startedAt)

    AddRunLog _
        "Summary", _
        summaryText

SafeExit:

    On Error Resume Next

    If Not wbSource Is Nothing Then
        CloseSourceWorkbook wbSource
    End If

    Application.StatusBar = oldStatusBar
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    Set sourceFiles = Nothing
    Set existingIndex = Nothing
    Set targetMap = Nothing
    Set wsTarget = Nothing

    On Error GoTo 0

    If errorNumber <> 0 Then

        AddRunLog _
            "Error", _
            "Compile_FPTK FAST ARRAY | Error " & _
            CStr(errorNumber) & _
            ": " & errorText

    End If

    Exit Sub

ErrorHandler:

    errorNumber = Err.Number
    errorText = Err.Description

    On Error Resume Next

    HeartbeatProgress _
        "ERROR", _
        "FPTK", _
        "Compile FPTK", _
        "Error " & CStr(errorNumber) & _
        ": " & errorText

    MsgBox _
        "Compile FPTK gagal." & _
        vbCrLf & vbCrLf & _
        "Error " & CStr(errorNumber) & _
        ": " & errorText, _
        vbCritical, _
        "Compile FPTK"

    On Error GoTo 0

    Resume SafeExit

End Sub

' ============================================================
' FAST ARRAY / BATCH HELPERS
' ============================================================
Private Function FPTK_FastTargetRowCount( _
    ByVal ws As Worksheet, _
    ByVal targetMap As Object) As Long

    Dim colKodeUnik As Long
    Dim colPosisi As Long

    Dim lastKodeRow As Long
    Dim lastPosisiRow As Long
    Dim lastDataRow As Long

    If ws Is Nothing Then Exit Function
    If targetMap Is Nothing Then Exit Function

    colKodeUnik = ColumnByKey( _
        targetMap, _
        "KODE UNIK")

    colPosisi = ColumnByKey( _
        targetMap, _
        "POSISI")

    If colKodeUnik > 0 Then

        lastKodeRow = ws.Cells( _
            ws.rows.Count, _
            colKodeUnik).End(xlUp).Row

    End If

    If colPosisi > 0 Then

        lastPosisiRow = ws.Cells( _
            ws.rows.Count, _
            colPosisi).End(xlUp).Row

    End If

    lastDataRow = Application.Max( _
        lastKodeRow, _
        lastPosisiRow)

    If lastDataRow >= FPTK_FIRST_DATA_ROW Then

        FPTK_FastTargetRowCount = _
            lastDataRow - _
            FPTK_FIRST_DATA_ROW + 1

    End If

End Function


Private Function FPTK_LoadTargetDataArray( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal columnCount As Long) As Variant

    Dim emptyData() As Variant

    If rowCount <= 0 Or _
       columnCount <= 0 Then

        ReDim emptyData(1 To 1, 1 To 1)
        FPTK_LoadTargetDataArray = emptyData
        Exit Function

    End If

    FPTK_LoadTargetDataArray = _
        ws.Cells( _
            FPTK_FIRST_DATA_ROW, _
            1).Resize( _
                rowCount, _
                columnCount).Value2

End Function

Private Function FPTK_BuildExistingIndexFromArray( _
    ByVal targetData As Variant, _
    ByVal rowCount As Long, _
    ByVal targetMap As Object) As Object

    Dim result As Object
    Dim rowIndex As Long
    Dim rowKey As String

    Set result = CreateObject( _
        "Scripting.Dictionary")

    result.CompareMode = vbTextCompare

    If rowCount <= 0 Then

        Set FPTK_BuildExistingIndexFromArray = result
        Exit Function

    End If

    For rowIndex = 1 To rowCount

        If rowIndex = 1 Or _
           rowIndex Mod 5000 = 0 Or _
           rowIndex = rowCount Then

            HeartbeatProgress _
                "RUNNING", _
                "FPTK", _
                "FPTK - Index Existing", _
                Format$(rowIndex, "#,##0") & "/" & _
                Format$(rowCount, "#,##0")

            DoEvents

        End If

        rowKey = FPTK_KeyFromTargetArray( _
            targetData, _
            rowIndex, _
            targetMap)

        If Len(rowKey) > 0 Then

            If Not result.Exists(rowKey) Then
                result.Add rowKey, rowIndex
            End If

        End If

    Next rowIndex

    Set FPTK_BuildExistingIndexFromArray = result

End Function

Private Function FPTK_KeyFromTargetArray( _
    ByVal targetData As Variant, _
    ByVal rowIndex As Long, _
    ByVal targetMap As Object) As String

    Dim colKodeUnik As Long
    Dim colPosisi As Long

    Dim kodeUnik As String
    Dim posisi As String

    colKodeUnik = ColumnByKey( _
        targetMap, _
        "KODE UNIK")

    colPosisi = ColumnByKey( _
        targetMap, _
        "POSISI")

    If colKodeUnik <= 0 Or _
       colPosisi <= 0 Then

        Exit Function

    End If

    kodeUnik = CStr( _
        targetData( _
            rowIndex, _
            colKodeUnik))

    posisi = CStr( _
        targetData( _
            rowIndex, _
            colPosisi))

    If Len(NormalizeDataKey(kodeUnik)) = 0 Or _
       Len(NormalizeDataKey(posisi)) = 0 Then

        Exit Function

    End If

    FPTK_KeyFromTargetArray = _
        NormalizeDataKey(kodeUnik) & _
        "||" & _
        NormalizeDataKey(posisi)

End Function

Private Sub FPTK_ExpandTargetDataColumns( _
    ByRef targetData As Variant, _
    ByVal rowCount As Long, _
    ByVal newColumnCount As Long)

    Dim oldColumnCount As Long
    Dim expandedData() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    If newColumnCount <= 0 Then Exit Sub
    If rowCount <= 0 Then Exit Sub

    oldColumnCount = UBound( _
        targetData, _
        2)

    If newColumnCount <= oldColumnCount Then
        Exit Sub
    End If

    ReDim expandedData( _
        1 To rowCount, _
        1 To newColumnCount)

    For rowIndex = 1 To rowCount

        For columnIndex = 1 To oldColumnCount

            expandedData( _
                rowIndex, _
                columnIndex) = _
                    targetData( _
                        rowIndex, _
                        columnIndex)

        Next columnIndex

    Next rowIndex

    targetData = expandedData

End Sub

Private Sub FPTK_ExpandInsertedRowsColumns( _
    ByRef insertedRows() As Variant, _
    ByVal insertedCount As Long, _
    ByVal newColumnCount As Long)

    Dim rowIndex As Long
    Dim oldRow As Variant
    Dim newRow As Variant
    Dim columnIndex As Long
    Dim oldColumnCount As Long

    If insertedCount <= 0 Then Exit Sub

    For rowIndex = 1 To insertedCount

        oldRow = insertedRows(rowIndex)

        If IsArray(oldRow) Then

            oldColumnCount = UBound(oldRow)

            If newColumnCount > oldColumnCount Then

                ReDim newRow( _
                    1 To newColumnCount)

                For columnIndex = 1 To oldColumnCount

                    newRow(columnIndex) = _
                        oldRow(columnIndex)

                Next columnIndex

                insertedRows(rowIndex) = newRow

            End If

        End If

    Next rowIndex

End Sub

Private Function FPTK_CreateBlankRow( _
    ByVal columnCount As Long) As Variant

    Dim rowData() As Variant

    ReDim rowData( _
        1 To columnCount)

    FPTK_CreateBlankRow = rowData

End Function

Private Function FPTK_ExtractTargetRow( _
    ByVal targetData As Variant, _
    ByVal rowIndex As Long, _
    ByVal columnCount As Long) As Variant

    Dim rowData() As Variant
    Dim columnIndex As Long

    ReDim rowData( _
        1 To columnCount)

    For columnIndex = 1 To columnCount

        rowData(columnIndex) = _
            targetData( _
                rowIndex, _
                columnIndex)

    Next columnIndex

    FPTK_ExtractTargetRow = rowData

End Function

Private Sub FPTK_WriteTargetRowBack( _
    ByRef targetData As Variant, _
    ByVal rowIndex As Long, _
    ByVal rowData As Variant, _
    ByVal columnCount As Long)

    Dim columnIndex As Long

    For columnIndex = 1 To columnCount

        targetData( _
            rowIndex, _
            columnIndex) = _
                rowData(columnIndex)

    Next columnIndex

End Sub
Private Function FPTK_FindReasonSourceColumn( _
    ByVal wsSource As Object, _
    ByVal headerRow As Long, _
    ByVal lastColumn As Long) As Long

    Dim columnIndex As Long
    Dim headerKey As String
    Dim currentPriority As Long
    Dim bestPriority As Long

    If wsSource Is Nothing Then Exit Function

    For columnIndex = 1 To lastColumn

        headerKey = _
            HeaderKeyText( _
                wsSource.Cells( _
                    headerRow, _
                    columnIndex).Value2)

        currentPriority = 0

        Select Case headerKey

            Case "ALASAN PERMINTAAN FPTK"
                currentPriority = 400

            Case "ALASAN FPTK"
                currentPriority = 300

            Case "ALASAN"
                currentPriority = 200

            Case "CATEGORY"
                currentPriority = 100

        End Select

        If currentPriority > bestPriority Then

            bestPriority = _
                currentPriority

            FPTK_FindReasonSourceColumn = _
                columnIndex

        End If

    Next columnIndex

End Function

Private Function FPTK_MapReasonDirect( _
    ByRef rowData As Variant, _
    ByVal targetColumn As Long, _
    ByVal sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceColumn As Long) As Boolean

    Dim sourceValue As Variant
    Dim currentValue As Variant

    If targetColumn <= 0 Then Exit Function
    If sourceColumn <= 0 Then Exit Function

    sourceValue = _
        sourceData( _
            sourceRow, _
            sourceColumn)

    If ValueIsBlankFPTK( _
        sourceValue) Then

        Exit Function

    End If

    If Not FPTK_IsValidSourceReasonValue( _
        sourceValue) Then

        Exit Function

    End If

    currentValue = _
        rowData( _
            targetColumn)

    If Not ValuesEqualFPTK( _
        currentValue, _
        sourceValue) Then

        rowData( _
            targetColumn) = _
                sourceValue

        FPTK_MapReasonDirect = _
            True

    End If

End Function

Private Function FPTK_IsValidSourceReasonValue( _
    ByVal rawValue As Variant) As Boolean

    Dim valueKey As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    valueKey = _
        FPTK_NormalizeLookupText( _
            rawValue)

    If Len(valueKey) = 0 Then Exit Function

    Select Case valueKey

        Case "Y", _
             "N", _
             "YA", _
             "TIDAK", _
             "YES", _
             "NO", _
             "TRUE", _
             "FALSE", _
             "1", _
             "0"

            Exit Function

    End Select

    FPTK_IsValidSourceReasonValue = True

End Function

Private Function FPTK_MapSourceIntoRowArray( _
    ByRef rowData As Variant, _
    ByVal targetMap As Object, _
    ByVal sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceKeys As Variant, _
    ByVal sourceFilePath As String, _
    ByVal sourceExcelRow As Long) As Boolean

    Dim sourceColumn As Long
    Dim headerKey As String
    Dim targetColumn As Long

    Dim sourceValue As Variant
    Dim targetValue As Variant
    Dim rowChanged As Boolean

    For sourceColumn = LBound(sourceKeys) _
        To UBound(sourceKeys)

        headerKey = CStr( _
            sourceKeys(sourceColumn))

        If Len(headerKey) > 0 Then

            If Not IsFPTKCompilerManagedHeaderKey( _
                headerKey) Then

                If targetMap.Exists(headerKey) Then

                    targetColumn = CLng( _
                        targetMap(headerKey))

                    sourceValue = sourceData( _
                        sourceRow, _
                        sourceColumn)

                    ' Status FPTK / FPTK Availability langsung
                    ' dinormalisasi saat mapping source.
                    If StrComp( _
                        headerKey, _
                        "FPTK AVAILABILITY", _
                        vbTextCompare) = 0 Then

                        sourceValue = _
                            STO_NormalizeAvailability( _
                                sourceValue)

                        If ValueIsBlankFPTK( _
                            sourceValue) Then

                            GoTo NextSourceColumn

                        End If

                    End If

                    If FPTK_IsStrictSourceDateKey( _
                        headerKey) Then

                        sourceValue = _
                            FPTK_ResolveSourceDateValue( _
                                sourceValue, _
                                headerKey, _
                                sourceData, _
                                sourceRow, _
                                sourceKeys, _
                                sourceFilePath, _
                                sourceExcelRow)

                    End If

                    ' ==================================================
                    ' CATEGORY di source = Alasan Permintaan FPTK.
                    '
                    ' Source recruiter boleh memakai nama:
                    ' - Category
                    ' - Alasan
                    ' - Alasan FPTK
                    ' - Alasan Permintaan FPTK
                    '
                    ' Tapi value Ya/Y/N/Tidak dan boolean sejenis
                    ' BUKAN alasan, sehingga tidak boleh overwrite.
                    '
                    ' Jika ada lebih dari satu kolom yang canonical-nya
                    ' CATEGORY, kolom yang berisi alasan valid akan tetap
                    ' diproses pada iterasi berikutnya.
                    ' ==================================================
                    If StrComp( _
                        headerKey, _
                        "CATEGORY", _
                        vbTextCompare) = 0 Then

                        If Not FPTK_IsValidSourceReasonValue( _
                            sourceValue) Then

                            GoTo NextSourceColumn

                        End If

                    End If

                    If Not ValueIsBlankFPTK( _
                        sourceValue) Then

                        targetValue = _
                            rowData(targetColumn)

                        If Not ValuesEqualFPTK( _
                            targetValue, _
                            sourceValue) Then

                            rowData(targetColumn) = _
                                sourceValue

                            rowChanged = True

                        End If

                    End If

                End If

            End If

        End If

NextSourceColumn:

    Next sourceColumn

    FPTK_MapSourceIntoRowArray = _
        rowChanged

End Function

Private Sub FPTK_StampAuditRowArray( _
    ByRef rowData As Variant, _
    ByVal targetMap As Object, _
    ByVal compileAction As String, _
    ByVal sourceFileName As String, _
    ByVal isInsert As Boolean)

    Dim colCreatedAt As Long
    Dim colLastUpdatedAt As Long
    Dim colLastAction As Long
    Dim colSourceFile As Long

    Dim stampTime As Date

    stampTime = Now

    colCreatedAt = ColumnByKey( _
        targetMap, _
        "CREATED AT")

    colLastUpdatedAt = ColumnByKey( _
        targetMap, _
        "LAST UPDATED AT")

    colLastAction = ColumnByKey( _
        targetMap, _
        "LAST COMPILE ACTION")

    colSourceFile = ColumnByKey( _
        targetMap, _
        "SOURCE FILE")

    If isInsert Then

        If colCreatedAt > 0 Then

            If ValueIsBlankFPTK( _
                rowData(colCreatedAt)) Then

                rowData(colCreatedAt) = _
                    stampTime

            End If

        End If

        If colLastUpdatedAt > 0 Then
            rowData(colLastUpdatedAt) = Empty
        End If

    Else

        If colLastUpdatedAt > 0 Then
            rowData(colLastUpdatedAt) = stampTime
        End If

    End If

    If colLastAction > 0 Then

        rowData(colLastAction) = _
            UCase$(Trim$(compileAction))

    End If

    If colSourceFile > 0 Then

        rowData(colSourceFile) = _
            sourceFileName

    End If

End Sub
Private Sub FPTK_RefreshDerivedRowArray( _
    ByRef rowData As Variant, _
    ByVal targetMap As Object)

    Dim colLevelNumber As Long
    Dim colTanggalReal As Long
    Dim colJumlahSLA As Long
    Dim colDeadline As Long
    Dim colDetailSLA As Long
    Dim colStatus As Long
    Dim colOffering As Long

    Dim colFPTKDateKode As Long
    Dim colWeekFPTK As Long
    Dim colMonthFPTK As Long

    Dim colCancelDate As Long
    Dim colWeekCancel As Long
    Dim colMonthCancel As Long

    Dim colWeekOffering As Long
    Dim colMonthOffering As Long

    Dim colPosisi As Long
    Dim colFilter As Long

    Dim slaDays As Long
    Dim parsedDate As Date

    colLevelNumber = ColumnByKey( _
        targetMap, _
        "LEVEL NUMBER")

    colTanggalReal = ColumnByKey( _
        targetMap, _
        "FPTK DATE REAL")

    colJumlahSLA = ColumnByKey( _
        targetMap, _
        "JUMLAH SLA")

    colDeadline = ColumnByKey( _
        targetMap, _
        "DEADLINE SLA")

    colDetailSLA = ColumnByKey( _
        targetMap, _
        "DETAIL SLA")

    colStatus = ColumnByKey( _
        targetMap, _
        "STATUS")

    colOffering = ColumnByKey( _
        targetMap, _
        "OFFERING DATE")

    colFPTKDateKode = ColumnByKey( _
        targetMap, _
        "FPTK DATE KODE")

    colWeekFPTK = ColumnByKey( _
        targetMap, _
        "WEEK FPTK DATE")

    colMonthFPTK = ColumnByKey( _
        targetMap, _
        "MONTH FPTK DATE")

    colCancelDate = ColumnByKey( _
        targetMap, _
        "FPTK CANCEL DATE")

    colWeekCancel = ColumnByKey( _
        targetMap, _
        "WEEK CANCEL DATE")

    colMonthCancel = ColumnByKey( _
        targetMap, _
        "MONTH CANCEL DATE")

    colWeekOffering = ColumnByKey( _
        targetMap, _
        "WEEK OFFERING DATE")

    colMonthOffering = ColumnByKey( _
        targetMap, _
        "MONTH OFFERING")

    colPosisi = ColumnByKey( _
        targetMap, _
        "POSISI")

    colFilter = ColumnByKey( _
        targetMap, _
        "FILTER KATEGORISASI FPTK")

    slaDays = 0

    If colLevelNumber > 0 Then

        If Not ValueIsBlankFPTK( _
            rowData( _
                colLevelNumber)) Then

            slaDays = _
                GetSLAByLevelFPTK( _
                    rowData( _
                        colLevelNumber))

        End If

    End If

    If slaDays <= 0 And _
       colJumlahSLA > 0 Then

        If IsNumeric( _
            rowData( _
                colJumlahSLA)) Then

            If CLng( _
                rowData( _
                    colJumlahSLA)) > 0 Then

                slaDays = _
                    CLng( _
                        rowData( _
                            colJumlahSLA))

            End If

        End If

    End If

    If colJumlahSLA > 0 And _
       slaDays > 0 Then

        rowData(colJumlahSLA) = _
            slaDays

    End If

    If colDeadline > 0 Then

        If colTanggalReal > 0 And _
           slaDays > 0 Then

            parsedDate = ParseFlexibleDate( _
                rowData(colTanggalReal))

            If parsedDate > 0 Then

                rowData(colDeadline) = _
                    parsedDate + slaDays

            Else

                rowData(colDeadline) = Empty

            End If

        Else

            rowData(colDeadline) = Empty

        End If

    End If

    If colDetailSLA > 0 And _
       colStatus > 0 Then

        rowData(colDetailSLA) = _
            CalculateSLAStatusFPTK( _
                rowData(colStatus), _
                FPTK_RowValue( _
                    rowData, _
                    colDeadline), _
                FPTK_RowValue( _
                    rowData, _
                    colOffering))

    End If

    FPTK_WriteWeekMonthToRowArray _
        rowData, _
        colFPTKDateKode, _
        colWeekFPTK, _
        colMonthFPTK

    FPTK_WriteWeekMonthToRowArray _
        rowData, _
        colCancelDate, _
        colWeekCancel, _
        colMonthCancel

    FPTK_WriteWeekMonthToRowArray _
        rowData, _
        colOffering, _
        colWeekOffering, _
        colMonthOffering

    If colFilter > 0 Then

        rowData(colFilter) = _
            CalculateFilterKategorisasiFPTK( _
                FPTK_RowValue( _
                    rowData, _
                    colPosisi), _
                FPTK_RowValue( _
                    rowData, _
                    colLevelNumber), _
                FPTK_RowValue( _
                    rowData, _
                    colFilter))

    End If

End Sub

Private Function FPTK_RowValue( _
    ByVal rowData As Variant, _
    ByVal columnIndex As Long) As Variant

    If columnIndex > 0 Then
        FPTK_RowValue = rowData(columnIndex)
    Else
        FPTK_RowValue = Empty
    End If

End Function
Private Sub FPTK_WriteWeekMonthToRowArray( _
    ByRef rowData As Variant, _
    ByVal sourceDateColumn As Long, _
    ByVal weekTargetColumn As Long, _
    ByVal monthTargetColumn As Long)

    Dim parsedDate As Date
    Dim weekNumber As Long

    If sourceDateColumn <= 0 Then Exit Sub

    parsedDate = ParseFlexibleDate( _
        rowData(sourceDateColumn))

    If parsedDate > 0 Then

        If weekTargetColumn > 0 Then

            On Error Resume Next

            weekNumber = _
                Application.WorksheetFunction. _
                    WeekNum(parsedDate, 1)

            If Err.Number <> 0 Then

                Err.Clear

                weekNumber = DatePart( _
                    "ww", _
                    parsedDate, _
                    vbSunday, _
                    vbFirstJan1)

            End If

            On Error GoTo 0

            rowData(weekTargetColumn) = _
                weekNumber

        End If

        If monthTargetColumn > 0 Then

            rowData(monthTargetColumn) = _
                Format$(parsedDate, "mmmm")

        End If

    Else

        If weekTargetColumn > 0 Then
            rowData(weekTargetColumn) = Empty
        End If

        If monthTargetColumn > 0 Then
            rowData(monthTargetColumn) = Empty
        End If

    End If

End Sub

Private Sub FPTK_WriteTargetDataBlock( _
    ByVal ws As Worksheet, _
    ByVal targetData As Variant, _
    ByVal rowCount As Long, _
    ByVal columnCount As Long)

    If rowCount <= 0 Then Exit Sub
    If columnCount <= 0 Then Exit Sub

    ws.Cells( _
        FPTK_FIRST_DATA_ROW, _
        1).Resize( _
            rowCount, _
            columnCount).Value2 = _
                targetData

End Sub

Private Sub FPTK_WriteInsertedRowsBlock( _
    ByVal ws As Worksheet, _
    ByRef insertedRows() As Variant, _
    ByVal insertedCount As Long, _
    ByVal existingRowCount As Long, _
    ByVal columnCount As Long)

    Dim outputData() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim rowData As Variant
    Dim firstWriteRow As Long

    If insertedCount <= 0 Then Exit Sub
    If columnCount <= 0 Then Exit Sub

    ReDim outputData( _
        1 To insertedCount, _
        1 To columnCount)

    For rowIndex = 1 To insertedCount

        rowData = insertedRows(rowIndex)

        For columnIndex = 1 To columnCount

            outputData( _
                rowIndex, _
                columnIndex) = _
                    rowData(columnIndex)

        Next columnIndex

    Next rowIndex

    firstWriteRow = _
        FPTK_FIRST_DATA_ROW + _
        existingRowCount

    ws.Cells( _
        firstWriteRow, _
        1).Resize( _
            insertedCount, _
            columnCount).Value2 = _
                outputData

End Sub

Private Function FPTK_FastElapsedText( _
    ByVal startedAt As Double) As String

    Dim elapsedSeconds As Double

    elapsedSeconds = Timer - startedAt

    If elapsedSeconds < 0 Then
        elapsedSeconds = elapsedSeconds + 86400
    End If

    If elapsedSeconds < 60 Then

        FPTK_FastElapsedText = _
            Format$(elapsedSeconds, "0.0") & _
            " detik"

    Else

        FPTK_FastElapsedText = _
            Format$(Int(elapsedSeconds / 60), "0") & _
            " menit " & _
            Format$(elapsedSeconds Mod 60, "0") & _
            " detik"

    End If

End Function

Private Function CollectFPTKSourceFiles( _
    ByVal folderPath As String) As Collection

    Dim result As New Collection
    Dim extensions As Variant
    Dim extensionItem As Variant

    Dim fileName As String
    Dim fullPath As String

    extensions = Array( _
        "*.xlsx", _
        "*.xlsm")

    For Each extensionItem In extensions

        fileName = Dir$( _
            folderPath & _
            CStr(extensionItem))

        Do While Len(fileName) > 0

            If Left$(fileName, 2) <> "~$" Then

                fullPath = _
                    folderPath & fileName

                If StrComp( _
                    fullPath, _
                    ThisWorkbook.FullName, _
                    vbTextCompare) <> 0 Then

                    result.Add fullPath

                End If

            End If

            fileName = Dir$()

        Loop

    Next extensionItem

    Set CollectFPTKSourceFiles = result

End Function

Private Function FPTKFileNameFromPath( _
    ByVal fullPath As String) As String

    Dim separatorPosition As Long

    separatorPosition = InStrRev( _
        fullPath, _
        "\")

    If separatorPosition = 0 Then

        separatorPosition = InStrRev( _
            fullPath, _
            "/")

    End If

    If separatorPosition > 0 Then

        FPTKFileNameFromPath = Mid$( _
            fullPath, _
            separatorPosition + 1)

    Else

        FPTKFileNameFromPath = fullPath

    End If

End Function

' ============================================================
' SCHEMA FPTK - SAFE OVERLAY
'
' Schema hanya dipakai untuk:
' - register header;
' - canonical alias;
' - hide/show kolom.
'
' Schema tidak dipakai untuk rebuild, reorder, clear,
' atau delete data saat compile.
' ============================================================
Private Sub SyncFPTKSchemaSafe( _
    ByVal applyVisibility As Boolean)

    Dim schemaErrorNumber As Long
    Dim schemaErrorText As String

    On Error GoTo SchemaError

    SchemaRegisterCurrentHeadersSafe _
        "FPTK"

    If applyVisibility Then

        SchemaApplyVisibilityOnly _
            "FPTK"

    End If

    Exit Sub

SchemaError:

    schemaErrorNumber = Err.Number
    schemaErrorText = Err.Description

    On Error Resume Next

    AddRunLog _
        "Warning", _
        "FPTK Schema | Error " & _
        CStr(schemaErrorNumber) & _
        ": " & schemaErrorText

    On Error GoTo 0

End Sub

' ============================================================
' LOAD EXISTING KEY DB KODE POSISI
' ============================================================
' ============================================================
' LOAD EXISTING KEY FPTK
' ============================================================
Private Sub LoadExistingKeysFPTK( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object, _
    ByRef keyDictionary As Object)

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim rowKey As String

    lastRow = LastTargetDataRow( _
        ws, headerMap)

    If lastRow < FPTK_FIRST_DATA_ROW Then Exit Sub

    For rowIndex = FPTK_FIRST_DATA_ROW To lastRow

        rowKey = BuildFPTKKeyFromSheet( _
            ws, rowIndex, headerMap)

        If Len(rowKey) > 0 Then

            If Not keyDictionary.Exists(rowKey) Then
                keyDictionary.Add rowKey, rowIndex
            End If

        End If

    Next rowIndex

End Sub

' ============================================================
' KEY DB KODE POSISI
'
' ============================================================
' KEY FPTK
' KEY = KODE UNIK + POSISI
' Jika tanggal real berubah tetapi Kode Unik dan Posisi sama,
' row lama akan di-update, bukan insert row baru.
' ============================================================
Private Function BuildFPTKKeyFromArray( _
    ByVal sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceKeys As Variant) As String

    Dim kodeUnik As String
    Dim posisi As String

    kodeUnik = CStr(ArrayValueByKey( _
        sourceData, sourceRow, _
        sourceKeys, "KODE UNIK"))

    posisi = CStr(ArrayValueByKey( _
        sourceData, sourceRow, _
        sourceKeys, "POSISI"))

    If Len(NormalizeDataKey(kodeUnik)) = 0 Or _
       Len(NormalizeDataKey(posisi)) = 0 Then

        Exit Function

    End If

    BuildFPTKKeyFromArray = _
        NormalizeDataKey(kodeUnik) & "||" & _
        NormalizeDataKey(posisi)

End Function

Private Function BuildFPTKKeyFromSheet( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal headerMap As Object) As String

    Dim kodeUnik As String
    Dim posisi As String

    kodeUnik = CStr(SheetValueByKey( _
        ws, rowIndex, _
        headerMap, "KODE UNIK"))

    posisi = CStr(SheetValueByKey( _
        ws, rowIndex, _
        headerMap, "POSISI"))

    If Len(NormalizeDataKey(kodeUnik)) = 0 Or _
       Len(NormalizeDataKey(posisi)) = 0 Then

        Exit Function

    End If

    BuildFPTKKeyFromSheet = _
        NormalizeDataKey(kodeUnik) & "||" & _
        NormalizeDataKey(posisi)

End Function

' ============================================================
' RESET KODE ANGKA + KODE UNIK DB KODE POSISI
' ============================================================
' ============================================================
' PROTECT / ENSURE HEADER FPTK
'
' Header FPTK tetap berada di row 1.
' Data pertama dimulai dari row 2.
' ============================================================
Private Sub EnsureFPTKHeaderRow1Layout( _
    ByVal ws As Worksheet)

    Dim row1IsHeader As Boolean
    Dim row2IsHeader As Boolean

    If ws Is Nothing Then
        Exit Sub
    End If

    row1IsHeader = _
        FPTK_RowLooksLikeHeader( _
            ws, _
            1)

    row2IsHeader = _
        FPTK_RowLooksLikeHeader( _
            ws, _
            2)

    If Not row1IsHeader And _
       row2IsHeader Then

        ws.rows(1).Delete _
            Shift:=xlUp

        AddRunLog _
            "Info", _
            "FPTK layout migrated: header row 2 -> row 1."

        row1IsHeader = True

    End If

    If row1IsHeader Then

        Do While FPTK_RowLooksLikeHeader( _
            ws, _
            2)

            ws.rows(2).Delete _
                Shift:=xlUp

        Loop

    End If

End Sub

Private Function FPTK_RowLooksLikeHeader( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long) As Boolean

    Dim headerMap As Object

    If ws Is Nothing Then Exit Function
    If rowIndex <= 0 Then Exit Function

    Set headerMap = BuildHeaderMap( _
        ws, _
        rowIndex)

    If headerMap Is Nothing Then
        Exit Function
    End If

    FPTK_RowLooksLikeHeader = _
        headerMap.Exists( _
            CanonicalHeaderKey( _
                "KODE UNIK")) And _
        headerMap.Exists( _
            CanonicalHeaderKey( _
                "POSISI"))

End Function
Public Sub ProtectFPTKHeader()

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            "FPTK")

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    EnsureFPTKHeaderRow1Layout _
        ws

    EnsureRequiredHeaders _
        ws

    ' Di tahap awal compile hanya gabungkan alias.
    ' Jangan memindahkan seluruh kolom karena row 1
    ' pada layout header row 1 langsung berisi header FPTK.
    FPTK_CleanupLegacyColumns _
        ws

    FPTK_ApplyAvailabilityValidation _
        ws

    FPTK_ApplyStatusValidation _
        ws

    StyleHeaderDynamic _
        ws

End Sub

' ============================================================
' FAST DERIVED VALUES - HANYA ROW INSERT / UPDATE
' ============================================================
Public Sub FillFPTKHelperFormulas()

    Dim ws As Worksheet
    Dim headerMap As Object
    Dim rowsToRefresh As Object

    Dim lastRow As Long
    Dim rowIndex As Long

    Set ws = GetOrCreateWorksheet( _
        "FPTK")

    ProtectFPTKHeader

    Set headerMap = BuildHeaderMap( _
        ws, _
        FPTK_HEADER_ROW)

    Set rowsToRefresh = CreateObject( _
        "Scripting.Dictionary")

    rowsToRefresh.CompareMode = _
        vbTextCompare

    lastRow = LastTargetDataRow( _
        ws, _
        headerMap)

    If lastRow >= FPTK_FIRST_DATA_ROW Then

        For rowIndex = FPTK_FIRST_DATA_ROW _
            To lastRow

            If RowHasPrimaryData( _
                ws, _
                rowIndex, _
                headerMap) Then

                MarkFPTKChangedRow _
                    rowsToRefresh, _
                    rowIndex

            End If

        Next rowIndex

    End If

    RefreshFPTKDerivedValuesForChangedRows _
        ws, _
        headerMap, _
        rowsToRefresh

    ' Re-apply STO setelah kategori dasar Level/CLAP dihitung.
    RefreshFPTKSTOCategorization _
        ws, _
        headerMap, _
        True

    AddRunLog _
        "Info", _
        "FillFPTKHelperFormulas done | Rows: " & _
        CStr(rowsToRefresh.Count)

    Set rowsToRefresh = Nothing
    Set headerMap = Nothing
    Set ws = Nothing

End Sub

Private Sub MarkFPTKChangedRow( _
    ByVal changedRows As Object, _
    ByVal rowIndex As Long)

    Dim rowKey As String

    If changedRows Is Nothing Then Exit Sub
    If rowIndex < FPTK_FIRST_DATA_ROW Then Exit Sub

    rowKey = CStr(rowIndex)

    If Not changedRows.Exists(rowKey) Then

        changedRows.Add _
            rowKey, _
            rowIndex

    End If

End Sub
Private Sub RefreshFPTKDerivedValuesForChangedRows( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object, _
    ByVal changedRows As Object)

    Dim rowKey As Variant
    Dim rowIndex As Long

    Dim colLevelNumber As Long
    Dim colTanggalReal As Long
    Dim colJumlahSLA As Long
    Dim colDeadline As Long
    Dim colDetailSLA As Long
    Dim colStatus As Long
    Dim colOffering As Long

    Dim colFPTKDateKode As Long
    Dim colWeekFPTK As Long
    Dim colMonthFPTK As Long

    Dim colCancelDate As Long
    Dim colWeekCancel As Long
    Dim colMonthCancel As Long

    Dim colWeekOffering As Long
    Dim colMonthOffering As Long

    Dim colPosisi As Long
    Dim colFilter As Long

    Dim slaDays As Long
    Dim processedCount As Long
    Dim totalChangedRows As Long

    If ws Is Nothing Then Exit Sub
    If headerMap Is Nothing Then Exit Sub
    If changedRows Is Nothing Then Exit Sub
    If changedRows.Count = 0 Then Exit Sub

    colLevelNumber = ColumnByKey( _
        headerMap, _
        "LEVEL NUMBER")

    colTanggalReal = ColumnByKey( _
        headerMap, _
        "FPTK DATE REAL")

    colJumlahSLA = ColumnByKey( _
        headerMap, _
        "JUMLAH SLA")

    colDeadline = ColumnByKey( _
        headerMap, _
        "DEADLINE SLA")

    colDetailSLA = ColumnByKey( _
        headerMap, _
        "DETAIL SLA")

    colStatus = ColumnByKey( _
        headerMap, _
        "STATUS")

    colOffering = ColumnByKey( _
        headerMap, _
        "OFFERING DATE")

    colFPTKDateKode = ColumnByKey( _
        headerMap, _
        "FPTK DATE KODE")

    colWeekFPTK = ColumnByKey( _
        headerMap, _
        "WEEK FPTK DATE")

    colMonthFPTK = ColumnByKey( _
        headerMap, _
        "MONTH FPTK DATE")

    colCancelDate = ColumnByKey( _
        headerMap, _
        "FPTK CANCEL DATE")

    colWeekCancel = ColumnByKey( _
        headerMap, _
        "WEEK CANCEL DATE")

    colMonthCancel = ColumnByKey( _
        headerMap, _
        "MONTH CANCEL DATE")

    colWeekOffering = ColumnByKey( _
        headerMap, _
        "WEEK OFFERING DATE")

    colMonthOffering = ColumnByKey( _
        headerMap, _
        "MONTH OFFERING")

    colPosisi = ColumnByKey( _
        headerMap, _
        "POSISI")

    colFilter = ColumnByKey( _
        headerMap, _
        "FILTER KATEGORISASI FPTK")

    totalChangedRows = changedRows.Count

    For Each rowKey In changedRows.keys

        processedCount = processedCount + 1

        If processedCount = 1 Or _
           processedCount Mod 100 = 0 Or _
           processedCount = totalChangedRows Then

            Application.StatusBar = _
                "FPTK | Refresh row " & _
                CStr(processedCount) & "/" & _
                CStr(totalChangedRows)

            HeartbeatProgress _
                "RUNNING", _
                "FPTK", _
                "FPTK - Refresh Derived", _
                CStr(processedCount) & "/" & _
                CStr(totalChangedRows)

            DoEvents

        End If

        rowIndex = CLng( _
            changedRows(rowKey))

        If RowHasPrimaryData( _
            ws, _
            rowIndex, _
            headerMap) Then

            ' ================================================
            ' SLA Ã¯Â¿Â½ sumber tanggal hanya FPTK Date (Real)
            ' ================================================
            If colLevelNumber > 0 And _
               colJumlahSLA > 0 Then

                slaDays = GetSLAByLevelFPTK( _
                    ws.Cells( _
                        rowIndex, _
                        colLevelNumber).Value2)

                ws.Cells( _
                    rowIndex, _
                    colJumlahSLA).Value2 = _
                        slaDays

            Else

                slaDays = 0

            End If

            If colDeadline > 0 Then

                If colTanggalReal > 0 And _
                   slaDays > 0 And _
                   IsDateValueFlexible( _
                        ws.Cells( _
                            rowIndex, _
                            colTanggalReal).Value2) Then

                    ws.Cells( _
                        rowIndex, _
                        colDeadline).Value = _
                            ParseFlexibleDate( _
                                ws.Cells( _
                                    rowIndex, _
                                    colTanggalReal).Value2) + _
                            slaDays

                    ws.Cells( _
                        rowIndex, _
                        colDeadline).numberFormat = _
                            "dd/mm/yyyy"

                Else

                    ws.Cells( _
                        rowIndex, _
                        colDeadline).ClearContents

                End If

            End If

            If colDetailSLA > 0 And _
               colStatus > 0 Then

                ws.Cells( _
                    rowIndex, _
                    colDetailSLA).Value2 = _
                        CalculateSLAStatusFPTK( _
                            ws.Cells( _
                                rowIndex, _
                                colStatus).Value2, _
                            ValueFromOptionalColumn( _
                                ws, _
                                rowIndex, _
                                colDeadline), _
                            ValueFromOptionalColumn( _
                                ws, _
                                rowIndex, _
                                colOffering))

            End If

            ' ================================================
            ' WEEK / MONTH Ã¯Â¿Â½ ditulis sebagai VALUE
            ' ================================================
            WriteFPTKWeekMonthValue _
                ws, _
                rowIndex, _
                colFPTKDateKode, _
                colWeekFPTK, _
                colMonthFPTK

            WriteFPTKWeekMonthValue _
                ws, _
                rowIndex, _
                colCancelDate, _
                colWeekCancel, _
                colMonthCancel

            WriteFPTKWeekMonthValue _
                ws, _
                rowIndex, _
                colOffering, _
                colWeekOffering, _
                colMonthOffering

            ' ================================================
            ' FILTER KATEGORISASI Ã¯Â¿Â½ ditulis sebagai VALUE
            ' ================================================
            If colFilter > 0 Then

                ws.Cells( _
                    rowIndex, _
                    colFilter).Value2 = _
                        CalculateFilterKategorisasiFPTK( _
                            ValueFromOptionalColumn( _
                                ws, _
                                rowIndex, _
                                colPosisi), _
                            ValueFromOptionalColumn( _
                                ws, _
                                rowIndex, _
                                colLevelNumber), _
                            ValueFromOptionalColumn( _
                                ws, _
                                rowIndex, _
                                colFilter))

            End If

        End If

    Next rowKey

End Sub

Public Sub SetTulangPunggungSTOFile()

    Dim picker As FileDialog
    Dim selectedPath As String

    Set picker = _
        Application.FileDialog( _
            msoFileDialogFilePicker)

    With picker

        .Title = _
            "Pilih file STO.xlsm"

        .AllowMultiSelect = False

        .Filters.Clear

        .Filters.Add _
            "Excel Workbook", _
            "*.xlsm;*.xlsx;*.xlsb;*.xls"

        If .Show <> -1 Then
            Exit Sub
        End If

        selectedPath = _
            CStr( _
                .SelectedItems(1))

    End With

    If Not STO_IsValidCompiledSourcePath( _
        selectedPath) Then

        MsgBox _
            "File yang dipilih bukan STO source yang valid." & _
            vbCrLf & vbCrLf & _
            "File wajib punya sheet FPTK dengan header:" & _
            vbCrLf & _
            "- Posisi" & _
            vbCrLf & _
            "- FPTK Date (Real)" & _
            vbCrLf & _
            "- Status FPTK / FPTK Availability", _
            vbExclamation, _
            "STO Categorization"

        Set picker = Nothing

        Exit Sub

    End If

    STO_SaveSettingValue _
        "STOCompiledFilePath", _
        selectedPath

    MsgBox _
        "File STO source sudah disimpan:" & _
        vbCrLf & _
        selectedPath, _
        vbInformation, _
        "STO Categorization"

    Set picker = Nothing

End Sub

Public Sub SetSTOFile()

    SetTulangPunggungSTOFile

End Sub
Public Sub RefreshFPTKSTOAssignmentsNow()

    Dim ws As Worksheet
    Dim headerMap As Object
    Dim sourcePath As String

    sourcePath = STO_ResolveSourcePath()

    If Len(sourcePath) = 0 Then
        SetSTOFile
        sourcePath = STO_ResolveSourcePath()
    End If

    If Len(sourcePath) = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("FPTK")
    Set headerMap = BuildHeaderMap(ws, FPTK_HEADER_ROW)

    ' STO source diproses 1-to-1. Tidak ada legacy delete otomatis.
    STO_SyncAllAssignmentsByPositionRealDate _
        ws, _
        True, _
        sourcePath

    Set headerMap = BuildHeaderMap(ws, FPTK_HEADER_ROW)

    FPTK_ApplyAvailabilityValidation ws

    MsgBox _
        "STO refresh selesai." & vbCrLf & vbCrLf & _
        "Source:" & vbCrLf & sourcePath, _
        vbInformation, _
        "STO Refresh"

    Set headerMap = Nothing
    Set ws = Nothing

End Sub

Public Sub Compile_STO()

    RefreshFPTKSTOAssignmentsNow

End Sub

Public Sub Diagnose_STO_Source()

    Dim sourcePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedByMacro As Boolean
    Dim headerRow As Long
    Dim colAvail As Long
    Dim colPosisi As Long
    Dim colNama As Long
    Dim lastRow As Long
    Dim r As Long
    Dim yCount As Long
    Dim nCount As Long
    Dim otherCount As Long
    Dim t As String

    sourcePath = STO_ResolveSourcePath()

    If Len(sourcePath) = 0 Then
        MsgBox "Source STO belum terdeteksi. Jalankan SetSTOFile dulu.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wb = Workbooks(STO_FileNameFromPath(sourcePath))
    On Error GoTo 0

    If wb Is Nothing Then
        Set wb = Workbooks.Open(sourcePath, ReadOnly:=True, UpdateLinks:=0)
        openedByMacro = True
    End If

    On Error Resume Next
    Set ws = wb.Worksheets("FPTK")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet FPTK tidak ditemukan di source STO.", vbCritical
        GoTo SafeExit
    End If

    headerRow = STO_FindCompiledHeaderRow(ws)
    If headerRow = 0 Then
        MsgBox "Header STO tidak ditemukan.", vbCritical
        GoTo SafeExit
    End If

    colAvail = STO_FindSourceAvailabilityColumn(ws, headerRow)
    colPosisi = FPTK_DirectFindHeader(ws, headerRow, "Posisi")
    colNama = FPTK_DirectFindHeader(ws, headerRow, "Nama Kandidat")

    If colAvail = 0 Then
        MsgBox "Status FPTK / FPTK Availability tidak ditemukan.", vbCritical
        GoTo SafeExit
    End If

    lastRow = ws.Cells(ws.rows.Count, colAvail).End(xlUp).Row

    For r = headerRow + 1 To lastRow
        t = STO_NormalizeAvailability(ws.Cells(r, colAvail).Value2)
        Select Case t
            Case "Y": yCount = yCount + 1
            Case "N": nCount = nCount + 1
            Case Else
                If Len(Trim$(CStr(ws.Cells(r, colAvail).Value2))) > 0 Then
                    otherCount = otherCount + 1
                End If
        End Select
    Next r

    MsgBox _
        "Source STO terdeteksi." & vbCrLf & vbCrLf & _
        "Header row: " & CStr(headerRow) & vbCrLf & _
        "Kolom Posisi: " & CStr(colPosisi) & vbCrLf & _
        "Kolom Nama Kandidat: " & CStr(colNama) & vbCrLf & _
        "Kolom Status FPTK/Availability: " & CStr(colAvail) & vbCrLf & _
        "YES (V/Ya/Y): " & CStr(yCount) & vbCrLf & _
        "NO (X/Tidak/N): " & CStr(nCount) & vbCrLf & _
        "Other nonblank: " & CStr(otherCount), _
        vbInformation, _
        "Diagnose STO Source"

SafeExit:
    On Error Resume Next
    If openedByMacro Then wb.Close SaveChanges:=False
    Set ws = Nothing
    Set wb = Nothing

End Sub

Public Sub RefreshFPTKSTONow()

    Dim ws As Worksheet
    Dim headerMap As Object
    Dim sourcePath As String

    Set ws = _
        ThisWorkbook.Worksheets( _
            "FPTK")

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    sourcePath = _
        STO_ResolveSourcePath()

    If Len(sourcePath) = 0 Then

        SetSTOFile

        sourcePath = _
            STO_ResolveSourcePath()

    End If

    If Len(sourcePath) = 0 Then
        Exit Sub
    End If

    RefreshFPTKSTOCategorization _
        ws, _
        headerMap, _
        False

    Set headerMap = Nothing
    Set ws = Nothing

End Sub
Private Function STO_NormalizeAvailability( _
    ByVal rawValue As Variant) As String

    Dim t As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    t = UCase$(Trim$(CStr(rawValue)))
    t = Replace$(t, ChrW(160), " ")
    t = Application.Trim(t)

    Select Case t

        Case "V", "Y", "YA", "YES", "TRUE", "1"
            STO_NormalizeAvailability = "Y"

        Case "X", "N", "TIDAK", "NO", "FALSE", "0"
            STO_NormalizeAvailability = "N"

    End Select

End Function

Private Function STO_FindSourceAvailabilityColumn( _
    ByVal ws As Object, _
    ByVal headerRow As Long) As Long

    Dim aliases As Variant
    Dim item As Variant
    Dim resultColumn As Long

    If ws Is Nothing Then Exit Function
    If headerRow <= 0 Then Exit Function

    aliases = Array( _
        "Status FPTK", _
        "FPTK Availability", _
        "Availability FPTK", _
        "FPTK Available", _
        "Available FPTK")

    For Each item In aliases

        resultColumn = _
            FPTK_DirectFindHeader( _
                ws, _
                headerRow, _
                CStr(item))

        If resultColumn > 0 Then
            STO_FindSourceAvailabilityColumn = resultColumn
            Exit Function
        End If

    Next item

End Function

Private Function STO_SourceWorksheetHasVX( _
    ByVal ws As Object, _
    ByVal headerRow As Long) As Boolean

    Dim availabilityColumn As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim availabilityText As String

    If ws Is Nothing Then Exit Function
    If headerRow <= 0 Then Exit Function

    availabilityColumn = _
        STO_FindSourceAvailabilityColumn( _
            ws, _
            headerRow)

    If availabilityColumn <= 0 Then Exit Function

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            availabilityColumn).End(xlUp).Row

    For rowIndex = headerRow + 1 To lastRow

        availabilityText = _
            STO_NormalizeAvailability( _
                ws.Cells( _
                    rowIndex, _
                    availabilityColumn).Value2)

        If availabilityText = "Y" Or _
           availabilityText = "N" Then

            STO_SourceWorksheetHasVX = True
            Exit Function

        End If

    Next rowIndex

End Function

Private Function FPTK_IsDedicatedSTOSourceFile( _
    ByVal fullPath As String) As Boolean

    Dim fileName As String
    Dim configuredPath As String

    fileName = UCase$(Trim$(FPTKFileNameFromPath(fullPath)))

    If fileName = "STO.XLSM" Or _
       fileName = "STO.XLSX" Then

        FPTK_IsDedicatedSTOSourceFile = True
        Exit Function

    End If

    If InStr(1, fileName, "TULANG PUNGGUNG", vbTextCompare) > 0 Then
        FPTK_IsDedicatedSTOSourceFile = True
        Exit Function
    End If

    configuredPath = _
        Trim$( _
            STO_GetSettingValue( _
                "STOCompiledFilePath"))

    If Len(configuredPath) > 0 Then
        If STO_PathsEqual(fullPath, configuredPath) Then
            FPTK_IsDedicatedSTOSourceFile = True
        End If
    End If

End Function

Private Function STO_PathsEqual( _
    ByVal pathA As String, _
    ByVal pathB As String) As Boolean

    Dim cleanA As String
    Dim cleanB As String

    cleanA = _
        Trim$( _
            Replace( _
                pathA, _
                "/", _
                "\"))

    cleanB = _
        Trim$( _
            Replace( _
                pathB, _
                "/", _
                "\"))

    Do While Right$(cleanA, 1) = "\"
        cleanA = Left$(cleanA, Len(cleanA) - 1)
    Loop

    Do While Right$(cleanB, 1) = "\"
        cleanB = Left$(cleanB, Len(cleanB) - 1)
    Loop

    STO_PathsEqual = _
        (StrComp( _
            cleanA, _
            cleanB, _
            vbTextCompare) = 0)

End Function
Private Sub STO_RemoveLegacyVDuplicates( _
    ByVal wsTarget As Worksheet, _
    Optional ByVal silentMode As Boolean = True, _
    Optional ByVal sourcePathOverride As String = vbNullString)

    Dim sourcePath As String
    Dim vMap As Object

    Dim colPosisi As Long
    Dim colNama As Long
    Dim colSourceFile As Long

    Dim lastRow As Long
    Dim rowIndex As Long

    Dim posisiKey As String
    Dim namaKey As String
    Dim pairKey As String

    Dim sourceFileText As String
    Dim sourceFileName As String

    Dim originalPairMap As Object
    Dim deleteRows As Collection
    Dim item As Variant

    Dim deletedCount As Long

    If wsTarget Is Nothing Then Exit Sub

    sourcePath = _
        Trim$(sourcePathOverride)

    If Len(sourcePath) = 0 Then

        sourcePath = _
            STO_ResolveSourcePath()

    End If

    If Len(sourcePath) = 0 Then Exit Sub

    Set vMap = _
        STO_LoadCandidateMap( _
            sourcePath)

    If vMap Is Nothing Then Exit Sub
    If vMap.Count = 0 Then
        Set vMap = Nothing
        Exit Sub
    End If

    colPosisi = _
        STO_FindTargetHeaderAny( _
            wsTarget, _
            Array("Posisi"))

    colNama = _
        STO_FindTargetHeaderAny( _
            wsTarget, _
            Array("Nama Kandidat"))

    colSourceFile = _
        STO_FindTargetHeaderAny( _
            wsTarget, _
            Array("Source File"))

    If colPosisi <= 0 Or _
       colNama <= 0 Or _
       colSourceFile <= 0 Then

        Set vMap = Nothing
        Exit Sub

    End If

    Set originalPairMap = _
        CreateObject( _
            "Scripting.Dictionary")

    originalPairMap.CompareMode = _
        vbTextCompare

    Set deleteRows = _
        New Collection

    lastRow = _
        Application.Max( _
            wsTarget.Cells( _
                wsTarget.rows.Count, _
                colPosisi).End( _
                    xlUp).Row, _
            wsTarget.Cells( _
                wsTarget.rows.Count, _
                colNama).End( _
                    xlUp).Row, _
            wsTarget.Cells( _
                wsTarget.rows.Count, _
                colSourceFile).End( _
                    xlUp).Row)

    ' --------------------------------------------------------
    ' PASS 1:
    ' Catat V pairs yang punya row asli NON-STO.
    ' --------------------------------------------------------
    For rowIndex = _
        FPTK_FIRST_DATA_ROW _
        To lastRow

        posisiKey = _
            STO_NormalizePosition( _
                wsTarget.Cells( _
                    rowIndex, _
                    colPosisi).Value2)

        namaKey = _
            STO_NormalizeCandidate( _
                wsTarget.Cells( _
                    rowIndex, _
                    colNama).Value2)

        If Len(posisiKey) > 0 And _
           Len(namaKey) > 0 Then

            pairKey = _
                posisiKey & _
                "|" & _
                namaKey

            If vMap.Exists( _
                pairKey) Then

                sourceFileText = _
                    Trim$( _
                        CStr( _
                            wsTarget.Cells( _
                                rowIndex, _
                                colSourceFile).Value2))

                sourceFileName = _
                    UCase$( _
                        STO_FileNameFromPath( _
                            sourceFileText))

                If sourceFileName <> "STO.XLSM" And _
                   sourceFileName <> "STO.XLSX" And _
                   UCase$(sourceFileText) <> "STO.XLSM" And _
                   UCase$(sourceFileText) <> "STO.XLSX" Then

                    If Not originalPairMap.Exists( _
                        pairKey) Then

                        originalPairMap.Add _
                            pairKey, _
                            rowIndex

                    End If

                End If

            End If

        End If

    Next rowIndex

    ' --------------------------------------------------------
    ' PASS 2:
    ' Kalau V pair sudah punya row asli, row hasil legacy
    ' generic INSERT dari STO.xlsm adalah duplicate -> delete.
    ' --------------------------------------------------------
    For rowIndex = _
        FPTK_FIRST_DATA_ROW _
        To lastRow

        posisiKey = _
            STO_NormalizePosition( _
                wsTarget.Cells( _
                    rowIndex, _
                    colPosisi).Value2)

        namaKey = _
            STO_NormalizeCandidate( _
                wsTarget.Cells( _
                    rowIndex, _
                    colNama).Value2)

        If Len(posisiKey) > 0 And _
           Len(namaKey) > 0 Then

            pairKey = _
                posisiKey & _
                "|" & _
                namaKey

            If vMap.Exists( _
                pairKey) And _
               originalPairMap.Exists( _
                pairKey) Then

                sourceFileText = _
                    Trim$( _
                        CStr( _
                            wsTarget.Cells( _
                                rowIndex, _
                                colSourceFile).Value2))

                sourceFileName = _
                    UCase$( _
                        STO_FileNameFromPath( _
                            sourceFileText))

                If sourceFileName = "STO.XLSM" Or _
                   sourceFileName = "STO.XLSX" Or _
                   UCase$(sourceFileText) = "STO.XLSM" Or _
                   UCase$(sourceFileText) = "STO.XLSX" Then

                    deleteRows.Add _
                        rowIndex

                End If

            End If

        End If

    Next rowIndex

    ' Delete bottom-up so row numbers stay valid.
    For rowIndex = deleteRows.Count _
        To 1 Step -1

        wsTarget.rows( _
            CLng( _
                deleteRows(rowIndex))).Delete

        deletedCount = _
            deletedCount + 1

    Next rowIndex

    If deletedCount > 0 Then

        AddRunLog _
            "Info", _
            "STO Legacy V Duplicate Cleanup | Deleted: " & _
            CStr(deletedCount)

    End If

    If Not silentMode Then

        MsgBox _
            "Legacy duplicate V cleanup selesai." & _
            vbCrLf & _
            "Deleted: " & _
            CStr(deletedCount), _
            vbInformation, _
            "STO Cleanup"

    End If

    Set deleteRows = Nothing
    Set originalPairMap = Nothing
    Set vMap = Nothing

End Sub
Private Function STO_FindExistingBlankSTOBySourceRow( _
    ByVal wsTarget As Worksheet, _
    ByVal colPosisi As Long, _
    ByVal colNama As Long, _
    ByVal colSourceFile As Long, _
    ByVal colAction As Long, _
    ByVal wantedPositionKey As String, _
    ByVal sourceFileName As String, _
    ByVal sourceRow As Long, _
    ByVal availabilityYN As String) As Long

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim posisiKey As String
    Dim namaKey As String
    Dim sourceFileKey As String
    Dim actionKey As String
    Dim wantedAction As String

    If wsTarget Is Nothing Then Exit Function
    If colPosisi <= 0 Then Exit Function
    If colAction <= 0 Then Exit Function
    If Len(wantedPositionKey) = 0 Then Exit Function

    wantedAction = _
        "INSERT STO " & _
        UCase$(Trim$(availabilityYN)) & _
        "|SRCROW=" & _
        CStr(sourceRow)

    lastRow = _
        wsTarget.Cells( _
            wsTarget.rows.Count, _
            colPosisi).End(xlUp).Row

    For rowIndex = FPTK_FIRST_DATA_ROW To lastRow

        posisiKey = _
            STO_NormalizePosition( _
                wsTarget.Cells( _
                    rowIndex, _
                    colPosisi).Value2)

        If StrComp( _
            posisiKey, _
            wantedPositionKey, _
            vbTextCompare) <> 0 Then

            GoTo NextRow

        End If

        If colNama > 0 Then

            namaKey = _
                STO_NormalizeCandidate( _
                    wsTarget.Cells( _
                        rowIndex, _
                        colNama).Value2)

            If Len(namaKey) > 0 Then
                GoTo NextRow
            End If

        End If

        actionKey = _
            UCase$( _
                Trim$( _
                    CStr( _
                        wsTarget.Cells( _
                            rowIndex, _
                            colAction).Value2)))

        If actionKey <> wantedAction Then
            GoTo NextRow
        End If

        If colSourceFile > 0 And _
           Len(sourceFileName) > 0 Then

            sourceFileKey = _
                UCase$( _
                    STO_FileNameFromPath( _
                        CStr( _
                            wsTarget.Cells( _
                                rowIndex, _
                                colSourceFile).Value2)))

            If StrComp( _
                sourceFileKey, _
                UCase$(sourceFileName), _
                vbTextCompare) <> 0 Then

                GoTo NextRow

            End If

        End If

        STO_FindExistingBlankSTOBySourceRow = _
            rowIndex

        Exit Function

NextRow:

    Next rowIndex

End Function

Private Sub STO_SyncVAssignmentsComplete( _
    ByVal wsTarget As Worksheet, _
    ByVal headerMap As Object, _
    Optional ByVal silentMode As Boolean = True, _
    Optional ByVal sourcePathOverride As String = vbNullString)

    ' Backward-compatible wrapper.
    ' Seluruh record STO sekarang diproses sekaligus agar
    ' Posisi + FPTK Date (Real) bisa di-match 1-to-1.
    STO_SyncAllAssignmentsByPositionRealDate _
        wsTarget, _
        silentMode, _
        sourcePathOverride

End Sub

Private Sub STO_SyncAllAssignmentsByPositionRealDate( _
    ByVal wsTarget As Worksheet, _
    Optional ByVal silentMode As Boolean = True, _
    Optional ByVal sourcePathOverride As String = vbNullString)

    Dim sourcePath As String
    Dim wbSTO As Workbook
    Dim wsSTO As Worksheet
    Dim openedByMacro As Boolean

    Dim sourceHeaderRow As Long
    Dim sourceHeaderMap As Object
    Dim targetHeaderMap As Object
    Dim targetBuckets As Object
    Dim usedTargetRows As Object

    Dim srcColPosisi As Long
    Dim srcColRealDate As Long
    Dim srcColAvailability As Long
    Dim srcColKodeUnik As Long
    Dim srcColNama As Long

    Dim tgtColKodeUnik As Long
    Dim tgtColPosisi As Long
    Dim tgtColNama As Long
    Dim tgtColRealDate As Long
    Dim tgtColFilter As Long
    Dim tgtColAvailability As Long
    Dim tgtColUpdated As Long
    Dim tgtColAction As Long
    Dim tgtColSourceFile As Long

    Dim lastSourceRow As Long
    Dim lastTargetRow As Long
    Dim sourceRow As Long
    Dim targetRow As Long

    Dim availabilityText As String
    Dim posisiKey As String
    Dim dateKey As String
    Dim matchKey As String
    Dim sourceFileName As String

    Dim sourceRecordCount As Long
    Dim matchedCount As Long
    Dim insertedCount As Long
    Dim existingInsertedCount As Long
    Dim skippedNoPositionCount As Long
    Dim blankDateCount As Long
    Dim duplicateKeyCount As Long
    Dim recognizedAvailabilityCount As Long
    Dim otherAvailabilityCount As Long
    Dim accountedCount As Long
    Dim finalSTOCount As Long

    Dim sourceKeySeen As Object

    If wsTarget Is Nothing Then Exit Sub

    sourcePath = Trim$(sourcePathOverride)

    If Len(sourcePath) = 0 Then
        sourcePath = STO_ResolveSourcePath()
    End If

    If Len(sourcePath) = 0 Then
        AddRunLog _
            "Warning", _
            "STO ALL Sync | Source STO belum ditemukan."
        Exit Sub
    End If

    sourceFileName = STO_FileNameFromPath(sourcePath)

    On Error Resume Next
    Set wbSTO = Workbooks(sourceFileName)
    On Error GoTo OpenFailed

    If wbSTO Is Nothing Then
        Set wbSTO = _
            Workbooks.Open( _
                fileName:=sourcePath, _
                UpdateLinks:=0, _
                ReadOnly:=True, _
                AddToMru:=False, _
                IgnoreReadOnlyRecommended:=True)
        openedByMacro = True
    End If

    On Error GoTo 0

    On Error Resume Next
    Set wsSTO = wbSTO.Worksheets("FPTK")
    On Error GoTo 0

    If wsSTO Is Nothing Then
        AddRunLog _
            "Warning", _
            "STO ALL Sync | Sheet FPTK tidak ditemukan."
        GoTo SafeExit
    End If

    sourceHeaderRow = STO_FindCompiledHeaderRow(wsSTO)

    If sourceHeaderRow = 0 Then
        AddRunLog _
            "Warning", _
            "STO ALL Sync | Header source tidak ditemukan."
        GoTo SafeExit
    End If

    Set sourceHeaderMap = _
        STO_BuildCompiledHeaderMap( _
            wsSTO, _
            sourceHeaderRow)

    Set targetHeaderMap = _
        STO_BuildCompiledHeaderMap( _
            wsTarget, _
            FPTK_HEADER_ROW)

    srcColPosisi = _
        FPTK_DirectFindHeader( _
            wsSTO, _
            sourceHeaderRow, _
            "Posisi")

    srcColRealDate = _
        FPTK_DirectFindHeader( _
            wsSTO, _
            sourceHeaderRow, _
            "FPTK Date (Real)")

    srcColAvailability = _
        STO_FindSourceAvailabilityColumn( _
            wsSTO, _
            sourceHeaderRow)

    srcColKodeUnik = _
        FPTK_DirectFindHeader( _
            wsSTO, _
            sourceHeaderRow, _
            "Kode Unik")

    srcColNama = _
        FPTK_DirectFindHeader( _
            wsSTO, _
            sourceHeaderRow, _
            "Nama Kandidat")

    tgtColKodeUnik = _
        FPTK_DirectFindHeader( _
            wsTarget, _
            FPTK_HEADER_ROW, _
            "Kode Unik")

    tgtColPosisi = _
        FPTK_DirectFindHeader( _
            wsTarget, _
            FPTK_HEADER_ROW, _
            "Posisi")

    tgtColNama = _
        FPTK_DirectFindHeader( _
            wsTarget, _
            FPTK_HEADER_ROW, _
            "Nama Kandidat")

    tgtColRealDate = _
        FPTK_DirectFindHeader( _
            wsTarget, _
            FPTK_HEADER_ROW, _
            "FPTK Date (Real)")

    tgtColFilter = _
        FPTK_DirectFindHeader( _
            wsTarget, _
            FPTK_HEADER_ROW, _
            "Filter Kategorisasi FPTK")

    tgtColAvailability = _
        STO_FindTargetHeaderAny( _
            wsTarget, _
            Array("FPTK Availability"))

    tgtColUpdated = _
        FPTK_DirectFindHeader( _
            wsTarget, _
            FPTK_HEADER_ROW, _
            "Last Updated At")

    tgtColAction = _
        FPTK_DirectFindHeader( _
            wsTarget, _
            FPTK_HEADER_ROW, _
            "Last Compile Action")

    tgtColSourceFile = _
        FPTK_DirectFindHeader( _
            wsTarget, _
            FPTK_HEADER_ROW, _
            "Source File")

    If srcColPosisi <= 0 Or _
       srcColRealDate <= 0 Or _
       tgtColPosisi <= 0 Or _
       tgtColRealDate <= 0 Or _
       tgtColFilter <= 0 Then

        AddRunLog _
            "Warning", _
            "STO ALL Sync | Header wajib: Posisi + FPTK Date (Real) + Filter."
        GoTo SafeExit
    End If

    ' ========================================================
    ' TARGET BUCKETS
    '
    ' BUKAN lagi key -> satu row.
    ' Sekarang key -> seluruh row target dengan key yang sama.
    ' Setiap source hanya boleh memakai satu target row yang
    ' belum dipakai pada run ini. Jadi duplicate key source
    ' tidak akan menumpuk ke row target yang sama.
    ' ========================================================
    Set targetBuckets = _
        STO_BuildTargetPositionRealDateBuckets( _
            wsTarget, _
            tgtColPosisi, _
            tgtColRealDate)

    Set usedTargetRows = _
        CreateObject("Scripting.Dictionary")
    usedTargetRows.CompareMode = vbTextCompare

    Set sourceKeySeen = _
        CreateObject("Scripting.Dictionary")
    sourceKeySeen.CompareMode = vbTextCompare

    lastSourceRow = _
        Application.Max( _
            wsSTO.Cells(wsSTO.rows.Count, srcColPosisi).End(xlUp).Row, _
            wsSTO.Cells(wsSTO.rows.Count, srcColRealDate).End(xlUp).Row)

    If srcColAvailability > 0 Then
        lastSourceRow = _
            Application.Max( _
                lastSourceRow, _
                wsSTO.Cells(wsSTO.rows.Count, srcColAvailability).End(xlUp).Row)
    End If

    If srcColKodeUnik > 0 Then
        lastSourceRow = _
            Application.Max( _
                lastSourceRow, _
                wsSTO.Cells(wsSTO.rows.Count, srcColKodeUnik).End(xlUp).Row)
    End If

    If srcColNama > 0 Then
        lastSourceRow = _
            Application.Max( _
                lastSourceRow, _
                wsSTO.Cells(wsSTO.rows.Count, srcColNama).End(xlUp).Row)
    End If

    lastTargetRow = _
        Application.Max( _
            wsTarget.Cells(wsTarget.rows.Count, tgtColPosisi).End(xlUp).Row, _
            wsTarget.Cells(wsTarget.rows.Count, tgtColRealDate).End(xlUp).Row)

    If tgtColKodeUnik > 0 Then
        lastTargetRow = _
            Application.Max( _
                lastTargetRow, _
                wsTarget.Cells(wsTarget.rows.Count, tgtColKodeUnik).End(xlUp).Row)
    End If

    If tgtColNama > 0 Then
        lastTargetRow = _
            Application.Max( _
                lastTargetRow, _
                wsTarget.Cells(wsTarget.rows.Count, tgtColNama).End(xlUp).Row)
    End If

    If tgtColSourceFile > 0 Then
        lastTargetRow = _
            Application.Max( _
                lastTargetRow, _
                wsTarget.Cells(wsTarget.rows.Count, tgtColSourceFile).End(xlUp).Row)
    End If

    If lastTargetRow < FPTK_FIRST_DATA_ROW Then
        lastTargetRow = FPTK_FIRST_DATA_ROW - 1
    End If

    For sourceRow = sourceHeaderRow + 1 To lastSourceRow

        posisiKey = _
            STO_NormalizePosition( _
                wsSTO.Cells( _
                    sourceRow, _
                    srcColPosisi).Value2)

        ' Definisi record STO = row source dengan Posisi terisi.
        ' Availability V/X bukan lagi syarat agar row diproses.
        If Len(posisiKey) = 0 Then
            skippedNoPositionCount = _
                skippedNoPositionCount + 1
            GoTo NextSourceRow
        End If

        sourceRecordCount = sourceRecordCount + 1

        availabilityText = vbNullString
        If srcColAvailability > 0 Then
            availabilityText = _
                STO_NormalizeAvailability( _
                    wsSTO.Cells( _
                        sourceRow, _
                        srcColAvailability).Value2)
        End If

        If availabilityText = "Y" Or _
           availabilityText = "N" Then
            recognizedAvailabilityCount = _
                recognizedAvailabilityCount + 1
        Else
            otherAvailabilityCount = _
                otherAvailabilityCount + 1
        End If

        dateKey = _
            NormalizeDateKey( _
                wsSTO.Cells( _
                    sourceRow, _
                    srcColRealDate).Value2)

        targetRow = 0

        If Len(dateKey) > 0 Then

            matchKey = _
                posisiKey & _
                "|" & _
                dateKey

            If sourceKeySeen.Exists(matchKey) Then
                duplicateKeyCount = duplicateKeyCount + 1
            Else
                sourceKeySeen.Add matchKey, True
            End If

            targetRow = _
                STO_TakeUnusedTargetRow( _
                    targetBuckets, _
                    matchKey, _
                    usedTargetRows)

            If targetRow > 0 Then

                wsTarget.Cells( _
                    targetRow, _
                    tgtColFilter).Value2 = _
                        "STO"

                usedTargetRows(CStr(targetRow)) = True
                matchedCount = matchedCount + 1
                GoTo NextSourceRow

            End If

        Else

            blankDateCount = blankDateCount + 1

            ' Tanggal kosong tidak punya key match yang aman.
            ' Marker hanya untuk mencegah insert ulang pada rerun.
            targetRow = _
                STO_FindExistingInsertedRowBySourceRowAny( _
                    wsTarget, _
                    tgtColSourceFile, _
                    tgtColAction, _
                    sourceFileName, _
                    sourceRow)

            If targetRow > 0 Then

                wsTarget.Cells( _
                    targetRow, _
                    tgtColFilter).Value2 = _
                        "STO"

                usedTargetRows(CStr(targetRow)) = True
                existingInsertedCount = _
                    existingInsertedCount + 1

                GoTo NextSourceRow

            End If

        End If

        ' ====================================================
        ' TIDAK ADA TARGET ROW YANG BELUM DIPAKAI -> INSERT.
        ' ====================================================
        lastTargetRow = lastTargetRow + 1
        targetRow = lastTargetRow

        STO_CopyCompiledRowByHeaders _
            wsSTO, _
            sourceRow, _
            sourceHeaderMap, _
            wsTarget, _
            targetRow, _
            targetHeaderMap

        wsTarget.Cells( _
            targetRow, _
            tgtColPosisi).Value = _
                wsSTO.Cells( _
                    sourceRow, _
                    srcColPosisi).Value

        wsTarget.Cells( _
            targetRow, _
            tgtColRealDate).Value = _
                wsSTO.Cells( _
                    sourceRow, _
                    srcColRealDate).Value

        wsTarget.Cells( _
            targetRow, _
            tgtColFilter).Value2 = _
                "STO"

        ' Availability hanya ditulis jika source dikenali sebagai Y/N.
        If tgtColAvailability > 0 Then
            If availabilityText = "Y" Or _
               availabilityText = "N" Then
                wsTarget.Cells( _
                    targetRow, _
                    tgtColAvailability).Value2 = _
                        availabilityText
            End If
        End If

        If tgtColUpdated > 0 Then
            wsTarget.Cells( _
                targetRow, _
                tgtColUpdated).Value = Now
        End If

        If tgtColSourceFile > 0 Then
            wsTarget.Cells( _
                targetRow, _
                tgtColSourceFile).Value2 = _
                    sourceFileName
        End If

        If tgtColAction > 0 Then
            wsTarget.Cells( _
                targetRow, _
                tgtColAction).Value2 = _
                    "INSERT STO|SRCROW=" & _
                    CStr(sourceRow)
        End If

        usedTargetRows(CStr(targetRow)) = True
        insertedCount = insertedCount + 1

NextSourceRow:

    Next sourceRow

    accountedCount = _
        matchedCount + _
        existingInsertedCount + _
        insertedCount

    finalSTOCount = _
        STO_CountFilterValue( _
            wsTarget, _
            tgtColFilter, _
            "STO")

    AddRunLog _
        "Info", _
        "STO ALL Sync 1-to-1 | Source Records: " & _
        CStr(sourceRecordCount) & _
        " | Accounted: " & CStr(accountedCount) & _
        " | Matched Existing: " & CStr(matchedCount) & _
        " | Existing Insert: " & CStr(existingInsertedCount) & _
        " | Inserted: " & CStr(insertedCount) & _
        " | Duplicate Posisi+Date Source: " & CStr(duplicateKeyCount) & _
        " | Blank Date: " & CStr(blankDateCount) & _
        " | Availability Y/N: " & CStr(recognizedAvailabilityCount) & _
        " | Availability Other/Blank: " & CStr(otherAvailabilityCount) & _
        " | Skipped No Position: " & CStr(skippedNoPositionCount) & _
        " | Final Filter STO: " & CStr(finalSTOCount)

    If Not silentMode Then
        MsgBox _
            "STO sync selesai." & vbCrLf & vbCrLf & _
            "Source records: " & CStr(sourceRecordCount) & vbCrLf & _
            "Accounted: " & CStr(accountedCount) & vbCrLf & _
            "Matched existing: " & CStr(matchedCount) & vbCrLf & _
            "Inserted: " & CStr(insertedCount) & vbCrLf & _
            "Duplicate Posisi + Date: " & CStr(duplicateKeyCount) & vbCrLf & _
            "Blank Date: " & CStr(blankDateCount) & vbCrLf & _
            "Final Filter STO: " & CStr(finalSTOCount), _
            vbInformation, _
            "STO Sync 1-to-1"
    End If

SafeExit:

    On Error Resume Next

    If openedByMacro Then
        wbSTO.Close SaveChanges:=False
    End If

    Set sourceKeySeen = Nothing
    Set usedTargetRows = Nothing
    Set targetBuckets = Nothing
    Set targetHeaderMap = Nothing
    Set sourceHeaderMap = Nothing
    Set wsSTO = Nothing
    Set wbSTO = Nothing

    On Error GoTo 0
    Exit Sub

OpenFailed:

    AddRunLog _
        "Warning", _
        "STO ALL Sync | Gagal membuka source: " & _
        sourcePath & _
        " | " & _
        Err.Description

    Err.Clear
    Resume SafeExit

End Sub

Private Function STO_BuildTargetPositionRealDateBuckets( _
    ByVal ws As Worksheet, _
    ByVal colPosisi As Long, _
    ByVal colRealDate As Long) As Object

    Dim resultMap As Object
    Dim rowBucket As Collection
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim posisiKey As String
    Dim dateKey As String
    Dim matchKey As String

    Set resultMap = _
        CreateObject("Scripting.Dictionary")
    resultMap.CompareMode = vbTextCompare

    If ws Is Nothing Or _
       colPosisi <= 0 Or _
       colRealDate <= 0 Then

        Set STO_BuildTargetPositionRealDateBuckets = resultMap
        Exit Function
    End If

    lastRow = _
        Application.Max( _
            ws.Cells(ws.rows.Count, colPosisi).End(xlUp).Row, _
            ws.Cells(ws.rows.Count, colRealDate).End(xlUp).Row)

    For rowIndex = FPTK_FIRST_DATA_ROW To lastRow

        posisiKey = _
            STO_NormalizePosition( _
                ws.Cells( _
                    rowIndex, _
                    colPosisi).Value2)

        dateKey = _
            NormalizeDateKey( _
                ws.Cells( _
                    rowIndex, _
                    colRealDate).Value2)

        If Len(posisiKey) > 0 And _
           Len(dateKey) > 0 Then

            matchKey = _
                posisiKey & _
                "|" & _
                dateKey

            If Not resultMap.Exists(matchKey) Then
                Set rowBucket = New Collection
                resultMap.Add matchKey, rowBucket
            Else
                Set rowBucket = resultMap(matchKey)
            End If

            rowBucket.Add rowIndex

        End If

    Next rowIndex

    Set STO_BuildTargetPositionRealDateBuckets = resultMap

End Function

Private Function STO_TakeUnusedTargetRow( _
    ByVal targetBuckets As Object, _
    ByVal matchKey As String, _
    ByVal usedTargetRows As Object) As Long

    Dim rowBucket As Collection
    Dim item As Variant
    Dim rowKey As String

    If targetBuckets Is Nothing Then Exit Function
    If usedTargetRows Is Nothing Then Exit Function
    If Len(matchKey) = 0 Then Exit Function

    If Not targetBuckets.Exists(matchKey) Then
        Exit Function
    End If

    Set rowBucket = targetBuckets(matchKey)

    For Each item In rowBucket

        rowKey = CStr(CLng(item))

        If Not usedTargetRows.Exists(rowKey) Then
            STO_TakeUnusedTargetRow = CLng(item)
            Exit Function
        End If

    Next item

End Function

Private Function STO_FindExistingInsertedRowBySourceRowAny( _
    ByVal wsTarget As Worksheet, _
    ByVal colSourceFile As Long, _
    ByVal colAction As Long, _
    ByVal sourceFileName As String, _
    ByVal sourceRow As Long) As Long

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim actionText As String
    Dim sourceFileKey As String
    Dim wantedMarker As String

    If wsTarget Is Nothing Then Exit Function
    If colAction <= 0 Then Exit Function

    wantedMarker = _
        "|SRCROW=" & _
        CStr(sourceRow)

    lastRow = _
        wsTarget.Cells( _
            wsTarget.rows.Count, _
            colAction).End(xlUp).Row

    For rowIndex = FPTK_FIRST_DATA_ROW To lastRow

        actionText = _
            UCase$( _
                Trim$( _
                    CStr( _
                        wsTarget.Cells( _
                            rowIndex, _
                            colAction).Value2)))

        If Left$(actionText, 10) = "INSERT STO" And _
           InStr(1, actionText, wantedMarker, vbTextCompare) > 0 Then

            If colSourceFile > 0 And _
               Len(sourceFileName) > 0 Then

                sourceFileKey = _
                    UCase$( _
                        STO_FileNameFromPath( _
                            CStr( _
                                wsTarget.Cells( _
                                    rowIndex, _
                                    colSourceFile).Value2)))

                If StrComp( _
                    sourceFileKey, _
                    UCase$(sourceFileName), _
                    vbTextCompare) <> 0 Then

                    GoTo NextRow
                End If

            End If

            STO_FindExistingInsertedRowBySourceRowAny = rowIndex
            Exit Function

        End If

NextRow:
    Next rowIndex

End Function

Private Function STO_CountFilterValue( _
    ByVal ws As Worksheet, _
    ByVal colFilter As Long, _
    ByVal wantedValue As String) As Long

    Dim lastRow As Long
    Dim rowIndex As Long

    If ws Is Nothing Then Exit Function
    If colFilter <= 0 Then Exit Function

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            colFilter).End(xlUp).Row

    For rowIndex = FPTK_FIRST_DATA_ROW To lastRow

        If StrComp( _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        rowIndex, _
                        colFilter).Value2)), _
            wantedValue, _
            vbTextCompare) = 0 Then

            STO_CountFilterValue = _
                STO_CountFilterValue + 1
        End If

    Next rowIndex

End Function

Private Sub STO_SyncXAssignments( _
    ByVal wsTarget As Worksheet, _
    ByVal headerMap As Object, _
    Optional ByVal silentMode As Boolean = True, _
    Optional ByVal sourcePathOverride As String = vbNullString)

    ' Backward compatibility only.
    ' X tidak diproses terpisah lagi karena seluruh source STO
    ' harus diproses satu kali secara 1-to-1.

End Sub

Private Function STO_FindTargetHeaderAny( _
    ByVal ws As Worksheet, _
    ByVal headerNames As Variant) As Long

    Dim item As Variant
    Dim resultColumn As Long

    If ws Is Nothing Then Exit Function

    For Each item In headerNames

        resultColumn = _
            FPTK_DirectFindHeader( _
                ws, _
                FPTK_HEADER_ROW, _
                CStr(item))

        If resultColumn > 0 Then

            STO_FindTargetHeaderAny = _
                resultColumn

            Exit Function

        End If

    Next item

End Function

Private Function STO_BuildExistingXTargetMap( _
    ByVal ws As Worksheet, _
    ByVal colRemark As Long) As Object

    Dim resultMap As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim remarkText As String
    Dim markerText As String
    Dim p As Long

    Set resultMap = _
        CreateObject( _
            "Scripting.Dictionary")

    resultMap.CompareMode = _
        vbTextCompare

    If ws Is Nothing Or _
       colRemark <= 0 Then

        Set STO_BuildExistingXTargetMap = _
            resultMap

        Exit Function

    End If

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            colRemark).End( _
                xlUp).Row

    For rowIndex = FPTK_FIRST_DATA_ROW _
        To lastRow

        remarkText = _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        rowIndex, _
                        colRemark).Value2))

        If UCase$(Left$( _
            remarkText, _
            9)) = _
           "STO_X|NO=" Then

            p = _
                InStr( _
                    1, _
                    remarkText, _
                    "|FPTK ASSIGN=", _
                    vbTextCompare)

            If p > 0 Then

                markerText = _
                    Left$( _
                        remarkText, _
                        p - 1)

            Else

                markerText = _
                    remarkText

            End If

            If Not resultMap.Exists( _
                UCase$(markerText)) Then

                resultMap.Add _
                    UCase$(markerText), _
                    rowIndex

            End If

        End If

    Next rowIndex

    Set STO_BuildExistingXTargetMap = _
        resultMap

End Function

Private Function STO_IsManagedXTargetRow( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal colRemark As Long) As Boolean

    Dim remarkText As String

    If ws Is Nothing Then Exit Function
    If rowIndex <= 0 Then Exit Function
    If colRemark <= 0 Then Exit Function

    remarkText = _
        UCase$( _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        rowIndex, _
                        colRemark).Value2)))

    STO_IsManagedXTargetRow = _
        (Left$( _
            remarkText, _
            9) = _
         "STO_X|NO=")

End Function

Private Function STO_SourcePositionText( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal colPosisi As Long, _
    ByVal colExcel As Long, _
    ByVal colArea As Long) As String

    Dim resultText As String
    Dim leftText As String
    Dim areaText As String

    If ws Is Nothing Then Exit Function

    If colPosisi > 0 Then

        resultText = _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        rowIndex, _
                        colPosisi).Value2))

    End If

    ' Kalau cached formula Posisi kosong atau malah formula text,
    ' bangun dari kolom excel + Area.
    If Len(resultText) = 0 Or _
       Left$(resultText, 1) = "=" Then

        If colExcel > 0 Then

            leftText = _
                Trim$( _
                    CStr( _
                        ws.Cells( _
                            rowIndex, _
                            colExcel).Value2))

        End If

        If colArea > 0 Then

            areaText = _
                Trim$( _
                    CStr( _
                        ws.Cells( _
                            rowIndex, _
                            colArea).Value2))

        End If

        resultText = _
            leftText & _
            areaText

    End If

    STO_SourcePositionText = _
        Trim$(resultText)

End Function

Private Function STO_SourceCellText( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long) As String

    If ws Is Nothing Then Exit Function
    If columnIndex <= 0 Then Exit Function

    STO_SourceCellText = _
        Trim$( _
            CStr( _
                ws.Cells( _
                    rowIndex, _
                    columnIndex).Value2))

End Function

Private Sub STO_CopySourceIfPresent( _
    ByVal wsSource As Worksheet, _
    ByVal sourceRow As Long, _
    ByVal sourceColumn As Long, _
    ByVal wsTarget As Worksheet, _
    ByVal targetRow As Long, _
    ByVal targetColumn As Long)

    Dim sourceValue As Variant

    If wsSource Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub
    If sourceColumn <= 0 Then Exit Sub
    If targetColumn <= 0 Then Exit Sub

    sourceValue = _
        wsSource.Cells( _
            sourceRow, _
            sourceColumn).Value

    If ValueIsBlankFPTK( _
        sourceValue) Then

        Exit Sub

    End If

    wsTarget.Cells( _
        targetRow, _
        targetColumn).Value = _
            sourceValue

End Sub

Private Function STO_LastSourceDataRowByColumns( _
    ByVal ws As Worksheet, _
    ByVal columnsList As Variant) As Long

    Dim item As Variant
    Dim currentRow As Long
    Dim bestRow As Long

    If ws Is Nothing Then Exit Function

    For Each item In columnsList

        If CLng(item) > 0 Then

            currentRow = _
                ws.Cells( _
                    ws.rows.Count, _
                    CLng(item)).End( _
                        xlUp).Row

            If currentRow > bestRow Then
                bestRow = currentRow
            End If

        End If

    Next item

    STO_LastSourceDataRowByColumns = _
        bestRow

End Function

Private Function STO_LookupKodePIC( _
    ByVal recruiterName As String) As String

    Dim ws As Worksheet
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim colRecruiter As Long
    Dim colKodePIC As Long
    Dim columnIndex As Long
    Dim rowIndex As Long

    Dim headerText As String
    Dim recruiterKey As String
    Dim rowRecruiterKey As String

    recruiterKey = _
        STO_NormalizeText( _
            recruiterName)

    If Len(recruiterKey) = 0 Then
        Exit Function
    End If

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            "Kode Dropdown")

    If ws Is Nothing Then

        Set ws = _
            ThisWorkbook.Worksheets( _
                "Kode Dropdown (tidak diubah)")

    End If

    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    lastColumn = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End( _
                xlToLeft).Column

    For columnIndex = 1 To lastColumn

        headerText = _
            STO_NormalizeText( _
                ws.Cells( _
                    1, _
                    columnIndex).Value2)

        If headerText = _
           "PIC RECRUITER" Then

            colRecruiter = _
                columnIndex

        ElseIf headerText = _
               "KODE PIC" Then

            colKodePIC = _
                columnIndex

        End If

    Next columnIndex

    If colRecruiter = 0 Or _
       colKodePIC = 0 Then

        Exit Function

    End If

    lastRow = _
        Application.Max( _
            ws.Cells( _
                ws.rows.Count, _
                colRecruiter).End( _
                    xlUp).Row, _
            ws.Cells( _
                ws.rows.Count, _
                colKodePIC).End( _
                    xlUp).Row)

    For rowIndex = 2 To lastRow

        rowRecruiterKey = _
            STO_NormalizeText( _
                ws.Cells( _
                    rowIndex, _
                    colRecruiter).Value2)

        If rowRecruiterKey = _
           recruiterKey Then

            STO_LookupKodePIC = _
                Trim$( _
                    CStr( _
                        ws.Cells( _
                            rowIndex, _
                            colKodePIC).Value2))

            Exit Function

        End If

    Next rowIndex

End Function

Private Sub STO_ApplyExistingFPTKTemplate( _
    ByVal ws As Worksheet, _
    ByVal targetRow As Long, _
    ByVal posisiText As String, _
    ByVal colPosisi As Long, _
    ByVal colKodeAngka As Long, _
    ByVal colBU As Long, _
    ByVal colDivisi As Long, _
    ByVal colDepartment As Long, _
    ByVal colLevel As Long, _
    ByVal colLevelNumber As Long, _
    ByVal colRemark As Long)

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim templateRow As Long
    Dim genericTemplateRow As Long

    Dim targetPositionKey As String
    Dim rowPositionKey As String

    If ws Is Nothing Then Exit Sub
    If targetRow <= 0 Then Exit Sub
    If colPosisi <= 0 Then Exit Sub

    targetPositionKey = _
        STO_NormalizePosition( _
            posisiText)

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            colPosisi).End( _
                xlUp).Row

    For rowIndex = FPTK_FIRST_DATA_ROW _
        To lastRow

        If rowIndex <> targetRow Then

            If Not STO_IsManagedXTargetRow( _
                ws, _
                rowIndex, _
                colRemark) Then

                rowPositionKey = _
                    STO_NormalizePosition( _
                        ws.Cells( _
                            rowIndex, _
                            colPosisi).Value2)

                If rowPositionKey = _
                   targetPositionKey Then

                    templateRow = _
                        rowIndex

                    Exit For

                ElseIf genericTemplateRow = 0 Then

                    If Left$( _
                        rowPositionKey, _
                        Len( _
                            "SALES TAKING ORDER STAFF (CHILLED)")) = _
                       "SALES TAKING ORDER STAFF (CHILLED)" Then

                        genericTemplateRow = _
                            rowIndex

                    End If

                End If

            End If

        End If

    Next rowIndex

    If templateRow = 0 Then
        templateRow = genericTemplateRow
    End If

    If templateRow = 0 Then Exit Sub

    STO_CopyTargetTemplateValue _
        ws, _
        templateRow, _
        targetRow, _
        colKodeAngka

    STO_CopyTargetTemplateValue _
        ws, _
        templateRow, _
        targetRow, _
        colBU

    STO_CopyTargetTemplateValue _
        ws, _
        templateRow, _
        targetRow, _
        colDivisi

    STO_CopyTargetTemplateValue _
        ws, _
        templateRow, _
        targetRow, _
        colDepartment

    STO_CopyTargetTemplateValue _
        ws, _
        templateRow, _
        targetRow, _
        colLevel

    STO_CopyTargetTemplateValue _
        ws, _
        templateRow, _
        targetRow, _
        colLevelNumber

End Sub

Private Sub STO_CopyTargetTemplateValue( _
    ByVal ws As Worksheet, _
    ByVal sourceRow As Long, _
    ByVal targetRow As Long, _
    ByVal columnIndex As Long)

    Dim sourceValue As Variant

    If ws Is Nothing Then Exit Sub
    If columnIndex <= 0 Then Exit Sub

    sourceValue = _
        ws.Cells( _
            sourceRow, _
            columnIndex).Value

    If ValueIsBlankFPTK( _
        sourceValue) Then

        Exit Sub

    End If

    ws.Cells( _
        targetRow, _
        columnIndex).Value = _
            sourceValue

End Sub
Private Sub RefreshFPTKSTOCategorization( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object, _
    Optional ByVal silentMode As Boolean = True, _
    Optional ByVal sourcePathOverride As String = vbNullString)

    ' Satu sumber kebenaran untuk STO:
    ' - setiap row source dengan Posisi terisi harus accounted;
    ' - match 1-to-1 berdasarkan Posisi + FPTK Date (Real);
    ' - jika target occurrence tidak tersedia -> INSERT;
    ' - existing row hanya mengubah Filter Kategorisasi FPTK.
    STO_SyncAllAssignmentsByPositionRealDate _
        ws, _
        silentMode, _
        sourcePathOverride

End Sub

Private Function STO_LoadPositionRealDateMap( _
    ByVal sourcePath As String) As Object

    Dim resultMap As Object
    Dim wbSTO As Workbook
    Dim wsSTO As Worksheet
    Dim openedByMacro As Boolean

    Dim headerRow As Long
    Dim colPosisi As Long
    Dim colRealDate As Long
    Dim colAvailability As Long

    Dim lastRow As Long
    Dim rowIndex As Long

    Dim availabilityText As String
    Dim posisiKey As String
    Dim dateKey As String
    Dim mapKey As String

    Set resultMap = _
        CreateObject( _
            "Scripting.Dictionary")

    resultMap.CompareMode = vbTextCompare

    If Len(Trim$(sourcePath)) = 0 Then
        Set STO_LoadPositionRealDateMap = resultMap
        Exit Function
    End If

    On Error Resume Next
    Set wbSTO = _
        Workbooks( _
            STO_FileNameFromPath( _
                sourcePath))
    On Error GoTo OpenFailed

    If wbSTO Is Nothing Then
        Set wbSTO = _
            Workbooks.Open( _
                fileName:=sourcePath, _
                UpdateLinks:=0, _
                ReadOnly:=True, _
                AddToMru:=False, _
                IgnoreReadOnlyRecommended:=True)
        openedByMacro = True
    End If

    On Error GoTo 0

    On Error Resume Next
    Set wsSTO = wbSTO.Worksheets("FPTK")
    On Error GoTo 0

    If wsSTO Is Nothing Then
        GoTo SafeExit
    End If

    headerRow = STO_FindCompiledHeaderRow(wsSTO)
    If headerRow = 0 Then GoTo SafeExit

    colPosisi = _
        FPTK_DirectFindHeader( _
            wsSTO, _
            headerRow, _
            "Posisi")

    colRealDate = _
        FPTK_DirectFindHeader( _
            wsSTO, _
            headerRow, _
            "FPTK Date (Real)")

    colAvailability = _
        STO_FindSourceAvailabilityColumn( _
            wsSTO, _
            headerRow)

    If colPosisi <= 0 Or _
       colRealDate <= 0 Or _
       colAvailability <= 0 Then
        GoTo SafeExit
    End If

    lastRow = _
        Application.Max( _
            wsSTO.Cells(wsSTO.rows.Count, colPosisi).End(xlUp).Row, _
            wsSTO.Cells(wsSTO.rows.Count, colRealDate).End(xlUp).Row, _
            wsSTO.Cells(wsSTO.rows.Count, colAvailability).End(xlUp).Row)

    For rowIndex = headerRow + 1 To lastRow

        availabilityText = _
            STO_NormalizeAvailability( _
                wsSTO.Cells( _
                    rowIndex, _
                    colAvailability).Value2)

        ' Y dan N sama-sama merupakan record STO.
        If availabilityText = "Y" Or _
           availabilityText = "N" Then

            posisiKey = _
                STO_NormalizePosition( _
                    wsSTO.Cells( _
                        rowIndex, _
                        colPosisi).Value2)

            dateKey = _
                NormalizeDateKey( _
                    wsSTO.Cells( _
                        rowIndex, _
                        colRealDate).Value2)

            If Len(posisiKey) > 0 And _
               Len(dateKey) > 0 Then

                mapKey = _
                    posisiKey & _
                    "|" & _
                    dateKey

                If Not resultMap.Exists(mapKey) Then
                    resultMap.Add mapKey, True
                End If

            End If

        End If

    Next rowIndex

SafeExit:

    On Error Resume Next
    If openedByMacro Then
        wbSTO.Close SaveChanges:=False
    End If
    Set wsSTO = Nothing
    Set wbSTO = Nothing
    On Error GoTo 0

    Set STO_LoadPositionRealDateMap = resultMap
    Exit Function

OpenFailed:

    AddRunLog _
        "Warning", _
        "STO Map Posisi+RealDate | Gagal membuka source: " & _
        sourcePath & _
        " | " & _
        Err.Description
    Err.Clear
    Resume SafeExit

End Function

Private Function STO_LoadCandidateMap( _
    ByVal sourcePath As String) As Object

    Dim resultMap As Object

    Dim wbSTO As Workbook
    Dim wsSTO As Worksheet
    Dim openedByMacro As Boolean

    Dim headerRow As Long
    Dim sourceHeaderMap As Object

    Dim colPosisi As Long
    Dim colNama As Long
    Dim colAvailability As Long

    Dim lastRow As Long
    Dim rowIndex As Long

    Dim availabilityText As String
    Dim posisiKey As String
    Dim namaKey As String
    Dim mapKey As String

    Set resultMap = _
        CreateObject( _
            "Scripting.Dictionary")

    resultMap.CompareMode = _
        vbTextCompare

    If Len(sourcePath) = 0 Then

        Set STO_LoadCandidateMap = _
            resultMap

        Exit Function

    End If

    On Error Resume Next

    Set wbSTO = _
        Workbooks( _
            STO_FileNameFromPath( _
                sourcePath))

    On Error GoTo 0

    If wbSTO Is Nothing Then

        On Error GoTo OpenFailed

        Set wbSTO = _
            Workbooks.Open( _
                fileName:=sourcePath, _
                UpdateLinks:=0, _
                ReadOnly:=True, _
                AddToMru:=False, _
                IgnoreReadOnlyRecommended:=True)

        openedByMacro = True

        On Error GoTo 0

    End If

    On Error Resume Next

    Set wsSTO = _
        wbSTO.Worksheets( _
            "FPTK")

    On Error GoTo 0

    If wsSTO Is Nothing Then
        GoTo SafeExit
    End If

    headerRow = _
        STO_FindCompiledHeaderRow( _
            wsSTO)

    If headerRow = 0 Then
        GoTo SafeExit
    End If

    Set sourceHeaderMap = _
        STO_BuildCompiledHeaderMap( _
            wsSTO, _
            headerRow)

    colPosisi = _
        STO_HeaderMapColumn( _
            sourceHeaderMap, _
            "POSISI")

    colNama = _
        STO_HeaderMapColumn( _
            sourceHeaderMap, _
            "NAMA KANDIDAT")

    colAvailability = _
        STO_HeaderMapColumn( _
            sourceHeaderMap, _
            "FPTK AVAILABILITY")

    If colPosisi = 0 Or _
       colAvailability = 0 Then

        GoTo SafeExit

    End If

    ' Tanpa Nama Kandidat, pair-map V tidak bisa dibentuk.
    ' Bukan berarti source STO invalid; sync utama tetap berjalan.
    If colNama = 0 Then
        GoTo SafeExit
    End If

    lastRow = _
        Application.Max( _
            wsSTO.Cells( _
                wsSTO.rows.Count, _
                colPosisi).End( _
                    xlUp).Row, _
            wsSTO.Cells( _
                wsSTO.rows.Count, _
                colNama).End( _
                    xlUp).Row, _
            wsSTO.Cells( _
                wsSTO.rows.Count, _
                colAvailability).End( _
                    xlUp).Row)

    For rowIndex = _
        headerRow + 1 _
        To lastRow

        availabilityText = _
            STO_NormalizeAvailability( _
                wsSTO.Cells( _
                    rowIndex, _
                    colAvailability).Value2)

        ' V / Ya / Y = available.
        If availabilityText = "Y" Then

            posisiKey = _
                STO_NormalizePosition( _
                    wsSTO.Cells( _
                        rowIndex, _
                        colPosisi).Value2)

            namaKey = _
                STO_NormalizeCandidate( _
                    wsSTO.Cells( _
                        rowIndex, _
                        colNama).Value2)

            If Len(posisiKey) > 0 And _
               Len(namaKey) > 0 Then

                mapKey = _
                    posisiKey & _
                    "|" & _
                    namaKey

                If Not resultMap.Exists( _
                    mapKey) Then

                    resultMap.Add _
                        mapKey, _
                        True

                End If

            End If

        End If

    Next rowIndex

SafeExit:

    On Error Resume Next

    If openedByMacro Then

        wbSTO.Close _
            SaveChanges:=False

    End If

    Set sourceHeaderMap = Nothing
    Set wsSTO = Nothing
    Set wbSTO = Nothing

    On Error GoTo 0

    Set STO_LoadCandidateMap = _
        resultMap

    Exit Function

OpenFailed:

    AddRunLog _
        "Warning", _
        "STO Categorization | Gagal membuka STO source: " & _
        sourcePath & _
        " | " & _
        Err.Description

    Err.Clear

    Resume SafeExit

End Function

Private Function STO_FindCompiledHeaderRow( _
    ByVal ws As Worksheet) As Long

    Dim rowIndex As Long
    Dim kodeCol As Long
    Dim posisiCol As Long
    Dim availCol As Long

    If ws Is Nothing Then Exit Function

    For rowIndex = 1 To 10

        kodeCol = _
            FPTK_DirectFindHeader( _
                ws, _
                rowIndex, _
                "Kode Unik")

        posisiCol = _
            FPTK_DirectFindHeader( _
                ws, _
                rowIndex, _
                "Posisi")

        availCol = _
            STO_FindSourceAvailabilityColumn( _
                ws, _
                rowIndex)

        If posisiCol > 0 And _
           availCol > 0 Then

            STO_FindCompiledHeaderRow = _
                rowIndex

            Exit Function

        End If

    Next rowIndex

End Function

Private Function STO_BuildCompiledHeaderMap( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long) As Object

    Dim resultMap As Object
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim headerText As String

    Set resultMap = _
        CreateObject( _
            "Scripting.Dictionary")

    resultMap.CompareMode = _
        vbTextCompare

    If ws Is Nothing Or _
       headerRow <= 0 Then

        Set STO_BuildCompiledHeaderMap = _
            resultMap

        Exit Function

    End If

    lastColumn = _
        ws.Cells( _
            headerRow, _
            ws.Columns.Count).End( _
                xlToLeft).Column

    For columnIndex = 1 To lastColumn

        headerText = _
            STO_NormalizeText( _
                ws.Cells( _
                    headerRow, _
                    columnIndex).Value2)

        Select Case headerText

            Case "STATUS FPTK", _
                 "FPTK AVAILABILITY", _
                 "AVAILABILITY FPTK", _
                 "FPTK AVAILABLE", _
                 "AVAILABLE FPTK"

                headerText = "FPTK AVAILABILITY"

        End Select

        If Len(headerText) > 0 Then

            If Not resultMap.Exists( _
                headerText) Then

                resultMap.Add _
                    headerText, _
                    columnIndex

            End If

        End If

    Next columnIndex

    Set STO_BuildCompiledHeaderMap = _
        resultMap

End Function

Private Function STO_HeaderMapColumn( _
    ByVal headerMap As Object, _
    ByVal headerName As String) As Long

    Dim keyText As String

    If headerMap Is Nothing Then Exit Function

    keyText = _
        STO_NormalizeText( _
            headerName)

    If headerMap.Exists( _
        keyText) Then

        STO_HeaderMapColumn = _
            CLng( _
                headerMap( _
                    keyText))

    End If

End Function
Private Function STO_FindUnusedBlankCandidateTargetRow( _
    ByVal wsTarget As Worksheet, _
    ByVal colPosisi As Long, _
    ByVal colNama As Long, _
    ByVal colAction As Long, _
    ByVal wantedPositionKey As String, _
    ByVal usedRows As Object) As Long

    Dim lastRow As Long
    Dim rowIndex As Long

    Dim posisiKey As String
    Dim namaKey As String
    Dim actionText As String
    Dim rowKey As String

    If wsTarget Is Nothing Then Exit Function
    If colPosisi <= 0 Or colNama <= 0 Then Exit Function
    If Len(wantedPositionKey) = 0 Then Exit Function

    lastRow = _
        Application.Max( _
            wsTarget.Cells(wsTarget.rows.Count, colPosisi).End(xlUp).Row, _
            wsTarget.Cells(wsTarget.rows.Count, colNama).End(xlUp).Row)

    For rowIndex = FPTK_FIRST_DATA_ROW To lastRow

        rowKey = CStr(rowIndex)

        If Not usedRows Is Nothing Then

            If usedRows.Exists(rowKey) Then
                GoTo NextRow
            End If

        End If

        posisiKey = _
            STO_NormalizePosition( _
                wsTarget.Cells( _
                    rowIndex, _
                    colPosisi).Value2)

        If StrComp( _
            posisiKey, _
            wantedPositionKey, _
            vbTextCompare) <> 0 Then

            GoTo NextRow

        End If

        namaKey = _
            STO_NormalizeCandidate( _
                wsTarget.Cells( _
                    rowIndex, _
                    colNama).Value2)

        If Len(namaKey) > 0 Then
            GoTo NextRow
        End If

        If colAction > 0 Then

            actionText = _
                UCase$( _
                    Trim$( _
                        CStr( _
                            wsTarget.Cells( _
                                rowIndex, _
                                colAction).Value2)))

            If Left$(actionText, 10) = "INSERT STO" Then
                GoTo NextRow
            End If

        End If

        STO_FindUnusedBlankCandidateTargetRow = rowIndex
        Exit Function

NextRow:

    Next rowIndex

End Function

Private Function STO_BuildTargetPositionRealDateMap( _
    ByVal ws As Worksheet, _
    ByVal colPosisi As Long, _
    ByVal colRealDate As Long) As Object

    Dim resultMap As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim posisiKey As String
    Dim dateKey As String
    Dim matchKey As String

    Set resultMap = _
        CreateObject( _
            "Scripting.Dictionary")

    resultMap.CompareMode = vbTextCompare

    If ws Is Nothing Or _
       colPosisi <= 0 Or _
       colRealDate <= 0 Then

        Set STO_BuildTargetPositionRealDateMap = resultMap
        Exit Function
    End If

    lastRow = _
        Application.Max( _
            ws.Cells(ws.rows.Count, colPosisi).End(xlUp).Row, _
            ws.Cells(ws.rows.Count, colRealDate).End(xlUp).Row)

    For rowIndex = FPTK_FIRST_DATA_ROW To lastRow

        posisiKey = _
            STO_NormalizePosition( _
                ws.Cells( _
                    rowIndex, _
                    colPosisi).Value2)

        dateKey = _
            NormalizeDateKey( _
                ws.Cells( _
                    rowIndex, _
                    colRealDate).Value2)

        If Len(posisiKey) > 0 And _
           Len(dateKey) > 0 Then

            matchKey = _
                posisiKey & _
                "|" & _
                dateKey

            If Not resultMap.Exists(matchKey) Then
                resultMap.Add _
                    matchKey, _
                    rowIndex
            End If

        End If

    Next rowIndex

    Set STO_BuildTargetPositionRealDateMap = resultMap

End Function

Private Function STO_BuildTargetPositionCandidateMap( _
    ByVal ws As Worksheet, _
    ByVal colPosisi As Long, _
    ByVal colNama As Long) As Object

    Dim resultMap As Object
    Dim lastRow As Long
    Dim rowIndex As Long

    Dim posisiKey As String
    Dim namaKey As String
    Dim pairKey As String

    Set resultMap = _
        CreateObject( _
            "Scripting.Dictionary")

    resultMap.CompareMode = _
        vbTextCompare

    If ws Is Nothing Or _
       colPosisi <= 0 Or _
       colNama <= 0 Then

        Set STO_BuildTargetPositionCandidateMap = _
            resultMap

        Exit Function

    End If

    lastRow = _
        Application.Max( _
            ws.Cells( _
                ws.rows.Count, _
                colPosisi).End( _
                    xlUp).Row, _
            ws.Cells( _
                ws.rows.Count, _
                colNama).End( _
                    xlUp).Row)

    For rowIndex = _
        FPTK_FIRST_DATA_ROW _
        To lastRow

        posisiKey = _
            STO_NormalizePosition( _
                ws.Cells( _
                    rowIndex, _
                    colPosisi).Value2)

        namaKey = _
            STO_NormalizeCandidate( _
                ws.Cells( _
                    rowIndex, _
                    colNama).Value2)

        If Len(posisiKey) > 0 And _
           Len(namaKey) > 0 Then

            pairKey = _
                posisiKey & _
                "|" & _
                namaKey

            If Not resultMap.Exists( _
                pairKey) Then

                resultMap.Add _
                    pairKey, _
                    rowIndex

            End If

        End If

    Next rowIndex

    Set STO_BuildTargetPositionCandidateMap = _
        resultMap

End Function


Private Function STO_BuildTargetKodeUnikMap( _
    ByVal ws As Worksheet, _
    ByVal colKodeUnik As Long) As Object

    Dim resultMap As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim kodeKey As String

    Set resultMap = _
        CreateObject( _
            "Scripting.Dictionary")

    resultMap.CompareMode = _
        vbTextCompare

    If ws Is Nothing Or _
       colKodeUnik <= 0 Then

        Set STO_BuildTargetKodeUnikMap = _
            resultMap

        Exit Function

    End If

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            colKodeUnik).End( _
                xlUp).Row

    For rowIndex = _
        FPTK_FIRST_DATA_ROW _
        To lastRow

        kodeKey = _
            STO_NormalizeText( _
                ws.Cells( _
                    rowIndex, _
                    colKodeUnik).Value2)

        If Len(kodeKey) > 0 Then

            If Not resultMap.Exists( _
                kodeKey) Then

                resultMap.Add _
                    kodeKey, _
                    rowIndex

            End If

        End If

    Next rowIndex

    Set STO_BuildTargetKodeUnikMap = _
        resultMap

End Function
Private Sub STO_CopyCompiledRowByHeaders( _
    ByVal wsSource As Worksheet, _
    ByVal sourceRow As Long, _
    ByVal sourceHeaderMap As Object, _
    ByVal wsTarget As Worksheet, _
    ByVal targetRow As Long, _
    ByVal targetHeaderMap As Object)

    Dim targetKey As Variant
    Dim sourceColumn As Long
    Dim targetColumn As Long
    Dim sourceValue As Variant

    If wsSource Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub
    If sourceHeaderMap Is Nothing Then Exit Sub
    If targetHeaderMap Is Nothing Then Exit Sub

    For Each targetKey In targetHeaderMap.keys

        Select Case UCase$( _
            CStr(targetKey))

            Case "FILTER KATEGORISASI FPTK", _
                 "FPTK AVAILABILITY", _
                 "CREATED AT", _
                 "LAST UPDATED AT", _
                 "LAST COMPILE ACTION"

                ' Diatur oleh compiler target.

            Case Else

                If sourceHeaderMap.Exists( _
                    CStr(targetKey)) Then

                    sourceColumn = _
                        CLng( _
                            sourceHeaderMap( _
                                CStr(targetKey)))

                    targetColumn = _
                        CLng( _
                            targetHeaderMap( _
                                CStr(targetKey)))

                    sourceValue = _
                        wsSource.Cells( _
                            sourceRow, _
                            sourceColumn).Value

                    ' Jangan hapus enrichment yang sudah ada
                    ' hanya karena source STO kosong.
                    If Not ValueIsBlankFPTK( _
                        sourceValue) Then

                        wsTarget.Cells( _
                            targetRow, _
                            targetColumn).Value = _
                                sourceValue

                    End If

                End If

        End Select

    Next targetKey

End Sub

Private Function STO_NormalizePosition( _
    ByVal rawValue As Variant) As String

    Const STO_PREFIX As String = _
        "SALES TAKING ORDER STAFF (CHILLED)"

    Dim sourceText As String
    Dim areaText As String
    Dim splitPosition As Long

    sourceText = _
        STO_NormalizeText( _
            rawValue)

    If Len(sourceText) = 0 Then
        Exit Function
    End If

    If Left$( _
        sourceText, _
        Len(STO_PREFIX)) <> _
       STO_PREFIX Then

        STO_NormalizePosition = _
            sourceText

        Exit Function

    End If

    splitPosition = _
        InStr( _
            Len(STO_PREFIX) + 1, _
            sourceText, _
            "-", _
            vbTextCompare)

    If splitPosition > 0 Then

        areaText = _
            Trim$( _
                Mid$( _
                    sourceText, _
                    splitPosition + 1))

    Else

        areaText = _
            Trim$( _
                Mid$( _
                    sourceText, _
                    Len(STO_PREFIX) + 1))

    End If

    Select Case areaText

        Case "JAKARTA PUSAT & UTARA", _
             "JAKARTA UTARA & PUSAT", _
             "JAKARTA UTARA"

            areaText = _
                "JAKARTA PUSAT & UTARA"

    End Select

    If Len(areaText) > 0 Then

        STO_NormalizePosition = _
            STO_PREFIX & _
            " - " & _
            areaText

    Else

        STO_NormalizePosition = _
            STO_PREFIX

    End If

End Function

Private Function STO_NormalizeCandidate( _
    ByVal rawValue As Variant) As String

    STO_NormalizeCandidate = _
        STO_NormalizeText( _
            rawValue)

End Function

Private Function STO_NormalizeText( _
    ByVal rawValue As Variant) As String

    Dim sourceText As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    sourceText = _
        CStr(rawValue)

    sourceText = _
        Replace( _
            sourceText, _
            ChrW(160), _
            " ")

    sourceText = _
        Replace( _
            sourceText, _
            vbCr, _
            " ")

    sourceText = _
        Replace( _
            sourceText, _
            vbLf, _
            " ")

    sourceText = _
        Trim$( _
            sourceText)

    Do While InStr( _
        1, _
        sourceText, _
        "  ", _
        vbBinaryCompare) > 0

        sourceText = _
            Replace( _
                sourceText, _
                "  ", _
                " ")

    Loop

    STO_NormalizeText = _
        UCase$( _
            sourceText)

End Function

Private Function STO_FindSourceSheet( _
    ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function

    For Each ws In wb.Worksheets

        If InStr( _
            1, _
            STO_NormalizeText( _
                ws.Name), _
            "FPTK STO", _
            vbTextCompare) > 0 Then

            Set STO_FindSourceSheet = _
                ws

            Exit Function

        End If

    Next ws

End Function

Private Function STO_FindHeaderRow( _
    ByVal ws As Worksheet) As Long

    Dim rowIndex As Long
    Dim lastColumn As Long

    Dim colPosisi As Long
    Dim colNama As Long

    If ws Is Nothing Then Exit Function

    For rowIndex = 1 To 10

        lastColumn = _
            ws.Cells( _
                rowIndex, _
                ws.Columns.Count).End( _
                    xlToLeft).Column

        colPosisi = _
            STO_FindHeaderColumn( _
                ws, _
                rowIndex, _
                "POSISI FPTK")

        colNama = _
            STO_FindHeaderColumn( _
                ws, _
                rowIndex, _
                "NAMA KANDIDAT")

        If colPosisi > 0 And _
           colNama > 0 Then

            STO_FindHeaderRow = _
                rowIndex

            Exit Function

        End If

    Next rowIndex

End Function

Private Function STO_FindHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal wantedHeader As String) As Long

    Dim lastColumn As Long
    Dim columnIndex As Long

    Dim wantedKey As String
    Dim currentKey As String

    If ws Is Nothing Then Exit Function
    If headerRow <= 0 Then Exit Function

    wantedKey = _
        STO_NormalizeText( _
            wantedHeader)

    lastColumn = _
        ws.Cells( _
            headerRow, _
            ws.Columns.Count).End( _
                xlToLeft).Column

    For columnIndex = 1 To lastColumn

        currentKey = _
            STO_NormalizeText( _
                ws.Cells( _
                    headerRow, _
                    columnIndex).Value2)

        If currentKey = _
           wantedKey Then

            STO_FindHeaderColumn = _
                columnIndex

            Exit Function

        End If

    Next columnIndex

End Function
Private Function STO_IsValidCompiledSourcePath( _
    ByVal sourcePath As String) As Boolean

    Dim wbCheck As Workbook
    Dim wsCheck As Worksheet
    Dim openedByMacro As Boolean

    Dim headerRow As Long
    Dim colPosisi As Long
    Dim colNama As Long
    Dim colRealDate As Long
    Dim colAvailability As Long

    If Len(Trim$(sourcePath)) = 0 Then
        Exit Function
    End If

    If Not STO_FileExists( _
        sourcePath) Then

        Exit Function

    End If

    On Error Resume Next

    Set wbCheck = _
        Workbooks( _
            STO_FileNameFromPath( _
                sourcePath))

    On Error GoTo InvalidSource

    If wbCheck Is Nothing Then

        Set wbCheck = _
            Workbooks.Open( _
                fileName:=sourcePath, _
                UpdateLinks:=0, _
                ReadOnly:=True, _
                AddToMru:=False, _
                IgnoreReadOnlyRecommended:=True)

        openedByMacro = True

    End If

    On Error GoTo 0

    On Error Resume Next

    Set wsCheck = _
        wbCheck.Worksheets( _
            "FPTK")

    On Error GoTo InvalidSource

    If wsCheck Is Nothing Then
        GoTo SafeExit
    End If

    headerRow = _
        STO_FindCompiledHeaderRow( _
            wsCheck)

    If headerRow = 0 Then
        GoTo SafeExit
    End If

    colPosisi = _
        FPTK_DirectFindHeader( _
            wsCheck, _
            headerRow, _
            "Posisi")

    colNama = _
        FPTK_DirectFindHeader( _
            wsCheck, _
            headerRow, _
            "Nama Kandidat")

    colRealDate = _
        FPTK_DirectFindHeader( _
            wsCheck, _
            headerRow, _
            "FPTK Date (Real)")

    colAvailability = _
        STO_FindSourceAvailabilityColumn( _
            wsCheck, _
            headerRow)

    STO_IsValidCompiledSourcePath = _
        (colPosisi > 0 And _
         colRealDate > 0 And _
         colAvailability > 0)

SafeExit:

    On Error Resume Next

    If openedByMacro Then

        wbCheck.Close _
            SaveChanges:=False

    End If

    Set wsCheck = Nothing
    Set wbCheck = Nothing

    On Error GoTo 0

    Exit Function

InvalidSource:

    STO_IsValidCompiledSourcePath = _
        False

    Err.Clear

    Resume SafeExit

End Function

Private Function STO_ResolveSourcePath() As String

    Dim candidatePath As String
    Dim settingPath As String
    Dim oldSettingPath As String
    Dim parentPath As String

    ' ========================================================
    ' PRIORITY 1:
    ' New dedicated setting.
    ' ========================================================
    settingPath = _
        Trim$( _
            STO_GetSettingValue( _
                "STOCompiledFilePath"))

    If STO_IsValidCompiledSourcePath( _
        settingPath) Then

        STO_ResolveSourcePath = _
            settingPath

        Exit Function

    End If

    ' ========================================================
    ' PRIORITY 2:
    ' STO.xlsm next to master workbook.
    ' ========================================================
    candidatePath = _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        "STO.xlsm"

    If STO_IsValidCompiledSourcePath( _
        candidatePath) Then

        STO_ResolveSourcePath = _
            candidatePath

        Exit Function

    End If

    candidatePath = _
        ThisWorkbook.Path & _
        Application.PathSeparator & _
        "STO.xlsx"

    If STO_IsValidCompiledSourcePath( _
        candidatePath) Then

        STO_ResolveSourcePath = _
            candidatePath

        Exit Function

    End If

    ' ========================================================
    ' PRIORITY 3:
    ' Parent folder.
    ' ========================================================
    parentPath = _
        STO_ParentFolder( _
            ThisWorkbook.Path)

    If Len(parentPath) > 0 Then

        candidatePath = _
            parentPath & _
            Application.PathSeparator & _
            "STO.xlsm"

        If STO_IsValidCompiledSourcePath( _
            candidatePath) Then

            STO_ResolveSourcePath = _
                candidatePath

            Exit Function

        End If

        candidatePath = _
            parentPath & _
            Application.PathSeparator & _
            "STO.xlsx"

        If STO_IsValidCompiledSourcePath( _
            candidatePath) Then

            STO_ResolveSourcePath = _
                candidatePath

            Exit Function

        End If

    End If

    ' ========================================================
    ' PRIORITY 4:
    ' Old setting key for backward compatibility ONLY.
    '
    ' Kalau sebelumnya STOFilePath menunjuk ke
    ' Tulang Punggung Sales Cimory Group.xlsx, file itu
    ' otomatis ditolak karena tidak punya compiled FPTK sheet
    ' dengan struktur yang diwajibkan.
    ' ========================================================
    oldSettingPath = _
        Trim$( _
            STO_GetSettingValue( _
                "STOFilePath"))

    If STO_IsValidCompiledSourcePath( _
        oldSettingPath) Then

        STO_ResolveSourcePath = _
            oldSettingPath

        Exit Function

    End If

End Function

Private Function STO_GetSettingValue( _
    ByVal settingName As String) As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            "Settings")

    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            1).End( _
                xlUp).Row

    For rowIndex = 1 To lastRow

        If StrComp( _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        rowIndex, _
                        1).Value2)), _
            settingName, _
            vbTextCompare) = 0 Then

            STO_GetSettingValue = _
                Trim$( _
                    CStr( _
                        ws.Cells( _
                            rowIndex, _
                            2).Value2))

            Exit Function

        End If

    Next rowIndex

End Function

Private Sub STO_SaveSettingValue( _
    ByVal settingName As String, _
    ByVal settingValue As String)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim targetRow As Long

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            "Settings")

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = _
            ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        ws.Name = _
            "Settings"

        ws.Cells(1, 1).Value2 = _
            "Setting"

        ws.Cells(1, 2).Value2 = _
            "Value"

    End If

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            1).End( _
                xlUp).Row

    For rowIndex = 2 To lastRow

        If StrComp( _
            Trim$( _
                CStr( _
                    ws.Cells( _
                        rowIndex, _
                        1).Value2)), _
            settingName, _
            vbTextCompare) = 0 Then

            targetRow = _
                rowIndex

            Exit For

        End If

    Next rowIndex

    If targetRow = 0 Then

        targetRow = _
            lastRow + 1

        If targetRow < 2 Then
            targetRow = 2
        End If

        ws.Cells( _
            targetRow, _
            1).Value2 = _
                settingName

    End If

    ws.Cells( _
        targetRow, _
        2).Value2 = _
            settingValue

End Sub

Private Function STO_FileExists( _
    ByVal filePath As String) As Boolean

    If Len(Trim$(filePath)) = 0 Then
        Exit Function
    End If

    On Error Resume Next

    STO_FileExists = _
        Len(Dir$( _
            filePath, _
            vbNormal Or _
            vbReadOnly Or _
            vbHidden Or _
            vbSystem)) > 0

    On Error GoTo 0

End Function

Private Function STO_FileNameFromPath( _
    ByVal filePath As String) As String

    Dim separatorPosition As Long

    separatorPosition = _
        InStrRev( _
            filePath, _
            Application.PathSeparator)

    If separatorPosition > 0 Then

        STO_FileNameFromPath = _
            Mid$( _
                filePath, _
                separatorPosition + 1)

    Else

        STO_FileNameFromPath = _
            filePath

    End If

End Function

Private Function STO_ParentFolder( _
    ByVal folderPath As String) As String

    Dim cleanPath As String
    Dim separatorPosition As Long

    cleanPath = _
        folderPath

    Do While Right$( _
        cleanPath, _
        1) = Application.PathSeparator

        cleanPath = _
            Left$( _
                cleanPath, _
                Len(cleanPath) - 1)

    Loop

    separatorPosition = _
        InStrRev( _
            cleanPath, _
            Application.PathSeparator)

    If separatorPosition > 0 Then

        STO_ParentFolder = _
            Left$( _
                cleanPath, _
                separatorPosition - 1)

    End If

End Function
Public Sub RefreshFPTKKategorisasiNow()

    Dim ws As Worksheet
    Dim headerMap As Object

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("FPTK")

    Set headerMap = BuildHeaderMap(ws, FPTK_HEADER_ROW)

    ' Refresh kategori dasar.
    FillFPTKHelperFormulas

    ' Re-apply STO dengan protection CLAP FGDP.
    Set headerMap = BuildHeaderMap(ws, FPTK_HEADER_ROW)

    RefreshFPTKSTOCategorization _
        ws, _
        headerMap, _
        True

    MsgBox _
        "Filter Kategorisasi FPTK berhasil di-refresh." & vbCrLf & vbCrLf & _
        "Priority:" & vbCrLf & _
        "1. CLAP FGDP (Cimory/Fresh)" & vbCrLf & _
        "2. STO" & vbCrLf & _
        "3. Level 1-2 / Level 3 / Level 4", _
        vbInformation, _
        "Refresh Kategorisasi"

    Set headerMap = Nothing
    Set ws = Nothing
    Exit Sub

ErrorHandler:

    MsgBox _
        "Refresh Filter Kategorisasi gagal." & vbCrLf & vbCrLf & _
        "Error " & CStr(Err.Number) & ": " & Err.Description, _
        vbCritical, _
        "Refresh Kategorisasi"

    Set headerMap = Nothing
    Set ws = Nothing

End Sub

Private Function CalculateFilterKategorisasiFPTK( _
    ByVal posisiValue As Variant, _
    ByVal levelValue As Variant, _
    Optional ByVal existingFilterValue As Variant) As String

    Dim posisiText As String
    Dim levelNumber As Long
    Dim existingFilterText As String


    ' ========================================================
    ' CLEAN VALUE
    ' ========================================================

    posisiText = _
        LCase$( _
            Trim$( _
                CStr(posisiValue)))


    existingFilterText = _
        UCase$( _
            Trim$( _
                CStr(existingFilterValue)))


    ' ========================================================
    ' PRIORITY 1 - CLAP FGDP
    '
    ' Posisi yang diawali "Cimory" atau "Fresh"
    ' tetap dianggap CLAP FGDP
    '
    ' Rule ini tetap paling tinggi
    ' ========================================================

    If Left$(posisiText, 6) = "cimory" Or _
       Left$(posisiText, 5) = "fresh" Then

        CalculateFilterKategorisasiFPTK = _
            "CLAP FGDP"

        Exit Function

    End If



    ' ========================================================
    ' PRIORITY 2 - STO
    '
    ' Jika FILTER KATEGORISASI dari file input
    ' mengandung kata STO
    '
    ' Contoh:
    ' STO
    ' STO MP
    ' STO JESS
    ' SALES STO
    '
    ' Semua dianggap STO
    ' ========================================================

    If InStr(1, existingFilterText, "STO", vbTextCompare) > 0 Then

        CalculateFilterKategorisasiFPTK = _
            "STO"

        Exit Function

    End If



    ' ========================================================
    ' PRIORITY 3 - LEVEL
    ' ========================================================

    If IsNumeric(levelValue) Then

        levelNumber = CLng(levelValue)

    Else

        levelNumber = _
            ExtractFirstNumberLocal( _
                CStr(levelValue))

    End If



    Select Case levelNumber


        Case 1, 2

            CalculateFilterKategorisasiFPTK = _
                "Level 1-2"


        Case 3

            CalculateFilterKategorisasiFPTK = _
                "Level 3"


        Case 4

            CalculateFilterKategorisasiFPTK = _
                "Level 4"


        Case Else

            CalculateFilterKategorisasiFPTK = _
                vbNullString


    End Select


End Function

Public Sub RefreshFPTKSLAFromRealDateNow()

    ' Recalculate seluruh SLA memakai:
    ' FPTK Date (Real) + Jumlah SLA.
    FillFPTKHelperFormulas

End Sub

Public Sub InitializeFPTKAuditForExistingRows()

    Dim ws As Worksheet
    Dim headerMap As Object

    Dim colCreatedAt As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim initializedAt As Date

    Set ws = GetOrCreateWorksheet( _
        "FPTK")

    ProtectFPTKHeader

    Set headerMap = BuildHeaderMap( _
        ws, _
        FPTK_HEADER_ROW)

    colCreatedAt = ColumnByKey( _
        headerMap, _
        "CREATED AT")

    If colCreatedAt = 0 Then Exit Sub

    lastRow = LastTargetDataRow( _
        ws, _
        headerMap)

    If lastRow < FPTK_FIRST_DATA_ROW Then
        Exit Sub
    End If

    initializedAt = Now

    For rowIndex = FPTK_FIRST_DATA_ROW _
        To lastRow

        If RowHasPrimaryData( _
            ws, _
            rowIndex, _
            headerMap) Then

            If Len(Trim$(CStr( _
                ws.Cells( _
                    rowIndex, _
                    colCreatedAt).Value2))) = 0 Then

                ws.Cells( _
                    rowIndex, _
                    colCreatedAt).Value = _
                        initializedAt

            End If

        End If

    Next rowIndex

    ApplyFPTKAuditFormats _
        ws, _
        headerMap

    AddRunLog _
        "Info", _
        "InitializeFPTKAuditForExistingRows done"

    Set headerMap = Nothing
    Set ws = Nothing

End Sub

Public Sub RefreshFPTKDateWeekMonthNow()

    Dim ws As Worksheet
    Dim headerMap As Object

    Set ws = _
        GetOrCreateWorksheet( _
            "FPTK")

    ProtectFPTKHeader

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    RefreshFPTKWeekMonthValues _
        ws, _
        headerMap

    ApplyFPTKDateFormats _
        ws, _
        headerMap

    AddRunLog _
        "Info", _
        "Refresh FPTK date/week/month completed"

    Set headerMap = Nothing
    Set ws = Nothing

End Sub

Public Sub RefreshFPTKWeekMonthValuesNow()

    Dim ws As Worksheet
    Dim headerMap As Object

    Set ws = GetOrCreateWorksheet("FPTK")

    ProtectFPTKHeader

    Set headerMap = BuildHeaderMap( _
        ws, _
        FPTK_HEADER_ROW)

    RefreshFPTKWeekMonthValues _
        ws, _
        headerMap

    AddRunLog _
        "Info", _
        "Refresh FPTK Week/Month values done"

    Set headerMap = Nothing
    Set ws = Nothing

End Sub
Private Function FPTK_ReadColumnArray( _
    ByVal ws As Worksheet, _
    ByVal columnIndex As Long, _
    ByVal rowCount As Long) As Variant

    Dim rawValue As Variant
    Dim outputData() As Variant

    If rowCount <= 0 Then Exit Function

    ReDim outputData( _
        1 To rowCount, _
        1 To 1)

    If ws Is Nothing Or _
       columnIndex <= 0 Then

        FPTK_ReadColumnArray = outputData
        Exit Function

    End If

    rawValue = _
        ws.Cells( _
            FPTK_FIRST_DATA_ROW, _
            columnIndex).Resize( _
                rowCount, _
                1).Value2

    If rowCount = 1 Then

        outputData(1, 1) = rawValue

    Else

        outputData = rawValue

    End If

    FPTK_ReadColumnArray = outputData

End Function
Private Sub FPTK_StoreUserLookupPair( _
    ByVal lookupMap As Object, _
    ByVal lookupKey As String, _
    ByVal resultPair As Variant)

    Dim existingPair As Variant
    Dim newHasValue As Boolean
    Dim oldHasValue As Boolean

    If lookupMap Is Nothing Then Exit Sub
    If Len(lookupKey) = 0 Then Exit Sub

    newHasValue = _
        Len(Trim$(CStr(resultPair(0)))) > 0 Or _
        Len(Trim$(CStr(resultPair(1)))) > 0

    If Not lookupMap.Exists(lookupKey) Then

        lookupMap.Add _
            lookupKey, _
            resultPair

        Exit Sub

    End If

    existingPair = _
        lookupMap( _
            lookupKey)

    oldHasValue = _
        Len(Trim$(CStr(existingPair(0)))) > 0 Or _
        Len(Trim$(CStr(existingPair(1)))) > 0

    If newHasValue Or _
       Not oldHasValue Then

        lookupMap(lookupKey) = _
            resultPair

    End If

End Sub
Private Sub FPTK_StoreLatestPositionRow( _
    ByVal lookupMap As Object, _
    ByVal lookupKey As String, _
    ByVal dbRow As Long, _
    ByVal dbData As Variant)

    Dim currentRow As Long
    Dim currentRank As Double
    Dim newRank As Double

    If lookupMap Is Nothing Then Exit Sub
    If Len(lookupKey) = 0 Then Exit Sub

    newRank = FPTK_PositionCodeRank( _
        dbData(dbRow, 1))

    If Not lookupMap.Exists(lookupKey) Then

        lookupMap.Add lookupKey, dbRow
        Exit Sub

    End If

    currentRow = CLng(lookupMap(lookupKey))

    currentRank = FPTK_PositionCodeRank( _
        dbData(currentRow, 1))

    If newRank > currentRank Or _
       (newRank = currentRank And dbRow > currentRow) Then

        lookupMap(lookupKey) = dbRow

    End If

End Sub

Private Function FPTK_MapMasterRow( _
    ByVal lookupMap As Object, _
    ByVal lookupKey As String) As Long

    If lookupMap Is Nothing Then Exit Function
    If Len(lookupKey) = 0 Then Exit Function

    If lookupMap.Exists(lookupKey) Then
        FPTK_MapMasterRow = CLng(lookupMap(lookupKey))
    End If

End Function

Private Function FPTK_PositionCodeRank( _
    ByVal rawValue As Variant) As Double

    Dim sourceText As String
    Dim digitText As String
    Dim characterIndex As Long
    Dim currentCharacter As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    sourceText = Trim$(CStr(rawValue))

    For characterIndex = 1 To Len(sourceText)

        currentCharacter = Mid$( _
            sourceText, characterIndex, 1)

        If currentCharacter >= "0" And _
           currentCharacter <= "9" Then

            digitText = digitText & _
                currentCharacter

        End If

    Next characterIndex

    If Len(digitText) > 0 Then

        On Error Resume Next
        FPTK_PositionCodeRank = CDbl(digitText)
        On Error GoTo 0

    End If

End Function

Private Function FPTK_FindBestUserRow( _
    ByVal dbData As Variant, _
    ByVal posisiKey As String, _
    ByVal direktoratKey As String, _
    ByVal buKey As String, _
    ByVal lokasiKey As String) As Long

    Const DB_COL_KODE As Long = 1
    Const DB_COL_POSITION As Long = 2
    Const DB_COL_LOCATION As Long = 3
    Const DB_COL_BUSINESS_UNIT As Long = 4
    Const DB_COL_USER_MANAGER As Long = 7
    Const DB_COL_INDIRECT_USER As Long = 8
    Const DB_COL_DIRECTORATE As Long = 9

    Dim dbRow As Long
    Dim score As Long
    Dim bestScore As Long
    Dim codeRank As Double
    Dim bestCodeRank As Double

    Dim dbPosisi As String
    Dim dbDirektorat As String
    Dim dbBU As String
    Dim dbLokasi As String
    Dim dbUser As String
    Dim dbIndirect As String

    If Not IsArray(dbData) Then Exit Function
    If Len(posisiKey) = 0 Then Exit Function

    For dbRow = 1 To UBound(dbData, 1)

        dbUser = FPTK_SafeCellText( _
            dbData(dbRow, DB_COL_USER_MANAGER))

        dbIndirect = FPTK_SafeCellText( _
            dbData(dbRow, DB_COL_INDIRECT_USER))

        If Len(dbUser) > 0 Or Len(dbIndirect) > 0 Then

            dbPosisi = FPTK_NormalizeLookupText( _
                dbData(dbRow, DB_COL_POSITION))

            If dbPosisi = posisiKey Then

                dbDirektorat = FPTK_NormalizeLookupText( _
                    dbData(dbRow, DB_COL_DIRECTORATE))

                dbBU = FPTK_NormalizeLookupText( _
                    dbData(dbRow, DB_COL_BUSINESS_UNIT))

                dbLokasi = FPTK_NormalizeLookupText( _
                    dbData(dbRow, DB_COL_LOCATION))

                score = 100

                If Len(buKey) > 0 And dbBU = buKey Then
                    score = score + 100
                End If

                If Len(direktoratKey) > 0 And _
                   dbDirektorat = direktoratKey Then

                    score = score + 80

                End If

                If Len(lokasiKey) > 0 And _
                   dbLokasi = lokasiKey Then

                    score = score + 40

                End If

                codeRank = FPTK_PositionCodeRank( _
                    dbData(dbRow, DB_COL_KODE))

                If score > bestScore Or _
                   (score = bestScore And _
                    codeRank > bestCodeRank) Then

                    bestScore = score
                    bestCodeRank = codeRank
                    FPTK_FindBestUserRow = dbRow

                End If

            End If

        End If

    Next dbRow

End Function
Private Sub RefreshFPTKReasonValues( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object)

    Dim colReason As Long
    Dim lastRow As Long
    Dim rowCount As Long
    Dim rowIndex As Long

    Dim reasonData As Variant
    Dim rawText As String

    If ws Is Nothing Then Exit Sub
    If headerMap Is Nothing Then Exit Sub

    colReason = _
        FPTK_TargetColumnByCanonical( _
            ws, _
            headerMap, _
            "CATEGORY")

    If colReason <= 0 Then Exit Sub

    lastRow = _
        LastTargetDataRow( _
            ws, _
            headerMap)

    If lastRow < FPTK_FIRST_DATA_ROW Then
        Exit Sub
    End If

    rowCount = _
        lastRow - _
        FPTK_FIRST_DATA_ROW + 1

    reasonData = _
        FPTK_ReadColumnArray( _
            ws, _
            colReason, _
            rowCount)

    For rowIndex = 1 To rowCount

        rawText = _
            FPTK_SafeCellText( _
                reasonData( _
                    rowIndex, _
                    1))

        ' Jangan kosongkan alasan valid yang sudah datang
        ' dari source hanya karena dropdown berubah.
        '
        ' Legacy value boolean yang jelas bukan alasan
        ' dibersihkan. Nilai alasan lainnya dipertahankan.
        If Len(rawText) > 0 Then

            If Not FPTK_IsValidSourceReasonValue( _
                rawText) Then

                reasonData( _
                    rowIndex, _
                    1) = Empty

            End If

        End If

    Next rowIndex

    ws.Cells( _
        FPTK_FIRST_DATA_ROW, _
        colReason).Resize( _
            rowCount, _
            1).Value2 = _
                reasonData

End Sub

Private Function FPTK_LoadAllowedReasonMap() As Object

    Const DROPDOWN_REASON_COL As Long = 3

    Dim resultMap As Object
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim rowIndex As Long

    Dim itemText As String
    Dim itemKey As String

    Set resultMap = _
        CreateObject( _
            "Scripting.Dictionary")

    resultMap.CompareMode = vbTextCompare

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            "Kode Dropdown")

    If ws Is Nothing Then
        Set ws = _
            ThisWorkbook.Worksheets( _
                "Kode Dropdown (tidak diubah)")
    End If

    On Error GoTo 0

    If ws Is Nothing Then

        Set FPTK_LoadAllowedReasonMap = _
            resultMap

        Exit Function

    End If

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            DROPDOWN_REASON_COL).End( _
                xlUp).Row

    For rowIndex = 2 To lastRow

        itemText = _
            FPTK_SafeCellText( _
                ws.Cells( _
                    rowIndex, _
                    DROPDOWN_REASON_COL).Value2)

        itemKey = _
            FPTK_NormalizeLookupText( _
                itemText)

        If Len(itemKey) > 0 Then

            If Not resultMap.Exists( _
                itemKey) Then

                resultMap.Add _
                    itemKey, _
                    itemText

            End If

        End If

    Next rowIndex

    Set FPTK_LoadAllowedReasonMap = _
        resultMap

    Set ws = Nothing

End Function

Private Function FPTK_TargetColumnByCanonical( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object, _
    ByVal canonicalKey As String) As Long

    Dim resultColumn As Long
    Dim lastColumn As Long
    Dim columnIndex As Long

    If ws Is Nothing Then Exit Function

    If Not headerMap Is Nothing Then

        resultColumn = _
            ColumnByKey( _
                headerMap, _
                canonicalKey)

    End If

    If resultColumn > 0 Then

        FPTK_TargetColumnByCanonical = _
            resultColumn

        Exit Function

    End If

    lastColumn = _
        ws.Cells( _
            FPTK_HEADER_ROW, _
            ws.Columns.Count).End( _
                xlToLeft).Column

    For columnIndex = 1 To lastColumn

        If CanonicalHeaderKey( _
            ws.Cells( _
                FPTK_HEADER_ROW, _
                columnIndex).Value2) = _
           canonicalKey Then

            FPTK_TargetColumnByCanonical = _
                columnIndex

            Exit Function

        End If

    Next columnIndex

End Function

Private Function FPTK_LoadMappingUserMap() As Object

    Dim resultMap As Object
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim rowIndex As Long

    Dim posisiKey As String
    Dim userText As String
    Dim indirectText As String

    Set resultMap = _
        CreateObject( _
            "Scripting.Dictionary")

    resultMap.CompareMode = vbTextCompare

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            "Mapping User")

    On Error GoTo 0

    If ws Is Nothing Then

        Set FPTK_LoadMappingUserMap = _
            resultMap

        Exit Function

    End If

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            1).End( _
                xlUp).Row

    For rowIndex = lastRow _
        To 2 Step -1

        posisiKey = _
            FPTK_NormalizeLookupText( _
                ws.Cells( _
                    rowIndex, _
                    1).Value2)

        userText = _
            FPTK_SafeCellText( _
                ws.Cells( _
                    rowIndex, _
                    2).Value2)

        indirectText = _
            FPTK_SafeCellText( _
                ws.Cells( _
                    rowIndex, _
                    3).Value2)

        directorateText = _
            FPTK_SafeCellText( _
                ws.Cells( _
                    rowIndex, _
                    4).Value2)

        If Len(posisiKey) > 0 Then

            If Not resultMap.Exists( _
                posisiKey) Then

                resultMap.Add _
                    posisiKey, _
                    Array( _
                        userText, _
                        indirectText, _
                        directorateText)

            End If

        End If

    Next rowIndex

    Set FPTK_LoadMappingUserMap = _
        resultMap

    Set ws = Nothing

End Function

Private Function FPTK_FindBestMasterRow( _
    ByVal dbData As Variant, _
    ByVal posisiKey As String, _
    ByVal buKey As String, _
    ByVal lokasiKey As String, _
    ByVal direktoratKey As String, _
    Optional ByVal requireUser As Boolean = False) As Long

    Const DB_COL_KODE As Long = 1
    Const DB_COL_POSITION As Long = 2
    Const DB_COL_LOCATION As Long = 3
    Const DB_COL_BUSINESS_UNIT As Long = 4
    Const DB_COL_USER_MANAGER As Long = 7
    Const DB_COL_INDIRECT_USER As Long = 8
    Const DB_COL_DIRECTORATE As Long = 9

    Dim rowIndex As Long
    Dim score As Long
    Dim bestScore As Long

    Dim codeRank As Double
    Dim bestCodeRank As Double

    Dim dbPosisiKey As String
    Dim dbBUKey As String
    Dim dbLokasiKey As String
    Dim dbDirektoratKey As String

    Dim hasUser As Boolean

    If Len(posisiKey) = 0 Then Exit Function

    bestScore = -1
    bestCodeRank = -1

    For rowIndex = 1 To UBound( _
        dbData, _
        1)

        dbPosisiKey = _
            FPTK_NormalizeLookupText( _
                dbData( _
                    rowIndex, _
                    DB_COL_POSITION))

        If dbPosisiKey = posisiKey Then

            hasUser = _
                Len(FPTK_SafeCellText( _
                    dbData( _
                        rowIndex, _
                        DB_COL_USER_MANAGER))) > 0 Or _
                Len(FPTK_SafeCellText( _
                    dbData( _
                        rowIndex, _
                        DB_COL_INDIRECT_USER))) > 0

            If Not requireUser Or _
               hasUser Then

                score = 1000

                dbBUKey = _
                    FPTK_NormalizeLookupText( _
                        dbData( _
                            rowIndex, _
                            DB_COL_BUSINESS_UNIT))

                dbLokasiKey = _
                    FPTK_NormalizeLookupText( _
                        dbData( _
                            rowIndex, _
                            DB_COL_LOCATION))

                dbDirektoratKey = _
                    FPTK_NormalizeLookupText( _
                        dbData( _
                            rowIndex, _
                            DB_COL_DIRECTORATE))

                If Len(buKey) > 0 And _
                   dbBUKey = buKey Then

                    score = score + 300

                End If

                If Len(lokasiKey) > 0 And _
                   dbLokasiKey = lokasiKey Then

                    score = score + 150

                End If

                If Len(direktoratKey) > 0 And _
                   dbDirektoratKey = direktoratKey Then

                    score = score + 50

                End If

                codeRank = _
                    FPTK_PositionCodeRank( _
                        dbData( _
                            rowIndex, _
                            DB_COL_KODE))

                If score > bestScore Or _
                   (score = bestScore And _
                    codeRank > bestCodeRank) Then

                    bestScore = score
                    bestCodeRank = codeRank

                    FPTK_FindBestMasterRow = _
                        rowIndex

                End If

            End If

        End If

    Next rowIndex

End Function
Private Function FPTK_DirectFindHeader( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal wantedHeader As String) As Long

    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim wantedKey As String

    If ws Is Nothing Then Exit Function

    wantedKey = FPTK_DirectNorm(wantedHeader)

    lastColumn = ws.Cells( _
        headerRow, ws.Columns.Count).End(xlToLeft).Column

    For columnIndex = 1 To lastColumn

        If FPTK_DirectNorm( _
            ws.Cells(headerRow, columnIndex).Value2) = wantedKey Then

            FPTK_DirectFindHeader = columnIndex
            Exit Function

        End If

    Next columnIndex

End Function

Private Function FPTK_DirectNorm( _
    ByVal rawValue As Variant) As String

    Dim sourceText As String
    Dim resultText As String
    Dim i As Long
    Dim ch As String

    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
        Exit Function
    End If

    sourceText = UCase$(Trim$(CStr(rawValue)))

    For i = 1 To Len(sourceText)

        ch = Mid$(sourceText, i, 1)

        If ch Like "[A-Z0-9]" Then
            resultText = resultText & ch
        End If

    Next i

    FPTK_DirectNorm = resultText

End Function

Private Function FPTK_DirectCodeRank( _
    ByVal rawValue As Variant) As Double

    Dim sourceText As String
    Dim digits As String
    Dim i As Long
    Dim ch As String

    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
        Exit Function
    End If

    sourceText = CStr(rawValue)

    For i = 1 To Len(sourceText)

        ch = Mid$(sourceText, i, 1)

        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        End If

    Next i

    If Len(digits) > 0 Then
        On Error Resume Next
        FPTK_DirectCodeRank = CDbl(digits)
        On Error GoTo 0
    End If

End Function
Private Sub RefreshFPTKUsersFromDBKodePosisi( _
    ByVal wsFPTK As Worksheet, _
    ByVal headerMap As Object)

    Const FPTK_HEADER As Long = 1
    Const FPTK_FIRST_ROW As Long = 2

    Const DB_COL_KODE As Long = 1
    Const DB_COL_POSITION As Long = 2
    Const DB_COL_LOCATION As Long = 3
    Const DB_COL_BUSINESS_UNIT As Long = 4
    Const DB_COL_USER As Long = 7
    Const DB_COL_INDIRECT As Long = 8
    Const DB_COL_DIRECTORATE As Long = 9

    Dim wsDB As Worksheet
    Dim wsMap As Worksheet

    Dim colKode As Long
    Dim colPosisi As Long
    Dim colBU As Long
    Dim colDirektorat As Long
    Dim colLokasi As Long
    Dim colUser As Long
    Dim colIndirect As Long

    Dim lastFPTKRow As Long
    Dim lastDBRow As Long
    Dim lastMapRow As Long

    Dim bestMapRow As Long
    Dim mapScore As Long
    Dim bestMapScore As Long
    Dim mapDirectorateKey As String
    Dim mapUserText As String
    Dim mapIndirectText As String

    Dim fptkRow As Long
    Dim dbRow As Long
    Dim mapRow As Long

    Dim posisiKey As String
    Dim buKey As String
    Dim lokasiKey As String
    Dim direktoratKey As String

    Dim dbPosisiKey As String
    Dim dbBUKey As String
    Dim dbLokasiKey As String
    Dim dbDirektoratKey As String

    Dim score As Long
    Dim bestScore As Long
    Dim codeRank As Double
    Dim bestCodeRank As Double

    Dim bestRow As Long
    Dim bestUserRow As Long
    Dim userScore As Long
    Dim bestUserScore As Long
    Dim userCodeRank As Double
    Dim bestUserCodeRank As Double

    Dim userText As String
    Dim indirectText As String
    Dim directorateText As String

    Dim matchedCount As Long
    Dim writtenUserCount As Long
    Dim mappingFallbackCount As Long
    Dim unmatchedCount As Long

    If wsFPTK Is Nothing Then Exit Sub

    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets("DB Kode Posisi")
    Set wsMap = ThisWorkbook.Worksheets("Mapping User")
    On Error GoTo 0

    If wsDB Is Nothing Then
        MsgBox _
            "Sheet DB Kode Posisi tidak ditemukan.", _
            vbExclamation, _
            "Refresh User FPTK"
        Exit Sub
    End If

    ' Cari kolom target langsung dari header row 1.
    ' Tidak memakai Schema/Canonical/HeaderMap.
    colKode = FPTK_DirectFindHeader( _
        wsFPTK, FPTK_HEADER, "KODE ANGKA")

    colPosisi = FPTK_DirectFindHeader( _
        wsFPTK, FPTK_HEADER, "POSISI")

    colBU = FPTK_DirectFindHeader( _
        wsFPTK, FPTK_HEADER, "BUSINESS UNIT")

    colDirektorat = FPTK_DirectFindHeader( _
        wsFPTK, FPTK_HEADER, "DIREKTORAT")

    colLokasi = FPTK_DirectFindHeader( _
        wsFPTK, FPTK_HEADER, "LOKASI KERJA")

    colUser = FPTK_DirectFindHeader( _
        wsFPTK, FPTK_HEADER, "USER MANAGER")

    colIndirect = FPTK_DirectFindHeader( _
        wsFPTK, FPTK_HEADER, "INDIRECT USER")

    If colPosisi = 0 Or _
       colUser = 0 Or _
       colIndirect = 0 Then

        MsgBox _
            "Kolom Posisi/User/Indirect User di FPTK tidak ditemukan.", _
            vbCritical, _
            "Refresh User FPTK"

        Exit Sub
    End If

    lastFPTKRow = wsFPTK.Cells( _
        wsFPTK.rows.Count, colPosisi).End(xlUp).Row

    lastDBRow = wsDB.Cells( _
        wsDB.rows.Count, DB_COL_POSITION).End(xlUp).Row

    If Not wsMap Is Nothing Then
        lastMapRow = wsMap.Cells( _
            wsMap.rows.Count, 1).End(xlUp).Row
    End If

    For fptkRow = FPTK_FIRST_ROW To lastFPTKRow

        posisiKey = FPTK_DirectNorm( _
            wsFPTK.Cells(fptkRow, colPosisi).Value2)

        If Len(posisiKey) = 0 Then
            GoTo NextFPTKRow
        End If

        buKey = ""
        lokasiKey = ""
        direktoratKey = ""

        If colBU > 0 Then
            buKey = FPTK_DirectNorm( _
                wsFPTK.Cells(fptkRow, colBU).Value2)
        End If

        If colLokasi > 0 Then
            lokasiKey = FPTK_DirectNorm( _
                wsFPTK.Cells(fptkRow, colLokasi).Value2)
        End If

        If colDirektorat > 0 Then
            direktoratKey = FPTK_DirectNorm( _
                wsFPTK.Cells(fptkRow, colDirektorat).Value2)
        End If

        bestRow = 0
        bestScore = -1
        bestCodeRank = -1

        bestUserRow = 0
        bestUserScore = -1
        bestUserCodeRank = -1

        For dbRow = 2 To lastDBRow

            dbPosisiKey = FPTK_DirectNorm( _
                wsDB.Cells(dbRow, DB_COL_POSITION).Value2)

            If dbPosisiKey = posisiKey Then

                score = 1000

                dbBUKey = FPTK_DirectNorm( _
                    wsDB.Cells(dbRow, DB_COL_BUSINESS_UNIT).Value2)

                dbLokasiKey = FPTK_DirectNorm( _
                    wsDB.Cells(dbRow, DB_COL_LOCATION).Value2)

                dbDirektoratKey = FPTK_DirectNorm( _
                    wsDB.Cells(dbRow, DB_COL_DIRECTORATE).Value2)

                If Len(buKey) > 0 And dbBUKey = buKey Then
                    score = score + 400
                End If

                If Len(lokasiKey) > 0 Then
                    If dbLokasiKey = lokasiKey Then
                        score = score + 250
                    ElseIf InStr(1, lokasiKey, dbLokasiKey, vbTextCompare) > 0 _
                        Or InStr(1, dbLokasiKey, lokasiKey, vbTextCompare) > 0 Then
                        score = score + 150
                    End If
                End If

                If Len(direktoratKey) > 0 _
                   And dbDirektoratKey = direktoratKey Then
                    score = score + 100
                End If

                codeRank = FPTK_DirectCodeRank( _
                    wsDB.Cells(dbRow, DB_COL_KODE).Value2)

                If score > bestScore _
                   Or (score = bestScore And codeRank > bestCodeRank) Then

                    bestScore = score
                    bestCodeRank = codeRank
                    bestRow = dbRow

                End If

                userText = Trim$(CStr( _
                    wsDB.Cells(dbRow, DB_COL_USER).Value2))

                indirectText = Trim$(CStr( _
                    wsDB.Cells(dbRow, DB_COL_INDIRECT).Value2))

                If Len(userText) > 0 Or Len(indirectText) > 0 Then

                    userScore = score
                    userCodeRank = codeRank

                    If userScore > bestUserScore _
                       Or (userScore = bestUserScore _
                           And userCodeRank > bestUserCodeRank) Then

                        bestUserScore = userScore
                        bestUserCodeRank = userCodeRank
                        bestUserRow = dbRow

                    End If

                End If

            End If

        Next dbRow

        userText = ""
        indirectText = ""

        If bestRow > 0 Then

            matchedCount = matchedCount + 1

            If colKode > 0 Then
                wsFPTK.Cells(fptkRow, colKode).Value = _
                    wsDB.Cells(bestRow, DB_COL_KODE).Value
            End If

            ' Directorate TIDAK BOLEH di-overwrite dari DB Kode Posisi.
            ' Nilai Directorate di FPTK selalu berasal dari file input/source.

            If bestUserRow > 0 Then

                userText = Trim$(CStr( _
                    wsDB.Cells(bestUserRow, DB_COL_USER).Value2))

                indirectText = Trim$(CStr( _
                    wsDB.Cells(bestUserRow, DB_COL_INDIRECT).Value2))

            End If

        Else

            unmatchedCount = unmatchedCount + 1

        End If

        ' ==================================================
        ' Fallback Mapping User A:D.
        '
        ' Jangan ambil row terakhir secara buta karena bisa
        ' ada duplicate Posisi dengan User/Indirect kosong.
        '
        ' Prioritas:
        ' 1. Posisi exact.
        ' 2. Directorate sama.
        ' 3. User + Indirect sama-sama terisi.
        ' ==================================================
        If Not wsMap Is Nothing Then

            If Len(userText) = 0 Or _
               Len(indirectText) = 0 Then

                bestMapRow = 0
                bestMapScore = -1

                For mapRow = 2 To lastMapRow

                    If FPTK_DirectNorm( _
                        wsMap.Cells( _
                            mapRow, _
                            1).Value2) = _
                       posisiKey Then

                        mapUserText = _
                            Trim$( _
                                CStr( _
                                    wsMap.Cells( _
                                        mapRow, _
                                        2).Value2))

                        mapIndirectText = _
                            Trim$( _
                                CStr( _
                                    wsMap.Cells( _
                                        mapRow, _
                                        3).Value2))

                        If Len(mapUserText) > 0 Or _
                           Len(mapIndirectText) > 0 Then

                            mapScore = 1000

                            mapDirectorateKey = _
                                FPTK_DirectNorm( _
                                    wsMap.Cells( _
                                        mapRow, _
                                        4).Value2)

                            If Len(direktoratKey) > 0 And _
                               mapDirectorateKey = _
                               direktoratKey Then

                                mapScore = _
                                    mapScore + 300

                            End If

                            If Len(mapUserText) > 0 And _
                               Len(mapIndirectText) > 0 Then

                                mapScore = _
                                    mapScore + 100

                            End If

                            If mapScore > bestMapScore Then

                                bestMapScore = _
                                    mapScore

                                bestMapRow = _
                                    mapRow

                            End If

                        End If

                    End If

                Next mapRow

                If bestMapRow > 0 Then

                    If Len(userText) = 0 Then

                        userText = _
                            Trim$( _
                                CStr( _
                                    wsMap.Cells( _
                                        bestMapRow, _
                                        2).Value2))

                    End If

                    If Len(indirectText) = 0 Then

                        indirectText = _
                            Trim$( _
                                CStr( _
                                    wsMap.Cells( _
                                        bestMapRow, _
                                        3).Value2))

                    End If

                    mappingFallbackCount = _
                        mappingFallbackCount + 1

                End If

            End If

        End If

        ' TULIS LANGSUNG KE FPTK.
        wsFPTK.Cells(fptkRow, colUser).Value = userText
        wsFPTK.Cells(fptkRow, colIndirect).Value = indirectText

        If Len(userText) > 0 Or Len(indirectText) > 0 Then
            writtenUserCount = writtenUserCount + 1
        End If

NextFPTKRow:

    Next fptkRow

    AddRunLog _
        "Info", _
        "DIRECT USER REFRESH | Matched DB: " & _
        CStr(matchedCount) & _
        " | Written: " & CStr(writtenUserCount) & _
        " | Mapping fallback: " & CStr(mappingFallbackCount) & _
        " | Unmatched: " & CStr(unmatchedCount)

End Sub

Private Function FPTK_FindDBKodePosisiColumn( _
    ByVal ws As Worksheet, _
    ByVal aliases As Variant, _
    Optional ByVal fallbackColumn As Long = 0) As Long

    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim aliasItem As Variant
    Dim headerKey As String

    If ws Is Nothing Then Exit Function

    lastColumn = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End( _
                xlToLeft).Column

    For columnIndex = 1 To lastColumn

        headerKey = _
            FPTK_NormalizeLookupText( _
                ws.Cells( _
                    1, _
                    columnIndex).Value2)

        For Each aliasItem In aliases

            If headerKey = _
               FPTK_NormalizeLookupText( _
                    aliasItem) Then

                FPTK_FindDBKodePosisiColumn = _
                    columnIndex

                Exit Function

            End If

        Next aliasItem

    Next columnIndex

    If fallbackColumn > 0 And _
       fallbackColumn <= ws.Columns.Count Then

        FPTK_FindDBKodePosisiColumn = _
            fallbackColumn

    End If

End Function

Private Function FPTK_LastUsedRowInDBKodePosisi( _
    ByVal ws As Worksheet, _
    ByVal codeColumn As Long, _
    ByVal positionColumn As Long) As Long

    Dim lastCodeRow As Long
    Dim lastPositionRow As Long

    If ws Is Nothing Then Exit Function

    If codeColumn > 0 Then

        lastCodeRow = _
            ws.Cells( _
                ws.rows.Count, _
                codeColumn).End( _
                    xlUp).Row

    End If

    If positionColumn > 0 Then

        lastPositionRow = _
            ws.Cells( _
                ws.rows.Count, _
                positionColumn).End( _
                    xlUp).Row

    End If

    If lastCodeRow > lastPositionRow Then

        FPTK_LastUsedRowInDBKodePosisi = _
            lastCodeRow

    Else

        FPTK_LastUsedRowInDBKodePosisi = _
            lastPositionRow

    End If

End Function
Private Function FPTK_NormalizePositionCodeKey( _
    ByVal rawValue As Variant) As String

    Dim sourceText As String
    Dim digitText As String
    Dim alphaNumericText As String

    Dim characterIndex As Long
    Dim currentCharacter As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    sourceText = _
        UCase$( _
            Trim$( _
                CStr(rawValue)))

    For characterIndex = 1 To Len(sourceText)

        currentCharacter = _
            Mid$( _
                sourceText, _
                characterIndex, _
                1)

        If currentCharacter >= "0" And _
           currentCharacter <= "9" Then

            digitText = _
                digitText & _
                currentCharacter

        End If

        If currentCharacter Like _
           "[A-Z0-9]" Then

            alphaNumericText = _
                alphaNumericText & _
                currentCharacter

        End If

    Next characterIndex

    If Len(digitText) > 0 Then

        On Error Resume Next

        FPTK_NormalizePositionCodeKey = _
            CStr( _
                CLng( _
                    CDbl(digitText)))

        On Error GoTo 0

        If Len( _
            FPTK_NormalizePositionCodeKey) > 0 Then

            Exit Function

        End If

    End If

    FPTK_NormalizePositionCodeKey = _
        alphaNumericText

End Function
Private Function FPTK_NormalizeLookupText( _
    ByVal rawValue As Variant) As String

    Dim sourceText As String
    Dim resultText As String

    Dim characterIndex As Long
    Dim currentCharacter As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    sourceText = _
        UCase$( _
            Trim$( _
                CStr(rawValue)))

    For characterIndex = 1 To Len(sourceText)

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

    FPTK_NormalizeLookupText = _
        resultText

End Function

Private Function FPTK_SafeCellText( _
    ByVal rawValue As Variant) As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    FPTK_SafeCellText = _
        Trim$( _
            CStr(rawValue))

End Function
Public Sub RefreshFPTKUsersNow( _
    Optional ByVal showMessage As Boolean = True)

    Dim ws As Worksheet
    Dim headerMap As Object

    Dim colPosisi As Long
    Dim colUser As Long
    Dim colIndirect As Long

    Dim lastRow As Long
    Dim rowIndex As Long

    Dim positionCount As Long
    Dim userFilledCount As Long
    Dim stillBlankCount As Long

    Set ws = _
        ThisWorkbook.Worksheets( _
            "FPTK")

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    RefreshFPTKUsersFromDBKodePosisi _
        ws, _
        headerMap

    RefreshFPTKReasonValues _
        ws, _
        headerMap

    ' Final STO guard:
    ' post-master refresh tidak boleh meninggalkan Filter STO blank.
    RefreshFPTKSTOCategorization _
        ws, _
        headerMap, _
        True

    colPosisi = _
        FPTK_DirectFindHeader( _
            ws, _
            FPTK_HEADER_ROW, _
            "Posisi")

    colUser = _
        FPTK_DirectFindHeader( _
            ws, _
            FPTK_HEADER_ROW, _
            "User (Manager)")

    colIndirect = _
        FPTK_DirectFindHeader( _
            ws, _
            FPTK_HEADER_ROW, _
            "Indirect User")

    If colPosisi > 0 Then

        lastRow = _
            ws.Cells( _
                ws.rows.Count, _
                colPosisi).End( _
                    xlUp).Row

        For rowIndex = FPTK_FIRST_DATA_ROW To lastRow

            If Len(Trim$(CStr( _
                ws.Cells( _
                    rowIndex, _
                    colPosisi).Value2))) > 0 Then

                positionCount = _
                    positionCount + 1

                If Len(Trim$(CStr( _
                    ws.Cells( _
                        rowIndex, _
                        colUser).Value2))) > 0 And _
                   Len(Trim$(CStr( _
                    ws.Cells( _
                        rowIndex, _
                        colIndirect).Value2))) > 0 Then

                    userFilledCount = _
                        userFilledCount + 1

                Else

                    stillBlankCount = _
                        stillBlankCount + 1

                End If

            End If

        Next rowIndex

    End If

    AddRunLog _
        "Info", _
        "POST DB USER REFRESH | Posisi: " & _
        CStr(positionCount) & _
        " | Full User+Indirect: " & _
        CStr(userFilledCount) & _
        " | Still blank: " & _
        CStr(stillBlankCount)

    If showMessage Then

        MsgBox _
            "Refresh User selesai." & _
            vbCrLf & vbCrLf & _
            "Row Posisi: " & _
            CStr(positionCount) & _
            vbCrLf & _
            "User + Indirect terisi lengkap: " & _
            CStr(userFilledCount) & _
            vbCrLf & _
            "Masih ada yang kosong: " & _
            CStr(stillBlankCount), _
            vbInformation, _
            "Refresh User FPTK"

    End If

    Set headerMap = Nothing
    Set ws = Nothing

End Sub

Public Sub RefreshFPTKUsersAfterMaster()

    ' Dipanggil Main Compiler SETELAH
    ' Compile_DB_Kode_Posisi_MultiFolders selesai.
    ' Refresh hanya Kode Angka/User/Indirect.
    ' Directorate tetap dari file input/source.
    RefreshFPTKUsersNow False

End Sub

Private Sub RefreshFPTKStatusValues( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object)

    Dim colStatus As Long
    Dim colCancelDate As Long
    Dim colOfferingDate As Long

    Dim lastRow As Long
    Dim rowCount As Long
    Dim rowIndex As Long

    Dim statusData As Variant
    Dim cancelData As Variant
    Dim offeringData As Variant

    Dim normalizedStatus As String
    Dim cancelDate As Date
    Dim offeringDate As Date

    If ws Is Nothing Then Exit Sub
    If headerMap Is Nothing Then Exit Sub

    colStatus = _
        ColumnByKey( _
            headerMap, _
            "STATUS")

    colCancelDate = _
        ColumnByKey( _
            headerMap, _
            "FPTK CANCEL DATE")

    colOfferingDate = _
        ColumnByKey( _
            headerMap, _
            "OFFERING DATE")

    If colStatus <= 0 Then Exit Sub

    lastRow = _
        LastTargetDataRow( _
            ws, _
            headerMap)

    If lastRow < FPTK_FIRST_DATA_ROW Then
        Exit Sub
    End If

    rowCount = _
        lastRow - _
        FPTK_FIRST_DATA_ROW + 1

    statusData = _
        FPTK_ReadColumnArray( _
            ws, _
            colStatus, _
            rowCount)

    cancelData = _
        FPTK_ReadColumnArray( _
            ws, _
            colCancelDate, _
            rowCount)

    offeringData = _
        FPTK_ReadColumnArray( _
            ws, _
            colOfferingDate, _
            rowCount)

    For rowIndex = 1 To rowCount

        normalizedStatus = _
            FPTK_NormalizeLookupText( _
                statusData( _
                    rowIndex, _
                    1))

        Select Case normalizedStatus

            Case "OP", "OPEN"
                statusData(rowIndex, 1) = "OP"

            Case "CLOSED", "CLOSE"
                statusData(rowIndex, 1) = "Closed"

            Case "CANCEL", _
                 "CANCELLED", _
                 "CANCELED"

                statusData(rowIndex, 1) = "Cancel"

            Case Else

                cancelDate = _
                    ParseFlexibleDate( _
                        cancelData( _
                            rowIndex, _
                            1))

                offeringDate = _
                    ParseFlexibleDate( _
                        offeringData( _
                            rowIndex, _
                            1))

                If cancelDate > 0 Then

                    statusData(rowIndex, 1) = _
                        "Cancel"

                ElseIf offeringDate > 0 Then

                    statusData(rowIndex, 1) = _
                        "Closed"

                Else

                    statusData(rowIndex, 1) = _
                        "OP"

                End If

        End Select

    Next rowIndex

    ws.Cells( _
        FPTK_FIRST_DATA_ROW, _
        colStatus).Resize( _
            rowCount, _
            1).Value2 = _
                statusData

    FPTK_ApplyStatusValidation _
        ws

End Sub

Private Sub FPTK_ApplyStatusValidation( _
    ByVal ws As Worksheet)

    Dim headerMap As Object
    Dim statusColumn As Long
    Dim lastRow As Long
    Dim validationRange As Range

    If ws Is Nothing Then Exit Sub

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    statusColumn = _
        ColumnByKey( _
            headerMap, _
            "STATUS")

    If statusColumn <= 0 Then
        GoTo SafeExit
    End If

    lastRow = _
        FPTK_LastUsedDataRowAny( _
            ws)

    If lastRow < FPTK_FIRST_DATA_ROW Then
        lastRow = FPTK_FIRST_DATA_ROW
    End If

    Set validationRange = _
        ws.Range( _
            ws.Cells( _
                FPTK_FIRST_DATA_ROW, _
                statusColumn), _
            ws.Cells( _
                lastRow + 1000, _
                statusColumn))

    On Error Resume Next

    validationRange.Validation.Delete

    validationRange.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="OP,Closed,Cancel"

    validationRange.Validation.IgnoreBlank = _
        True

    validationRange.Validation.InCellDropdown = _
        True

    validationRange.Validation.ShowError = _
        True

    validationRange.Validation.ErrorTitle = _
        "Status tidak valid"

    validationRange.Validation.ErrorMessage = _
        "Pilih OP, Closed, atau Cancel."

    On Error GoTo 0

SafeExit:

    Set validationRange = Nothing
    Set headerMap = Nothing

End Sub

Public Sub RefreshFPTKStatusNow()

    Dim ws As Worksheet
    Dim headerMap As Object

    Set ws = _
        GetOrCreateWorksheet( _
            "FPTK")

    ProtectFPTKHeader

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    RefreshFPTKStatusValues _
        ws, _
        headerMap

    Set headerMap = Nothing
    Set ws = Nothing

    MsgBox _
        "Status FPTK sudah dinormalisasi menjadi " & _
        "OP, Closed, atau Cancel.", _
        vbInformation, _
        "Refresh Status FPTK"

End Sub


Private Sub RefreshFPTKKodeBUValues( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object)

    Dim colKodePIC As Long
    Dim colKodeBU As Long

    Dim lastRow As Long
    Dim rowCount As Long
    Dim rowIndex As Long

    Dim kodePICData As Variant
    Dim kodeBUData() As Variant

    If ws Is Nothing Then Exit Sub
    If headerMap Is Nothing Then Exit Sub

    colKodePIC = _
        ColumnByKey( _
            headerMap, _
            "KODE PIC")

    colKodeBU = _
        ColumnByKey( _
            headerMap, _
            "KODE BU")

    If colKodePIC <= 0 Or _
       colKodeBU <= 0 Then

        AddRunLog _
            "Warning", _
            "FPTK Kode BU refresh skipped | " & _
            "Kode PIC / Kode BU column missing."

        Exit Sub

    End If

    lastRow = _
        LastTargetDataRow( _
            ws, _
            headerMap)

    If lastRow < FPTK_FIRST_DATA_ROW Then
        Exit Sub
    End If

    rowCount = _
        lastRow - _
        FPTK_FIRST_DATA_ROW + 1

    kodePICData = _
        ws.Cells( _
            FPTK_FIRST_DATA_ROW, _
            colKodePIC).Resize( _
                rowCount, _
                1).Value2

    ReDim kodeBUData( _
        1 To rowCount, _
        1 To 1)

    If rowCount = 1 Then

        kodeBUData(1, 1) = _
            FPTK_GetBUCodeFromPIC( _
                kodePICData)

    Else

        For rowIndex = 1 To rowCount

            kodeBUData( _
                rowIndex, _
                1) = _
                    FPTK_GetBUCodeFromPIC( _
                        kodePICData( _
                            rowIndex, _
                            1))

        Next rowIndex

    End If

    ws.Cells( _
        FPTK_FIRST_DATA_ROW, _
        colKodeBU).Resize( _
            rowCount, _
            1).Value2 = _
                kodeBUData

    AddRunLog _
        "Info", _
        "FPTK Kode BU refreshed from Kode PIC | Rows: " & _
        CStr(rowCount)

End Sub

Private Function FPTK_GetBUCodeFromPIC( _
    ByVal rawKodePIC As Variant) As String

    Dim kodeText As String

    If IsError(rawKodePIC) Or _
       IsNull(rawKodePIC) Or _
       IsEmpty(rawKodePIC) Then

        Exit Function

    End If

    kodeText = _
        UCase$( _
            Trim$( _
                CStr(rawKodePIC)))

    kodeText = _
        Replace$( _
            kodeText, _
            " ", _
            vbNullString)

    kodeText = _
        Replace$( _
            kodeText, _
            "-", _
            vbNullString)

    kodeText = _
        Replace$( _
            kodeText, _
            "_", _
            vbNullString)

    kodeText = _
        Replace$( _
            kodeText, _
            ".", _
            vbNullString)

    Select Case True

        Case Left$(kodeText, 4) = "CORP"
            FPTK_GetBUCodeFromPIC = "HO"

        Case Left$(kodeText, 4) = "JESS"
            FPTK_GetBUCodeFromPIC = "JESS"

        Case Left$(kodeText, 3) = "CMD"
            FPTK_GetBUCodeFromPIC = "CMD"

        Case Left$(kodeText, 3) = "BHC"
            FPTK_GetBUCodeFromPIC = "BHC"

        ' Alias lama BCH tetap dinormalisasi menjadi BHC.
        Case Left$(kodeText, 3) = "BCH"
            FPTK_GetBUCodeFromPIC = "BHC"

        Case Left$(kodeText, 3) = "ARC"
            FPTK_GetBUCodeFromPIC = "ARC"

        Case Left$(kodeText, 2) = "MS"
            FPTK_GetBUCodeFromPIC = "MS"

        Case Left$(kodeText, 2) = "MP"
            FPTK_GetBUCodeFromPIC = "MP"

        Case Left$(kodeText, 2) = "MB"
            FPTK_GetBUCodeFromPIC = "MB"

        Case Left$(kodeText, 2) = "HO"
            FPTK_GetBUCodeFromPIC = "HO"

    End Select

End Function

Public Sub RefreshFPTKKodeBUNow()

    Dim ws As Worksheet
    Dim headerMap As Object

    Set ws = _
        GetOrCreateWorksheet( _
            "FPTK")

    ProtectFPTKHeader

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    RefreshFPTKKodeBUValues _
        ws, _
        headerMap

    Set headerMap = Nothing
    Set ws = Nothing

    MsgBox _
        "Kode BU sudah diperbarui dari Kode PIC.", _
        vbInformation, _
        "Refresh Kode BU"

End Sub


Private Sub RefreshFPTKSLADerivedValues( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object)

    Dim colLevelNumber As Long
    Dim colTanggalReal As Long
    Dim colJumlahSLA As Long
    Dim colDeadline As Long
    Dim colDetailSLA As Long
    Dim colStatus As Long
    Dim colOffering As Long

    Dim lastRow As Long
    Dim rowCount As Long
    Dim rowIndex As Long

    Dim levelData As Variant
    Dim realDateData As Variant
    Dim jumlahSLAData As Variant
    Dim deadlineData As Variant
    Dim detailSLAData As Variant
    Dim statusData As Variant
    Dim offeringData As Variant

    Dim slaDays As Long
    Dim parsedDate As Date
    Dim offeringValue As Variant

    If ws Is Nothing Then Exit Sub
    If headerMap Is Nothing Then Exit Sub

    colLevelNumber = _
        ColumnByKey(headerMap, "LEVEL NUMBER")

    colTanggalReal = _
        ColumnByKey(headerMap, "FPTK DATE REAL")

    colJumlahSLA = _
        ColumnByKey(headerMap, "JUMLAH SLA")

    colDeadline = _
        ColumnByKey(headerMap, "DEADLINE SLA")

    colDetailSLA = _
        ColumnByKey(headerMap, "DETAIL SLA")

    colStatus = _
        ColumnByKey(headerMap, "STATUS")

    colOffering = _
        ColumnByKey(headerMap, "OFFERING DATE")

    If colTanggalReal <= 0 Or _
       colJumlahSLA <= 0 Or _
       colDeadline <= 0 Then

        AddRunLog _
            "Warning", _
            "FPTK SLA refresh skipped | Missing column."

        Exit Sub

    End If

    lastRow = _
        LastTargetDataRow( _
            ws, _
            headerMap)

    If lastRow < FPTK_FIRST_DATA_ROW Then Exit Sub

    rowCount = _
        lastRow - _
        FPTK_FIRST_DATA_ROW + 1

    levelData = _
        FPTK_ReadColumnArray( _
            ws, _
            colLevelNumber, _
            rowCount)

    realDateData = _
        FPTK_ReadColumnArray( _
            ws, _
            colTanggalReal, _
            rowCount)

    jumlahSLAData = _
        FPTK_ReadColumnArray( _
            ws, _
            colJumlahSLA, _
            rowCount)

    deadlineData = _
        FPTK_ReadColumnArray( _
            ws, _
            colDeadline, _
            rowCount)

    detailSLAData = _
        FPTK_ReadColumnArray( _
            ws, _
            colDetailSLA, _
            rowCount)

    statusData = _
        FPTK_ReadColumnArray( _
            ws, _
            colStatus, _
            rowCount)

    offeringData = _
        FPTK_ReadColumnArray( _
            ws, _
            colOffering, _
            rowCount)

    For rowIndex = 1 To rowCount

        slaDays = 0

        If colLevelNumber > 0 Then

            If Not ValueIsBlankFPTK( _
                levelData(rowIndex, 1)) Then

                slaDays = _
                    GetSLAByLevelFPTK( _
                        levelData(rowIndex, 1))

            End If

        End If

        If slaDays <= 0 Then

            If IsNumeric( _
                jumlahSLAData(rowIndex, 1)) Then

                If CLng( _
                    jumlahSLAData(rowIndex, 1)) > 0 Then

                    slaDays = _
                        CLng( _
                            jumlahSLAData(rowIndex, 1))

                End If

            End If

        End If

        If slaDays > 0 Then

            jumlahSLAData(rowIndex, 1) = _
                slaDays

            parsedDate = _
                ParseFlexibleDate( _
                    realDateData(rowIndex, 1))

            If parsedDate > 0 Then

                deadlineData(rowIndex, 1) = _
                    DateAdd( _
                        "d", _
                        slaDays, _
                        parsedDate)

            Else

                deadlineData(rowIndex, 1) = Empty

            End If

        Else

            deadlineData(rowIndex, 1) = Empty

        End If

        If colOffering > 0 Then
            offeringValue = offeringData(rowIndex, 1)
        Else
            offeringValue = Empty
        End If

        If colDetailSLA > 0 And _
           colStatus > 0 Then

            detailSLAData(rowIndex, 1) = _
                CalculateSLAStatusFPTK( _
                    statusData(rowIndex, 1), _
                    deadlineData(rowIndex, 1), _
                    offeringValue)

        End If

    Next rowIndex

    ws.Cells( _
        FPTK_FIRST_DATA_ROW, _
        colJumlahSLA).Resize( _
            rowCount, _
            1).Value2 = _
                jumlahSLAData

    ws.Cells( _
        FPTK_FIRST_DATA_ROW, _
        colDeadline).Resize( _
            rowCount, _
            1).Value = _
                deadlineData

    ws.Columns( _
        colDeadline).numberFormat = _
            "dd/mm/yyyy"

    If colDetailSLA > 0 Then

        ws.Cells( _
            FPTK_FIRST_DATA_ROW, _
            colDetailSLA).Resize( _
                rowCount, _
                1).Value2 = _
                    detailSLAData

    End If

    AddRunLog _
        "Info", _
        "FPTK SLA refresh completed | Rows: " & _
        CStr(rowCount)

End Sub

Private Function FPTK_ArrayValueSafe( _
    ByVal sourceArray As Variant, _
    ByVal rowIndex As Long, _
    ByVal sourceColumn As Long) As Variant

    If sourceColumn <= 0 Then
        Exit Function
    End If

    If Not IsArray(sourceArray) Then
        Exit Function
    End If

    FPTK_ArrayValueSafe = _
        sourceArray( _
            rowIndex, _
            1)

End Function

Public Sub RefreshFPTKSLANow()

    Dim ws As Worksheet
    Dim headerMap As Object

    Set ws = _
        GetOrCreateWorksheet( _
            "FPTK")

    ProtectFPTKHeader

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    RefreshFPTKSLADerivedValues _
        ws, _
        headerMap

    ApplyFPTKDateFormats _
        ws, _
        headerMap

    Set headerMap = Nothing
    Set ws = Nothing

    MsgBox _
        "Jumlah SLA dan Deadline pemenuhan SLA " & _
        "sudah dihitung ulang.", _
        vbInformation, _
        "Refresh SLA FPTK"

End Sub

Private Sub RefreshFPTKWeekMonthValues( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object)

    Dim colFPTKDateKode As Long
    Dim colWeekFPTK As Long
    Dim colMonthFPTK As Long

    Dim colCancelDate As Long
    Dim colWeekCancel As Long
    Dim colMonthCancel As Long

    Dim colOfferingDate As Long
    Dim colWeekOffering As Long
    Dim colMonthOffering As Long

    Dim lastRow As Long
    Dim rowIndex As Long

    If ws Is Nothing Then Exit Sub
    If headerMap Is Nothing Then Exit Sub

    colFPTKDateKode = ColumnByKey( _
        headerMap, _
        "FPTK DATE KODE")

    colWeekFPTK = ColumnByKey( _
        headerMap, _
        "WEEK FPTK DATE")

    colMonthFPTK = ColumnByKey( _
        headerMap, _
        "MONTH FPTK DATE")

    colCancelDate = ColumnByKey( _
        headerMap, _
        "FPTK CANCEL DATE")

    colWeekCancel = ColumnByKey( _
        headerMap, _
        "WEEK CANCEL DATE")

    colMonthCancel = ColumnByKey( _
        headerMap, _
        "MONTH CANCEL DATE")

    colOfferingDate = ColumnByKey( _
        headerMap, _
        "OFFERING DATE")

    colWeekOffering = ColumnByKey( _
        headerMap, _
        "WEEK OFFERING DATE")

    colMonthOffering = ColumnByKey( _
        headerMap, _
        "MONTH OFFERING")

    lastRow = LastTargetDataRow( _
        ws, _
        headerMap)

    If lastRow < FPTK_FIRST_DATA_ROW Then
        Exit Sub
    End If

    For rowIndex = FPTK_FIRST_DATA_ROW _
        To lastRow

        If RowHasPrimaryData( _
            ws, _
            rowIndex, _
            headerMap) Then

            ' Week FPTK + Month FPTK memakai
            ' FPTK Date (Kode), sesuai sistem sebelumnya.
            WriteFPTKWeekMonthValue _
                ws, _
                rowIndex, _
                colFPTKDateKode, _
                colWeekFPTK, _
                colMonthFPTK

            WriteFPTKWeekMonthValue _
                ws, _
                rowIndex, _
                colCancelDate, _
                colWeekCancel, _
                colMonthCancel

            WriteFPTKWeekMonthValue _
                ws, _
                rowIndex, _
                colOfferingDate, _
                colWeekOffering, _
                colMonthOffering

        End If

    Next rowIndex

End Sub
Private Sub WriteFPTKWeekMonthValue( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal sourceDateColumn As Long, _
    ByVal weekTargetColumn As Long, _
    ByVal monthTargetColumn As Long)

    Dim parsedDate As Date
    Dim weekNumber As Long

    If sourceDateColumn <= 0 Then Exit Sub

    If IsDateValueFlexible( _
        ws.Cells( _
            rowIndex, _
            sourceDateColumn).Value2) Then

        parsedDate = ParseFlexibleDate( _
            ws.Cells( _
                rowIndex, _
                sourceDateColumn).Value2)

        If weekTargetColumn > 0 Then

            On Error Resume Next

            weekNumber = _
                Application.WorksheetFunction. _
                    WeekNum(parsedDate, 1)

            If Err.Number <> 0 Then

                Err.Clear

                weekNumber = DatePart( _
                    "ww", _
                    parsedDate, _
                    vbSunday, _
                    vbFirstJan1)

            End If

            On Error GoTo 0

            ws.Cells( _
                rowIndex, _
                weekTargetColumn).Value2 = _
                weekNumber

        End If

        If monthTargetColumn > 0 Then

            ws.Cells( _
                rowIndex, _
                monthTargetColumn).Value2 = _
                Format$(parsedDate, "mmmm")

        End If

    Else

        If weekTargetColumn > 0 Then

            ws.Cells( _
                rowIndex, _
                weekTargetColumn).ClearContents

        End If

        If monthTargetColumn > 0 Then

            ws.Cells( _
                rowIndex, _
                monthTargetColumn).ClearContents

        End If

    End If

End Sub
Private Sub RefreshFilterKategorisasiFPTK( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object)

    Dim colKodeUnik As Long
    Dim colPosisi As Long
    Dim colLevelNumber As Long
    Dim colFilter As Long
    Dim lastRow As Long

    Dim kodeOffset As Long
    Dim posisiOffset As Long
    Dim levelOffset As Long

    Dim formulaText As String
    Dim targetRange As Range

    If ws Is Nothing Then Exit Sub
    If headerMap Is Nothing Then Exit Sub

    colKodeUnik = ColumnByKey( _
        headerMap, _
        "KODE UNIK")

    colPosisi = ColumnByKey( _
        headerMap, _
        "POSISI")

    colLevelNumber = ColumnByKey( _
        headerMap, _
        "LEVEL NUMBER")

    colFilter = ColumnByKey( _
        headerMap, _
        "FILTER KATEGORISASI FPTK")

    If colKodeUnik = 0 Then Exit Sub
    If colPosisi = 0 Then Exit Sub
    If colLevelNumber = 0 Then Exit Sub
    If colFilter = 0 Then Exit Sub

    lastRow = LastTargetDataRow( _
        ws, _
        headerMap)

    If lastRow < FPTK_FIRST_DATA_ROW Then
        Exit Sub
    End If

    kodeOffset = colKodeUnik - colFilter
    posisiOffset = colPosisi - colFilter
    levelOffset = colLevelNumber - colFilter

    formulaText = _
        "=IF(UPPER(TRIM(RC))=""STO"",""STO"",IF(" & _
        "RC[" & CStr(kodeOffset) & "]="""","""",IF(OR("

    formulaText = formulaText & _
        "LEFT(LOWER(TRIM(RC[" & _
        CStr(posisiOffset) & _
        "])),6)=""cimory"","

    formulaText = formulaText & _
        "LEFT(LOWER(TRIM(RC[" & _
        CStr(posisiOffset) & _
        "])),5)=""fresh""),"

    formulaText = formulaText & _
        """CLAP FGDP"",IF(OR("

    formulaText = formulaText & _
        "RC[" & CStr(levelOffset) & "]=1,"

    formulaText = formulaText & _
        "RC[" & CStr(levelOffset) & "]=2),"

    formulaText = formulaText & _
        """Level 1-2"",IF("

    formulaText = formulaText & _
        "RC[" & CStr(levelOffset) & "]=3,"

    formulaText = formulaText & _
        """Level 3"",IF("

    formulaText = formulaText & _
        "RC[" & CStr(levelOffset) & "]=4,"

    formulaText = formulaText & _
        """Level 4"",""""))))))"

    Set targetRange = ws.Range( _
        ws.Cells( _
            FPTK_FIRST_DATA_ROW, _
            colFilter), _
        ws.Cells( _
            lastRow, _
            colFilter))

    targetRange.FormulaR1C1 = _
        formulaText

    Set targetRange = Nothing

End Sub

Private Function GetSLAByLevelFPTK( _
    ByVal levelValue As Variant) As Long

    Dim levelNumber As Long

    If IsNumeric(levelValue) Then
        levelNumber = CLng(levelValue)
    Else
        levelNumber = ExtractFirstNumberLocal( _
            CStr(levelValue))
    End If

    Select Case levelNumber

        Case 1 To 3
            GetSLAByLevelFPTK = 30

        Case 4
            GetSLAByLevelFPTK = 45

        Case 5
            GetSLAByLevelFPTK = 60

        Case Else
            GetSLAByLevelFPTK = 30

    End Select

End Function

Private Function CalculateSLAStatusFPTK( _
    ByVal statusValue As Variant, _
    ByVal deadlineValue As Variant, _
    ByVal offeringValue As Variant) As String

    Dim statusText As String
    Dim deadlineDate As Date
    Dim offeringDate As Date

    statusText = LCase$( _
        Trim$(CStr(statusValue)))

    If IsDateValueFlexible(deadlineValue) Then
        deadlineDate = ParseFlexibleDate(deadlineValue)
    End If

    If IsDateValueFlexible(offeringValue) Then
        offeringDate = ParseFlexibleDate(offeringValue)
    End If

    Select Case statusText

        Case "op", "open"

            If deadlineDate > 0 And _
               deadlineDate < Date Then

                CalculateSLAStatusFPTK = _
                    "OP tidak lulus SLA"

            Else

                CalculateSLAStatusFPTK = _
                    "OP belum lewat SLA"

            End If

        Case "closed", "close"

            If deadlineDate > 0 And _
               offeringDate > 0 And _
               deadlineDate < offeringDate Then

                CalculateSLAStatusFPTK = _
                    "Closed tidak lulus SLA"

            Else

                CalculateSLAStatusFPTK = _
                    "Closed lulus SLA"

            End If

        Case "cancel", "cancel fptk", "cancelled"

            CalculateSLAStatusFPTK = _
                "Cancel FPTK"

        Case Else

            If deadlineDate > 0 And _
               deadlineDate < Date Then

                CalculateSLAStatusFPTK = _
                    "OP tidak lulus SLA"

            Else

                CalculateSLAStatusFPTK = _
                    "OP belum lewat SLA"

            End If

    End Select

End Function

' ============================================================
' HEADER MAP
' ============================================================
Private Function BuildHeaderMap( _
    ByVal ws As Object, _
    ByVal headerRow As Long) As Object

    Dim result As Object
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim headerKey As String

    Set result = CreateObject( _
        "Scripting.Dictionary")

    result.CompareMode = vbTextCompare

    lastColumn = LastHeaderColumn( _
        ws, headerRow)

    For columnIndex = 1 To lastColumn

        headerKey = CanonicalHeaderKey( _
            ws.Cells( _
                headerRow, columnIndex).Value2)

        If Len(headerKey) > 0 Then

            If Not result.Exists(headerKey) Then
                result.Add headerKey, columnIndex
            End If

        End If

    Next columnIndex

    Set BuildHeaderMap = result

End Function

Private Function BuildColumnKeyArray( _
    ByVal ws As Object, _
    ByVal headerRow As Long, _
    ByVal lastColumn As Long) As Variant

    Dim result() As String
    Dim columnIndex As Long

    ReDim result(1 To lastColumn)

    For columnIndex = 1 To lastColumn

        result(columnIndex) = CanonicalHeaderKey( _
            ws.Cells( _
                headerRow, columnIndex).Value2)

    Next columnIndex

    BuildColumnKeyArray = result

End Function

Private Function ColumnByKey( _
    ByVal headerMap As Object, _
    ByVal headerKey As String) As Long

    Dim canonicalKey As String

    If headerMap Is Nothing Then Exit Function

    canonicalKey = CanonicalHeaderKey(headerKey)

    If headerMap.Exists(canonicalKey) Then
        ColumnByKey = CLng(headerMap(canonicalKey))
    End If

End Function

' ============================================================
' ENSURE TARGET HEADER
' ============================================================
Private Sub EnsureRequiredHeaders( _
    ByVal ws As Worksheet)

    Dim requiredHeaders As Variant
    Dim headerItem As Variant
    Dim headerMap As Object

    Dim headerText As String
    Dim headerKey As String

    Dim existingColumn As Long
    Dim newColumn As Long

    requiredHeaders = _
        RequiredFPTKHeaders()

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    For Each headerItem In requiredHeaders

        headerText = _
            Trim$( _
                CStr(headerItem))

        headerKey = _
            CanonicalHeaderKey( _
                headerText)

        If Len(headerKey) > 0 Then

            If Not headerMap.Exists( _
                headerKey) Then

                existingColumn = _
                    FPTK_FindEquivalentTargetColumn( _
                        ws, _
                        headerText, _
                        headerKey)

                If existingColumn > 0 Then

                    headerMap.Add _
                        headerKey, _
                        existingColumn

                Else

                    newColumn = _
                        LastHeaderColumn( _
                            ws, _
                            FPTK_HEADER_ROW) + 1

                    If newColumn < 1 Then
                        newColumn = 1
                    End If

                    ws.Cells( _
                        FPTK_HEADER_ROW, _
                        newColumn).Value = _
                            headerText

                    headerMap.Add _
                        headerKey, _
                        newColumn

                End If

            End If

        End If

    Next headerItem

End Sub
Private Sub EnsureTargetHeadersFromSource( _
    ByVal wsTarget As Worksheet, _
    ByVal targetMap As Object, _
    ByVal wsSource As Object, _
    ByVal sourceHeaderRow As Long, _
    ByVal sourceLastColumn As Long)

    Dim sourceColumn As Long
    Dim sourceHeader As String
    Dim headerKey As String

    Dim existingColumn As Long
    Dim newColumn As Long

    For sourceColumn = 1 _
        To sourceLastColumn

        sourceHeader = _
            FPTK_SafeHeaderText( _
                wsSource.Cells( _
                    sourceHeaderRow, _
                    sourceColumn).Value2)

        If Len(sourceHeader) > 0 Then

            If Not FPTK_IsDeprecatedSourceHeader( _
                sourceHeader) Then

                headerKey = _
                    CanonicalHeaderKey( _
                        sourceHeader)

                If Len(headerKey) > 0 Then

                    If Not targetMap.Exists( _
                        headerKey) Then

                        existingColumn = _
                            FPTK_FindEquivalentTargetColumn( _
                                wsTarget, _
                                sourceHeader, _
                                headerKey)

                        If existingColumn > 0 Then

                            targetMap.Add _
                                headerKey, _
                                existingColumn

                        Else

                            newColumn = _
                                LastHeaderColumn( _
                                    wsTarget, _
                                    FPTK_HEADER_ROW) + 1

                            wsTarget.Cells( _
                                FPTK_HEADER_ROW, _
                                newColumn).Value = _
                                    sourceHeader

                            targetMap.Add _
                                headerKey, _
                                newColumn

                        End If

                    End If

                End If

            End If

        End If

    Next sourceColumn

End Sub

' ============================================================
' FIND EQUIVALENT FPTK TARGET COLUMN
'
' Prioritas pencarian:
' 1. CanonicalHeaderKey sama.
' 2. Header tanpa catatan instruksi sama.
'
' Contoh yang dianggap satu field:
' MODEL
' MODEL (LIHAT DROPDOWN)
' MODEL - LIHAT DROPDOWN
'
' Header baru hanya dibuat jika benar-benar tidak ada
' kolom tujuan yang ekuivalen.
' ============================================================
Private Function FPTK_FindEquivalentTargetColumn( _
    ByVal wsTarget As Worksheet, _
    ByVal sourceHeader As String, _
    ByVal sourceCanonical As String) As Long

    Dim lastColumn As Long
    Dim columnIndex As Long

    Dim targetHeader As String
    Dim targetCanonical As String

    Dim sourceBaseKey As String
    Dim targetBaseKey As String

    If wsTarget Is Nothing Then
        Exit Function
    End If

    lastColumn = _
        LastHeaderColumn( _
            wsTarget, _
            FPTK_HEADER_ROW)

    sourceBaseKey = _
        FPTK_HeaderInstructionBaseKey( _
            sourceHeader)

    For columnIndex = 1 _
        To lastColumn

        targetHeader = _
            FPTK_SafeHeaderText( _
                wsTarget.Cells( _
                    FPTK_HEADER_ROW, _
                    columnIndex).Value2)

        If Len(targetHeader) > 0 Then

            targetCanonical = _
                CanonicalHeaderKey( _
                    targetHeader)

            If Len(targetCanonical) > 0 Then

                If StrComp( _
                    targetCanonical, _
                    sourceCanonical, _
                    vbTextCompare) = 0 Then

                    FPTK_FindEquivalentTargetColumn = _
                        columnIndex

                    Exit Function

                End If

            End If

            targetBaseKey = _
                FPTK_HeaderInstructionBaseKey( _
                    targetHeader)

            If Len(sourceBaseKey) > 0 Then

                If sourceBaseKey = _
                   targetBaseKey Then

                    FPTK_FindEquivalentTargetColumn = _
                        columnIndex

                    Exit Function

                End If

            End If

        End If

    Next columnIndex

End Function

Private Function FPTK_HeaderInstructionBaseKey( _
    ByVal rawHeader As String) As String

    Dim baseText As String

    baseText = _
        FPTK_HeaderInstructionBaseText( _
            rawHeader)

    FPTK_HeaderInstructionBaseKey = _
        Replace$( _
            HeaderKeyText( _
                baseText), _
            " ", _
            vbNullString)

End Function

Private Function FPTK_HeaderInstructionBaseText( _
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
        " (LIHAT SHEET", _
        " [LIHAT SHEET", _
        " - LIHAT SHEET", _
        " LIHAT SHEET", _
        " (DROPDOWN", _
        " [DROPDOWN", _
        " - DROPDOWN", _
        " (COPY VALUE DARI", _
        " [COPY VALUE DARI", _
        " - COPY VALUE DARI", _
        " COPY VALUE DARI", _
        " (ISI SESUAI", _
        " [ISI SESUAI", _
        " - ISI SESUAI")

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

    FPTK_HeaderInstructionBaseText = _
        Trim$( _
            resultText)

End Function

Private Function FPTK_SafeHeaderText( _
    ByVal rawValue As Variant) As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    FPTK_SafeHeaderText = _
        Trim$( _
            CStr(rawValue))

End Function
' ============================================================
' FPTK AVAILABILITY
'
' Aturan:
' - header berada pada row 2;
' - data dimulai row 3;
' - kolom selalu dipindahkan ke paling kanan;
' - value memakai dropdown Y / N;
' - value existing tetap dipertahankan;
' - variasi Yes/No lama dinormalisasi ke Y/N.
' ============================================================
Private Sub FPTK_CleanupLegacyColumns( _
    ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    FPTK_ConsolidateAliasGroup _
        ws, _
        "CATEGORY", _
        "Alasan Permintaan FPTK"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "DIVISI", _
        "Divisi"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "FILTER KATEGORISASI FPTK", _
        "Filter Kategorisasi FPTK"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "LOKASI KERJA", _
        "Lokasi Kerja"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "JUMLAH SLA", _
        "Jumlah SLA"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "KETERANGAN LULUS SLA", _
        "Keterangan Lulus SLA"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "KETERANGAN TIDAK LULUS SLA", _
        "Keterangan Tidak Lulus SLA"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "KETERANGAN CANCEL", _
        "Keterangan Cancel"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "KEBUTUHAN LAPTOP", _
        "Kebutuhan Laptop"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "LOKASI HR", _
        "Lokasi HR"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "STATUS KARYAWAN", _
        "Status Karyawan"

    FPTK_ConsolidateAliasGroup _
        ws, _
        "KODE BU", _
        "Kode BU"

    FPTK_DeleteDeprecatedColumns _
        ws

End Sub

Private Sub FPTK_ConsolidateAliasGroup( _
    ByVal ws As Worksheet, _
    ByVal canonicalKey As String, _
    ByVal preferredHeader As String)

    Dim matchingColumns As New Collection

    Dim lastColumn As Long
    Dim lastRow As Long

    Dim columnIndex As Long
    Dim preferredColumn As Long
    Dim duplicateColumn As Long

    Dim headerText As String
    Dim headerKey As String
    Dim preferredKey As String

    If ws Is Nothing Then Exit Sub

    preferredKey = _
        HeaderKeyText( _
            preferredHeader)

    lastColumn = _
        LastHeaderColumn( _
            ws, _
            FPTK_HEADER_ROW)

    For columnIndex = 1 To lastColumn

        headerText = _
            FPTK_SafeHeaderText( _
                ws.Cells( _
                    FPTK_HEADER_ROW, _
                    columnIndex).Value2)

        headerKey = _
            CanonicalHeaderKey( _
                headerText)

        If StrComp( _
            headerKey, _
            canonicalKey, _
            vbTextCompare) = 0 Then

            matchingColumns.Add _
                columnIndex

            If HeaderKeyText( _
                headerText) = _
               preferredKey Then

                preferredColumn = _
                    columnIndex

            End If

        End If

    Next columnIndex

    If matchingColumns.Count = 0 Then Exit Sub

    If preferredColumn = 0 Then

        preferredColumn = _
            CLng( _
                matchingColumns(1))

    End If

    lastRow = _
        FPTK_LastUsedDataRowAny( _
            ws)

    For columnIndex = 1 To matchingColumns.Count

        duplicateColumn = _
            CLng( _
                matchingColumns( _
                    columnIndex))

        If duplicateColumn <> preferredColumn Then

            FPTK_MergeColumnValuesSafe _
                ws, _
                preferredColumn, _
                duplicateColumn, _
                lastRow, _
                preferredHeader

        End If

    Next columnIndex

    For columnIndex = _
        matchingColumns.Count _
        To 1 Step -1

        duplicateColumn = _
            CLng( _
                matchingColumns( _
                    columnIndex))

        If duplicateColumn <> preferredColumn Then

            FPTK_DeleteDataColumnRange _
                ws, _
                duplicateColumn, _
                lastRow

            If duplicateColumn < preferredColumn Then

                preferredColumn = _
                    preferredColumn - 1

            End If

        End If

    Next columnIndex

    ws.Cells( _
        FPTK_HEADER_ROW, _
        preferredColumn).Value = _
            preferredHeader

End Sub

Private Sub FPTK_MergeColumnValuesSafe( _
    ByVal ws As Worksheet, _
    ByVal primaryColumn As Long, _
    ByVal duplicateColumn As Long, _
    ByVal lastRow As Long, _
    ByVal fieldName As String)

    Dim primaryData As Variant
    Dim duplicateData As Variant
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim conflictCount As Long
    Dim changed As Boolean

    If lastRow < FPTK_FIRST_DATA_ROW Then
        Exit Sub
    End If

    rowCount = _
        lastRow - _
        FPTK_FIRST_DATA_ROW + 1

    primaryData = _
        ws.Cells( _
            FPTK_FIRST_DATA_ROW, _
            primaryColumn).Resize( _
                rowCount, _
                1).Value2

    duplicateData = _
        ws.Cells( _
            FPTK_FIRST_DATA_ROW, _
            duplicateColumn).Resize( _
                rowCount, _
                1).Value2

    For rowIndex = 1 _
        To rowCount

        If ValueIsBlankFPTK( _
            primaryData( _
                rowIndex, _
                1)) Then

            If Not ValueIsBlankFPTK( _
                duplicateData( _
                    rowIndex, _
                    1)) Then

                primaryData( _
                    rowIndex, _
                    1) = _
                        duplicateData( _
                            rowIndex, _
                            1)

                changed = True

            End If

        ElseIf Not ValueIsBlankFPTK( _
            duplicateData( _
                rowIndex, _
                1)) Then

            If Not ValuesEqualFPTK( _
                primaryData( _
                    rowIndex, _
                    1), _
                duplicateData( _
                    rowIndex, _
                    1)) Then

                conflictCount = _
                    conflictCount + 1

            End If

        End If

    Next rowIndex

    If changed Then

        ws.Cells( _
            FPTK_FIRST_DATA_ROW, _
            primaryColumn).Resize( _
                rowCount, _
                1).Value2 = _
                    primaryData

    End If

    If conflictCount > 0 Then

        AddRunLog _
            "Warning", _
            "FPTK alias conflict | Field: " & _
            fieldName & _
            " | Rows: " & _
            CStr(conflictCount) & _
            " | Primary value retained."

    End If

End Sub
Private Sub FPTK_DeleteDeprecatedColumns( _
    ByVal ws As Worksheet)

    Dim lastColumn As Long
    Dim lastRow As Long
    Dim columnIndex As Long
    Dim headerText As String

    If ws Is Nothing Then Exit Sub

    lastColumn = _
        LastHeaderColumn( _
            ws, _
            FPTK_HEADER_ROW)

    lastRow = _
        FPTK_LastUsedDataRowAny( _
            ws)

    For columnIndex = lastColumn _
        To 1 Step -1

        headerText = _
            FPTK_SafeHeaderText( _
                ws.Cells( _
                    FPTK_HEADER_ROW, _
                    columnIndex).Value2)

        If FPTK_IsDeprecatedSourceHeader( _
            headerText) Then

            FPTK_DeleteDataColumnRange _
                ws, _
                columnIndex, _
                lastRow

            AddRunLog _
                "Info", _
                "FPTK deprecated column removed: " & _
                headerText

        End If

    Next columnIndex

End Sub

Private Function FPTK_IsDeprecatedSourceHeader( _
    ByVal rawHeader As String) As Boolean

    Dim headerKey As String

    headerKey = _
        HeaderKeyText( _
            rawHeader)

    Select Case headerKey

        Case "MONTH FPTK DATE REAL", _
             "YEAR FPTK", _
             "YEAR FPTK DATE", _
             "FPTK YEAR"

            FPTK_IsDeprecatedSourceHeader = _
                True

    End Select

End Function

Private Function FPTK_LastUsedDataRowAny( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    If ws Is Nothing Then
        Exit Function
    End If

    On Error Resume Next

    Set foundCell = _
        ws.Cells.Find( _
            What:="*", _
            After:=ws.Cells( _
                FPTK_HEADER_ROW, _
                1), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False)

    On Error GoTo 0

    If foundCell Is Nothing Then

        FPTK_LastUsedDataRowAny = _
            FPTK_HEADER_ROW

    Else

        FPTK_LastUsedDataRowAny = _
            foundCell.Row

    End If

End Function
Public Sub RepairFPTKLegacyColumnsNow()

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            "FPTK")

    On Error GoTo 0

    If ws Is Nothing Then

        MsgBox _
            "Sheet FPTK tidak ditemukan.", _
            vbExclamation, _
            "Repair FPTK Header"

        Exit Sub

    End If

    EnsureFPTKHeaderRow1Layout _
        ws

    EnsureRequiredHeaders _
        ws

    FPTK_CleanupLegacyColumns _
        ws

    FPTK_EnsureSpecialColumnOrder _
        ws

    FPTK_ApplyAvailabilityValidation _
        ws

    StyleHeaderDynamic _
        ws

    MsgBox _
        "Header FPTK row 1 sudah dirapikan.", _
        vbInformation, _
        "Repair FPTK Header"

End Sub

Private Sub FPTK_MoveCanonicalBefore( _
    ByVal ws As Worksheet, _
    ByVal sourceKey As String, _
    ByVal beforeKey As String)

    Dim headerMap As Object
    Dim sourceColumn As Long
    Dim beforeColumn As Long

    If ws Is Nothing Then Exit Sub

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    sourceColumn = _
        ColumnByKey( _
            headerMap, _
            sourceKey)

    beforeColumn = _
        ColumnByKey( _
            headerMap, _
            beforeKey)

    If sourceColumn > 0 And _
       beforeColumn > 0 Then

        If sourceColumn <> _
           beforeColumn - 1 Then

            FPTK_MoveColumnBefore _
                ws, _
                sourceColumn, _
                beforeColumn

        End If

    End If

    Set headerMap = Nothing

End Sub
Private Sub FPTK_StandardizeFinalHeaderNames( _
    ByVal ws As Worksheet)

    Dim headerMap As Object
    Dim columnIndex As Long

    If ws Is Nothing Then Exit Sub

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    columnIndex = _
        ColumnByKey( _
            headerMap, _
            "CATEGORY")

    If columnIndex > 0 Then
        ws.Cells(FPTK_HEADER_ROW, columnIndex).Value = _
            "Alasan Permintaan FPTK"
    End If

    columnIndex = _
        ColumnByKey( _
            headerMap, _
            "DIVISI")

    If columnIndex > 0 Then
        ws.Cells(FPTK_HEADER_ROW, columnIndex).Value = "Divisi"
    End If

    columnIndex = _
        ColumnByKey( _
            headerMap, _
            "FILTER KATEGORISASI FPTK")

    If columnIndex > 0 Then
        ws.Cells(FPTK_HEADER_ROW, columnIndex).Value = _
            "Filter Kategorisasi FPTK"
    End If

    columnIndex = _
        ColumnByKey( _
            headerMap, _
            "WEEK FPTK DATE")

    If columnIndex > 0 Then
        ws.Cells(FPTK_HEADER_ROW, columnIndex).Value = _
            "Week FPTK Date (Kode)"
    End If

    columnIndex = _
        ColumnByKey( _
            headerMap, _
            "STATUS KARYAWAN")

    If columnIndex > 0 Then
        ws.Cells(FPTK_HEADER_ROW, columnIndex).Value = _
            "Status Karyawan"
    End If

    columnIndex = _
        ColumnByKey( _
            headerMap, _
            "KODE BU")

    If columnIndex > 0 Then
        ws.Cells(FPTK_HEADER_ROW, columnIndex).Value = _
            "Kode BU"
    End If

    columnIndex = _
        ColumnByKey( _
            headerMap, _
            "FPTK AVAILABILITY")

    If columnIndex > 0 Then
        ws.Cells(FPTK_HEADER_ROW, columnIndex).Value = _
            "FPTK Availability"
    End If

    Set headerMap = Nothing

End Sub

Private Sub FPTK_EnsureSpecialColumnOrder( _
    ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    FPTK_MoveCanonicalBefore _
        ws, _
        "DIVISI", _
        "DEPARTMENT"

    FPTK_MoveCanonicalBefore _
        ws, _
        "FILTER KATEGORISASI FPTK", _
        "VACANCY"

    FPTK_MoveCanonicalBefore _
        ws, _
        "STATUS KARYAWAN", _
        "KODE BU"

    FPTK_MoveCanonicalBefore _
        ws, _
        "KODE BU", _
        "FPTK AVAILABILITY"

    FPTK_MoveCanonicalBefore _
        ws, _
        "FPTK AVAILABILITY", _
        "CREATED AT"

    FPTK_StandardizeFinalHeaderNames _
        ws

End Sub

Private Sub FPTK_DeleteDataColumnRange( _
    ByVal ws As Worksheet, _
    ByVal columnIndex As Long, _
    ByVal lastRow As Long)

    Dim deleteRange As Range

    If ws Is Nothing Then Exit Sub
    If columnIndex <= 0 Then Exit Sub

    If lastRow < FPTK_HEADER_ROW Then
        lastRow = FPTK_HEADER_ROW
    End If

    Set deleteRange = _
        ws.Range( _
            ws.Cells( _
                FPTK_HEADER_ROW, _
                columnIndex), _
            ws.Cells( _
                lastRow, _
                columnIndex))

    deleteRange.Delete _
        Shift:=xlToLeft

    Set deleteRange = Nothing

End Sub

Private Sub FPTK_MoveColumnBefore( _
    ByVal ws As Worksheet, _
    ByVal sourceColumn As Long, _
    ByVal beforeColumn As Long)

    Dim lastRow As Long
    Dim insertColumn As Long
    Dim sourceRange As Range
    Dim insertRange As Range

    If ws Is Nothing Then Exit Sub
    If sourceColumn <= 0 Then Exit Sub
    If beforeColumn <= 0 Then Exit Sub
    If sourceColumn = beforeColumn Then Exit Sub
    If sourceColumn = beforeColumn - 1 Then Exit Sub

    lastRow = _
        FPTK_LastUsedDataRowAny( _
            ws)

    If lastRow < FPTK_HEADER_ROW Then
        lastRow = FPTK_HEADER_ROW
    End If

    If sourceColumn < beforeColumn Then
        insertColumn = beforeColumn - 1
    Else
        insertColumn = beforeColumn
    End If

    Set sourceRange = _
        ws.Range( _
            ws.Cells( _
                FPTK_HEADER_ROW, _
                sourceColumn), _
            ws.Cells( _
                lastRow, _
                sourceColumn))

    Set insertRange = _
        ws.Range( _
            ws.Cells( _
                FPTK_HEADER_ROW, _
                insertColumn), _
            ws.Cells( _
                lastRow, _
                insertColumn))

    sourceRange.Cut

    insertRange.Insert _
        Shift:=xlToRight

    Application.CutCopyMode = False

    Set insertRange = Nothing
    Set sourceRange = Nothing

End Sub

Private Sub FPTK_EnsureAvailabilityRightmost( _
    ByVal ws As Worksheet)

    ' Nama procedure dipertahankan supaya call lama tetap valid.
    ' Posisi final:
    ' Kode BU | FPTK Availability | Created At
    FPTK_EnsureSpecialColumnOrder ws

End Sub

Private Sub FPTK_ApplyAvailabilityValidation( _
    ByVal ws As Worksheet)

    Dim headerMap As Object

    Dim availabilityColumn As Long
    Dim lastRow As Long
    Dim validationLastRow As Long

    Dim targetRange As Range

    If ws Is Nothing Then
        Exit Sub
    End If

    Set headerMap = _
        BuildHeaderMap( _
            ws, _
            FPTK_HEADER_ROW)

    availabilityColumn = _
        ColumnByKey( _
            headerMap, _
            "FPTK AVAILABILITY")

    If availabilityColumn <= 0 Then
        Exit Sub
    End If

    lastRow = _
        LastTargetDataRow( _
            ws, _
            headerMap)

    validationLastRow = _
        lastRow + 5000

    If validationLastRow < 10000 Then

        validationLastRow = 10000

    End If

    If validationLastRow > _
       ws.rows.Count Then

        validationLastRow = _
            ws.rows.Count

    End If

    Set targetRange = _
        ws.Range( _
            ws.Cells( _
                FPTK_FIRST_DATA_ROW, _
                availabilityColumn), _
            ws.Cells( _
                validationLastRow, _
                availabilityColumn))

    On Error Resume Next

    targetRange.Validation.Delete

    On Error GoTo 0

    targetRange.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="Y,N"

    With targetRange.Validation

        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = _
            "FPTK Availability"

        .InputMessage = _
            "Pilih Y atau N."

        .ErrorTitle = _
            "Value tidak valid"

        .ErrorMessage = _
            "FPTK Availability hanya boleh diisi Y atau N."

    End With

    FPTK_NormalizeAvailabilityValues _
        ws, _
        availabilityColumn, _
        lastRow

    Set targetRange = Nothing
    Set headerMap = Nothing

End Sub

Private Sub FPTK_NormalizeAvailabilityValues( _
    ByVal ws As Worksheet, _
    ByVal availabilityColumn As Long, _
    ByVal lastRow As Long)

    Dim availabilityData As Variant

    Dim rowCount As Long
    Dim rowIndex As Long

    Dim rawText As String
    Dim normalizedValue As String
    Dim changed As Boolean

    If ws Is Nothing Then
        Exit Sub
    End If

    If availabilityColumn <= 0 Then
        Exit Sub
    End If

    If lastRow < FPTK_FIRST_DATA_ROW Then
        Exit Sub
    End If

    rowCount = _
        lastRow - _
        FPTK_FIRST_DATA_ROW + 1

    availabilityData = _
        ws.Cells( _
            FPTK_FIRST_DATA_ROW, _
            availabilityColumn).Resize( _
                rowCount, _
                1).Value2

    For rowIndex = 1 _
        To rowCount

        If Not IsError( _
            availabilityData( _
                rowIndex, _
                1)) Then

            rawText = _
                UCase$( _
                    Trim$( _
                        CStr( _
                            availabilityData( _
                                rowIndex, _
                                1))))

            normalizedValue = _
                vbNullString

            Select Case rawText

                Case "Y", _
                     "YES", _
                     "YA", _
                     "TRUE", _
                     "1", _
                     "V"

                    normalizedValue = "Y"

                Case "N", _
                     "NO", _
                     "TIDAK", _
                     "FALSE", _
                     "0", _
                     "X"

                    normalizedValue = "N"

            End Select

            If Len(normalizedValue) > 0 Then

                If rawText <> _
                   normalizedValue Then

                    availabilityData( _
                        rowIndex, _
                        1) = _
                            normalizedValue

                    changed = True

                End If

            End If

        End If

    Next rowIndex

    If changed Then

        ws.Cells( _
            FPTK_FIRST_DATA_ROW, _
            availabilityColumn).Resize( _
                rowCount, _
                1).Value2 = _
                    availabilityData

    End If

End Sub
Public Sub RefreshFPTKAvailabilityNow()

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets( _
            "FPTK")

    On Error GoTo 0

    If ws Is Nothing Then

        MsgBox _
            "Sheet FPTK tidak ditemukan.", _
            vbExclamation, _
            "FPTK Availability"

        Exit Sub

    End If

    EnsureFPTKHeaderRow1Layout _
        ws

    EnsureRequiredHeaders _
        ws

    FPTK_EnsureAvailabilityRightmost _
        ws

    FPTK_ApplyAvailabilityValidation _
        ws

    StyleHeaderDynamic _
        ws

    MsgBox _
        "Kolom FPTK Availability sudah dibuat " & _
        "sebelum Created At dengan pilihan Y/N.", _
        vbInformation, _
        "FPTK Availability"

End Sub
Private Function RequiredFPTKHeaders() As Variant

    Dim headerText As String

    headerText = _
        "Kode PIC|FPTK Date (Real)|Kode Angka|" & _
        "FPTK Date (Kode)|Kode Unik|Posisi|" & _
        "Business Unit|Direktorat|Divisi|" & _
        "Department|Level FPTK|Level Number|" & _
        "Alasan Permintaan FPTK|Category FPTK|PIC Recruiter"

    headerText = headerText & _
        "|Filter Kategorisasi FPTK|Vacancy|Status|" & _
        "Week FPTK Date (Kode)|Month FPTK Date|" & _
        "FPTK Cancel Date|Week Cancel Date|" & _
        "Month Cancel Date|Offering Date|" & _
        "Week Offering Date|Month Offering"

    headerText = headerText & _
        "|Jumlah SLA|Deadline pemenuhan SLA|" & _
        "Detail SLA|Keterangan Lulus SLA|" & _
        "Keterangan Tidak Lulus SLA|" & _
        "Keterangan Cancel|Nama Kandidat|" & _
        "Estimasi Join|Kebutuhan Laptop"

    headerText = headerText & _
        "|Lokasi Onboarding|" & _
        "Tanggal Upload ke Website|" & _
        "User (Manager)|Indirect User|" & _
        "Lokasi Kerja|Lokasi HR|Status Karyawan|" & _
        "Kode BU|FPTK Availability|Created At|" & _
        "Last Updated At|Last Compile Action|" & _
        "Source File|Remark"

    RequiredFPTKHeaders = _
        Split( _
            headerText, _
            "|")

End Function

' ============================================================
' CANONICAL HEADER ALIAS
' ============================================================
Private Function CanonicalHeaderKey( _
    ByVal rawHeader As Variant) As String

    Dim keyText As String
    Dim schemaKey As String

    keyText = HeaderKeyText(rawHeader)

    Select Case keyText

        Case "KODE PIC"
            CanonicalHeaderKey = "KODE PIC"

        Case "FPTK DATE REAL"
            CanonicalHeaderKey = "FPTK DATE REAL"

        Case "KODE ANGKA", _
             "KODE", _
             "POSITION CODE", _
             "KODE POSISI"

            CanonicalHeaderKey = "KODE ANGKA"

        Case "FPTK DATE PAKE INI", _
             "FPTK DATE KODE", _
             "FPTK DATE CODE"

            CanonicalHeaderKey = "FPTK DATE KODE"

        Case "KODE UNIK"
            CanonicalHeaderKey = "KODE UNIK"

        Case "POSISI", _
             "POSITION", _
             "POSISI KEBUTUHAN TA", _
             "NAMA POSISI"

            CanonicalHeaderKey = "POSISI"

        Case "BUSINESS UNIT", _
             "PT BUSINESS UNIT", _
             "PT UNIT"

            CanonicalHeaderKey = "BUSINESS UNIT"

        Case "DIREKTORAT", _
             "DIRECTORATE"

            CanonicalHeaderKey = "DIREKTORAT"

        Case "DIVISI", _
             "DIVISION", _
             "DIVISI SESUAI SO", _
             "DIVISI SESUAI SO 2", _
             "DIVISION CHRIS"

            CanonicalHeaderKey = "DIVISI"

        Case "DEPARTMENT", _
             "DEPARTEMEN", _
             "DEPARTMENT CHRIS"

            CanonicalHeaderKey = "DEPARTMENT"

        Case "LEVEL FPTK", _
             "LEVEL POSISI"

            CanonicalHeaderKey = "LEVEL FPTK"

        Case "LEVEL NUMBER", _
             "LEVEL NUMERIC"

            CanonicalHeaderKey = "LEVEL NUMBER"

        Case "CATEGORY", _
             "ALASAN", _
             "ALASAN FPTK", _
             "ALASAN PERMINTAAN FPTK"

            CanonicalHeaderKey = "CATEGORY"

        Case "CATEGORY FPTK"
            CanonicalHeaderKey = "CATEGORY FPTK"

        Case "PIC RECRUITER", _
             "PIC REKRUTER", _
             "RECRUITER", _
             "REKRUTER"

            CanonicalHeaderKey = "PIC RECRUITER"

        Case "FILTER KATEGORISASI FPTK", _
             "FILTER KATEGORISASI FPTK2", _
             "FILTER KATEGORISASI FPTK 2", _
             "FILTER KATEGORI FPTK", _
             "KATEGORISASI FPTK"

            CanonicalHeaderKey = _
                "FILTER KATEGORISASI FPTK"

        Case "VACANCY", _
             "JUMLAH POSISI"

            CanonicalHeaderKey = "VACANCY"

        Case "STATUS", _
             "STATUS REKRUTMEN", _
             "STATUS VACANCY"

            CanonicalHeaderKey = "STATUS"

        Case "WEEK FPTK DATE KODE", _
             "WEEK FPTK DATE CODE", _
             "WEEK FPTK DATE"

            CanonicalHeaderKey = _
                "WEEK FPTK DATE"

        Case "MONTH FPTK DATE"
            CanonicalHeaderKey = _
                "MONTH FPTK DATE"

        Case "FPTK CANCEL DATE", _
             "CANCEL DATE"

            CanonicalHeaderKey = _
                "FPTK CANCEL DATE"

        Case "WEEK CANCEL DATE"
            CanonicalHeaderKey = _
                "WEEK CANCEL DATE"

        Case "MONTH CANCEL DATE"
            CanonicalHeaderKey = _
                "MONTH CANCEL DATE"

        Case "OFFERING DATE", _
             "TANGGAL OFFERING"

            CanonicalHeaderKey = _
                "OFFERING DATE"

        Case "WEEK OFFERING DATE"
            CanonicalHeaderKey = _
                "WEEK OFFERING DATE"

        Case "MONTH OFFERING"
            CanonicalHeaderKey = _
                "MONTH OFFERING"

        Case "JUMLAH SLA", _
             "SLA DAYS", _
             "SLA"

            CanonicalHeaderKey = _
                "JUMLAH SLA"

        Case "DEADLINE PEMENUHAN SLA", _
             "DEADLINE SLA"

            CanonicalHeaderKey = _
                "DEADLINE SLA"

        Case "DETAIL SLA", _
             "STATUS SLA"

            CanonicalHeaderKey = _
                "DETAIL SLA"

        Case "KETERANGAN 1", _
             "KETERANGAN LULUS SLA"

            CanonicalHeaderKey = _
                "KETERANGAN LULUS SLA"

        Case "KETERANGAN 0", _
             "KETERANGAN TIDAK LULUS SLA"

            CanonicalHeaderKey = _
                "KETERANGAN TIDAK LULUS SLA"

        Case "KETERANGAN CANCEL", _
             "KETERANGAN CANCEL KOSONG"

            CanonicalHeaderKey = _
                "KETERANGAN CANCEL"

        Case "NAMA KANDIDAT", _
             "KANDIDAT"

            CanonicalHeaderKey = _
                "NAMA KANDIDAT"

        Case "ESTIMASI JOIN", _
             "ESTIMATED JOIN DATE"

            CanonicalHeaderKey = _
                "ESTIMASI JOIN"

        Case "KEBUTUHAN LAPTOP", _
             "KEBUTUHAN LAPTOP V"

            CanonicalHeaderKey = _
                "KEBUTUHAN LAPTOP"

        Case "LOKASI ONBOARDING"
            CanonicalHeaderKey = _
                "LOKASI ONBOARDING"

        Case "TANGGAL UPLOAD KE WEBSITE", _
             "TGL UPLOAD KE WEBSITE"

            CanonicalHeaderKey = _
                "TANGGAL UPLOAD WEBSITE"

        Case "USER MANAGER", _
             "USER", _
             "MANAGER"

            CanonicalHeaderKey = _
                "USER MANAGER"

        Case "INDIRECT USER"
            CanonicalHeaderKey = _
                "INDIRECT USER"

        Case "LOKASI KERJA", _
             "LOCATION", _
             "PENEMPATAN"

            CanonicalHeaderKey = _
                "LOKASI KERJA"

        Case "LOKASI HR", _
             "HR LOCATION", _
             "LOKASI RECRUITER"

            CanonicalHeaderKey = _
                "LOKASI HR"

        Case "STATUS KARYAWAN", _
             "EMPLOYMENT STATUS", _
             "EMPLOYEE STATUS", _
             "STATUS EMPLOYEE"

            CanonicalHeaderKey = _
                "STATUS KARYAWAN"

        Case "KODE BU", _
             "KODE BU2", _
             "KODE BU 2"

            CanonicalHeaderKey = "KODE BU"

        Case "CREATED AT", _
             "CREATED DATE", _
             "CREATED DATETIME"

            CanonicalHeaderKey = "CREATED AT"

        Case "LAST UPDATED AT", _
             "LAST UPDATE", _
             "LAST UPDATED", _
             "UPDATED AT"

            CanonicalHeaderKey = "LAST UPDATED AT"

        Case "LAST COMPILE ACTION", _
             "COMPILE ACTION", _
             "LAST ACTION"

            CanonicalHeaderKey = _
                "LAST COMPILE ACTION"

        Case "SOURCE FILE", _
             "LAST SOURCE FILE", _
             "SOURCE FILENAME"

            CanonicalHeaderKey = "SOURCE FILE"

        Case "REMARK", _
             "REMARKS"

            CanonicalHeaderKey = "REMARK"

        Case "FPTK AVAILABILITY", _
             "AVAILABILITY FPTK", _
             "FPTK AVAILABLE", _
             "AVAILABLE FPTK", _
             "STATUS FPTK"

            CanonicalHeaderKey = _
                "FPTK AVAILABILITY"

        Case Else

            ' Header yang tidak tercakup alias lokal akan
            ' dicocokkan ke Canonical Header di Schema Config.
            On Error Resume Next

            schemaKey = SchemaCanonicalHeaderKey( _
                "FPTK", _
                rawHeader)

            On Error GoTo 0

            If Len(Trim$(schemaKey)) > 0 And _
               StrComp( _
                    schemaKey, _
                    keyText, _
                    vbTextCompare) <> 0 Then

                CanonicalHeaderKey = _
                    CanonicalHeaderKey(schemaKey)

            Else

                CanonicalHeaderKey = keyText

            End If

    End Select

End Function

Private Function HeaderKeyText( _
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
    sourceText = Replace$(sourceText, Chr$(160), " ")

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

    HeaderKeyText = Application.Trim( _
        resultText)

End Function

' ============================================================
' SOURCE / TARGET ROW HELPERS
' ============================================================
Private Function FindHeaderRowDynamic( _
    ByVal ws As Object) As Long

    Dim rowIndex As Long
    Dim headerMap As Object
    Dim anchorCount As Long

    For rowIndex = 1 To 10

        Set headerMap = BuildHeaderMap( _
            ws, rowIndex)

        anchorCount = 0

        If headerMap.Exists("KODE PIC") Then
            anchorCount = anchorCount + 1
        End If

        If headerMap.Exists("KODE UNIK") Then
            anchorCount = anchorCount + 1
        End If

        If headerMap.Exists("POSISI") Then
            anchorCount = anchorCount + 1
        End If

        If headerMap.Exists("FPTK DATE REAL") Then
            anchorCount = anchorCount + 1
        End If

        If anchorCount >= 2 Then

            FindHeaderRowDynamic = rowIndex
            Exit Function

        End If

    Next rowIndex

End Function

Private Function LastHeaderColumn( _
    ByVal ws As Object, _
    ByVal headerRow As Long) As Long

    Dim lastColumn As Long

    On Error Resume Next

    lastColumn = ws.Cells( _
        headerRow, _
        ws.Columns.Count).End(xlToLeft).Column

    On Error GoTo 0

    If lastColumn < 1 Then
        lastColumn = 1
    End If

    LastHeaderColumn = lastColumn

End Function
Private Sub FPTK_ExposeAllSourceRows( _
    ByVal ws As Object, _
    ByVal headerRow As Long)

    Dim listObjectItem As Object
    Dim lastUsedRow As Long
    Dim lastUsedColumn As Long

    If ws Is Nothing Then Exit Sub

    On Error Resume Next

    ' --------------------------------------------------------
    ' Worksheet AutoFilter.
    ' --------------------------------------------------------
    If ws.FilterMode Then

        ws.ShowAllData

    End If

    ' --------------------------------------------------------
    ' Excel Table / ListObject filters.
    ' Beberapa source memakai table filter, jadi clear juga.
    ' --------------------------------------------------------
    For Each listObjectItem In ws.ListObjects

        Err.Clear

        If Not listObjectItem.AutoFilter Is Nothing Then

            If listObjectItem.AutoFilter.FilterMode Then

                listObjectItem.AutoFilter.ShowAllData

            End If

        End If

    Next listObjectItem

    ' --------------------------------------------------------
    ' Manual hidden rows.
    '
    ' Ini hanya pada workbook source yang dibuka read-only.
    ' Workbook ditutup SaveChanges:=False, sehingga tampilan
    ' asli user tidak berubah.
    ' --------------------------------------------------------
    lastUsedRow = _
        FPTK_LastPhysicalSourceRow( _
            ws)

    lastUsedColumn = _
        FPTK_LastPhysicalSourceColumn( _
            ws, _
            headerRow)

    If lastUsedRow > headerRow Then

        ws.rows( _
            CStr(headerRow + 1) & _
            ":" & _
            CStr(lastUsedRow)).Hidden = _
                False

    End If

    ' Column hidden tidak memengaruhi Range.Value2, tetapi
    ' sengaja tidak diubah agar tidak perlu menyentuh layout.
    If lastUsedColumn < 0 Then
        lastUsedColumn = 0
    End If

    Err.Clear
    On Error GoTo 0

End Sub

Private Function FPTK_LastPhysicalSourceRow( _
    ByVal ws As Object) As Long

    Dim foundCell As Object

    If ws Is Nothing Then Exit Function

    On Error Resume Next

    Set foundCell = _
        ws.Cells.Find( _
            What:="*", _
            After:=ws.Cells(1, 1), _
            LookIn:=-4163, _
            LookAt:=2, _
            SearchOrder:=1, _
            SearchDirection:=2, _
            MatchCase:=False)

    If Not foundCell Is Nothing Then

        FPTK_LastPhysicalSourceRow = _
            foundCell.Row

    End If

    Set foundCell = Nothing

    On Error GoTo 0

End Function

Private Function FPTK_LastPhysicalSourceColumn( _
    ByVal ws As Object, _
    ByVal headerRow As Long) As Long

    Dim foundCell As Object

    If ws Is Nothing Then Exit Function

    On Error Resume Next

    Set foundCell = _
        ws.Cells.Find( _
            What:="*", _
            After:=ws.Cells(1, 1), _
            LookIn:=-4163, _
            LookAt:=2, _
            SearchOrder:=2, _
            SearchDirection:=2, _
            MatchCase:=False)

    If Not foundCell Is Nothing Then

        FPTK_LastPhysicalSourceColumn = _
            foundCell.Column

    ElseIf headerRow > 0 Then

        FPTK_LastPhysicalSourceColumn = _
            ws.Cells( _
                headerRow, _
                ws.Columns.Count).End( _
                    xlToLeft).Column

    End If

    Set foundCell = Nothing

    On Error GoTo 0

End Function
Private Function LastSourceDataRow( _
    ByVal ws As Object, _
    ByVal headerRow As Long, _
    ByVal sourceKeys As Variant) As Long

    Dim preferredKeys As Variant
    Dim keyItem As Variant
    Dim sourceColumn As Long
    Dim candidateRow As Long
    Dim resultRow As Long
    Dim physicalLastRow As Long

    preferredKeys = Array( _
        "KODE UNIK", _
        "KODE PIC", _
        "POSISI", _
        "FPTK DATE REAL")

    resultRow = headerRow

    For Each keyItem In preferredKeys

        sourceColumn = ColumnIndexInKeyArray( _
            sourceKeys, _
            CStr(keyItem))

        If sourceColumn > 0 Then

            candidateRow = _
                ws.Cells( _
                    ws.rows.Count, _
                    sourceColumn).End( _
                        xlUp).Row

            If candidateRow > resultRow Then

                resultRow = _
                    candidateRow

            End If

        End If

    Next keyItem

    ' Extra guard:
    ' gunakan last physical populated row bila lebih bawah.
    ' Filter/hidden status tidak menjadi dasar pemotongan data.
    physicalLastRow = _
        FPTK_LastPhysicalSourceRow( _
            ws)

    If physicalLastRow > resultRow Then

        resultRow = _
            physicalLastRow

    End If

    LastSourceDataRow = _
        resultRow

End Function

Private Function LastTargetDataRow( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object) As Long

    Dim preferredKeys As Variant
    Dim keyItem As Variant
    Dim targetColumn As Long
    Dim candidateRow As Long
    Dim resultRow As Long

    preferredKeys = Array( _
        "KODE UNIK", _
        "KODE PIC", _
        "POSISI", _
        "FPTK DATE REAL")

    resultRow = FPTK_HEADER_ROW

    For Each keyItem In preferredKeys

        targetColumn = ColumnByKey( _
            headerMap, CStr(keyItem))

        If targetColumn > 0 Then

            candidateRow = ws.Cells( _
                ws.rows.Count, _
                targetColumn).End(xlUp).Row

            If candidateRow > resultRow Then
                resultRow = candidateRow
            End If

        End If

    Next keyItem

    LastTargetDataRow = resultRow

End Function

Private Function GetNextEmptyRowDynamic( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object) As Long

    Dim lastRow As Long

    lastRow = LastTargetDataRow( _
        ws, headerMap)

    If lastRow < FPTK_FIRST_DATA_ROW Then

        GetNextEmptyRowDynamic = _
            FPTK_FIRST_DATA_ROW

    Else

        GetNextEmptyRowDynamic = _
            lastRow + 1

    End If

End Function

Private Function ColumnIndexInKeyArray( _
    ByVal sourceKeys As Variant, _
    ByVal requestedKey As String) As Long

    Dim columnIndex As Long
    Dim canonicalKey As String

    canonicalKey = CanonicalHeaderKey( _
        requestedKey)

    For columnIndex = LBound(sourceKeys) _
        To UBound(sourceKeys)

        If StrComp( _
            CStr(sourceKeys(columnIndex)), _
            canonicalKey, _
            vbTextCompare) = 0 Then

            ColumnIndexInKeyArray = _
                columnIndex

            Exit Function

        End If

    Next columnIndex

End Function

' ============================================================
' READ / WRITE MAPPED VALUES
' ============================================================
Private Function WriteMappedRowFromArray( _
    ByVal wsTarget As Worksheet, _
    ByVal targetRow As Long, _
    ByVal targetMap As Object, _
    ByVal sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceKeys As Variant) As Boolean

    Dim sourceColumn As Long
    Dim headerKey As String
    Dim targetColumn As Long

    Dim sourceValue As Variant
    Dim targetValue As Variant
    Dim rowChanged As Boolean

    For sourceColumn = LBound(sourceKeys) _
        To UBound(sourceKeys)

        headerKey = CStr( _
            sourceKeys(sourceColumn))

        If Len(headerKey) > 0 Then

            ' Kolom audit dan derived value dikelola compiler.
            ' Nilai source tidak boleh menimpa hasil compiler.
            If Not IsFPTKCompilerManagedHeaderKey( _
                headerKey) Then

                If targetMap.Exists(headerKey) Then

                    targetColumn = CLng( _
                        targetMap(headerKey))

                    sourceValue = sourceData( _
                        sourceRow, _
                        sourceColumn)

                    If Not ValueIsBlankFPTK( _
                        sourceValue) Then

                        targetValue = wsTarget.Cells( _
                            targetRow, _
                            targetColumn).Value2

                        If Not ValuesEqualFPTK( _
                            targetValue, _
                            sourceValue) Then

                            wsTarget.Cells( _
                                targetRow, _
                                targetColumn).Value2 = _
                                sourceValue

                            rowChanged = True

                        End If

                    End If

                End If

            End If

        End If

    Next sourceColumn

    WriteMappedRowFromArray = _
        rowChanged

End Function

Private Function ValuesEqualFPTK( _
    ByVal firstValue As Variant, _
    ByVal secondValue As Variant) As Boolean

    Dim firstText As String
    Dim secondText As String

    If IsError(firstValue) Or _
       IsError(secondValue) Then

        ValuesEqualFPTK = _
            (IsError(firstValue) And _
             IsError(secondValue))

        Exit Function

    End If

    If IsNull(firstValue) Or _
       IsEmpty(firstValue) Then

        firstText = vbNullString

    Else

        firstText = CStr(firstValue)

    End If

    If IsNull(secondValue) Or _
       IsEmpty(secondValue) Then

        secondText = vbNullString

    Else

        secondText = CStr(secondValue)

    End If

    If IsNumeric(firstValue) And _
       IsNumeric(secondValue) Then

        ValuesEqualFPTK = _
            (Abs( _
                CDbl(firstValue) - _
                CDbl(secondValue)) < 0.0000001)

    Else

        ValuesEqualFPTK = _
            (StrComp( _
                Trim$(firstText), _
                Trim$(secondText), _
                vbBinaryCompare) = 0)

    End If

End Function

Private Function IsFPTKCompilerManagedHeaderKey( _
    ByVal headerKey As String) As Boolean

    Select Case CanonicalHeaderKey( _
        headerKey)

        Case "CREATED AT", _
             "LAST UPDATED AT", _
             "LAST COMPILE ACTION", _
             "SOURCE FILE", _
             "FILTER KATEGORISASI FPTK", _
             "WEEK FPTK DATE", _
             "MONTH FPTK DATE", _
             "WEEK CANCEL DATE", _
             "MONTH CANCEL DATE", _
             "WEEK OFFERING DATE", _
             "MONTH OFFERING", _
             "JUMLAH SLA", _
             "DEADLINE SLA", _
             "DETAIL SLA"

            IsFPTKCompilerManagedHeaderKey = True

    End Select

End Function

Private Sub StampFPTKAudit( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal headerMap As Object, _
    ByVal compileAction As String, _
    ByVal sourceFileName As String, _
    ByVal isInsert As Boolean)

    Dim colCreatedAt As Long
    Dim colLastUpdatedAt As Long
    Dim colLastAction As Long
    Dim colSourceFile As Long

    Dim stampTime As Date

    If ws Is Nothing Then Exit Sub
    If headerMap Is Nothing Then Exit Sub

    stampTime = Now

    colCreatedAt = ColumnByKey( _
        headerMap, _
        "CREATED AT")

    colLastUpdatedAt = ColumnByKey( _
        headerMap, _
        "LAST UPDATED AT")

    colLastAction = ColumnByKey( _
        headerMap, _
        "LAST COMPILE ACTION")

    colSourceFile = ColumnByKey( _
        headerMap, _
        "SOURCE FILE")

    If isInsert And _
       colCreatedAt > 0 Then

        If Len(Trim$(CStr( _
            ws.Cells( _
                rowIndex, _
                colCreatedAt).Value2))) = 0 Then

            ws.Cells( _
                rowIndex, _
                colCreatedAt).Value = _
                stampTime

        End If

    End If

    If colLastUpdatedAt > 0 Then

        ws.Cells( _
            rowIndex, _
            colLastUpdatedAt).Value = _
                stampTime

    End If

    If colLastAction > 0 Then

        ws.Cells( _
            rowIndex, _
            colLastAction).Value2 = _
                UCase$(Trim$(compileAction))

    End If

    If colSourceFile > 0 Then

        ws.Cells( _
            rowIndex, _
            colSourceFile).Value2 = _
                sourceFileName

    End If

End Sub

Private Function ValueIsBlankFPTK( _
    ByVal rawValue As Variant) As Boolean

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        ValueIsBlankFPTK = True
        Exit Function

    End If

    ValueIsBlankFPTK = _
        (Len(Trim$(CStr(rawValue))) = 0)

End Function

Private Function ArrayValueByKey( _
    ByVal sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceKeys As Variant, _
    ByVal requestedKey As String) As Variant

    Dim sourceColumn As Long

    sourceColumn = ColumnIndexInKeyArray( _
        sourceKeys, requestedKey)

    If sourceColumn > 0 Then

        ArrayValueByKey = sourceData( _
            sourceRow, sourceColumn)

    Else

        ArrayValueByKey = Empty

    End If

End Function

Private Function SheetValueByKey( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal headerMap As Object, _
    ByVal requestedKey As String) As Variant

    Dim targetColumn As Long

    targetColumn = ColumnByKey( _
        headerMap, requestedKey)

    If targetColumn > 0 Then

        SheetValueByKey = ws.Cells( _
            rowIndex, targetColumn).Value2

    Else

        SheetValueByKey = Empty

    End If

End Function

Private Function ValueFromOptionalColumn( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long) As Variant

    If columnIndex > 0 Then

        ValueFromOptionalColumn = _
            ws.Cells( _
                rowIndex, columnIndex).Value2

    Else

        ValueFromOptionalColumn = Empty

    End If

End Function

' ============================================================
' SKIP INSTRUCTION / DUPLICATE HEADER ROW
' ============================================================
Private Function RowIsInstruction( _
    ByVal sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceKeys As Variant) As Boolean

    Dim kodePIC As String
    Dim kodeUnik As String
    Dim posisi As String
    Dim checkText As String

    kodePIC = CStr(ArrayValueByKey( _
        sourceData, sourceRow, _
        sourceKeys, "KODE PIC"))

    kodeUnik = CStr(ArrayValueByKey( _
        sourceData, sourceRow, _
        sourceKeys, "KODE UNIK"))

    posisi = CStr(ArrayValueByKey( _
        sourceData, sourceRow, _
        sourceKeys, "POSISI"))

    checkText = UCase$( _
        Trim$(kodePIC & " " & _
              kodeUnik & " " & posisi))

    If Len(checkText) = 0 Then

        RowIsInstruction = True
        Exit Function

    End If

    If InStr(1, checkText, _
        "RUMUS", vbTextCompare) > 0 Then

        RowIsInstruction = True

    ElseIf InStr(1, checkText, _
        "DROPDOWN", vbTextCompare) > 0 Then

        RowIsInstruction = True

    ElseIf InStr(1, checkText, _
        "ISI SESUAI", vbTextCompare) > 0 Then

        RowIsInstruction = True

    ElseIf InStr(1, checkText, _
        "KODE PIC", vbTextCompare) > 0 And _
           InStr(1, checkText, _
        "KODE UNIK", vbTextCompare) > 0 Then

        RowIsInstruction = True

    End If

End Function

Private Sub RemoveDuplicateHeaderRows( _
    ByVal ws As Worksheet)

    Dim checkCount As Long

    For checkCount = 1 To 5

        If RowLooksLikeHeader(ws, 2) Then

            ws.rows(2).Delete Shift:=xlUp

        Else

            Exit For

        End If

    Next checkCount

End Sub

Private Function RowLooksLikeHeader( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long) As Boolean

    Dim headerMap As Object
    Dim anchorCount As Long

    Set headerMap = BuildHeaderMap( _
        ws, rowIndex)

    If headerMap.Exists("KODE PIC") Then
        anchorCount = anchorCount + 1
    End If

    If headerMap.Exists("KODE UNIK") Then
        anchorCount = anchorCount + 1
    End If

    If headerMap.Exists("POSISI") Then
        anchorCount = anchorCount + 1
    End If

    If headerMap.Exists("FPTK DATE REAL") Then
        anchorCount = anchorCount + 1
    End If

    RowLooksLikeHeader = _
        (anchorCount >= 2)

End Function

Private Function RowHasPrimaryData( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal headerMap As Object) As Boolean

    Dim kodePIC As String
    Dim posisi As String
    Dim kodeUnik As String

    kodePIC = Trim$(CStr( _
        SheetValueByKey( _
            ws, rowIndex, _
            headerMap, "KODE PIC")))

    posisi = Trim$(CStr( _
        SheetValueByKey( _
            ws, rowIndex, _
            headerMap, "POSISI")))

    kodeUnik = Trim$(CStr( _
        SheetValueByKey( _
            ws, rowIndex, _
            headerMap, "KODE UNIK")))

    RowHasPrimaryData = _
        (Len(kodePIC) > 0 Or _
         Len(posisi) > 0 Or _
         Len(kodeUnik) > 0)

End Function


' ============================================================
' VALIDASI RECORD SOURCE - SAFE
'
' Fungsi ini hanya menolak row helper/dropdown dari source.
' Tidak ada penghapusan row pada sheet target.
' ============================================================
Private Function FPTK_IsValidSourceRecord( _
    ByVal sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceKeys As Variant) As Boolean

    Dim kodeUnik As Variant
    Dim posisi As Variant

    If RowIsInstruction( _
        sourceData, _
        sourceRow, _
        sourceKeys) Then

        Exit Function

    End If

    kodeUnik = ArrayValueByKey( _
        sourceData, _
        sourceRow, _
        sourceKeys, _
        "KODE UNIK")

    posisi = ArrayValueByKey( _
        sourceData, _
        sourceRow, _
        sourceKeys, _
        "POSISI")

    If ValueIsBlankFPTK(kodeUnik) Then
        Exit Function
    End If

    If ValueIsBlankFPTK(posisi) Then
        Exit Function
    End If

    ' Kode Unik FPTK yang benar memiliki huruf dan angka.
    ' Ini menolak helper seperti MPRec20, CORPRec09, atau 45901.
    If Not FPTK_TextHasLetterAndDigit( _
        CStr(kodeUnik)) Then

        Exit Function

    End If

    FPTK_IsValidSourceRecord = True

End Function

Private Function FPTK_TextHasLetterAndDigit( _
    ByVal sourceText As String) As Boolean

    Dim normalizedText As String
    Dim currentCharacter As String
    Dim characterIndex As Long

    Dim hasLetter As Boolean
    Dim hasDigit As Boolean

    normalizedText = UCase$( _
        Trim$(sourceText))

    normalizedText = Replace$( _
        normalizedText, _
        Chr$(160), _
        vbNullString)

    normalizedText = Replace$( _
        normalizedText, _
        " ", _
        vbNullString)

    For characterIndex = 1 To Len(normalizedText)

        currentCharacter = Mid$( _
            normalizedText, _
            characterIndex, _
            1)

        If currentCharacter >= "A" And _
           currentCharacter <= "Z" Then

            hasLetter = True

        ElseIf currentCharacter >= "0" And _
               currentCharacter <= "9" Then

            hasDigit = True

        End If

        If hasLetter And hasDigit Then

            FPTK_TextHasLetterAndDigit = True
            Exit Function

        End If

    Next characterIndex

End Function

' ============================================================
' RESET DATE RESOLVER
' ============================================================
Private Sub FPTK_ResetDateResolverCache()

    gFPTKDateSettingsLoaded = False

    ' DMY dipakai sebagai fallback agar sama dengan baseline
    ' sebelumnya bila Month/Week source tidak tersedia.
    gFPTKDefaultDateOrder = "DMY"

    gFPTKDateFormatOverrides = vbNullString

    gFPTKAmbiguousDateCount = 0
    gFPTKResolvedBySourceMonthCount = 0
    gFPTKResolvedBySourceWeekCount = 0
    gFPTKResolvedByMonthWeekCount = 0
    gFPTKSourceDateConflictCount = 0
    gFPTKResolvedByFallbackCount = 0

End Sub
Private Function FPTK_IsStrictSourceDateKey( _
    ByVal headerKey As String) As Boolean

    Select Case CanonicalHeaderKey( _
        headerKey)

        Case "FPTK DATE REAL", _
             "FPTK DATE KODE", _
             "FPTK CANCEL DATE", _
             "OFFERING DATE"

            FPTK_IsStrictSourceDateKey = True

    End Select

End Function

' ============================================================
' RESOLVE SOURCE DATE
'
' Rule:
' - first > 12  : DMY
' - second > 12 : MDY
' - both <= 12  : cek Month/Week terkait
'
' Month/Week source tidak disalin sebagai hasil.
' Keduanya hanya menentukan DMY atau MDY.
' ============================================================
Private Function FPTK_ResolveSourceDateValue( _
    ByVal rawValue As Variant, _
    ByVal dateFieldKey As String, _
    ByVal sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceKeys As Variant, _
    ByVal sourceFilePath As String, _
    ByVal sourceExcelRow As Long) As Variant

    Dim parsedDate As Date
    Dim dmyDate As Date
    Dim mdyDate As Date
    Dim isAmbiguous As Boolean

    Dim weekKey As String
    Dim monthKey As String

    Dim expectedWeek As Long
    Dim expectedMonth As Long

    Dim dateFromWeek As Date
    Dim dateFromMonth As Date
    Dim selectedDate As Date

    Dim sourceOverride As String
    Dim defaultOrder As String

    If ValueIsBlankFPTK(rawValue) Then

        FPTK_ResolveSourceDateValue = rawValue
        Exit Function

    End If

    If FPTK_TryParseDateByRules( _
        rawValue, _
        parsedDate, _
        isAmbiguous) Then

        FPTK_ResolveSourceDateValue = parsedDate
        Exit Function

    End If

    If Not isAmbiguous Then

        FPTK_ResolveSourceDateValue = rawValue
        Exit Function

    End If

    If Not FPTK_TryBuildAmbiguousCandidates( _
        rawValue, _
        dmyDate, _
        mdyDate) Then

        FPTK_ResolveSourceDateValue = rawValue
        Exit Function

    End If

    gFPTKAmbiguousDateCount = _
        gFPTKAmbiguousDateCount + 1

    FPTK_GetCompanionDateKeys _
        dateFieldKey, _
        weekKey, _
        monthKey

    expectedWeek = _
        FPTK_ParseExpectedWeek( _
            FPTK_SourceArrayValueByKey( _
                sourceData, _
                sourceRow, _
                sourceKeys, _
                weekKey))

    expectedMonth = _
        FPTK_ParseExpectedMonth( _
            FPTK_SourceArrayValueByKey( _
                sourceData, _
                sourceRow, _
                sourceKeys, _
                monthKey))

    dateFromMonth = _
        FPTK_SelectCandidateByMonth( _
            dmyDate, _
            mdyDate, _
            expectedMonth)

    dateFromWeek = _
        FPTK_SelectCandidateByWeek( _
            dmyDate, _
            mdyDate, _
            expectedWeek)

    If dateFromMonth > 0 And _
       dateFromWeek > 0 Then

        If CLng(dateFromMonth) = _
           CLng(dateFromWeek) Then

            selectedDate = dateFromMonth

            gFPTKResolvedByMonthWeekCount = _
                gFPTKResolvedByMonthWeekCount + 1

        Else

            gFPTKSourceDateConflictCount = _
                gFPTKSourceDateConflictCount + 1

            AddRunLog _
                "Warning", _
                "FPTK date clue conflict | File: " & _
                FPTKFileNameFromPath(sourceFilePath) & _
                " | Row: " & _
                CStr(sourceExcelRow) & _
                " | Field: " & _
                CanonicalHeaderKey(dateFieldKey) & _
                " | Raw: " & _
                Trim$(CStr(rawValue)) & _
                " | Source month chooses: " & _
                Format$(dateFromMonth, "dd/mm/yyyy") & _
                " | Source week chooses: " & _
                Format$(dateFromWeek, "dd/mm/yyyy")

        End If

    ElseIf dateFromMonth > 0 Then

        selectedDate = dateFromMonth

        gFPTKResolvedBySourceMonthCount = _
            gFPTKResolvedBySourceMonthCount + 1

    ElseIf dateFromWeek > 0 Then

        selectedDate = dateFromWeek

        gFPTKResolvedBySourceWeekCount = _
            gFPTKResolvedBySourceWeekCount + 1

    End If

    If selectedDate > 0 Then

        FPTK_ResolveSourceDateValue = selectedDate
        Exit Function

    End If

    ' Fallback per file/folder bila tersedia.
    sourceOverride = _
        FPTK_GetSourceDateOverride( _
            sourceFilePath)

    If sourceOverride = "DMY" Then

        selectedDate = dmyDate

    ElseIf sourceOverride = "MDY" Then

        selectedDate = mdyDate

    Else

        defaultOrder = _
            FPTK_GetDefaultDateOrder()

        If defaultOrder = "MDY" Then

            selectedDate = mdyDate

        ElseIf defaultOrder = "DMY" Then

            selectedDate = dmyDate

        End If

    End If

    If selectedDate > 0 Then

        gFPTKResolvedByFallbackCount = _
            gFPTKResolvedByFallbackCount + 1

        AddRunLog _
            "Warning", _
            "FPTK ambiguous date fallback | File: " & _
            FPTKFileNameFromPath(sourceFilePath) & _
            " | Row: " & _
            CStr(sourceExcelRow) & _
            " | Field: " & _
            CanonicalHeaderKey(dateFieldKey) & _
            " | Raw: " & _
            Trim$(CStr(rawValue)) & _
            " | Source month: " & _
            IIf( _
                expectedMonth > 0, _
                CStr(expectedMonth), _
                "blank") & _
            " | Source week: " & _
            IIf( _
                expectedWeek > 0, _
                CStr(expectedWeek), _
                "blank") & _
            " | Selected: " & _
            Format$(selectedDate, "dd/mm/yyyy")

        FPTK_ResolveSourceDateValue = selectedDate
        Exit Function

    End If

    ' AUTO dan tidak ada clue yang membedakan:
    ' raw dipertahankan supaya tidak salah menebak.
    AddRunLog _
        "Warning", _
        "FPTK ambiguous date unresolved | File: " & _
        FPTKFileNameFromPath(sourceFilePath) & _
        " | Row: " & _
        CStr(sourceExcelRow) & _
        " | Field: " & _
        CanonicalHeaderKey(dateFieldKey) & _
        " | Raw: " & _
        Trim$(CStr(rawValue)) & _
        " | DMY: " & _
        Format$(dmyDate, "dd/mm/yyyy") & _
        " | MDY: " & _
        Format$(mdyDate, "dd/mm/yyyy")

    FPTK_ResolveSourceDateValue = rawValue

End Function
Private Function FPTK_SelectCandidateByMonth( _
    ByVal dmyDate As Date, _
    ByVal mdyDate As Date, _
    ByVal expectedMonth As Long) As Date

    Dim dmyMatches As Boolean
    Dim mdyMatches As Boolean

    If expectedMonth < 1 Or _
       expectedMonth > 12 Then

        Exit Function

    End If

    dmyMatches = _
        (Month(dmyDate) = expectedMonth)

    mdyMatches = _
        (Month(mdyDate) = expectedMonth)

    If dmyMatches Xor mdyMatches Then

        If dmyMatches Then

            FPTK_SelectCandidateByMonth = dmyDate

        Else

            FPTK_SelectCandidateByMonth = mdyDate

        End If

    End If

End Function
Private Function FPTK_SelectCandidateByWeek( _
    ByVal dmyDate As Date, _
    ByVal mdyDate As Date, _
    ByVal expectedWeek As Long) As Date

    Dim dmyMatches As Boolean
    Dim mdyMatches As Boolean

    If expectedWeek < 1 Or _
       expectedWeek > 53 Then

        Exit Function

    End If

    dmyMatches = _
        (FPTK_WeekNumber(dmyDate) = expectedWeek)

    mdyMatches = _
        (FPTK_WeekNumber(mdyDate) = expectedWeek)

    If dmyMatches Xor mdyMatches Then

        If dmyMatches Then

            FPTK_SelectCandidateByWeek = dmyDate

        Else

            FPTK_SelectCandidateByWeek = mdyDate

        End If

    End If

End Function

' ============================================================
' PARSE TANPA CONTEXT
'
' Return True bila tanggal sudah pasti.
' isAmbiguous=True bila angka pertama dan kedua sama-sama <=12.
' ============================================================
Private Function FPTK_TryParseDateByRules( _
    ByVal rawValue As Variant, _
    ByRef parsedDate As Date, _
    ByRef isAmbiguous As Boolean) As Boolean

    Dim textValue As String
    Dim dateParts() As String

    Dim firstValue As Long
    Dim secondValue As Long
    Dim thirdValue As Long

    Dim candidateDate As Date

    parsedDate = 0
    isAmbiguous = False

    On Error GoTo InvalidDate

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    ' Serial Excel sudah tidak ambigu.
    If IsNumeric(rawValue) Then

        If CDbl(rawValue) > 0 Then

            parsedDate = _
                dateValue( _
                    CDate(CDbl(rawValue)))

            FPTK_TryParseDateByRules = True

        End If

        Exit Function

    End If

    textValue = _
        FPTK_CleanDateText( _
            CStr(rawValue))

    If Len(textValue) = 0 Then
        Exit Function
    End If

    dateParts = Split(textValue, "/")

    If UBound(dateParts) = 2 Then

        If IsNumeric(dateParts(0)) And _
           IsNumeric(dateParts(1)) And _
           IsNumeric(dateParts(2)) Then

            firstValue = CLng(dateParts(0))
            secondValue = CLng(dateParts(1))
            thirdValue = CLng(dateParts(2))

            ' YYYY/MM/DD
            If Len(dateParts(0)) = 4 Then

                If FPTK_TryMakeDate( _
                    firstValue, _
                    secondValue, _
                    thirdValue, _
                    candidateDate) Then

                    parsedDate = candidateDate
                    FPTK_TryParseDateByRules = True

                End If

                Exit Function

            End If

            If thirdValue < 100 Then
                thirdValue = thirdValue + 2000
            End If

            ' Angka pertama >12 = pasti DMY.
            If firstValue > 12 And _
               firstValue <= 31 And _
               secondValue >= 1 And _
               secondValue <= 12 Then

                If FPTK_TryMakeDate( _
                    thirdValue, _
                    secondValue, _
                    firstValue, _
                    candidateDate) Then

                    parsedDate = candidateDate
                    FPTK_TryParseDateByRules = True

                End If

                Exit Function

            End If

            ' Angka kedua >12 = pasti MDY.
            If secondValue > 12 And _
               secondValue <= 31 And _
               firstValue >= 1 And _
               firstValue <= 12 Then

                If FPTK_TryMakeDate( _
                    thirdValue, _
                    firstValue, _
                    secondValue, _
                    candidateDate) Then

                    parsedDate = candidateDate
                    FPTK_TryParseDateByRules = True

                End If

                Exit Function

            End If

            ' Angka sama menghasilkan tanggal yang sama.
            If firstValue = secondValue And _
               firstValue >= 1 And _
               firstValue <= 12 Then

                If FPTK_TryMakeDate( _
                    thirdValue, _
                    secondValue, _
                    firstValue, _
                    candidateDate) Then

                    parsedDate = candidateDate
                    FPTK_TryParseDateByRules = True

                End If

                Exit Function

            End If

            ' Keduanya <=12 dan berbeda = ambigu.
            If firstValue >= 1 And _
               firstValue <= 12 And _
               secondValue >= 1 And _
               secondValue <= 12 Then

                isAmbiguous = True
                Exit Function

            End If

        End If

    End If

    ' Tanggal dengan nama bulan tidak dianggap ambigu.
    If FPTK_TextContainsLetter(textValue) Then

        If IsDate(textValue) Then

            parsedDate = _
                dateValue(CDate(textValue))

            FPTK_TryParseDateByRules = True

        End If

    End If

    Exit Function

InvalidDate:

    parsedDate = 0
    FPTK_TryParseDateByRules = False

End Function
Private Function FPTK_TryBuildAmbiguousCandidates( _
    ByVal rawValue As Variant, _
    ByRef dmyDate As Date, _
    ByRef mdyDate As Date) As Boolean

    Dim textValue As String
    Dim dateParts() As String

    Dim firstValue As Long
    Dim secondValue As Long
    Dim yearValue As Long

    textValue = _
        FPTK_CleanDateText( _
            CStr(rawValue))

    dateParts = Split(textValue, "/")

    If UBound(dateParts) <> 2 Then
        Exit Function
    End If

    If Not IsNumeric(dateParts(0)) Or _
       Not IsNumeric(dateParts(1)) Or _
       Not IsNumeric(dateParts(2)) Then

        Exit Function

    End If

    firstValue = CLng(dateParts(0))
    secondValue = CLng(dateParts(1))
    yearValue = CLng(dateParts(2))

    If yearValue < 100 Then
        yearValue = yearValue + 2000
    End If

    If firstValue < 1 Or _
       firstValue > 12 Or _
       secondValue < 1 Or _
       secondValue > 12 Or _
       firstValue = secondValue Then

        Exit Function

    End If

    ' DMY: first = day, second = month.
    If Not FPTK_TryMakeDate( _
        yearValue, _
        secondValue, _
        firstValue, _
        dmyDate) Then

        Exit Function

    End If

    ' MDY: first = month, second = day.
    If Not FPTK_TryMakeDate( _
        yearValue, _
        firstValue, _
        secondValue, _
        mdyDate) Then

        Exit Function

    End If

    FPTK_TryBuildAmbiguousCandidates = True

End Function
Private Function FPTK_TryMakeDate( _
    ByVal yearValue As Long, _
    ByVal monthValue As Long, _
    ByVal dayValue As Long, _
    ByRef resultDate As Date) As Boolean

    Dim candidateDate As Date

    On Error GoTo InvalidDate

    candidateDate = DateSerial( _
        yearValue, _
        monthValue, _
        dayValue)

    If Year(candidateDate) <> yearValue Or _
       Month(candidateDate) <> monthValue Or _
       Day(candidateDate) <> dayValue Then

        Exit Function

    End If

    resultDate = candidateDate
    FPTK_TryMakeDate = True

    Exit Function

InvalidDate:

    FPTK_TryMakeDate = False

End Function
Private Function FPTK_CleanDateText( _
    ByVal rawText As String) As String

    Dim resultText As String
    Dim spacePosition As Long

    resultText = Trim$(rawText)

    resultText = Replace$( _
        resultText, _
        ".", _
        "/")

    resultText = Replace$( _
        resultText, _
        "-", _
        "/")

    resultText = Replace$( _
        resultText, _
        "\", _
        "/")

    ' Buang bagian jam untuk numeric date + time.
    spacePosition = InStr( _
        1, _
        resultText, _
        " ", _
        vbBinaryCompare)

    If spacePosition > 0 And _
       InStr( _
            1, _
            Left$(resultText, spacePosition - 1), _
            "/", _
            vbBinaryCompare) > 0 Then

        resultText = _
            Left$( _
                resultText, _
                spacePosition - 1)

    End If

    FPTK_CleanDateText = Trim$(resultText)

End Function
Private Function FPTK_TextContainsLetter( _
    ByVal textValue As String) As Boolean

    Dim characterIndex As Long
    Dim characterCode As Long

    For characterIndex = 1 To Len(textValue)

        characterCode = AscW( _
            UCase$( _
                Mid$( _
                    textValue, _
                    characterIndex, _
                    1)))

        If characterCode >= 65 And _
           characterCode <= 90 Then

            FPTK_TextContainsLetter = True
            Exit Function

        End If

    Next characterIndex

End Function
Private Sub FPTK_GetCompanionDateKeys( _
    ByVal dateFieldKey As String, _
    ByRef weekKey As String, _
    ByRef monthKey As String)

    weekKey = vbNullString
    monthKey = vbNullString

    Select Case CanonicalHeaderKey( _
        dateFieldKey)

        Case "FPTK DATE REAL", _
             "FPTK DATE KODE"

            weekKey = "WEEK FPTK DATE"
            monthKey = "MONTH FPTK DATE"

        Case "FPTK CANCEL DATE"

            weekKey = "WEEK CANCEL DATE"
            monthKey = "MONTH CANCEL DATE"

        Case "OFFERING DATE"

            weekKey = "WEEK OFFERING DATE"
            monthKey = "MONTH OFFERING"

    End Select

End Sub
Private Function FPTK_SourceArrayValueByKey( _
    ByVal sourceData As Variant, _
    ByVal sourceRow As Long, _
    ByVal sourceKeys As Variant, _
    ByVal requestedKey As String) As Variant

    Dim sourceColumn As Long

    If Len(requestedKey) = 0 Then
        Exit Function
    End If

    sourceColumn = ColumnIndexInKeyArray( _
        sourceKeys, _
        requestedKey)

    If sourceColumn > 0 Then

        FPTK_SourceArrayValueByKey = _
            sourceData( _
                sourceRow, _
                sourceColumn)

    End If

End Function
Private Function FPTK_ParseExpectedWeek( _
    ByVal rawValue As Variant) As Long

    Dim textValue As String
    Dim digitText As String
    Dim characterIndex As Long
    Dim currentCharacter As String
    Dim weekValue As Long

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    If IsNumeric(rawValue) Then

        weekValue = CLng(rawValue)

    Else

        textValue = CStr(rawValue)

        For characterIndex = 1 To Len(textValue)

            currentCharacter = Mid$( _
                textValue, _
                characterIndex, _
                1)

            If currentCharacter >= "0" And _
               currentCharacter <= "9" Then

                digitText = _
                    digitText & currentCharacter

            ElseIf Len(digitText) > 0 Then

                Exit For

            End If

        Next characterIndex

        If Len(digitText) > 0 Then
            weekValue = CLng(digitText)
        End If

    End If

    If weekValue >= 1 And _
       weekValue <= 53 Then

        FPTK_ParseExpectedWeek = weekValue

    End If

End Function
Private Function FPTK_ParseExpectedMonth( _
    ByVal rawValue As Variant) As Long

    Dim monthText As String
    Dim monthValue As Long

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    If IsNumeric(rawValue) Then

        monthValue = CLng(rawValue)

        If monthValue >= 1 And _
           monthValue <= 12 Then

            FPTK_ParseExpectedMonth = monthValue

        End If

        Exit Function

    End If

    monthText = UCase$( _
        Trim$(CStr(rawValue)))

    Select Case Left$(monthText, 3)

        Case "JAN"
            monthValue = 1

        Case "FEB"
            monthValue = 2

        Case "MAR"
            monthValue = 3

        Case "APR"
            monthValue = 4

        Case "MEI", "MAY"
            monthValue = 5

        Case "JUN"
            monthValue = 6

        Case "JUL"
            monthValue = 7

        Case "AGU", "AUG"
            monthValue = 8

        Case "SEP"
            monthValue = 9

        Case "OKT", "OCT"
            monthValue = 10

        Case "NOV"
            monthValue = 11

        Case "DES", "DEC"
            monthValue = 12

    End Select

    FPTK_ParseExpectedMonth = monthValue

End Function
Private Function FPTK_WeekNumber( _
    ByVal dateValueInput As Date) As Long

    On Error Resume Next

    FPTK_WeekNumber = _
        Application.WorksheetFunction. _
            WeekNum(dateValueInput, 1)

    If Err.Number <> 0 Then

        Err.Clear

        FPTK_WeekNumber = _
            DatePart( _
                "ww", _
                dateValueInput, _
                vbSunday, _
                vbFirstJan1)

    End If

    On Error GoTo 0

End Function

' ============================================================
' OPTIONAL SETTINGS
'
' Settings:
' DefaultDateOrder   = DMY / MDY / AUTO
' DateFormatOverrides = Citeureup=DMY;NamaFileLain=MDY
' ============================================================
Private Function FPTK_GetSourceDateOverride( _
    ByVal sourceFilePath As String) As String

    Dim rulesText As String
    Dim normalizedRules As String
    Dim ruleItem As Variant

    Dim separatorPosition As Long
    Dim matchText As String
    Dim dateOrder As String
    Dim lowerPath As String

    FPTK_LoadDateResolverSettings

    rulesText = gFPTKDateFormatOverrides

    normalizedRules = rulesText
    normalizedRules = Replace$( _
        normalizedRules, _
        vbCrLf, _
        ";")

    normalizedRules = Replace$( _
        normalizedRules, _
        vbCr, _
        ";")

    normalizedRules = Replace$( _
        normalizedRules, _
        vbLf, _
        ";")

    normalizedRules = Replace$( _
        normalizedRules, _
        "|", _
        ";")

    lowerPath = LCase$(sourceFilePath)

    For Each ruleItem In _
        Split(normalizedRules, ";")

        separatorPosition = InStr( _
            1, _
            CStr(ruleItem), _
            "=", _
            vbBinaryCompare)

        If separatorPosition > 1 Then

            matchText = LCase$( _
                Trim$( _
                    Left$( _
                        CStr(ruleItem), _
                        separatorPosition - 1)))

            dateOrder = _
                FPTK_NormalizeDateOrder( _
                    Mid$( _
                        CStr(ruleItem), _
                        separatorPosition + 1))

            If Len(matchText) > 0 And _
               Len(dateOrder) > 0 Then

                If InStr( _
                    1, _
                    lowerPath, _
                    matchText, _
                    vbTextCompare) > 0 Then

                    FPTK_GetSourceDateOverride = _
                        dateOrder

                    Exit Function

                End If

            End If

        End If

    Next ruleItem

End Function
Private Function FPTK_GetDefaultDateOrder() As String

    FPTK_LoadDateResolverSettings

    FPTK_GetDefaultDateOrder = _
        gFPTKDefaultDateOrder

End Function
Private Sub FPTK_LoadDateResolverSettings()

    Dim configuredOrder As String

    If gFPTKDateSettingsLoaded Then
        Exit Sub
    End If

    configuredOrder = _
        FPTK_NormalizeDateOrder( _
            FPTK_ReadSettingValueLocal( _
                "DefaultDateOrder"))

    If Len(configuredOrder) > 0 Then

        gFPTKDefaultDateOrder = _
            configuredOrder

    Else

        ' Sama dengan baseline aman sebelumnya.
        gFPTKDefaultDateOrder = "DMY"

    End If

    gFPTKDateFormatOverrides = _
        FPTK_ReadSettingValueLocal( _
            "DateFormatOverrides")

    gFPTKDateSettingsLoaded = True

End Sub
Private Function FPTK_ReadSettingValueLocal( _
    ByVal settingName As String) As String

    Dim ws As Worksheet
    Dim foundCell As Range

    Dim lastRow As Long
    Dim lastColumn As Long

    Dim rowIndex As Long
    Dim columnIndex As Long

    Dim keyText As String
    Dim requestedKey As String

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets( _
        "Settings")

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets( _
            "Setting")

    End If

    On Error GoTo 0

    If ws Is Nothing Then
        Exit Function
    End If

    On Error Resume Next

    Set foundCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    On Error GoTo 0

    If foundCell Is Nothing Then
        Exit Function
    End If

    lastRow = foundCell.Row
    Set foundCell = Nothing

    On Error Resume Next

    Set foundCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    On Error GoTo 0

    If foundCell Is Nothing Then
        Exit Function
    End If

    lastColumn = foundCell.Column

    If lastColumn > 5 Then
        lastColumn = 5
    End If

    requestedKey = _
        FPTK_NormalizeSettingKey( _
            settingName)

    For rowIndex = 1 To lastRow

        For columnIndex = 1 To lastColumn

            keyText = _
                FPTK_NormalizeSettingKey( _
                    ws.Cells( _
                        rowIndex, _
                        columnIndex).Value2)

            If keyText = requestedKey Then

                FPTK_ReadSettingValueLocal = _
                    Trim$(CStr( _
                        ws.Cells( _
                            rowIndex, _
                            columnIndex + 1).Value2))

                Exit Function

            End If

        Next columnIndex

    Next rowIndex

End Function
Private Function FPTK_NormalizeSettingKey( _
    ByVal rawValue As Variant) As String

    Dim resultText As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    resultText = UCase$( _
        Trim$(CStr(rawValue)))

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

    FPTK_NormalizeSettingKey = resultText

End Function
Private Function FPTK_NormalizeDateOrder( _
    ByVal rawValue As String) As String

    Dim resultText As String

    resultText = UCase$( _
        Trim$(rawValue))

    Select Case resultText

        Case "DMY", _
             "DDMMYYYY", _
             "DD/MM/YYYY"

            FPTK_NormalizeDateOrder = "DMY"

        Case "MDY", _
             "MMDDYYYY", _
             "MM/DD/YYYY"

            FPTK_NormalizeDateOrder = "MDY"

        Case "AUTO"
            FPTK_NormalizeDateOrder = "AUTO"

    End Select

End Function

' ============================================================
' DATE HELPERS
' ============================================================
Private Sub ApplyFPTKDateFormats( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object)

    Dim dateKeys As Variant
    Dim keyItem As Variant
    Dim targetColumn As Long

    dateKeys = Array( _
        "FPTK DATE REAL", _
        "FPTK DATE KODE", _
        "FPTK CANCEL DATE", _
        "OFFERING DATE", _
        "DEADLINE SLA", _
        "ESTIMASI JOIN", _
        "TANGGAL UPLOAD WEBSITE")

    For Each keyItem In dateKeys

        targetColumn = ColumnByKey( _
            headerMap, CStr(keyItem))

        If targetColumn > 0 Then

            ws.Columns( _
                targetColumn).numberFormat = _
                "dd/mm/yyyy"

        End If

    Next keyItem

End Sub

Private Sub ApplyFPTKAuditFormats( _
    ByVal ws As Worksheet, _
    ByVal headerMap As Object)

    Dim auditDateKeys As Variant
    Dim keyItem As Variant
    Dim targetColumn As Long

    auditDateKeys = Array( _
        "CREATED AT", _
        "LAST UPDATED AT")

    For Each keyItem In auditDateKeys

        targetColumn = ColumnByKey( _
            headerMap, _
            CStr(keyItem))

        If targetColumn > 0 Then

            ws.Columns( _
                targetColumn).numberFormat = _
                "dd/mm/yyyy hh:mm:ss"

        End If

    Next keyItem

End Sub
Private Function ParseFlexibleDate( _
    ByVal rawValue As Variant) As Date

    Dim parsedDate As Date
    Dim dmyDate As Date
    Dim mdyDate As Date
    Dim isAmbiguous As Boolean
    Dim defaultOrder As String

    If FPTK_TryParseDateByRules( _
        rawValue, _
        parsedDate, _
        isAmbiguous) Then

        ParseFlexibleDate = parsedDate
        Exit Function

    End If

    If isAmbiguous Then

        If FPTK_TryBuildAmbiguousCandidates( _
            rawValue, _
            dmyDate, _
            mdyDate) Then

            defaultOrder = _
                FPTK_GetDefaultDateOrder()

            If defaultOrder = "MDY" Then

                ParseFlexibleDate = mdyDate

            ElseIf defaultOrder = "DMY" Then

                ParseFlexibleDate = dmyDate

            End If

        End If

    End If

End Function
Private Function IsDateValueFlexible( _
    ByVal rawValue As Variant) As Boolean

    IsDateValueFlexible = _
        (ParseFlexibleDate(rawValue) <> 0)

End Function
Private Function NormalizeDateKey( _
    ByVal rawValue As Variant) As String

    Dim parsedDate As Date

    parsedDate = ParseFlexibleDate(rawValue)

    If parsedDate <> 0 Then

        NormalizeDateKey = _
            Format$(parsedDate, "yyyymmdd")

    Else

        NormalizeDateKey = _
            NormalizeDataKey(rawValue)

    End If

End Function
Private Function DateCodeText( _
    ByVal rawValue As Variant) As String

    Dim parsedDate As Date

    parsedDate = ParseFlexibleDate(rawValue)

    If parsedDate <> 0 Then

        DateCodeText = _
            Format$(parsedDate, "ddmmyy")

    End If

End Function

' ============================================================
' NORMALIZE DATA KEY
' ============================================================
Private Function NormalizeDataKey( _
    ByVal rawValue As Variant) As String

    Dim resultText As String

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    resultText = UCase$( _
        Trim$(CStr(rawValue)))

    resultText = Replace$( _
        resultText, Chr$(160), " ")

    resultText = Replace$( _
        resultText, vbTab, " ")

    resultText = Replace$( _
        resultText, vbCr, " ")

    resultText = Replace$( _
        resultText, vbLf, " ")

    resultText = Replace$( _
        resultText, ChrW$(&H2010), "-")

    resultText = Replace$( _
        resultText, ChrW$(&H2011), "-")

    resultText = Replace$( _
        resultText, ChrW$(&H2012), "-")

    resultText = Replace$( _
        resultText, ChrW$(&H2013), "-")

    resultText = Replace$( _
        resultText, ChrW$(&H2014), "-")

    resultText = Replace$( _
        resultText, ChrW$(&H2212), "-")

    Do While InStr( _
        1, resultText, "  ", _
        vbBinaryCompare) > 0

        resultText = Replace$( _
            resultText, "  ", " ")

    Loop

    NormalizeDataKey = resultText

End Function

Private Function ExtractFirstNumberLocal( _
    ByVal sourceText As String) As Long

    Dim characterIndex As Long
    Dim numberText As String
    Dim currentCharacter As String

    For characterIndex = 1 To Len(sourceText)

        currentCharacter = Mid$( _
            sourceText, characterIndex, 1)

        If currentCharacter Like "#" Then

            numberText = numberText & _
                currentCharacter

        ElseIf Len(numberText) > 0 Then

            Exit For

        End If

    Next characterIndex

    If Len(numberText) > 0 Then
        ExtractFirstNumberLocal = CLng(numberText)
    End If

End Function

' ============================================================
' STYLE HEADER
' ============================================================
Private Sub StyleHeaderDynamic( _
    ByVal ws As Worksheet)

    Dim lastColumn As Long

    lastColumn = LastHeaderColumn( _
        ws, _
        FPTK_HEADER_ROW)

    If lastColumn < 1 Then Exit Sub

    With ws.Range( _
        ws.Cells(FPTK_HEADER_ROW, 1), _
        ws.Cells(FPTK_HEADER_ROW, lastColumn))

        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True

    End With

    ws.rows(FPTK_HEADER_ROW).RowHeight = 36

End Sub

' ============================================================
' FOLDER PATH
' ============================================================
Private Function EnsureTrailingSeparator( _
    ByVal folderPath As String) As String

    folderPath = Trim$(folderPath)

    If Len(folderPath) = 0 Then Exit Function

    If Right$(folderPath, 1) <> "\" And _
       Right$(folderPath, 1) <> "/" Then

        folderPath = folderPath & _
            Application.PathSeparator

    End If

    EnsureTrailingSeparator = folderPath

End Function





















