Attribute VB_Name = "m04_TableManagement"
Option Explicit

' ============================================================
' REPORT TABLE MANAGER
'
' DB Kode Posisi dan FPTK WAJIB tetap range biasa.
' Module ini mencegah kedua sheet tersebut dibuat menjadi
' Excel Table/ListObject, meskipun procedure ResizeOrCreateTable
' dipanggil dari module lain.
' ============================================================

Private Const SHEET_DB_KODE_POSISI As String = "DB Kode Posisi"
Private Const SHEET_FPTK As String = "FPTK"

' ============================================================
' MAIN
' ============================================================
Public Sub EnsureReportTables()
    Dim ws As Worksheet

    SchemaB02PrepareAll

    ' DB Kode Posisi tidak boleh menjadi Table.
    Set ws = GetOrCreateWorksheet(SHEET_DB_KODE_POSISI)
    SafeUnlistAllTables ws

    ' FPTK tidak boleh menjadi Table.
    Set ws = GetOrCreateWorksheet(SHEET_FPTK)
    SafeUnlistAllTables ws
    RT_FinalizeFPTKDuplicateHeaders ws

    ' DB Sourcing mengikuti last row dan last active header.
    ' Tidak lagi dibatasi A:BM karena schema dapat berubah.
    ResizeOrCreateTableDynamic _
        "DB Sourcing", _
        "Table5", _
        1, _
        2, _
        "A"
    ResizeOrCreateTableDynamic "Monitoring Sourcing", "Table1", 1, 2, "F"
    ResizeOrCreateTableUsed "Blacklist Candidate", "tbl_Blacklist_Candidate", "A", "I", 1, 2
    ResizeOrCreateTableUsed "funneling", "tbl_funneling", "A", "AG", 1, 2
    ResizeOrCreateTableUsed "funneling - Updated", "tbl_funneling_Updated", "A", "AG", 1, 2
    ResizeOrCreateTable "Run_Log", "tbl_Run_Log", "A", "D", 1, 2, "A"
End Sub

' ============================================================
' CLEANUP MANUAL
' Jalankan sekali setelah mengganti module lama.
' ============================================================
Public Sub Cleanup_DBKodePosisi_And_FPTK_Tables()
    Dim ws As Worksheet

    Set ws = GetOrCreateWorksheet(SHEET_DB_KODE_POSISI)
    SafeUnlistAllTables ws

    Set ws = GetOrCreateWorksheet(SHEET_FPTK)
    SafeUnlistAllTables ws
    RT_FinalizeFPTKDuplicateHeaders ws

    MsgBox "Excel Table sudah dilepas dan header FPTK sudah dirapikan.", _
           vbInformation, "Cleanup Table"
End Sub

Public Sub FinalizeFPTKHeadersFromTableManager()

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_FPTK)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet FPTK tidak ditemukan.", _
               vbExclamation, "Finalize FPTK"
        Exit Sub
    End If

    SafeUnlistAllTables ws
    RT_FinalizeFPTKDuplicateHeaders ws
    SchemaEnsureSheet SHEET_FPTK

    MsgBox "Header alias FPTK sudah digabung dan dibersihkan.", _
           vbInformation, "Finalize FPTK"

End Sub

Private Sub RT_FinalizeFPTKDuplicateHeaders(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    RT_MergeFPTKAliasGroup _
        ws, _
        "Week FPTK Date (Kode)", _
        Array("Week FPTK Date", "Week FPTK Date (CODE)")

    RT_MergeFPTKAliasGroup _
        ws, _
        "Jumlah SLA", _
        Array("SLA Days")

    RT_MergeFPTKAliasGroup _
        ws, _
        "Deadline pemenuhan SLA", _
        Array("Deadline SLA", "SLA Deadline")

    RT_MergeFPTKAliasGroup _
        ws, _
        "Detail SLA", _
        Array("SLA Status", "Status SLA")

    RT_MergeFPTKAliasGroup _
        ws, _
        "FPTK Date (Kode)", _
        Array("FPTK Date")

    RT_MergeFPTKAliasGroup _
        ws, _
        "Level FPTK", _
        Array("Level FPTK (Sesuai SO)")

    RT_MergeFPTKAliasGroup _
        ws, _
        "Kode BU", _
        Array("Kode BU2", "Kode BU 2")

End Sub

Private Sub RT_MergeFPTKAliasGroup(ByVal ws As Worksheet, _
                                   ByVal primaryHeader As String, _
                                   ByVal aliasHeaders As Variant)

    Dim primaryColumn As Long
    Dim aliasColumn As Long
    Dim aliasItem As Variant
    Dim lastRow As Long

    primaryColumn = RT_FindExactHeaderColumn(ws, primaryHeader)

    ' Bila kolom utama belum ada, ubah alias pertama menjadi header utama.
    If primaryColumn = 0 Then
        For Each aliasItem In aliasHeaders
            aliasColumn = RT_FindExactHeaderColumn(ws, CStr(aliasItem))
            If aliasColumn > 0 Then
                ws.Cells(1, aliasColumn).Value = primaryHeader
                primaryColumn = aliasColumn
                Exit For
            End If
        Next aliasItem
    End If

    If primaryColumn = 0 Then Exit Sub

    lastRow = RT_GetLastUsedRowByValue(ws)

    Do
        aliasColumn = 0

        For Each aliasItem In aliasHeaders
            aliasColumn = RT_FindRightmostHeaderColumn( _
                ws, CStr(aliasItem), primaryColumn)

            If aliasColumn > 0 Then Exit For
        Next aliasItem

        If aliasColumn = 0 Then Exit Do

        RT_CopyAliasIntoPrimary _
            ws, primaryColumn, aliasColumn, lastRow

        ws.Columns(aliasColumn).Delete Shift:=xlToLeft

        If aliasColumn < primaryColumn Then
            primaryColumn = primaryColumn - 1
        End If
    Loop

End Sub

Private Sub RT_CopyAliasIntoPrimary(ByVal ws As Worksheet, _
                                    ByVal primaryColumn As Long, _
                                    ByVal aliasColumn As Long, _
                                    ByVal lastRow As Long)

    Dim rowIndex As Long
    Dim primaryValue As Variant
    Dim aliasValue As Variant

    If lastRow < 2 Then Exit Sub

    For rowIndex = 2 To lastRow
        primaryValue = ws.Cells(rowIndex, primaryColumn).Formula
        aliasValue = ws.Cells(rowIndex, aliasColumn).Formula

        If Len(Trim$(CStr(primaryValue))) = 0 Then
            If Len(Trim$(CStr(aliasValue))) > 0 Then
                ws.Cells(rowIndex, primaryColumn).Formula = aliasValue
            End If
        End If
    Next rowIndex

End Sub

Private Function RT_FindExactHeaderColumn(ByVal ws As Worksheet, _
                                          ByVal headerName As String) As Long

    Dim schemaColumn As Long
    Dim requestedKey As String

    schemaColumn = SchemaFindColumnAuto( _
        ws, headerName, False)

    If schemaColumn > 0 Then
        RT_FindExactHeaderColumn = schemaColumn
        Exit Function
    End If
    Dim currentKey As String
    Dim lastColumn As Long
    Dim columnIndex As Long

    requestedKey = RT_NormalizeHeaderKey(headerName)
    lastColumn = RT_GetLastUsedColumnInRow(ws, 1)

    For columnIndex = 1 To lastColumn
        currentKey = RT_NormalizeHeaderKey( _
            ws.Cells(1, columnIndex).Value2)

        If StrComp(currentKey, requestedKey, vbTextCompare) = 0 Then
            RT_FindExactHeaderColumn = columnIndex
            Exit Function
        End If
    Next columnIndex

End Function

Private Function RT_FindRightmostHeaderColumn(ByVal ws As Worksheet, _
                                              ByVal headerName As String, _
                                              ByVal excludedColumn As Long) As Long

    Dim requestedKey As String
    Dim currentKey As String
    Dim lastColumn As Long
    Dim columnIndex As Long

    requestedKey = RT_NormalizeHeaderKey(headerName)
    lastColumn = RT_GetLastUsedColumnInRow(ws, 1)

    For columnIndex = lastColumn To 1 Step -1
        If columnIndex <> excludedColumn Then
            currentKey = RT_NormalizeHeaderKey( _
                ws.Cells(1, columnIndex).Value2)

            If StrComp(currentKey, requestedKey, vbTextCompare) = 0 Then
                RT_FindRightmostHeaderColumn = columnIndex
                Exit Function
            End If
        End If
    Next columnIndex

End Function

Private Function RT_NormalizeHeaderKey(ByVal rawValue As Variant) As String

    Dim sourceText As String
    Dim resultText As String
    Dim characterIndex As Long
    Dim currentCharacter As String
    Dim characterCode As Long

    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
        Exit Function
    End If

    sourceText = UCase$(CStr(rawValue))
    sourceText = Replace$(sourceText, Chr$(160), " ")

    For characterIndex = 1 To Len(sourceText)
        currentCharacter = Mid$(sourceText, characterIndex, 1)
        characterCode = AscW(currentCharacter)

        If characterCode >= 48 And characterCode <= 57 Then
            resultText = resultText & currentCharacter
        ElseIf characterCode >= 65 And characterCode <= 90 Then
            resultText = resultText & currentCharacter
        End If
    Next characterIndex

    RT_NormalizeHeaderKey = resultText

End Function

' ============================================================
' SHEET YANG DILARANG MENJADI TABLE
' ============================================================
Private Function IsNoTableSheet(ByVal sheetName As String) As Boolean
    If StrComp(sheetName, SHEET_DB_KODE_POSISI, vbTextCompare) = 0 Then
        IsNoTableSheet = True
        Exit Function
    End If

    If StrComp(sheetName, SHEET_FPTK, vbTextCompare) = 0 Then
        IsNoTableSheet = True
    End If
End Function

' ============================================================
' SAFE UNLIST
' Unlist mempertahankan data dan format.
' Tidak memakai ListObject.Delete agar data tidak ikut terhapus.
' ============================================================
Public Sub SafeUnlistAllTables(ByVal ws As Worksheet)
    Dim i As Long
    Dim lo As ListObject

    If ws Is Nothing Then Exit Sub

    For i = ws.ListObjects.Count To 1 Step -1
        Set lo = Nothing

        On Error Resume Next
        Set lo = ws.ListObjects(i)

        If Not lo Is Nothing Then
            lo.Unlist
        End If

        If Err.Number <> 0 Then
            Debug.Print "Gagal Unlist | Sheet: " & ws.Name & _
                        " | Error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next i

    If ws.AutoFilterMode Then
        On Error Resume Next
        ws.AutoFilterMode = False
        On Error GoTo 0
    End If
End Sub

' ============================================================
' REMOVE OVERLAPPING TABLES
' ============================================================
Public Sub RemoveOverlappingTables(ByVal ws As Worksheet, _
                                   ByVal tableRange As Range, _
                                   ByVal tableName As String)
    Dim i As Long
    Dim lo As ListObject

    If ws Is Nothing Then Exit Sub
    If tableRange Is Nothing Then Exit Sub

    For i = ws.ListObjects.Count To 1 Step -1
        Set lo = Nothing

        On Error Resume Next
        Set lo = ws.ListObjects(i)
        On Error GoTo 0

        If Not lo Is Nothing Then
            If StrComp(lo.Name, tableName, vbTextCompare) <> 0 Then
                If Not Intersect(tableRange, lo.Range) Is Nothing Then
                    On Error Resume Next
                    lo.Unlist

                    If Err.Number <> 0 Then
                        Debug.Print "Gagal Unlist overlapping Table | Sheet: " & _
                                    ws.Name & " | Table: " & lo.Name & _
                                    " | Error: " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    Next i
End Sub

' ============================================================
' HEADER CHECK
' Tidak lagi membuat Column1, Column2, dan seterusnya.
' ============================================================
Public Sub PrepareSheetHeaderIfBlank(ByVal ws As Worksheet, _
                                     ByVal headerRow As Long, _
                                     ByVal firstCol As Long, _
                                     ByVal lastCol As Long)
    ' Sengaja tidak mengisi header kosong dengan "Column1", dst.
    ' Table hanya dibuat bila header row memang sudah berisi header asli.
End Sub

Private Function HeaderRowIsReady(ByVal ws As Worksheet, _
                                  ByVal headerRow As Long, _
                                  ByVal firstCol As Long, _
                                  ByVal lastCol As Long) As Boolean
    HeaderRowIsReady = RowHasAnyValue(ws, headerRow, firstCol, lastCol)
End Function

' ============================================================
' FIXED TABLE RANGE
' ============================================================
Public Sub ResizeOrCreateTable(ByVal sheetName As String, _
                               ByVal tableName As String, _
                               ByVal firstCol As String, _
                               ByVal lastCol As String, _
                               ByVal headerRow As Long, _
                               ByVal firstDataRow As Long, _
                               ByVal keyColumn As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lastRow As Long
    Dim tableRange As Range
    Dim firstColNumber As Long
    Dim lastColNumber As Long

    Set ws = GetOrCreateWorksheet(sheetName)

    If IsNoTableSheet(sheetName) Then
        SafeUnlistAllTables ws
        Exit Sub
    End If

    lastRow = GetLastRowInColumnByValue(ws, keyColumn)
    If lastRow < firstDataRow Then lastRow = firstDataRow

    Set tableRange = ws.Range(firstCol & headerRow & ":" & lastCol & lastRow)

    firstColNumber = tableRange.Column
    lastColNumber = tableRange.Column + tableRange.Columns.Count - 1

    If Not HeaderRowIsReady(ws, headerRow, firstColNumber, lastColNumber) Then
        Debug.Print "Table tidak dibuat karena header kosong | Sheet: " & _
                    sheetName & " | Table: " & tableName
        Exit Sub
    End If

    RemoveOverlappingTables ws, tableRange, tableName
    Set lo = GetListObjectByName(ws, tableName)

    If Not lo Is Nothing Then
        On Error Resume Next
        lo.Resize tableRange

        If Err.Number <> 0 Then
            Err.Clear
            lo.Unlist
            Set lo = Nothing
        End If
        On Error GoTo 0
    End If

    If lo Is Nothing Then
        On Error Resume Next
        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                    Source:=tableRange, _
                                    XlListObjectHasHeaders:=xlYes)

        If Not lo Is Nothing Then lo.Name = tableName

        If Err.Number <> 0 Then
            Debug.Print "Gagal membuat Table | Sheet: " & sheetName & _
                        " | Table: " & tableName & _
                        " | Error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

' ============================================================
' USED-RANGE TABLE
' ============================================================
Public Sub ResizeOrCreateTableUsed(ByVal sheetName As String, _
                                   ByVal tableName As String, _
                                   ByVal firstCol As String, _
                                   ByVal lastCol As String, _
                                   ByVal headerRow As Long, _
                                   ByVal firstDataRow As Long)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lastRow As Long
    Dim tableRange As Range
    Dim firstColNumber As Long
    Dim lastColNumber As Long

    Set ws = GetOrCreateWorksheet(sheetName)

    If IsNoTableSheet(sheetName) Then
        SafeUnlistAllTables ws
        Exit Sub
    End If

    lastRow = GetLastUsedRowByValue(ws)
    If lastRow < firstDataRow Then lastRow = firstDataRow

    Set tableRange = ws.Range(firstCol & headerRow & ":" & lastCol & lastRow)

    firstColNumber = tableRange.Column
    lastColNumber = tableRange.Column + tableRange.Columns.Count - 1

    If Not HeaderRowIsReady(ws, headerRow, firstColNumber, lastColNumber) Then
        Debug.Print "Table tidak dibuat karena header kosong | Sheet: " & _
                    sheetName & " | Table: " & tableName
        Exit Sub
    End If

    RemoveOverlappingTables ws, tableRange, tableName
    Set lo = GetListObjectByName(ws, tableName)

    If Not lo Is Nothing Then
        On Error Resume Next
        lo.Resize tableRange

        If Err.Number <> 0 Then
            Err.Clear
            lo.Unlist
            Set lo = Nothing
        End If
        On Error GoTo 0
    End If

    If lo Is Nothing Then
        On Error Resume Next
        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                    Source:=tableRange, _
                                    XlListObjectHasHeaders:=xlYes)

        If Not lo Is Nothing Then lo.Name = tableName

        If Err.Number <> 0 Then
            Debug.Print "Gagal membuat Table | Sheet: " & sheetName & _
                        " | Table: " & tableName & _
                        " | Error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

' ============================================================
' DYNAMIC TABLE RANGE
' ============================================================
Public Sub ResizeOrCreateTableDynamic(ByVal sheetName As String, _
                                      ByVal tableName As String, _
                                      ByVal headerRow As Long, _
                                      ByVal firstDataRow As Long, _
                                      ByVal keyColumn As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tableRange As Range

    Set ws = GetOrCreateWorksheet(sheetName)

    If IsNoTableSheet(sheetName) Then
        SafeUnlistAllTables ws
        Exit Sub
    End If

    lastRow = MaxLong(GetLastRowInColumnByValue(ws, keyColumn), _
                      GetLastUsedRowByValue(ws))

    lastCol = MaxLong(GetLastUsedColumnInRow(ws, headerRow), _
                      GetLastUsedColumnByValue(ws))

    If lastCol < 1 Then lastCol = 1
    If lastRow < firstDataRow Then lastRow = firstDataRow

    Set tableRange = ws.Range(ws.Cells(headerRow, 1), _
                              ws.Cells(lastRow, lastCol))

    If Not HeaderRowIsReady(ws, headerRow, 1, lastCol) Then
        Debug.Print "Table tidak dibuat karena header kosong | Sheet: " & _
                    sheetName & " | Table: " & tableName
        Exit Sub
    End If

    RemoveOverlappingTables ws, tableRange, tableName
    Set lo = GetListObjectByName(ws, tableName)

    If Not lo Is Nothing Then
        On Error Resume Next
        lo.Resize tableRange

        If Err.Number <> 0 Then
            Err.Clear
            lo.Unlist
            Set lo = Nothing
        End If
        On Error GoTo 0
    End If

    If lo Is Nothing Then
        On Error Resume Next
        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, _
                                    Source:=tableRange, _
                                    XlListObjectHasHeaders:=xlYes)

        If Not lo Is Nothing Then lo.Name = tableName

        If Err.Number <> 0 Then
            Debug.Print "Gagal membuat Table | Sheet: " & sheetName & _
                        " | Table: " & tableName & _
                        " | Error: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

Private Function RT_GetLastUsedRowByValue( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    If ws Is Nothing Then
        RT_GetLastUsedRowByValue = 1
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

        RT_GetLastUsedRowByValue = 1

    Else

        RT_GetLastUsedRowByValue = _
            foundCell.Row

    End If

End Function

Private Function RT_GetLastUsedColumnInRow( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long) As Long

    If ws Is Nothing Then
        RT_GetLastUsedColumnInRow = 1
        Exit Function
    End If

    RT_GetLastUsedColumnInRow = _
        ws.Cells( _
            rowNumber, _
            ws.Columns.Count).End(xlToLeft).Column

    If RT_GetLastUsedColumnInRow < 1 Then
        RT_GetLastUsedColumnInRow = 1
    End If

End Function

' ============================================================
' GET TABLE BY NAME
' ============================================================
Public Function GetListObjectByName(ByVal ws As Worksheet, _
                                    ByVal tableName As String) As ListObject
    Dim lo As ListObject

    If ws Is Nothing Then Exit Function

    For Each lo In ws.ListObjects
        If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
            Set GetListObjectByName = lo
            Exit Function
        End If
    Next lo

    Set GetListObjectByName = Nothing
End Function







