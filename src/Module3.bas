Attribute VB_Name = "Module3"
Option Explicit

Private Const TRACE_FPTK_SHEET As String = "FPTK"
Private Const TRACE_LOG_SHEET As String = "FPTK_Corruption_Trace"
Private Const TRACE_HEADER_ROW As Long = 1
Private Const TRACE_FIRST_DATA_ROW As Long = 2

Private mTraceBaseline As Long
Private mTraceReady As Boolean

' ============================================================
' START TRACE
' ============================================================
Public Sub FPTKTrace_ResetAndClean( _
    Optional ByVal stageName As String = "Manual Reset")

    Dim ws As Worksheet
    Dim beforeCount As Long
    Dim deletedCount As Long
    Dim afterCount As Long

    Set ws = FPTKTrace_GetFPTKSheet()

    If ws Is Nothing Then Exit Sub

    beforeCount = FPTKTrace_CountCorruptRowsInternal(ws)

    deletedCount = _
        FPTKTrace_DeleteCorruptRowsInternal( _
            ws, _
            "RESET CLEAN")

    afterCount = FPTKTrace_CountCorruptRowsInternal(ws)

    mTraceBaseline = afterCount
    mTraceReady = True

    FPTKTrace_LogEvent _
        stageName, _
        "RESET", _
        afterCount, _
        "Before=" & CStr(beforeCount) & _
        " | Deleted=" & CStr(deletedCount)

End Sub

' ============================================================
' CHECKPOINT
' Stop compile tepat saat row corrupt pertama kali muncul lagi.
' ============================================================
Public Sub FPTKTrace_Checkpoint( _
    ByVal stageName As String)

    Dim ws As Worksheet
    Dim currentCount As Long

    Set ws = FPTKTrace_GetFPTKSheet()

    If ws Is Nothing Then Exit Sub

    If Not mTraceReady Then

        mTraceBaseline = _
            FPTKTrace_CountCorruptRowsInternal(ws)

        mTraceReady = True

    End If

    currentCount = _
        FPTKTrace_CountCorruptRowsInternal(ws)

    FPTKTrace_LogEvent _
        stageName, _
        "CHECK", _
        currentCount, _
        "Baseline=" & CStr(mTraceBaseline)

    If currentCount > mTraceBaseline Then

        FPTKTrace_DumpCorruptRows _
            ws, _
            stageName

        Err.Raise _
            vbObjectError + 9101, _
            "FPTKTrace_Checkpoint", _
            "ROW FPTK CORRUPT PERTAMA MUNCUL SETELAH:" & _
            vbCrLf & vbCrLf & _
            stageName & _
            vbCrLf & vbCrLf & _
            "Baseline: " & CStr(mTraceBaseline) & _
            vbCrLf & _
            "Current: " & CStr(currentCount) & _
            vbCrLf & vbCrLf & _
            "Lihat sheet '" & _
            TRACE_LOG_SHEET & _
            "'."

    End If

End Sub

' ============================================================
' MANUAL TOOLS
' ============================================================
Public Sub FPTKTrace_CheckNow()

    Dim ws As Worksheet
    Dim corruptCount As Long

    Set ws = FPTKTrace_GetFPTKSheet()

    If ws Is Nothing Then

        MsgBox _
            "Sheet FPTK tidak ditemukan.", _
            vbExclamation, _
            "FPTK Trace"

        Exit Sub

    End If

    corruptCount = _
        FPTKTrace_CountCorruptRowsInternal(ws)

    FPTKTrace_DumpCorruptRows _
        ws, _
        "MANUAL CHECK"

    MsgBox _
        "Corrupt/helper rows: " & _
        CStr(corruptCount) & _
        vbCrLf & _
        "Detail ada di sheet " & _
        TRACE_LOG_SHEET & _
        ".", _
        vbInformation, _
        "FPTK Trace"

End Sub

Public Sub FPTKTrace_CleanNow()

    Dim ws As Worksheet
    Dim deletedCount As Long

    Set ws = FPTKTrace_GetFPTKSheet()

    If ws Is Nothing Then

        MsgBox _
            "Sheet FPTK tidak ditemukan.", _
            vbExclamation, _
            "FPTK Trace"

        Exit Sub

    End If

    FPTKTrace_DumpCorruptRows _
        ws, _
        "BEFORE MANUAL CLEAN"

    deletedCount = _
        FPTKTrace_DeleteCorruptRowsInternal( _
            ws, _
            "MANUAL CLEAN")

    mTraceBaseline = _
        FPTKTrace_CountCorruptRowsInternal(ws)

    mTraceReady = True

    MsgBox _
        "Row helper/corrupt dihapus: " & _
        CStr(deletedCount) & _
        vbCrLf & _
        "Sisa: " & _
        CStr(mTraceBaseline), _
        vbInformation, _
        "FPTK Trace"

End Sub

' ============================================================
' CORRUPT DETECTION
'
' Sangat konservatif.
' Contoh target:
' Date Real = CORPRec09
' Kode Angka = MPRec19
' Kode Unik = MPRecruiter02
' Posisi = MPRec20
' PIC Recruiter = CORPRecruiter13
'
' Row normal seperti CORPRec1514146023 tidak kena.
' ============================================================
Private Function FPTKTrace_RowIsCorrupt( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long) As Boolean

    Dim colDateReal As Long
    Dim colKodeAngka As Long
    Dim colKodeUnik As Long
    Dim colPosisi As Long
    Dim colPIC As Long

    Dim dateRealText As String
    Dim kodeAngkaText As String
    Dim kodeUnikText As String
    Dim posisiText As String
    Dim picText As String

    Dim score As Long

    colDateReal = FPTKTrace_FindHeader(ws, "FPTK Date (Real)")
    colKodeAngka = FPTKTrace_FindHeader(ws, "Kode Angka")
    colKodeUnik = FPTKTrace_FindHeader(ws, "Kode Unik")
    colPosisi = FPTKTrace_FindHeader(ws, "Posisi")
    colPIC = FPTKTrace_FindHeader(ws, "PIC Recruiter")

    If colKodeUnik <= 0 Or colPosisi <= 0 Then Exit Function

    dateRealText = FPTKTrace_CellText(ws, rowIndex, colDateReal)
    kodeAngkaText = FPTKTrace_CellText(ws, rowIndex, colKodeAngka)
    kodeUnikText = FPTKTrace_CellText(ws, rowIndex, colKodeUnik)
    posisiText = FPTKTrace_CellText(ws, rowIndex, colPosisi)
    picText = FPTKTrace_CellText(ws, rowIndex, colPIC)

    If Len(kodeUnikText) = 0 And _
       Len(posisiText) = 0 Then

        Exit Function

    End If

    If Len(kodeUnikText) > 0 Then

        If FPTKTrace_CountDigits(kodeUnikText) < 3 Then
            score = score + 2
        End If

    End If

    If FPTKTrace_LooksLikePICCode(dateRealText) Then
        score = score + 1
    End If

    If Len(kodeAngkaText) > 0 Then

        If Not IsNumeric(kodeAngkaText) Then

            If FPTKTrace_LooksLikePICCode(kodeAngkaText) Then
                score = score + 1
            End If

        End If

    End If

    If FPTKTrace_LooksLikePICCode(posisiText) Then
        score = score + 2
    End If

    If FPTKTrace_LooksLikePICCode(picText) Then
        score = score + 1
    End If

    If score >= 3 Then
        FPTKTrace_RowIsCorrupt = True
    End If

End Function

Private Function FPTKTrace_CountCorruptRowsInternal( _
    ByVal ws As Worksheet) As Long

    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = FPTKTrace_LastUsedRow(ws)

    For rowIndex = TRACE_FIRST_DATA_ROW To lastRow

        If FPTKTrace_RowIsCorrupt(ws, rowIndex) Then

            FPTKTrace_CountCorruptRowsInternal = _
                FPTKTrace_CountCorruptRowsInternal + 1

        End If

    Next rowIndex

End Function

Private Function FPTKTrace_DeleteCorruptRowsInternal( _
    ByVal ws As Worksheet, _
    ByVal reasonText As String) As Long

    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = FPTKTrace_LastUsedRow(ws)

    For rowIndex = lastRow To TRACE_FIRST_DATA_ROW Step -1

        If FPTKTrace_RowIsCorrupt(ws, rowIndex) Then

            FPTKTrace_LogRow _
                ws, _
                rowIndex, _
                reasonText

            ws.rows(rowIndex).Delete

            FPTKTrace_DeleteCorruptRowsInternal = _
                FPTKTrace_DeleteCorruptRowsInternal + 1

        End If

    Next rowIndex

End Function

Private Sub FPTKTrace_DumpCorruptRows( _
    ByVal ws As Worksheet, _
    ByVal stageName As String)

    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = FPTKTrace_LastUsedRow(ws)

    For rowIndex = TRACE_FIRST_DATA_ROW To lastRow

        If FPTKTrace_RowIsCorrupt(ws, rowIndex) Then

            FPTKTrace_LogRow _
                ws, _
                rowIndex, _
                stageName

        End If

    Next rowIndex

End Sub

' ============================================================
' PIC CODE DETECTION
' ============================================================
Private Function FPTKTrace_LooksLikePICCode( _
    ByVal rawText As String) As Boolean

    Dim textValue As String
    Dim i As Long
    Dim ch As String

    textValue = UCase$(Trim$(rawText))
    textValue = Replace$(textValue, " ", vbNullString)

    If Len(textValue) < 4 Or Len(textValue) > 14 Then Exit Function
    If FPTKTrace_CountDigits(textValue) > 0 Then Exit Function

    For i = 1 To Len(textValue)

        ch = Mid$(textValue, i, 1)

        If ch < "A" Or ch > "Z" Then Exit Function

    Next i

    Select Case True
        Case Left$(textValue, 4) = "CORP"
            FPTKTrace_LooksLikePICCode = True
        Case Left$(textValue, 4) = "JESS"
            FPTKTrace_LooksLikePICCode = True
        Case Left$(textValue, 3) = "CMD"
            FPTKTrace_LooksLikePICCode = True
        Case Left$(textValue, 3) = "BHC"
            FPTKTrace_LooksLikePICCode = True
        Case Left$(textValue, 3) = "ARC"
            FPTKTrace_LooksLikePICCode = True
        Case Left$(textValue, 2) = "MP"
            FPTKTrace_LooksLikePICCode = True
        Case Left$(textValue, 2) = "MS"
            FPTKTrace_LooksLikePICCode = True
    End Select

End Function

Private Function FPTKTrace_CountDigits( _
    ByVal sourceText As String) As Long

    Dim i As Long
    Dim ch As String

    For i = 1 To Len(sourceText)

        ch = Mid$(sourceText, i, 1)

        If ch >= "0" And ch <= "9" Then

            FPTKTrace_CountDigits = _
                FPTKTrace_CountDigits + 1

        End If

    Next i

End Function

' ============================================================
' HEADER / CELL HELPERS
' ============================================================
Private Function FPTKTrace_FindHeader( _
    ByVal ws As Worksheet, _
    ByVal requestedHeader As String) As Long

    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim wanted As String
    Dim currentKey As String

    wanted = FPTKTrace_NormalizeHeader(requestedHeader)

    lastColumn = _
        ws.Cells( _
            TRACE_HEADER_ROW, _
            ws.Columns.Count).End(xlToLeft).Column

    For columnIndex = 1 To lastColumn

        currentKey = _
            FPTKTrace_NormalizeHeader( _
                ws.Cells( _
                    TRACE_HEADER_ROW, _
                    columnIndex).Value2)

        If StrComp(currentKey, wanted, vbTextCompare) = 0 Then

            FPTKTrace_FindHeader = columnIndex
            Exit Function

        End If

    Next columnIndex

End Function

Private Function FPTKTrace_NormalizeHeader( _
    ByVal rawValue As Variant) As String

    Dim sourceText As String
    Dim resultText As String
    Dim i As Long
    Dim ch As String
    Dim codeValue As Long

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    sourceText = UCase$(CStr(rawValue))
    sourceText = Replace$(sourceText, Chr$(160), " ")

    For i = 1 To Len(sourceText)

        ch = Mid$(sourceText, i, 1)
        codeValue = AscW(ch)

        If codeValue >= 48 And codeValue <= 57 Then
            resultText = resultText & ch
        ElseIf codeValue >= 65 And codeValue <= 90 Then
            resultText = resultText & ch
        End If

    Next i

    FPTKTrace_NormalizeHeader = resultText

End Function

Private Function FPTKTrace_CellText( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long) As String

    If columnIndex <= 0 Then Exit Function

    If IsError(ws.Cells(rowIndex, columnIndex).Value2) Then Exit Function

    FPTKTrace_CellText = _
        Trim$(CStr(ws.Cells(rowIndex, columnIndex).Value2))

End Function

Private Function FPTKTrace_LastUsedRow( _
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
        FPTKTrace_LastUsedRow = TRACE_HEADER_ROW
    Else
        FPTKTrace_LastUsedRow = foundCell.Row
    End If

End Function

Private Function FPTKTrace_GetFPTKSheet() As Worksheet

    On Error Resume Next

    Set FPTKTrace_GetFPTKSheet = _
        ThisWorkbook.Worksheets(TRACE_FPTK_SHEET)

    On Error GoTo 0

End Function

' ============================================================
' LOGGING
' ============================================================
Private Sub FPTKTrace_LogEvent( _
    ByVal stageName As String, _
    ByVal eventType As String, _
    ByVal corruptCount As Long, _
    ByVal detailText As String)

    Dim wsLog As Worksheet
    Dim NextRow As Long

    Set wsLog = FPTKTrace_GetOrCreateLogSheet()

    NextRow = _
        wsLog.Cells( _
            wsLog.rows.Count, _
            1).End(xlUp).Row + 1

    If NextRow < 2 Then NextRow = 2

    wsLog.Cells(NextRow, 1).Value = Now
    wsLog.Cells(NextRow, 2).Value2 = stageName
    wsLog.Cells(NextRow, 3).Value2 = eventType
    wsLog.Cells(NextRow, 4).Value2 = corruptCount
    wsLog.Cells(NextRow, 5).Value2 = detailText

    On Error Resume Next

    AddRunLog _
        "Trace", _
        "FPTK CORRUPTION | " & _
        stageName & _
        " | " & _
        eventType & _
        " | Count=" & _
        CStr(corruptCount) & _
        " | " & _
        detailText

    On Error GoTo 0

End Sub

Private Sub FPTKTrace_LogRow( _
    ByVal ws As Worksheet, _
    ByVal sourceRow As Long, _
    ByVal stageName As String)

    Dim wsLog As Worksheet
    Dim NextRow As Long

    Set wsLog = FPTKTrace_GetOrCreateLogSheet()

    NextRow = _
        wsLog.Cells( _
            wsLog.rows.Count, _
            1).End(xlUp).Row + 1

    If NextRow < 2 Then NextRow = 2

    wsLog.Cells(NextRow, 1).Value = Now
    wsLog.Cells(NextRow, 2).Value2 = stageName
    wsLog.Cells(NextRow, 3).Value2 = sourceRow

    wsLog.Cells(NextRow, 4).Value2 = _
        FPTKTrace_CellText( _
            ws, sourceRow, _
            FPTKTrace_FindHeader(ws, "Kode PIC"))

    wsLog.Cells(NextRow, 5).Value2 = _
        FPTKTrace_CellText( _
            ws, sourceRow, _
            FPTKTrace_FindHeader(ws, "FPTK Date (Real)"))

    wsLog.Cells(NextRow, 6).Value2 = _
        FPTKTrace_CellText( _
            ws, sourceRow, _
            FPTKTrace_FindHeader(ws, "Kode Angka"))

    wsLog.Cells(NextRow, 7).Value2 = _
        FPTKTrace_CellText( _
            ws, sourceRow, _
            FPTKTrace_FindHeader(ws, "Kode Unik"))

    wsLog.Cells(NextRow, 8).Value2 = _
        FPTKTrace_CellText( _
            ws, sourceRow, _
            FPTKTrace_FindHeader(ws, "Posisi"))

    wsLog.Cells(NextRow, 9).Value2 = _
        FPTKTrace_CellText( _
            ws, sourceRow, _
            FPTKTrace_FindHeader(ws, "PIC Recruiter"))

    wsLog.Cells(NextRow, 10).Value2 = _
        FPTKTrace_CellText( _
            ws, sourceRow, _
            FPTKTrace_FindHeader(ws, "Last Compile Action"))

    wsLog.Cells(NextRow, 11).Value2 = _
        FPTKTrace_CellText( _
            ws, sourceRow, _
            FPTKTrace_FindHeader(ws, "Source File"))

End Sub

Private Function FPTKTrace_GetOrCreateLogSheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TRACE_LOG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = _
            ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        ws.Name = TRACE_LOG_SHEET

        ws.Cells(1, 1).Value2 = "Timestamp"
        ws.Cells(1, 2).Value2 = "Stage"
        ws.Cells(1, 3).Value2 = "Row / Event"
        ws.Cells(1, 4).Value2 = "Kode PIC / Count"
        ws.Cells(1, 5).Value2 = "FPTK Date Real / Detail"
        ws.Cells(1, 6).Value2 = "Kode Angka"
        ws.Cells(1, 7).Value2 = "Kode Unik"
        ws.Cells(1, 8).Value2 = "Posisi"
        ws.Cells(1, 9).Value2 = "PIC Recruiter"
        ws.Cells(1, 10).Value2 = "Last Compile Action"
        ws.Cells(1, 11).Value2 = "Source File"

        ws.rows(1).Font.Bold = True

    End If

    Set FPTKTrace_GetOrCreateLogSheet = ws

End Function


