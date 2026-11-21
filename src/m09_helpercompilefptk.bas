Attribute VB_Name = "m09_helpercompilefptk"
Option Explicit

Private Const FPTK_SHEET_NAME As String = "FPTK"
Private Const FPTK_HEADER_ROW As Long = 1
Private Const FPTK_FIRST_DATA_ROW As Long = 2

' ============================================================
' REFRESH FPTK DERIVED VALUES
'
' Module ini berdiri sendiri dan tidak bergantung pada function
' private di module Compile_FPTK.
'
' Hasil ditulis sebagai VALUE, bukan formula.
' Semua kolom dicari berdasarkan nama header.
' ============================================================
Public Sub RefreshFPTKDerivedValues()
    SchemaB02Prepare "FPTK"

    Dim ws As Worksheet
    Dim headerMap As Object
    Dim lastRow As Long
    Dim rowIndex As Long

    Dim colKodePIC As Long
    Dim colKodeUnik As Long
    Dim colPosisi As Long
    Dim colLevelNumber As Long
    Dim colFilterKategori As Long

    Dim colFPTKDateKode As Long
    Dim colWeekFPTK As Long
    Dim colMonthFPTK As Long

    Dim colCancelDate As Long
    Dim colWeekCancel As Long
    Dim colMonthCancel As Long

    Dim colOfferingDate As Long
    Dim colWeekOffering As Long
    Dim colMonthOffering As Long

    Dim colFPTKDateReal As Long
    Dim colJumlahSLA As Long
    Dim colDeadline As Long
    Dim colDetailSLA As Long
    Dim colStatus As Long

    Dim positionText As String
    Dim positionLower As String
    Dim statusText As String

    Dim levelNumber As Long
    Dim slaDays As Long

    Dim sourceDate As Date
    Dim deadlineDate As Date
    Dim offeringDate As Date

    On Error GoTo ErrorHandler

    Set ws = Nothing

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets( _
        FPTK_SHEET_NAME)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Err.Raise vbObjectError + 7001, _
                  "RefreshFPTKDerivedValues", _
                  "Sheet FPTK tidak ditemukan."
    End If

    Application.StatusBar = _
        "Finalizing FPTK helper values..."

    FHV_UnlistAllTables ws

    ' --------------------------------------------------------
    ' GABUNG HEADER ALIAS KE HEADER UTAMA
    ' --------------------------------------------------------
    FHV_MergeAliasGroup _
        ws, _
        "Kode BU", _
        Array("Kode BU2", "Kode BU 2")

    FHV_MergeAliasGroup _
        ws, _
        "Jumlah SLA", _
        Array("SLA Days")

    FHV_MergeAliasGroup _
        ws, _
        "Deadline pemenuhan SLA", _
        Array("SLA Deadline", "Deadline SLA")

    FHV_MergeAliasGroup _
        ws, _
        "Detail SLA", _
        Array("SLA Status", "Status SLA")

    FHV_MergeAliasGroup _
        ws, _
        "FPTK Date (Kode)", _
        Array( _
            "FPTK Date", _
            "FPTK Date (CODE)", _
            "FPTK Date (Pake Ini)")

    FHV_MergeAliasGroup _
        ws, _
        "Level FPTK", _
        Array("Level FPTK (Sesuai SO)")

    FHV_MergeAliasGroup _
        ws, _
        "Week FPTK Date (Kode)", _
        Array( _
            "Week FPTK Date", _
            "Week FPTK Date (CODE)")

    ' --------------------------------------------------------
    ' PASTIKAN HEADER TARGET ADA
    ' --------------------------------------------------------
    FHV_EnsureHeader _
        ws, "Filter Kategorisasi FPTK"

    FHV_EnsureHeader _
        ws, "Week FPTK Date (Kode)"

    FHV_EnsureHeader _
        ws, "Month FPTK Date"

    FHV_EnsureHeader _
        ws, "Week Cancel Date"

    FHV_EnsureHeader _
        ws, "Month Cancel Date"

    FHV_EnsureHeader _
        ws, "Week Offering Date"

    FHV_EnsureHeader _
        ws, "Month Offering"

    FHV_EnsureHeader _
        ws, "Jumlah SLA"

    FHV_EnsureHeader _
        ws, "Deadline pemenuhan SLA"

    FHV_EnsureHeader _
        ws, "Detail SLA"

    Set headerMap = FHV_BuildHeaderMap(ws)

    ' --------------------------------------------------------
    ' AMBIL NOMOR KOLOM BERDASARKAN NAMA HEADER
    ' --------------------------------------------------------
    colKodePIC = FHV_MapColumn( _
        headerMap, "Kode PIC")

    colKodeUnik = FHV_MapColumn( _
        headerMap, "Kode Unik")

    colPosisi = FHV_MapColumn( _
        headerMap, "Posisi")

    colLevelNumber = FHV_MapColumn( _
        headerMap, "Level Number")

    colFilterKategori = FHV_MapColumn( _
        headerMap, "Filter Kategorisasi FPTK")

    colFPTKDateKode = FHV_MapColumn( _
        headerMap, "FPTK Date (Kode)")

    colWeekFPTK = FHV_MapColumn( _
        headerMap, "Week FPTK Date (Kode)")

    colMonthFPTK = FHV_MapColumn( _
        headerMap, "Month FPTK Date")

    colCancelDate = FHV_MapColumn( _
        headerMap, "FPTK Cancel Date")

    colWeekCancel = FHV_MapColumn( _
        headerMap, "Week Cancel Date")

    colMonthCancel = FHV_MapColumn( _
        headerMap, "Month Cancel Date")

    colOfferingDate = FHV_MapColumn( _
        headerMap, "Offering Date")

    colWeekOffering = FHV_MapColumn( _
        headerMap, "Week Offering Date")

    colMonthOffering = FHV_MapColumn( _
        headerMap, "Month Offering")

    colFPTKDateReal = FHV_MapColumn( _
        headerMap, "FPTK Date (Real)")

    colJumlahSLA = FHV_MapColumn( _
        headerMap, "Jumlah SLA")

    colDeadline = FHV_MapColumn( _
        headerMap, "Deadline pemenuhan SLA")

    colDetailSLA = FHV_MapColumn( _
        headerMap, "Detail SLA")

    colStatus = FHV_MapColumn( _
        headerMap, "Status")

    lastRow = FHV_LastUsedRow(ws)

    If lastRow < FPTK_FIRST_DATA_ROW Then
        GoTo SafeExit
    End If

    ' --------------------------------------------------------
    ' ISI VALUE DINAMIS
    ' --------------------------------------------------------
    For rowIndex = FPTK_FIRST_DATA_ROW _
        To lastRow

        If FHV_RowHasData( _
            ws, rowIndex, _
            colKodePIC, _
            colKodeUnik, _
            colPosisi) Then

            ' =================================================
            ' FILTER KATEGORISASI FPTK
            '
            ' =IF(OR(LEFT(Posisi,6)="cimory",
            '        LEFT(Posisi,5)="fresh"),
            '     "CLAP FGDP",
            '     IF(OR(Level Number=1,Level Number=2),
            '        "Level 1-2",
            '        IF(Level Number=3,"Level 3",
            '           IF(Level Number=4,"Level 4",""))))
            ' =================================================
            positionText = FHV_CellText( _
                ws, rowIndex, colPosisi)

            positionLower = LCase$( _
                Trim$(positionText))

            levelNumber = FHV_FirstNumber( _
                FHV_CellText( _
                    ws, rowIndex, _
                    colLevelNumber))

            If Left$(positionLower, 6) = "cimory" Or _
               Left$(positionLower, 5) = "fresh" Then

                FHV_SetValue _
                    ws, rowIndex, _
                    colFilterKategori, _
                    "CLAP FGDP"

            Else

                Select Case levelNumber

                    Case 1, 2
                        FHV_SetValue _
                            ws, rowIndex, _
                            colFilterKategori, _
                            "Level 1-2"

                    Case 3
                        FHV_SetValue _
                            ws, rowIndex, _
                            colFilterKategori, _
                            "Level 3"

                    Case 4
                        FHV_SetValue _
                            ws, rowIndex, _
                            colFilterKategori, _
                            "Level 4"

                    Case Else
                        FHV_ClearValue _
                            ws, rowIndex, _
                            colFilterKategori

                End Select

            End If

            ' =================================================
            ' WEEK + MONTH FPTK DATE (KODE)
            ' =================================================
            FHV_WriteWeekAndMonth _
                ws, rowIndex, _
                colFPTKDateKode, _
                colWeekFPTK, _
                colMonthFPTK

            ' =================================================
            ' WEEK + MONTH CANCEL DATE
            ' =================================================
            FHV_WriteWeekAndMonth _
                ws, rowIndex, _
                colCancelDate, _
                colWeekCancel, _
                colMonthCancel

            ' =================================================
            ' WEEK + MONTH OFFERING DATE
            ' =================================================
            FHV_WriteWeekAndMonth _
                ws, rowIndex, _
                colOfferingDate, _
                colWeekOffering, _
                colMonthOffering

            ' =================================================
            ' JUMLAH SLA
            '
            ' Existing numeric value dipertahankan.
            ' Bila kosong:
            ' Level 1-3 = 30
            ' Level 4   = 45
            ' Level 5   = 60
            ' =================================================
            slaDays = 0

            If colJumlahSLA > 0 Then

                If IsNumeric(ws.Cells( _
                    rowIndex, _
                    colJumlahSLA).Value2) Then

                    slaDays = CLng(ws.Cells( _
                        rowIndex, _
                        colJumlahSLA).Value2)

                End If

                If slaDays <= 0 Then

                    slaDays = FHV_SLAByLevel( _
                        levelNumber)

                    ws.Cells( _
                        rowIndex, _
                        colJumlahSLA).Value2 = _
                        slaDays

                End If

            End If

            ' =================================================
            ' DEADLINE PEMENUHAN SLA
            ' = FPTK DATE (REAL) + JUMLAH SLA
            ' =================================================
            If colDeadline > 0 Then

                If FHV_TryParseDate( _
                    ws.Cells( _
                        rowIndex, _
                        colFPTKDateReal).Value2, _
                    sourceDate) And _
                   slaDays > 0 Then

                    deadlineDate = _
                        DateAdd( _
                            "d", _
                            slaDays, _
                            sourceDate)

                    ws.Cells( _
                        rowIndex, _
                        colDeadline).Value = _
                        deadlineDate

                    ws.Cells( _
                        rowIndex, _
                        colDeadline).numberFormat = _
                        "dd/mm/yyyy"

                Else

                    FHV_ClearValue _
                        ws, rowIndex, _
                        colDeadline

                End If

            End If

            ' =================================================
            ' DETAIL SLA
            '
            ' =IF(Status="OP",
            '     IF(Deadline<TODAY(),
            '        "OP Tidak Lulus SLA",
            '        "OP Belum Lewat SLA"),
            '     IF(Status="Closed",
            '        IF(Deadline<Offering Date,
            '           "Closed Tidak Lulus SLA",
            '           "Closed Lulus SLA"),
            '        "Cancel"))
            ' =================================================
            statusText = LCase$( _
                Trim$(FHV_CellText( _
                    ws, rowIndex, _
                    colStatus)))

            deadlineDate = 0
            offeringDate = 0

            If colDeadline > 0 Then

                Call FHV_TryParseDate( _
                    ws.Cells( _
                        rowIndex, _
                        colDeadline).Value2, _
                    deadlineDate)

            End If

            If colOfferingDate > 0 Then

                Call FHV_TryParseDate( _
                    ws.Cells( _
                        rowIndex, _
                        colOfferingDate).Value2, _
                    offeringDate)

            End If

            Select Case statusText

                Case "op", "open"

                    If deadlineDate < Date Then

                        FHV_SetValue _
                            ws, rowIndex, _
                            colDetailSLA, _
                            "OP Tidak Lulus SLA"

                    Else

                        FHV_SetValue _
                            ws, rowIndex, _
                            colDetailSLA, _
                            "OP Belum Lewat SLA"

                    End If

                Case "closed", "close"

                    If deadlineDate < offeringDate Then

                        FHV_SetValue _
                            ws, rowIndex, _
                            colDetailSLA, _
                            "Closed Tidak Lulus SLA"

                    Else

                        FHV_SetValue _
                            ws, rowIndex, _
                            colDetailSLA, _
                            "Closed Lulus SLA"

                    End If

                Case Else

                    FHV_SetValue _
                        ws, rowIndex, _
                        colDetailSLA, _
                        "Cancel"

            End Select

        End If

    Next rowIndex

    ' --------------------------------------------------------
    ' PENGAMAN TERAKHIR: HAPUS ALIAS LAGI
    ' --------------------------------------------------------
    FHV_MergeAliasGroup _
        ws, "Kode BU", _
        Array("Kode BU2", "Kode BU 2")

    FHV_MergeAliasGroup _
        ws, "Jumlah SLA", _
        Array("SLA Days")

    FHV_MergeAliasGroup _
        ws, "Deadline pemenuhan SLA", _
        Array("SLA Deadline", "Deadline SLA")

    FHV_MergeAliasGroup _
        ws, "Detail SLA", _
        Array("SLA Status", "Status SLA")

    FHV_MergeAliasGroup _
        ws, "FPTK Date (Kode)", _
        Array( _
            "FPTK Date", _
            "FPTK Date (CODE)", _
            "FPTK Date (Pake Ini)")

    FHV_MergeAliasGroup _
        ws, "Level FPTK", _
        Array("Level FPTK (Sesuai SO)")

    FHV_MergeAliasGroup _
        ws, "Week FPTK Date (Kode)", _
        Array( _
            "Week FPTK Date", _
            "Week FPTK Date (CODE)")

    FHV_ApplyDateFormats ws
    SchemaEnsureSheet "FPTK"

SafeExit:

    Application.StatusBar = False
    Set headerMap = Nothing
    Exit Sub

ErrorHandler:

    Application.StatusBar = False

    MsgBox _
        "Refresh FPTK helper values gagal." & _
        vbCrLf & _
        "Error " & CStr(Err.Number) & _
        ": " & Err.Description, _
        vbCritical, _
        "FPTK Helper Values"

End Sub

Public Sub RefreshFPTKDerivedValues_Manual()

    RefreshFPTKDerivedValues

    MsgBox _
        "Filter kategori, week/month, deadline, dan Detail SLA " & _
        "sudah diperbarui.", _
        vbInformation, _
        "FPTK Helper Values"

End Sub

' ============================================================
' HEADER MAP
' ============================================================
Private Function FHV_BuildHeaderMap( _
    ByVal ws As Worksheet) As Object

    Dim result As Object
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim headerKey As String

    Set result = CreateObject( _
        "Scripting.Dictionary")

    result.CompareMode = vbTextCompare

    lastColumn = FHV_LastHeaderColumn(ws)

    For columnIndex = 1 To lastColumn

        headerKey = FHV_NormalizeHeader( _
            ws.Cells( _
                FPTK_HEADER_ROW, _
                columnIndex).Value2)

        If Len(headerKey) > 0 Then

            If Not result.Exists( _
                headerKey) Then

                result.Add _
                    headerKey, _
                    columnIndex

            End If

        End If

    Next columnIndex

    Set FHV_BuildHeaderMap = result

End Function

Private Function FHV_MapColumn( _
    ByVal headerMap As Object, _
    ByVal headerName As String) As Long

    Dim headerKey As String

    If headerMap Is Nothing Then Exit Function

    headerKey = SchemaB02CanonicalKey( _
        "FPTK", headerName)

    If headerMap.Exists(headerKey) Then

        FHV_MapColumn = CLng( _
            headerMap(headerKey))

    End If

End Function

Private Sub FHV_EnsureHeader( _
    ByVal ws As Worksheet, _
    ByVal headerName As String)

    Dim columnNumber As Long

    columnNumber = SchemaB02FindColumn( _
        ws, headerName, True)

End Sub

Private Function FHV_FindHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal headerName As String) As Long

    Dim schemaColumn As Long
    Dim requestedKey As String

    schemaColumn = SchemaB02FindColumn( _
        ws, headerName, False)

    If schemaColumn > 0 Then
        FHV_FindHeaderColumn = schemaColumn
        Exit Function
    End If
    Dim currentKey As String
    Dim lastColumn As Long
    Dim columnIndex As Long

    requestedKey = FHV_NormalizeHeader( _
        headerName)

    lastColumn = FHV_LastHeaderColumn(ws)

    For columnIndex = 1 To lastColumn

        currentKey = FHV_NormalizeHeader( _
            ws.Cells( _
                FPTK_HEADER_ROW, _
                columnIndex).Value2)

        If StrComp( _
            currentKey, _
            requestedKey, _
            vbTextCompare) = 0 Then

            FHV_FindHeaderColumn = _
                columnIndex

            Exit Function

        End If

    Next columnIndex

End Function

Private Function FHV_FindRightmostHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal headerName As String, _
    ByVal excludedColumn As Long) As Long

    Dim requestedKey As String
    Dim currentKey As String
    Dim lastColumn As Long
    Dim columnIndex As Long

    requestedKey = FHV_NormalizeHeader( _
        headerName)

    lastColumn = FHV_LastHeaderColumn(ws)

    For columnIndex = lastColumn _
        To 1 Step -1

        If columnIndex <> excludedColumn Then

            currentKey = FHV_NormalizeHeader( _
                ws.Cells( _
                    FPTK_HEADER_ROW, _
                    columnIndex).Value2)

            If StrComp( _
                currentKey, _
                requestedKey, _
                vbTextCompare) = 0 Then

                FHV_FindRightmostHeaderColumn = _
                    columnIndex

                Exit Function

            End If

        End If

    Next columnIndex

End Function

Private Function FHV_NormalizeHeader( _
    ByVal rawValue As Variant) As String

    Dim sourceText As String
    Dim resultText As String
    Dim characterIndex As Long
    Dim currentCharacter As String
    Dim characterCode As Long

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    sourceText = UCase$(CStr(rawValue))
    sourceText = Replace$( _
        sourceText, Chr$(160), " ")

    For characterIndex = 1 To Len(sourceText)

        currentCharacter = Mid$( _
            sourceText, _
            characterIndex, 1)

        characterCode = AscW( _
            currentCharacter)

        If characterCode >= 48 And _
           characterCode <= 57 Then

            resultText = resultText & _
                currentCharacter

        ElseIf characterCode >= 65 And _
               characterCode <= 90 Then

            resultText = resultText & _
                currentCharacter

        End If

    Next characterIndex

    FHV_NormalizeHeader = resultText

End Function

' ============================================================
' ALIAS MERGE
' ============================================================
Private Sub FHV_MergeAliasGroup( _
    ByVal ws As Worksheet, _
    ByVal primaryHeader As String, _
    ByVal aliasHeaders As Variant)

    Dim primaryColumn As Long
    Dim aliasColumn As Long
    Dim aliasItem As Variant
    Dim lastRow As Long

    primaryColumn = FHV_FindHeaderColumn( _
        ws, primaryHeader)

    If primaryColumn = 0 Then

        For Each aliasItem In aliasHeaders

            aliasColumn = FHV_FindHeaderColumn( _
                ws, CStr(aliasItem))

            If aliasColumn > 0 Then

                ws.Cells( _
                    FPTK_HEADER_ROW, _
                    aliasColumn).Value = _
                    primaryHeader

                primaryColumn = aliasColumn
                Exit For

            End If

        Next aliasItem

    End If

    If primaryColumn = 0 Then Exit Sub

    lastRow = FHV_LastUsedRow(ws)

    Do

        aliasColumn = 0

        For Each aliasItem In aliasHeaders

            aliasColumn = _
                FHV_FindRightmostHeaderColumn( _
                    ws, _
                    CStr(aliasItem), _
                    primaryColumn)

            If aliasColumn > 0 Then
                Exit For
            End If

        Next aliasItem

        If aliasColumn = 0 Then
            Exit Do
        End If

        FHV_CopyAliasIntoPrimary _
            ws, _
            primaryColumn, _
            aliasColumn, _
            lastRow

        ws.Columns(aliasColumn).Delete _
            Shift:=xlToLeft

        If aliasColumn < primaryColumn Then
            primaryColumn = primaryColumn - 1
        End If

    Loop

End Sub

Private Sub FHV_CopyAliasIntoPrimary( _
    ByVal ws As Worksheet, _
    ByVal primaryColumn As Long, _
    ByVal aliasColumn As Long, _
    ByVal lastRow As Long)

    Dim rowIndex As Long
    Dim primaryValue As Variant
    Dim aliasValue As Variant

    For rowIndex = FPTK_FIRST_DATA_ROW _
        To lastRow

        primaryValue = ws.Cells( _
            rowIndex, _
            primaryColumn).Formula

        aliasValue = ws.Cells( _
            rowIndex, _
            aliasColumn).Formula

        If Len(Trim$(CStr( _
            primaryValue))) = 0 Then

            If Len(Trim$(CStr( _
                aliasValue))) > 0 Then

                ws.Cells( _
                    rowIndex, _
                    primaryColumn).Formula = _
                    aliasValue

            End If

        End If

    Next rowIndex

End Sub

' ============================================================
' WEEK + MONTH
' ============================================================
Private Sub FHV_WriteWeekAndMonth( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal sourceDateColumn As Long, _
    ByVal weekTargetColumn As Long, _
    ByVal monthTargetColumn As Long)

    Dim parsedDate As Date

    If sourceDateColumn <= 0 Then Exit Sub

    If FHV_TryParseDate( _
        ws.Cells( _
            rowIndex, _
            sourceDateColumn).Value2, _
        parsedDate) Then

        If weekTargetColumn > 0 Then

            On Error Resume Next

            ws.Cells( _
                rowIndex, _
                weekTargetColumn).Value2 = _
                Application.WorksheetFunction. _
                    WeekNum(parsedDate, 1)

            If Err.Number <> 0 Then

                Err.Clear

                ws.Cells( _
                    rowIndex, _
                    weekTargetColumn).Value2 = _
                    DatePart( _
                        "ww", _
                        parsedDate, _
                        vbSunday, _
                        vbFirstJan1)

            End If

            On Error GoTo 0

        End If

        If monthTargetColumn > 0 Then

            ws.Cells( _
                rowIndex, _
                monthTargetColumn).Value2 = _
                Format$(parsedDate, "mmmm")

        End If

    Else

        FHV_ClearValue _
            ws, rowIndex, _
            weekTargetColumn

        FHV_ClearValue _
            ws, rowIndex, _
            monthTargetColumn

    End If

End Sub

' ============================================================
' DATE PARSER DD-MM-YYYY / DD/MM/YYYY / EXCEL SERIAL
' ============================================================
Private Function FHV_TryParseDate( _
    ByVal rawValue As Variant, _
    ByRef parsedDate As Date) As Boolean

    Dim textValue As String
    Dim dateParts() As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long

    On Error GoTo InvalidDate

    parsedDate = 0

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    If VarType(rawValue) = vbDate Then

        parsedDate = dateValue(CDate(rawValue))
        FHV_TryParseDate = True
        Exit Function

    End If

    If IsNumeric(rawValue) Then

        If CDbl(rawValue) > 0 Then

            parsedDate = dateValue( _
                CDate(CDbl(rawValue)))

            FHV_TryParseDate = True
            Exit Function

        End If

    End If

    textValue = Trim$(CStr(rawValue))

    If Len(textValue) = 0 Then
        Exit Function
    End If

    textValue = Replace$( _
        textValue, "-", "/")

    textValue = Replace$( _
        textValue, ".", "/")

    textValue = Replace$( _
        textValue, "\", "/")

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

            parsedDate = DateSerial( _
                yearValue, _
                monthValue, _
                dayValue)

            FHV_TryParseDate = True
            Exit Function

        End If

    End If

    If IsDate(textValue) Then

        parsedDate = dateValue( _
            CDate(textValue))

        FHV_TryParseDate = True
        Exit Function

    End If

InvalidDate:

    parsedDate = 0
    FHV_TryParseDate = False

End Function

' ============================================================
' GENERAL HELPERS
' ============================================================
Private Function FHV_LastUsedRow( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

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

        FHV_LastUsedRow = _
            FPTK_HEADER_ROW

    Else

        FHV_LastUsedRow = _
            foundCell.Row

    End If

End Function

Private Function FHV_LastHeaderColumn( _
    ByVal ws As Worksheet) As Long

    FHV_LastHeaderColumn = ws.Cells( _
        FPTK_HEADER_ROW, _
        ws.Columns.Count).End(xlToLeft).Column

    If FHV_LastHeaderColumn < 1 Then
        FHV_LastHeaderColumn = 1
    End If

End Function

Private Function FHV_RowHasData( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal colKodePIC As Long, _
    ByVal colKodeUnik As Long, _
    ByVal colPosisi As Long) As Boolean

    If Len(FHV_CellText( _
        ws, rowIndex, colKodePIC)) > 0 Then

        FHV_RowHasData = True
        Exit Function

    End If

    If Len(FHV_CellText( _
        ws, rowIndex, colKodeUnik)) > 0 Then

        FHV_RowHasData = True
        Exit Function

    End If

    If Len(FHV_CellText( _
        ws, rowIndex, colPosisi)) > 0 Then

        FHV_RowHasData = True

    End If

End Function

Private Function FHV_CellText( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long) As String

    If columnIndex <= 0 Then Exit Function

    If IsError(ws.Cells( _
        rowIndex, columnIndex).Value) Then

        Exit Function

    End If

    FHV_CellText = Trim$(CStr( _
        ws.Cells( _
            rowIndex, _
            columnIndex).Value2))

End Function

Private Function FHV_FirstNumber( _
    ByVal sourceText As String) As Long

    Dim characterIndex As Long
    Dim currentCharacter As String
    Dim numberText As String

    For characterIndex = 1 To Len(sourceText)

        currentCharacter = Mid$( _
            sourceText, _
            characterIndex, 1)

        If currentCharacter >= "0" And _
           currentCharacter <= "9" Then

            numberText = numberText & _
                currentCharacter

        ElseIf Len(numberText) > 0 Then

            Exit For

        End If

    Next characterIndex

    If Len(numberText) > 0 Then

        FHV_FirstNumber = CLng(numberText)

    End If

End Function

Private Function FHV_SLAByLevel( _
    ByVal levelNumber As Long) As Long

    Select Case levelNumber

        Case 1 To 3
            FHV_SLAByLevel = 30

        Case 4
            FHV_SLAByLevel = 45

        Case 5
            FHV_SLAByLevel = 60

        Case Else
            FHV_SLAByLevel = 0

    End Select

End Function

Private Sub FHV_SetValue( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long, _
    ByVal newValue As Variant)

    If columnIndex <= 0 Then Exit Sub

    ws.Cells( _
        rowIndex, _
        columnIndex).Value2 = _
        newValue

End Sub

Private Sub FHV_ClearValue( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long)

    If columnIndex <= 0 Then Exit Sub

    ws.Cells( _
        rowIndex, _
        columnIndex).ClearContents

End Sub

Private Sub FHV_UnlistAllTables( _
    ByVal ws As Worksheet)

    Dim tableIndex As Long

    For tableIndex = ws.ListObjects.Count _
        To 1 Step -1

        On Error Resume Next
        ws.ListObjects(tableIndex).Unlist
        On Error GoTo 0

    Next tableIndex

    If ws.AutoFilterMode Then

        On Error Resume Next
        ws.AutoFilterMode = False
        On Error GoTo 0

    End If

End Sub

Private Sub FHV_ApplyDateFormats( _
    ByVal ws As Worksheet)

    Dim dateHeaders As Variant
    Dim headerItem As Variant
    Dim columnIndex As Long

    dateHeaders = Array( _
        "FPTK Date (Real)", _
        "FPTK Date (Kode)", _
        "FPTK Cancel Date", _
        "Offering Date", _
        "Deadline pemenuhan SLA", _
        "Estimasi Join", _
        "Tanggal Upload ke Website")

    For Each headerItem In dateHeaders

        columnIndex = FHV_FindHeaderColumn( _
            ws, CStr(headerItem))

        If columnIndex > 0 Then

            ws.Columns( _
                columnIndex).numberFormat = _
                "dd/mm/yyyy"

        End If

    Next headerItem

End Sub




