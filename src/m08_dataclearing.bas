Attribute VB_Name = "m08_dataclearing"
Option Explicit

' ============================================================
' MANUAL CLEAR ONLY
'
' ClearReportData:
' - boleh dijalankan dari button Clear;
' - TIDAK boleh berjalan saat compiler aktif;
' - tetap dapat menghapus data FPTK;
' - header FPTK dipertahankan.
'
' FPTK first data row: 2
' ============================================================
Public Sub ClearReportData()

    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    Dim sheetNames As Variant
    Dim sheetIndex As Long
    Dim currentSheetName As String
    Dim firstDataRow As Long

    Dim errorNumber As Long
    Dim errorDescription As String

    ' Pengaman:
    ' walaupun ada module lain yang tidak sengaja memanggil
    ' ClearReportData saat compile, proses clear diblokir.
    If gSchemaCompileInProgress Then

        On Error Resume Next

        AddRunLog _
            "Warning", _
            "ClearReportData diblokir karena compiler sedang berjalan."

        On Error GoTo 0

        Exit Sub

    End If

    On Error GoTo ErrorHandler

    oldCalculation = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Membersihkan data report..."

    sheetNames = Array( _
        "FPTK", _
        "DB Kode Posisi", _
        "DB Sourcing", _
        "Blacklist Candidate", _
        "Funneling", _
        "Funneling - Updated")

    For sheetIndex = LBound(sheetNames) _
        To UBound(sheetNames)

        currentSheetName = _
            CStr(sheetNames(sheetIndex))

        If StrComp( _
            currentSheetName, _
            "FPTK", _
            vbTextCompare) = 0 Then

            firstDataRow = 2

        Else

            firstDataRow = 2

        End If

        Application.StatusBar = _
            "Membersihkan sheet: " & _
            currentSheetName

        ResetDataSheet _
            currentSheetName, _
            firstDataRow

        DoEvents

    Next sheetIndex

    ' Monitoring Sourcing tidak dibersihkan agar evidence
    ' dan hyperlink tetap dapat dipertahankan.
    ClearRunLog

SafeExit:

    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.StatusBar = False

    Exit Sub

ErrorHandler:

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    AddRunLog _
        "Error", _
        "ClearReportData | Error " & _
        CStr(errorNumber) & ": " & _
        errorDescription

    On Error GoTo 0

    Resume SafeExit

End Sub

Private Sub ResetDataSheet( _
    ByVal sheetName As String, _
    ByVal firstDataRow As Long)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tableIndex As Long

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets( _
        sheetName)

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    For tableIndex = ws.ListObjects.Count _
        To 1 Step -1

        On Error Resume Next
        ws.ListObjects(tableIndex).Unlist
        On Error GoTo 0

    Next tableIndex

    lastRow = LastUsedRowAny(ws)

    If lastRow < firstDataRow Then Exit Sub

    ws.rows( _
        CStr(firstDataRow) & ":" & _
        CStr(lastRow)).Delete _
        Shift:=xlUp

End Sub

Private Sub ClearRunLog()

    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets( _
        "Run_Log")

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    lastRow = LastUsedRowAny(ws)

    If lastRow >= 2 Then

        ws.Range( _
            ws.Cells(2, 1), _
            ws.Cells(lastRow, 4)).ClearContents

    End If

End Sub

Private Function LastUsedRowAny( _
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

        LastUsedRowAny = 1

    Else

        LastUsedRowAny = foundCell.Row

    End If

End Function


