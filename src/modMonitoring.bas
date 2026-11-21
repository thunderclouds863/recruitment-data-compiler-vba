Attribute VB_Name = "modMonitoring"
Option Explicit

Public SelectedEvidenceFolder As String

Private Const SOURCE_HEADER_ROW As Long = 1
Private Const SOURCE_FIRST_DATA_ROW As Long = 2

' ==================== NORMALIZER ====================
Private Function NormalizeKey(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Trim(s)
    NormalizeKey = UCase(s)
End Function

' ==================== DATE PARSER ====================
Private Function ParseDMY(ByVal v As Variant) As Date
    On Error GoTo fail
    
    If IsEmpty(v) Then GoTo fail
    If IsDate(v) Then ParseDMY = CDate(v): Exit Function
    
    Dim s As String: s = Trim(CStr(v))
    If Len(s) = 0 Then GoTo fail
    
    s = Replace(s, "-", "/"): s = Replace(s, ".", "/"): s = Replace(s, " ", "/")
    Dim p() As String: p = Split(s, "/")
    If UBound(p) <> 2 Then GoTo fail
    
    Dim d As Integer, m As Integer, Y As Integer
    d = CInt(p(0)): m = CInt(p(1)): Y = CInt(p(2))
    ParseDMY = DateSerial(Y, m, d)
    Exit Function
fail:
    ParseDMY = 0
End Function

' ==================== GET OR CREATE SHEET ====================
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function WSExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    WSExists = Not ws Is Nothing
End Function

Private Function GetFPTKWorksheet() As Worksheet
    On Error Resume Next
    Set GetFPTKWorksheet = ThisWorkbook.Worksheets("FPTK")
    On Error GoTo 0
End Function

' ==================== FIND HEADER ====================
Private Function FindHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal headerName As String) As Long

    Dim c As Range

    If ws Is Nothing Then Exit Function

    Set c = ws.rows(SOURCE_HEADER_ROW).Find( _
        What:=headerName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If Not c Is Nothing Then
        FindHeaderColumn = c.Column
    End If

End Function


Private Function FindHeaderColumnLike( _
    ByVal ws As Worksheet, _
    ByVal partialName As String) As Long

    Dim columnIndex As Long
    Dim lastColumn As Long
    Dim headerText As String

    If ws Is Nothing Then Exit Function

    lastColumn = ws.Cells( _
        SOURCE_HEADER_ROW, _
        ws.Columns.Count).End(xlToLeft).Column

    For columnIndex = 1 To lastColumn

        headerText = CStr( _
            ws.Cells( _
                SOURCE_HEADER_ROW, _
                columnIndex).Value2)

        If InStr(1, headerText, partialName, vbTextCompare) > 0 Then
            FindHeaderColumnLike = columnIndex
            Exit Function
        End If

    Next columnIndex

End Function


' ==================== HELPERS ====================
Private Function GetVal(ByRef arr As Variant, ByVal idx As Long, ByVal col As Long) As Variant
    If col > 0 And col <= UBound(arr, 2) Then
        GetVal = arr(idx, col)
    Else
        GetVal = ""
    End If
End Function

Private Function IsLolosFun(ByVal val As String) As Boolean
    val = Trim(val)
    If Len(val) = 0 Then Exit Function
    If InStr(1, val, "V", vbTextCompare) > 0 Then IsLolosFun = True
    If InStr(1, val, "Lolos", vbTextCompare) > 0 Then IsLolosFun = True
    If InStr(1, val, "Yes", vbTextCompare) > 0 Then IsLolosFun = True
End Function

Private Function GetStartOfWeek(ByVal dt As Date) As Date
    GetStartOfWeek = dt - Weekday(dt, vbMonday) + 1
End Function

Private Function GetEndOfWeek(ByVal dt As Date) As Date
    GetEndOfWeek = GetStartOfWeek(dt) + 6
End Function

Private Function ColumnLetter(ByVal colNum As Long) As String
    If colNum <= 26 Then
        ColumnLetter = Chr(64 + colNum)
    Else
        ColumnLetter = Chr(64 + Int((colNum - 1) / 26)) & Chr(65 + ((colNum - 1) Mod 26))
    End If
End Function

Private Function SafeTextReport(ByVal rawValue As Variant) As String

    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then
        Exit Function
    End If

    SafeTextReport = Trim$(CStr(rawValue))

End Function

Private Function ExtractFirstNumberReport(ByVal rawText As String) As Long

    Dim characterIndex As Long
    Dim numberText As String
    Dim oneCharacter As String

    For characterIndex = 1 To Len(rawText)

        oneCharacter = Mid$(rawText, characterIndex, 1)

        If oneCharacter >= "0" And oneCharacter <= "9" Then
            numberText = numberText & oneCharacter
        ElseIf Len(numberText) > 0 Then
            Exit For
        End If

    Next characterIndex

    If Len(numberText) > 0 Then
        ExtractFirstNumberReport = CLng(numberText)
    End If

End Function

Private Function GetFilterKategorisasiReport( _
    ByVal existingFilter As Variant, _
    ByVal posisiValue As Variant, _
    ByVal levelValue As Variant) As String

    Dim existingText As String
    Dim posisiText As String
    Dim levelNumber As Long

    existingText = SafeTextReport(existingFilter)

    If Len(existingText) > 0 Then
        GetFilterKategorisasiReport = existingText
        Exit Function
    End If

    posisiText = LCase$(SafeTextReport(posisiValue))

    If Left$(posisiText, 6) = "cimory" Or _
       Left$(posisiText, 5) = "fresh" Then

        GetFilterKategorisasiReport = "CLAP FGDP"
        Exit Function

    End If

    If IsNumeric(levelValue) Then
        levelNumber = CLng(levelValue)
    Else
        levelNumber = ExtractFirstNumberReport( _
            SafeTextReport(levelValue))
    End If

    Select Case levelNumber
        Case 1, 2
            GetFilterKategorisasiReport = "Level 1-2"
        Case 3
            GetFilterKategorisasiReport = "Level 3"
        Case 4
            GetFilterKategorisasiReport = "Level 4"
    End Select

End Function

' ============================================================
' GET LAST DATE FROM DB SOURCING
' ============================================================
Private Function GetLastDateFromDBSourcing() As Date
    Dim wsDB As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim maxDate As Date
    Dim dt As Date
    
    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets("DB Sourcing")
    On Error GoTo 0
    
    If wsDB Is Nothing Then
        GetLastDateFromDBSourcing = Date
        Exit Function
    End If
    Dim sourcingDateColumn As Long

    sourcingDateColumn = SchemaFindColumnAuto( _
        wsDB, "Sourcing Date", False)

    If sourcingDateColumn = 0 Then
        sourcingDateColumn = FindHeaderColumnLike( _
            wsDB, "Tanggal Sourcing")
    End If

    If sourcingDateColumn = 0 Then
        GetLastDateFromDBSourcing = Date
        Exit Function
    End If

    lastRow = wsDB.Cells( _
        wsDB.rows.Count, _
        sourcingDateColumn).End(xlUp).Row

    maxDate = DateSerial(2026, 1, 1)

    For r = SOURCE_FIRST_DATA_ROW To lastRow
        dt = ParseDMY(wsDB.Cells(r, sourcingDateColumn).Value)
        If dt > maxDate Then maxDate = dt
    Next r
    
    If maxDate < DateSerial(2026, 1, 2) Then
        GetLastDateFromDBSourcing = Date
    Else
        GetLastDateFromDBSourcing = maxDate
    End If
End Function

' ============================================================
' CREATE MONITORING SOURCING - 7 HARI (SENIN-MINGGU)
' ============================================================
Public Sub CreateMonitoringSheet()
    SchemaB02PrepareAll
    Dim t As Single
    t = Timer
    
    Dim wsMon As Worksheet
    Set wsMon = GetOrCreateSheet("Monitoring Sourcing")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Building Monitoring Sourcing..."
    
    Dim wsFPTK As Worksheet, wsDB As Worksheet
    Set wsFPTK = GetFPTKWorksheet()
    If wsFPTK Is Nothing Then
        MsgBox "Sheet FPTK tidak ditemukan!", vbCritical
        GoTo CleanExitMon
    End If
    
    EnsureDBSourcing
    Set wsDB = ThisWorkbook.Worksheets("DB Sourcing")
    
    ' ===== FIND COLUMNS IN FPTK =====
    Dim colKodeFP As Long, colPosFP As Long
    Dim colPIC As Long, colStatusFP As Long
    Dim colFPTKDateReal As Long
    Dim colFilterFP As Long
    Dim colLevelNumberMon As Long
    
    colKodeFP = FindHeaderColumn(wsFPTK, "Kode Unik")
    If colKodeFP = 0 Then colKodeFP = FindHeaderColumnLike(wsFPTK, "Kode Unik")
    
    colPosFP = FindHeaderColumn(wsFPTK, "Posisi")
    If colPosFP = 0 Then colPosFP = FindHeaderColumn(wsFPTK, "Posisi")
    If colPosFP = 0 Then colPosFP = FindHeaderColumnLike(wsFPTK, "Posisi")
    
    colPIC = FindHeaderColumn(wsFPTK, "PIC Recruiter")
    If colPIC = 0 Then colPIC = FindHeaderColumnLike(wsFPTK, "PIC")
    
    colStatusFP = FindHeaderColumn(wsFPTK, "Status")
    If colStatusFP = 0 Then colStatusFP = FindHeaderColumnLike(wsFPTK, "Status")
    
    colFPTKDateReal = FindHeaderColumn(wsFPTK, "FPTK Date (Real)")
    If colFPTKDateReal = 0 Then colFPTKDateReal = FindHeaderColumnLike(wsFPTK, "FPTK DATE")
    If colFPTKDateReal = 0 Then colFPTKDateReal = FindHeaderColumnLike(wsFPTK, "FPTK Date")

    colFilterFP = FindHeaderColumn(wsFPTK, "Filter Kategorisasi FPTK")
    If colFilterFP = 0 Then colFilterFP = FindHeaderColumnLike(wsFPTK, "Filter Kategorisasi")

    colLevelNumberMon = FindHeaderColumn(wsFPTK, "Level Number")
    If colLevelNumberMon = 0 Then colLevelNumberMon = FindHeaderColumnLike(wsFPTK, "Level Number")
    
    If colKodeFP = 0 Then
        MsgBox "Kolom 'Kode Unik' tidak ditemukan di sheet FPTK!", vbCritical
        GoTo CleanExitMon
    End If
    
    ' ===== FIND COLUMNS IN DB SOURCING =====
    Dim colKodeDB As Long
    Dim colTglDB As Long
    
    colKodeDB = FindHeaderColumn(wsDB, "Kode Unik (copy value dari FPTK)")
    If colKodeDB = 0 Then colKodeDB = FindHeaderColumn(wsDB, "Kode Unik")
    If colKodeDB = 0 Then colKodeDB = FindHeaderColumnLike(wsDB, "Kode Unik")
    
    If colKodeDB = 0 Then
        MsgBox "Kolom 'Kode Unik' tidak ditemukan di sheet DB Sourcing!", vbCritical
        GoTo CleanExitMon
    End If
    colTglDB = SchemaFindColumnAuto( _
        wsDB, "Sourcing Date", False)

    If colTglDB = 0 Then
        colTglDB = FindHeaderColumnLike( _
            wsDB, "Tanggal Sourcing")
    End If

    If colTglDB = 0 Then
        MsgBox _
            "Kolom Sourcing Date tidak ditemukan di DB Sourcing.", _
            vbCritical
        GoTo CleanExitMon
    End If
    
    ' ===== BUILD CACHE FROM DB SOURCING =====
    Dim cache As Object
    Set cache = CreateObject("Scripting.Dictionary")
    
    Dim lastDB As Long
    lastDB = wsDB.Cells(wsDB.rows.Count, colKodeDB).End(xlUp).Row
    
    Dim minDate As Date, maxDate As Date
    minDate = DateSerial(2026, 1, 1)
    maxDate = DateSerial(2026, 1, 1)
    
    Dim r As Long
    For r = SOURCE_FIRST_DATA_ROW To lastDB
        Dim kd As String
        kd = NormalizeKey(wsDB.Cells(r, colKodeDB).Value)
        If Len(kd) = 0 Then GoTo NextRowDB
        
        ' ===== AMBIL TANGGAL DARI KOLOM B =====
        Dim dt As Date
        dt = ParseDMY(wsDB.Cells(r, colTglDB).Value)
        
        If dt <> 0 Then
            If dt < minDate Then minDate = dt
            If dt > maxDate Then maxDate = dt
            
            Dim key As String
            key = kd & "|" & Format(dt, "yyyy-mm-dd")
            If cache.Exists(key) Then
                cache(key) = cache(key) + 1
            Else
                cache.Add key, 1
            End If
        End If
NextRowDB:
    Next r
    
    ' ===== TANGGAL TERAKHIR =====
    Dim lastDateDB As Date
    lastDateDB = GetLastDateFromDBSourcing()
    
    ' ===== TANGGAL MULAI =====
    Dim startDate As Date
    startDate = DateSerial(2026, 1, 1)
    
    Dim firstMonday As Date
    firstMonday = startDate
    Do While Weekday(firstMonday, vbMonday) <> 1
        firstMonday = firstMonday + 1
    Loop
    
    ' ===== END DATE ADALAH MINGGU TERAKHIR =====
    Dim lastSunday As Date
    lastSunday = GetEndOfWeek(lastDateDB)
    
    ' ===== TOTAL WEEKS (SENIN-MINGGU) =====
    Dim totalWeeks As Long
    totalWeeks = Int((lastSunday - firstMonday) / 7) + 1
    If totalWeeks > 53 Then totalWeeks = 53
    If totalWeeks < 1 Then totalWeeks = 1
    
    ' ===== BUILD HEADER =====
    BuildMonitoringHeaders7Days wsMon, firstMonday, totalWeeks, lastDateDB
    
    ' ===== LOAD FPTK =====
    Dim lastFP As Long
    lastFP = wsFPTK.Cells(wsFPTK.rows.Count, colKodeFP).End(xlUp).Row
    
    Dim maxColFP As Long
    maxColFP = Application.WorksheetFunction.Max( _
        colKodeFP, colPosFP, colPIC, colStatusFP, _
        colFPTKDateReal, colFilterFP, colLevelNumberMon)
    If maxColFP < 1 Then maxColFP = colKodeFP
    
    Dim arrFP As Variant
    arrFP = wsFPTK.Range(wsFPTK.Cells(SOURCE_FIRST_DATA_ROW, 1), wsFPTK.Cells(lastFP, maxColFP)).Value
    
    ' ===== TOTAL COLUMNS: 6 FIXED + (totalWeeks * 15) ====
    ' 15 = 7 hari ï¿½ 2 kolom (CV Count + Evidence) + 1 kolom Total
    Dim totalCols As Long
    totalCols = 6 + (totalWeeks * 15)
    
    Dim outArr() As Variant
    ReDim outArr(1 To UBound(arrFP, 1), 1 To totalCols)
    
    Dim i As Long, outRow As Long
    outRow = 0
    
    For i = 1 To UBound(arrFP, 1)
        Dim kode As String
        kode = NormalizeKey(arrFP(i, colKodeFP))
        
        If Len(kode) > 0 Then
            outRow = outRow + 1
            
            ' ===== KOLOM FIXED =====
            outArr(outRow, 1) = kode
            
            If colPosFP > 0 And colPosFP <= UBound(arrFP, 2) Then
                outArr(outRow, 2) = arrFP(i, colPosFP)
            End If
            
            If colPIC > 0 And colPIC <= UBound(arrFP, 2) Then
                outArr(outRow, 3) = arrFP(i, colPIC)
            End If
            
            If colStatusFP > 0 And colStatusFP <= UBound(arrFP, 2) Then
                outArr(outRow, 4) = arrFP(i, colStatusFP)
            End If
            
            If colFPTKDateReal > 0 And colFPTKDateReal <= UBound(arrFP, 2) Then
                Dim fptkDateVal As Variant
                fptkDateVal = arrFP(i, colFPTKDateReal)
                If IsDate(fptkDateVal) Then
                    outArr(outRow, 5) = CDate(fptkDateVal)
                ElseIf Len(Trim(CStr(fptkDateVal))) > 0 Then
                    outArr(outRow, 5) = ParseDMY(fptkDateVal)
                Else
                    outArr(outRow, 5) = ""
                End If
            End If
            
            outArr(outRow, 6) = _
                GetFilterKategorisasiReport( _
                    GetVal(arrFP, i, colFilterFP), _
                    GetVal(arrFP, i, colPosFP), _
                    GetVal(arrFP, i, colLevelNumberMon))

            ' ===== DATA PER HARI (7 HARI) =====
            Dim colIdx As Long: colIdx = 7
            Dim w As Long, d As Long
            
            For w = 0 To totalWeeks - 1
                Dim weekStart As Date: weekStart = firstMonday + (w * 7)
                Dim weekTotal As Long: weekTotal = 0
                
                For d = 0 To 6  ' 7 hari: Senin(0) sampai Minggu(6)
                    Dim dayDate As Date: dayDate = weekStart + d
                    
                    If dayDate > lastDateDB Then
                        outArr(outRow, colIdx) = ""
                        colIdx = colIdx + 1
                        outArr(outRow, colIdx) = ""
                        colIdx = colIdx + 1
                        GoTo NextDay
                    End If
                    
                    Dim cacheKey As String
                    cacheKey = kode & "|" & Format(dayDate, "yyyy-mm-dd")
                    
                    Dim cvCount As Long
                    cvCount = 0
                    If cache.Exists(cacheKey) Then
                        cvCount = cache(cacheKey)
                    End If
                    
                    outArr(outRow, colIdx) = cvCount
                    weekTotal = weekTotal + cvCount
                    colIdx = colIdx + 1
                    outArr(outRow, colIdx) = ""
                    colIdx = colIdx + 1
NextDay:
                Next d
                
                outArr(outRow, colIdx) = weekTotal
                colIdx = colIdx + 1
            Next w
        End If
    Next i
    
    ' ===== WRITE TO SHEET =====
    wsMon.Range("A2").Resize(wsMon.rows.Count - 1, totalCols).ClearContents
    If outRow > 0 Then
        wsMon.Range("A2").Resize(outRow, totalCols).Value = outArr
    End If
    
    wsMon.Range(wsMon.Columns(1), wsMon.Columns(totalCols)).AutoFit
    
CleanExitMon:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    If outRow > 0 Then
        MsgBox "Monitoring Sourcing selesai!" & vbCrLf & _
               "Posisi: " & outRow & " | Weeks: " & totalWeeks & vbCrLf & _
               "Periode: " & Format(firstMonday, "dd/mm/yyyy") & " s.d. " & Format(lastSunday, "dd/mm/yyyy") & vbCrLf & _
               "(7 hari per minggu)", vbInformation
    Else
        MsgBox "Tidak ada data FPTK yang ditemukan!", vbExclamation
    End If
End Sub

' ==================== BUILD HEADERS - 7 HARI ====================
Private Sub BuildMonitoringHeaders7Days(ByVal ws As Worksheet, ByVal firstMonday As Date, ByVal totalWeeks As Long, ByVal lastDate As Date)
    ws.Cells.Clear
    
    ' ===== 6 KOLOM FIXED =====
    ws.Cells(1, 1).Value = "Kode Unik"
    ws.Cells(1, 2).Value = "Posisi"
    ws.Cells(1, 3).Value = "PIC Recruiter"
    ws.Cells(1, 4).Value = "Status FPTK"
    ws.Cells(1, 5).Value = "FPTK Date (Real)"
    ws.Cells(1, 6).Value = "Filter Kategorisasi FPTK"
    ws.Columns(5).numberFormat = "dd/mm/yyyy"
    ws.Columns(5).columnWidth = 14
    ws.Columns(6).columnWidth = 23
    
    Dim daysOfWeek As Variant
    daysOfWeek = Array("Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min")
    
    Dim col As Long: col = 7
    Dim w As Long, d As Long
    Dim evCounter As Long
    evCounter = 0
    
    For w = 0 To totalWeeks - 1
        Dim weekStart As Date: weekStart = firstMonday + (w * 7)
        Dim weekEnd As Date: weekEnd = weekStart + 6
        
        For d = 0 To 6  ' 7 hari
            Dim dayDate As Date: dayDate = weekStart + d
            
            If dayDate > lastDate Then
                ws.Cells(1, col).Value = ""
                ws.Cells(1, col).columnWidth = 11
                col = col + 1
                ws.Cells(1, col).Value = ""
                ws.Cells(1, col).columnWidth = 35
                col = col + 1
                GoTo NextHeaderDay7
            End If
            
            ' ===== TANGGAL + HARI =====
            ws.Cells(1, col).Value = daysOfWeek(d) & " " & Format(dayDate, "dd-mmm-yy")
            ws.Cells(1, col).columnWidth = 14
            ws.Cells(1, col).HorizontalAlignment = xlCenter
            
            ' ===== WARNA UNTUK WEEKEND =====
            If d >= 5 Then  ' Sabtu & Minggu
                ws.Cells(1, col).Interior.Color = RGB(255, 230, 230)  ' Light red
            End If
            col = col + 1
            
            ' ===== EVIDENCE =====
            evCounter = evCounter + 1
            ws.Cells(1, col).Value = "Ev" & evCounter
            ws.Cells(1, col).columnWidth = 35
            ws.Cells(1, col).HorizontalAlignment = xlCenter
            
            If d >= 5 Then  ' Sabtu & Minggu
                ws.Cells(1, col).Interior.Color = RGB(255, 230, 230)  ' Light red
            End If
            col = col + 1
            
NextHeaderDay7:
        Next d
        
        ' ===== TOTAL =====
        ws.Cells(1, col).Value = "Total " & Format(weekStart, "dd-mmm") & " s.d. " & Format(weekEnd, "dd-mmm")
        ws.Cells(1, col).columnWidth = 20
        ws.Cells(1, col).HorizontalAlignment = xlCenter
        col = col + 1
    Next w
    
    Dim lastCol As Long: lastCol = col - 1
    
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Font.Size = 9
        .Interior.Color = RGB(44, 62, 80)
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.rows(1).RowHeight = 36
    
    ' ===== HIGHLIGHT TOTAL COLUMNS =====
    Dim totalCol As Long
    For totalCol = 7 To lastCol
        Dim hdr As String
        hdr = CStr(ws.Cells(1, totalCol).Value)
        If InStr(1, hdr, "Total", vbTextCompare) > 0 Then
            ws.Columns(totalCol).Interior.Color = RGB(230, 240, 250)
        End If
    Next totalCol
End Sub

' ============================================================
' REFRESH REPORT CALCULATIONS
' ============================================================
Public Sub RefreshReportCalculations()
    SchemaB02PrepareAll "FPTK"
    SchemaB02PrepareAll "DB Sourcing"
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Refreshing reports..."
    
    CreateMonitoringSheet
    UpdatefunnelingSheet
    
    Application.StatusBar = "Reports refreshed!"
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Semua report berhasil di-refresh!", vbInformation
    
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error refreshing reports: " & Err.Description, vbCritical
End Sub

' ============================================================
' UPDATE FUNNELING SHEET
' ============================================================
Public Sub UpdatefunnelingSheet()
    SchemaB02PrepareAll
    Dim t As Single
    t = Timer
    
    Dim wsFun As Worksheet
    If Not WSExists("funneling") Then
        Set wsFun = GetOrCreateSheet("funneling")
    Else
        Set wsFun = ThisWorkbook.Worksheets("funneling")
    End If
    
    wsFun.Cells.Clear
    wsFun.Cells.ClearFormats
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Update funneling..."
    
    Dim wsFPTK As Worksheet, wsDB As Worksheet
    Set wsFPTK = GetFPTKWorksheet()
    If wsFPTK Is Nothing Then GoTo CleanExitFun
    
    EnsureDBSourcing
    Set wsDB = ThisWorkbook.Worksheets("DB Sourcing")
    
    ' ===== FIND COLUMNS IN FPTK =====
    Dim colKodeUnikFP As Long, colKodePIC As Long, colFPTKDateReal As Long
    Dim colKodeAngka As Long, colFPTKDate As Long
    Dim colPosisi As Long, colBU As Long, colDirektorat As Long, colDivisi As Long
    Dim colDepartment As Long, colLevelFPTK As Long, colLevelNumber As Long
    Dim colCategory As Long, colCategoryFPTK As Long
    Dim colFilterKategorisasi As Long, colPIC As Long
    Dim colVacancy As Long, colStatusFP As Long
    
    colKodeUnikFP = FindHeaderColumn(wsFPTK, "Kode Unik")
    colKodePIC = FindHeaderColumn(wsFPTK, "Kode PIC")
    colFPTKDateReal = FindHeaderColumn(wsFPTK, "FPTK Date (Real)")
    If colFPTKDateReal = 0 Then colFPTKDateReal = FindHeaderColumnLike(wsFPTK, "FPTK Date")
    colKodeAngka = FindHeaderColumn(wsFPTK, "Kode Angka")
    colFPTKDate = FindHeaderColumn(wsFPTK, "FPTK Date (Kode)")
    If colFPTKDate = 0 Then colFPTKDate = FindHeaderColumnLike(wsFPTK, "FPTK Date")
    colPosisi = FindHeaderColumn(wsFPTK, "Posisi")
    If colPosisi = 0 Then colPosisi = FindHeaderColumnLike(wsFPTK, "Posisi")
    colBU = FindHeaderColumnLike(wsFPTK, "Business Unit")
    colDirektorat = FindHeaderColumnLike(wsFPTK, "Direktorat")
    colDivisi = FindHeaderColumnLike(wsFPTK, "Divisi")
    colDepartment = FindHeaderColumnLike(wsFPTK, "Department")
    colLevelFPTK = FindHeaderColumnLike(wsFPTK, "Level")
    colLevelNumber = FindHeaderColumnLike(wsFPTK, "Level Number")
    colCategory = FindHeaderColumnLike(wsFPTK, "Category")
    colCategoryFPTK = FindHeaderColumnLike(wsFPTK, "Category FPTK")
    colFilterKategorisasi = FindHeaderColumn(wsFPTK, "Filter Kategorisasi FPTK")
    If colFilterKategorisasi = 0 Then colFilterKategorisasi = FindHeaderColumnLike(wsFPTK, "Filter Kategorisasi")
    colPIC = FindHeaderColumn(wsFPTK, "PIC Recruiter")
    If colPIC = 0 Then colPIC = FindHeaderColumnLike(wsFPTK, "PIC")
    colVacancy = FindHeaderColumnLike(wsFPTK, "Vacancy")
    colStatusFP = FindHeaderColumnLike(wsFPTK, "Status")
    
    If colKodeUnikFP = 0 Then GoTo CleanExitFun
    
    ' ===== BUILD FPTK DICTIONARY =====
    Dim dictFPTK As Object
    Set dictFPTK = CreateObject("Scripting.Dictionary")
    
    Dim lastFP As Long
    lastFP = wsFPTK.Cells(wsFPTK.rows.Count, colKodeUnikFP).End(xlUp).Row
    
    Dim maxColFP As Long: maxColFP = 1
    Dim colArr As Variant
    colArr = Array(colKodeUnikFP, colKodePIC, colFPTKDateReal, colKodeAngka, colFPTKDate, _
                   colPosisi, colBU, colDirektorat, colDivisi, colDepartment, colLevelFPTK, _
                   colLevelNumber, colCategory, colCategoryFPTK, colFilterKategorisasi, _
                   colPIC, colVacancy, colStatusFP)
    Dim cIdx As Variant
    For Each cIdx In colArr
        If CLng(cIdx) > maxColFP Then maxColFP = CLng(cIdx)
    Next cIdx
    
    Dim arrFP As Variant
    arrFP = wsFPTK.Range(wsFPTK.Cells(SOURCE_FIRST_DATA_ROW, 1), wsFPTK.Cells(lastFP, maxColFP)).Value
    
    Dim i As Long
    For i = 1 To UBound(arrFP, 1)
        Dim fpKode As String
        fpKode = NormalizeKey(arrFP(i, colKodeUnikFP))
        If Len(fpKode) > 0 And Not dictFPTK.Exists(fpKode) Then
            dictFPTK.Add fpKode, i
        End If
    Next i
    
    ' ===== FIND COLUMNS IN DB SOURCING =====
    Dim colKodeDB As Long
    colKodeDB = FindHeaderColumn(wsDB, "Kode Unik (copy value dari FPTK)")
    If colKodeDB = 0 Then colKodeDB = FindHeaderColumn(wsDB, "Kode Unik")
    If colKodeDB = 0 Then GoTo CleanExitFun
    
    Dim colSourcingFL As Long, colSourcingHR As Long, colShortlist As Long
    Dim colPsikotes As Long, colHRInterview As Long, colTechnicalTest As Long
    Dim colMarketVisit As Long, colUserInterview As Long, colPanelInterview As Long
    Dim colReferenceCheck As Long, colMCU As Long, colOffering As Long, colDay1 As Long
    
    colSourcingFL = FindHeaderColumn(wsDB, "Sourcin Freelance")
    If colSourcingFL = 0 Then colSourcingFL = FindHeaderColumnLike(wsDB, "Freelance")
    colSourcingHR = FindHeaderColumn(wsDB, "Sourcing HR")
    colShortlist = FindHeaderColumn(wsDB, "Shortlist CV")
    colPsikotes = FindHeaderColumn(wsDB, "Psikotes")
    colHRInterview = FindHeaderColumn(wsDB, "HR Interview")
    colTechnicalTest = FindHeaderColumn(wsDB, "Technical Test/ Case Study")
    If colTechnicalTest = 0 Then colTechnicalTest = FindHeaderColumnLike(wsDB, "Technical")
    colMarketVisit = FindHeaderColumn(wsDB, "Market Visit")
    colUserInterview = FindHeaderColumn(wsDB, "User Interview")
    colPanelInterview = FindHeaderColumn(wsDB, "Panel Interview")
    colReferenceCheck = FindHeaderColumn(wsDB, "Reference Check")
    colMCU = FindHeaderColumn(wsDB, "MCU")
    colOffering = FindHeaderColumn(wsDB, "Offering")
    colDay1 = FindHeaderColumn(wsDB, "Day 1")
    
    ' ===== BUILD FUNNEL DICTIONARY =====
    Dim dictFunnel As Object
    Set dictFunnel = CreateObject("Scripting.Dictionary")
    
    Dim lastDB As Long
    lastDB = wsDB.Cells(wsDB.rows.Count, colKodeDB).End(xlUp).Row
    
    Dim r As Long
    For r = SOURCE_FIRST_DATA_ROW To lastDB
        Dim dbKode As String
        dbKode = NormalizeKey(wsDB.Cells(r, colKodeDB).Value)
        If Len(dbKode) = 0 Then GoTo NextDB
        
        If Not dictFunnel.Exists(dbKode) Then
            Dim counts(1 To 22) As Long
            dictFunnel.Add dbKode, counts
        End If
        
        Dim cnt As Variant
        cnt = dictFunnel(dbKode)
        
        If colSourcingFL > 0 Then If IsLolosFun(CStr(wsDB.Cells(r, colSourcingFL).Value)) Then cnt(1) = cnt(1) + 1
        If colSourcingHR > 0 Then If IsLolosFun(CStr(wsDB.Cells(r, colSourcingHR).Value)) Then cnt(2) = cnt(2) + 1
        If colShortlist > 0 Then If IsLolosFun(CStr(wsDB.Cells(r, colShortlist).Value)) Then cnt(3) = cnt(3) + 1
        
        If colPsikotes > 0 Then
            Dim vp As String: vp = Trim(CStr(wsDB.Cells(r, colPsikotes).Value))
            If Len(vp) > 0 Then cnt(4) = cnt(4) + 1
            If IsLolosFun(vp) Then cnt(5) = cnt(5) + 1
        End If
        If colHRInterview > 0 Then
            Dim vh As String: vh = Trim(CStr(wsDB.Cells(r, colHRInterview).Value))
            If Len(vh) > 0 Then cnt(6) = cnt(6) + 1
            If IsLolosFun(vh) Then cnt(7) = cnt(7) + 1
        End If
        If colTechnicalTest > 0 Then
            Dim vt As String: vt = Trim(CStr(wsDB.Cells(r, colTechnicalTest).Value))
            If Len(vt) > 0 Then cnt(8) = cnt(8) + 1
            If IsLolosFun(vt) Then cnt(9) = cnt(9) + 1
        End If
        If colMarketVisit > 0 Then
            Dim vm As String: vm = Trim(CStr(wsDB.Cells(r, colMarketVisit).Value))
            If Len(vm) > 0 Then cnt(10) = cnt(10) + 1
            If IsLolosFun(vm) Then cnt(11) = cnt(11) + 1
        End If
        If colUserInterview > 0 Then
            Dim vu As String: vu = Trim(CStr(wsDB.Cells(r, colUserInterview).Value))
            If Len(vu) > 0 Then cnt(12) = cnt(12) + 1
            If IsLolosFun(vu) Then cnt(13) = cnt(13) + 1
        End If
        If colPanelInterview > 0 Then
            Dim vpa As String: vpa = Trim(CStr(wsDB.Cells(r, colPanelInterview).Value))
            If Len(vpa) > 0 Then cnt(14) = cnt(14) + 1
            If IsLolosFun(vpa) Then cnt(15) = cnt(15) + 1
        End If
        If colReferenceCheck > 0 Then
            Dim vr As String: vr = Trim(CStr(wsDB.Cells(r, colReferenceCheck).Value))
            If Len(vr) > 0 Then cnt(16) = cnt(16) + 1
            If IsLolosFun(vr) Then cnt(17) = cnt(17) + 1
        End If
        If colMCU > 0 Then
            Dim vmc As String: vmc = Trim(CStr(wsDB.Cells(r, colMCU).Value))
            If Len(vmc) > 0 Then cnt(18) = cnt(18) + 1
            If IsLolosFun(vmc) Then cnt(19) = cnt(19) + 1
        End If
        If colOffering > 0 Then
            Dim vo As String: vo = Trim(CStr(wsDB.Cells(r, colOffering).Value))
            If Len(vo) > 0 Then cnt(20) = cnt(20) + 1
            If IsLolosFun(vo) Then cnt(21) = cnt(21) + 1
        End If
        If colDay1 > 0 Then
            Dim vd As String: vd = Trim(CStr(wsDB.Cells(r, colDay1).Value))
            If Len(vd) > 0 Then cnt(22) = cnt(22) + 1
        End If
        
        dictFunnel(dbKode) = cnt
        
NextDB:
    Next r
    
    ' ===== BUILD HEADER & OUTPUT =====
    BuildFunnelingHeaders wsFun
    
    Dim targetRow As Long: targetRow = 1
    Dim fptkKeys As Variant: fptkKeys = dictFPTK.keys
    Dim ki As Long
    
    For ki = 0 To UBound(fptkKeys)
        Dim kodeUnik As String: kodeUnik = fptkKeys(ki)
        Dim fpIdx As Long: fpIdx = dictFPTK(kodeUnik)
        targetRow = targetRow + 1
        
        With wsFun
            .Cells(targetRow, 1).Value = targetRow - 1
            .Cells(targetRow, 2).Value = GetVal(arrFP, fpIdx, colKodePIC)
            
            Dim fptkDateRealVal As Variant
            fptkDateRealVal = GetVal(arrFP, fpIdx, colFPTKDateReal)
            If IsDate(fptkDateRealVal) Then
                .Cells(targetRow, 3).Value = CDate(fptkDateRealVal)
            Else
                .Cells(targetRow, 3).Value = fptkDateRealVal
            End If
            
            .Cells(targetRow, 4).Value = GetVal(arrFP, fpIdx, colKodeAngka)
            
            Dim fptkDateVal As Variant
            fptkDateVal = GetVal(arrFP, fpIdx, colFPTKDate)
            If IsDate(fptkDateVal) Then
                .Cells(targetRow, 5).Value = CDate(fptkDateVal)
            Else
                .Cells(targetRow, 5).Value = fptkDateVal
            End If
            
            .Cells(targetRow, 6).Value = kodeUnik
            .Cells(targetRow, 7).Value = GetVal(arrFP, fpIdx, colPosisi)
            .Cells(targetRow, 8).Value = GetVal(arrFP, fpIdx, colBU)
            .Cells(targetRow, 9).Value = GetVal(arrFP, fpIdx, colDirektorat)
            .Cells(targetRow, 10).Value = GetVal(arrFP, fpIdx, colDivisi)
            .Cells(targetRow, 11).Value = GetVal(arrFP, fpIdx, colDepartment)
            .Cells(targetRow, 12).Value = GetVal(arrFP, fpIdx, colLevelFPTK)
            .Cells(targetRow, 13).Value = GetVal(arrFP, fpIdx, colLevelNumber)
            .Cells(targetRow, 14).Value = GetVal(arrFP, fpIdx, colCategory)
            .Cells(targetRow, 15).Value = GetVal(arrFP, fpIdx, colCategoryFPTK)
            .Cells(targetRow, 16).Value = _
                GetFilterKategorisasiReport( _
                    GetVal(arrFP, fpIdx, colFilterKategorisasi), _
                    GetVal(arrFP, fpIdx, colPosisi), _
                    GetVal(arrFP, fpIdx, colLevelNumber))
            .Cells(targetRow, 17).Value = GetVal(arrFP, fpIdx, colPIC)
            .Cells(targetRow, 18).Value = CLng(val(GetVal(arrFP, fpIdx, colVacancy)))
            .Cells(targetRow, 19).Value = GetVal(arrFP, fpIdx, colStatusFP)

            If dictFunnel.Exists(kodeUnik) Then
                cnt = dictFunnel(kodeUnik)
                .Cells(targetRow, 20).Value = cnt(1)
                .Cells(targetRow, 21).Value = cnt(2)
                .Cells(targetRow, 22).Value = cnt(3)
                .Cells(targetRow, 23).Value = cnt(4)
                .Cells(targetRow, 24).Value = cnt(5)
                .Cells(targetRow, 25).Value = cnt(6)
                .Cells(targetRow, 26).Value = cnt(7)
                .Cells(targetRow, 27).Value = cnt(8)
                .Cells(targetRow, 28).Value = cnt(9)
                .Cells(targetRow, 29).Value = cnt(10)
                .Cells(targetRow, 30).Value = cnt(11)
                .Cells(targetRow, 31).Value = cnt(12)
                .Cells(targetRow, 32).Value = cnt(13)
                .Cells(targetRow, 33).Value = cnt(14)
                .Cells(targetRow, 34).Value = cnt(15)
                .Cells(targetRow, 35).Value = cnt(16)
                .Cells(targetRow, 36).Value = cnt(17)
                .Cells(targetRow, 37).Value = cnt(18)
                .Cells(targetRow, 38).Value = cnt(19)
                .Cells(targetRow, 39).Value = cnt(20)
                .Cells(targetRow, 40).Value = cnt(21)
                .Cells(targetRow, 41).Value = cnt(22)
                .Cells(targetRow, 42).Value = ""
            End If
        End With
    Next ki
    
    Dim c As Long
    For c = 1 To 42
        Dim hdr As String
        hdr = CStr(wsFun.Cells(1, c).Value)
        If InStr(1, hdr, "Date", vbTextCompare) > 0 Then
            wsFun.Columns(c).numberFormat = "dd/mm/yyyy"
        End If
    Next c
    
    wsFun.Columns("A:AP").AutoFit
    
CleanExitFun:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If dictFPTK.Count > 0 Then
        MsgBox "funneling updated!" & vbCrLf & _
               "Posisi: " & dictFPTK.Count & vbCrLf & _
               "Time: " & Format(Timer - t, "0.0") & "s", vbInformation
    End If
End Sub

Private Sub BuildFunnelingHeaders(ByVal ws As Worksheet)
    Dim headers As Variant
    headers = Array( _
        "No.", "Kode PIC", "FPTK Date Real", "Kode Angka", "FPTK Date", _
        "Kode Unik", "Posisi", "Business Unit", "Direktorat", _
        "Divisi", "Department", "Level FPTK", "Level Number", _
        "Category", "Category FPTK", "Filter Kategorisasi FPTK", _
        "PIC Recruiter", "Vacancy", "Status", _
        "Sourcing FL", "Lolos Sourcing HR", "Shortlisted User", _
        "Send Psikotes", "Lulus Psikotes", _
        "HR Interview", "Lulus HR Interview", _
        "Technical Case", "Lulus Technical Case", _
        "Market Visit", "Lulus Market Visit", _
        "User Interview", "Lulus User Interview", _
        "Panel Interview", "Lulus Panel Interview", _
        "Reference Check", "Lulus Reference Check", _
        "Proses MCU", "Lolos MCU", _
        "Offering", "Lolos Offering", _
        "Day One", "Lolos Day One")
    
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    
    Dim lastCol As Long: lastCol = UBound(headers) + 1
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(44, 62, 80)
        .Font.Color = vbWhite
        .WrapText = True
        .HorizontalAlignment = xlCenter
    End With
    ws.rows(1).RowHeight = 36
End Sub




