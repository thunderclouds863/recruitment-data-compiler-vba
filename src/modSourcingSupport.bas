Attribute VB_Name = "modSourcingSupport"
Option Explicit

Public Function WSExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    WSExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    If WSExists(sheetName) Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    Else
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Public Sub InitSourcingCV()
    EnsureDBSourcing
    EnsureCVRawText
End Sub
Public Sub EnsureDBSourcing()

    Dim ws As Worksheet
    Dim lastColumn As Long

    EnsureSchemaConfig
    SchemaEnsureSheet "DB Sourcing"

    Set ws = GetOrCreateSheet( _
        "DB Sourcing")

    lastColumn = ws.Cells( _
        1, ws.Columns.Count). _
        End(xlToLeft).Column

    If lastColumn > 0 Then
        FormatHeader ws, lastColumn
    End If

End Sub
Public Sub EnsureCVRawText()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("CV Raw Text")
    If Len(ws.Cells(1, 1).Value) = 0 Then
        SetHeader ws, 1, "Timestamp"
        SetHeader ws, 2, "CV Path"
        SetHeader ws, 3, "File Name"
        SetHeader ws, 4, "Raw Text Length"
        SetHeader ws, 5, "Raw Text"
        FormatHeader ws, 5
        ws.Columns("A:E").AutoFit
    End If
End Sub

Public Sub SetHeader(ByVal ws As Worksheet, ByVal colNum As Long, ByVal headerText As String)
    ws.Cells(1, colNum).Value = headerText
End Sub

Public Sub FormatHeader(ByVal ws As Worksheet, ByVal colCount As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
        .Font.Bold = True
        .Interior.Color = RGB(44, 62, 80)
        .Font.Color = vbWhite
        .WrapText = True
    End With
    ws.rows(1).RowHeight = 32
End Sub
Public Function FindHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal headerText As String) As Long

    If ws Is Nothing Then Exit Function

    FindHeaderColumn = SchemaFindColumnAuto( _
        ws, headerText, False)

End Function
Public Sub SetByHeader(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal headerText As String, ByVal valueText As Variant)
    Dim c As Long
    c = FindHeaderColumn(ws, headerText)
    If c > 0 Then ws.Cells(rowNum, c).Value = valueText
End Sub

Public Function GetByHeader(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal headerText As String) As Variant
    Dim c As Long
    c = FindHeaderColumn(ws, headerText)
    If c > 0 Then
        GetByHeader = ws.Cells(rowNum, c).Value
    Else
        GetByHeader = ""
    End If
End Function

Public Function GetFPTKWorksheet() As Worksheet
    Dim ws As Worksheet
    If WSExists("FPTK") Then
        Set GetFPTKWorksheet = ThisWorkbook.Worksheets("FPTK")
        Exit Function
    End If
    For Each ws In ThisWorkbook.Worksheets
        If FindHeaderColumn(ws, "Kode Unik") > 0 Then
            If FindHeaderColumn(ws, "Posisi") > 0 Then
                Set GetFPTKWorksheet = ws
                Exit Function
            End If
        End If
    Next ws
End Function

Public Function SearchFPTK(ByVal keyword As String) As Collection
    Dim c As New Collection
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim colKode As Long, colPosisi As Long, colStatus As Long, colRecruiter As Long
    Set ws = GetFPTKWorksheet()
    If ws Is Nothing Then Set SearchFPTK = c: Exit Function
    colKode = FindHeaderColumn(ws, "Kode Unik")
    colPosisi = FindHeaderColumn(ws, "Posisi")
    colStatus = FindHeaderColumn(ws, "Status")
    colRecruiter = FindHeaderColumn(ws, "PIC Recruiter")
    If colKode = 0 Or colPosisi = 0 Then Set SearchFPTK = c: Exit Function
    keyword = LCase(Trim(keyword))
    lastRow = ws.Cells(ws.rows.Count, colKode).End(xlUp).Row
    Dim k As String, p As String, st As String, rc As String
    Dim md As String, hS As String
    For r = 2 To lastRow
        k = CStr(ws.Cells(r, colKode).Value)
        p = CStr(ws.Cells(r, colPosisi).Value)
        st = ""
        If colStatus > 0 Then st = CStr(ws.Cells(r, colStatus).Value)
        rc = ""
        If colRecruiter > 0 Then rc = CStr(ws.Cells(r, colRecruiter).Value)
        md = DetermineModelRekrutmen(p, "")
        hS = LCase(k & " " & p & " " & st & " " & rc & " " & md)
        If keyword = "" Or InStr(1, hS, keyword) > 0 Then
            c.Add k & "|" & p & "|" & st & "|" & rc & "|" & md
        End If
    Next r
    Set SearchFPTK = c
End Function

Public Function GetFPTKDetail(ByVal kodeUnik As String) As Variant
    Dim arr(4) As Variant
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim colKode As Long, colPosisi As Long, colStatus As Long, colRecruiter As Long
    Set ws = GetFPTKWorksheet()
    If ws Is Nothing Then GetFPTKDetail = arr: Exit Function
    colKode = FindHeaderColumn(ws, "Kode Unik")
    colPosisi = FindHeaderColumn(ws, "Posisi")
    colStatus = FindHeaderColumn(ws, "Status")
    colRecruiter = FindHeaderColumn(ws, "PIC Recruiter")
    If colKode = 0 Or colPosisi = 0 Then GetFPTKDetail = arr: Exit Function
    lastRow = ws.Cells(ws.rows.Count, colKode).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(ws.Cells(r, colKode).Value) = kodeUnik Then
            arr(0) = ws.Cells(r, colKode).Value
            arr(1) = ws.Cells(r, colPosisi).Value
            arr(2) = ""
            If colStatus > 0 Then arr(2) = ws.Cells(r, colStatus).Value
            arr(3) = ""
            If colRecruiter > 0 Then arr(3) = ws.Cells(r, colRecruiter).Value
            arr(4) = DetermineModelRekrutmen(CStr(arr(1)), "")
            GetFPTKDetail = arr
            Exit Function
        End If
    Next r
    GetFPTKDetail = arr
End Function

Public Function DetermineModelRekrutmen(ByVal posisi As String, Optional ByVal existingModel As Variant) As String
    Dim mapModel As String
    If Len(Trim(CStr(existingModel))) > 0 Then
        DetermineModelRekrutmen = CStr(existingModel)
        Exit Function
    End If
    mapModel = GetModelFromFlowMap(posisi)
    If Len(mapModel) > 0 Then
        DetermineModelRekrutmen = mapModel
        Exit Function
    End If
    posisi = LCase(posisi)
    If InStr(posisi, "fgdp") > 0 Then DetermineModelRekrutmen = "Model 1": Exit Function
    If InStr(posisi, "clap") > 0 Then DetermineModelRekrutmen = "Model 2": Exit Function
    If InStr(posisi, "sales taking order") > 0 Or InStr(posisi, "sto") > 0 Then DetermineModelRekrutmen = "Model 4": Exit Function
    If InStr(posisi, "area sales supervisor") > 0 Or InStr(posisi, "ass") > 0 Then DetermineModelRekrutmen = "Model 3": Exit Function
    If InStr(posisi, "manager") > 0 Then DetermineModelRekrutmen = "Model 2": Exit Function
    DetermineModelRekrutmen = "Model 1"
End Function

Public Function GetModelFromFlowMap(ByVal posisi As String) As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim colPos As Long, colModel As Long
    If Not WSExists("Flow Map Model") Then Exit Function
    Set ws = ThisWorkbook.Worksheets("Flow Map Model")
    colPos = FindHeaderColumn(ws, "Posisi")
    colModel = FindHeaderColumn(ws, "Model Rekrutmen")
    If colModel = 0 Then colModel = FindHeaderColumn(ws, "Model")
    If colPos = 0 Or colModel = 0 Then Exit Function
    lastRow = ws.Cells(ws.rows.Count, colPos).End(xlUp).Row
    For r = 2 To lastRow
        If InStr(1, posisi, CStr(ws.Cells(r, colPos).Value), vbTextCompare) > 0 Then
            GetModelFromFlowMap = CStr(ws.Cells(r, colModel).Value)
            Exit Function
        End If
    Next r
ErrHandler:
End Function

Public Function PickFileCV() As String
    On Error Resume Next
    With Application.FileDialog(3)
        .Title = "Pilih CV Kandidat"
        .Filters.Clear
        .Filters.Add "CV Files", "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.pdf"
        .AllowMultiSelect = False
        If .Show = -1 Then PickFileCV = CStr(.SelectedItems(1))
    End With
End Function

Public Function IsImageFile(ByVal filePath As String) As Boolean
    Dim ext As String
    ext = LCase(Mid(filePath, InStrRev(filePath, ".") + 1))
    IsImageFile = (ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "bmp" Or ext = "gif")
End Function

Public Function ShortFileName(ByVal filePath As String, Optional ByVal maxLen As Long = 35) As String
    Dim p As Long
    p = InStrRev(filePath, Application.PathSeparator)
    If p > 0 Then
        ShortFileName = Mid(filePath, p + 1)
    Else
        ShortFileName = filePath
    End If
    If Len(ShortFileName) > maxLen Then
        ShortFileName = Left(ShortFileName, maxLen - 3) & "..."
    End If
End Function

Public Sub SaveCVRawText(ByVal filePath As String, ByVal rawText As String)
    EnsureCVRawText
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("CV Raw Text")
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = filePath
    ws.Cells(r, 3).Value = ShortFileName(filePath, 80)
    ws.Cells(r, 4).Value = Len(rawText)
    ws.Cells(r, 5).Value = rawText
    ws.Cells(r, 5).WrapText = False
    ws.Columns("A:E").AutoFit
End Sub

Public Function ReadPDFAutoCopy(ByVal pdfPath As String) As String
    On Error GoTo ErrHandler
    Dim txt As String
    txt = ReadPDFAutoCopyFocused(pdfPath)
    If Len(Trim(txt)) < 50 Then
        MsgBox "PDF tidak terbaca. Coba copy manual (Ctrl+A, Ctrl+C) lalu klik OK.", vbInformation
        txt = GetClipboardTextSafe()
    End If
    ReadPDFAutoCopy = txt
    Exit Function
ErrHandler:
    ReadPDFAutoCopy = ""
End Function

Private Function ReadPDFAutoCopyFocused(ByVal pdfPath As String) As String
    On Error GoTo ErrHandler
    Dim i As Long, clipText As String
    ClearClipboardText
    Shell "cmd /c start """" " & Chr(34) & pdfPath & Chr(34), vbNormalFocus
    Application.Wait Now + TimeValue("00:00:04")
    AppActivatePDFWindow pdfPath
    Application.Wait Now + TimeValue("00:00:01")
    Application.SendKeys "^a", True
    Application.Wait Now + TimeValue("00:00:01")
    Application.SendKeys "^c", True
    Application.Wait Now + TimeValue("00:00:01")
    For i = 1 To 6
        clipText = GetClipboardTextSafe()
        If Len(Trim(clipText)) > 20 Then Exit For
        AppActivatePDFWindow pdfPath
        Application.SendKeys "^c", True
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    ReadPDFAutoCopyFocused = clipText
    Exit Function
ErrHandler:
    ReadPDFAutoCopyFocused = ""
End Function

Private Function AppActivatePDFWindow(ByVal pdfPath As String) As Boolean
    On Error Resume Next
    Dim wsh As Object, fileName As String
    Set wsh = CreateObject("WScript.Shell")
    fileName = ShortFileName(pdfPath, 200)
    AppActivatePDFWindow = wsh.AppActivate(fileName)
    If AppActivatePDFWindow Then Exit Function
    AppActivatePDFWindow = wsh.AppActivate("Adobe Acrobat")
    If AppActivatePDFWindow Then Exit Function
    AppActivatePDFWindow = wsh.AppActivate("Adobe Reader")
    If AppActivatePDFWindow Then Exit Function
    AppActivatePDFWindow = wsh.AppActivate("Acrobat")
    If AppActivatePDFWindow Then Exit Function
    AppActivatePDFWindow = wsh.AppActivate("Microsoft Edge")
    If AppActivatePDFWindow Then Exit Function
    AppActivatePDFWindow = wsh.AppActivate("Google Chrome")
End Function

Private Sub ClearClipboardText()
    On Error Resume Next
    Dim dataObj As Object
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText ""
    dataObj.PutInClipboard
End Sub

Public Function GetClipboardTextSafe() As String
    On Error GoTo ErrHandler
    Dim dataObj As Object
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.GetFromClipboard
    GetClipboardTextSafe = dataObj.GetText
    Exit Function
ErrHandler:
    GetClipboardTextSafe = ""
End Function

Public Function GetUniversityTierCV(ByVal univName As String) As String
    Dim u As String
    u = LCase(Trim(univName))
    If Len(u) = 0 Then Exit Function
    Select Case u
        Case "universitas indonesia", "universitas gadjah mada", "institut teknologi bandung"
            GetUniversityTierCV = "Top 3 PTN"
        Case "universitas airlangga", "universitas padjadjaran", "universitas diponegoro", "universitas brawijaya", "institut pertanian bogor", "universitas sebelas maret", "institut teknologi sepuluh nopember"
            GetUniversityTierCV = "Top 10 PTN"
        Case "telkom university", "universitas bina nusantara", "universitas gunadarma"
            GetUniversityTierCV = "Top 10 PTS"
        Case Else
            GetUniversityTierCV = "Lainnya"
    End Select
End Function

Public Function GetIPKTierCV(ByVal ipkText As String) As String
    Dim gpa As Double
    If Len(Trim(ipkText)) = 0 Then Exit Function
    gpa = val(Replace(Trim(ipkText), ",", "."))
    If gpa >= 3.5 Then
        GetIPKTierCV = "Lebih dari 3,5"
    ElseIf gpa >= 3.2 Then
        GetIPKTierCV = "Lebih dari 3,2"
    ElseIf gpa > 0 Then
        GetIPKTierCV = "Kurang dari 3,2"
    End If
End Function

Public Function ParseCVTextSmart(ByVal rawText As String) As Variant
    Dim info(11) As String
    Dim lines() As String, line As String, nl As String, low As String
    Dim i As Long, j As Long, regex As Object, oneLine As String
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    rawText = Replace(rawText, vbCr, vbLf)
    rawText = Replace(rawText, vbTab, vbLf)
    Do While InStr(rawText, vbLf & vbLf & vbLf) > 0
        rawText = Replace(rawText, vbLf & vbLf & vbLf, vbLf & vbLf)
    Loop
    If Len(Trim(rawText)) = 0 Then ParseCVTextSmart = info: Exit Function
    lines = Split(rawText, vbLf)
    oneLine = Replace(Replace(rawText, vbCr, " "), vbLf, " | ")

    ' 0 Nama
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If i < UBound(lines) Then nl = Trim(lines(i + 1)) Else nl = ""
        low = LCase(line)
        If StartsWithAny(low, Array("nama", "name", "full name", "candidate name")) Then
            info(0) = CleanTextValue(ExtractAfterSep(line))
            If Len(info(0)) = 0 And LooksLikeName(nl) Then info(0) = nl
            If Len(info(0)) > 0 Then Exit For
        End If
    Next i
    If Len(info(0)) = 0 Then
        For i = 0 To Application.Min(UBound(lines), 35)
            line = CleanTextValue(Trim(lines(i)))
            If LooksLikeName(line) Then
                info(0) = line
                Exit For
            End If
        Next i
    End If

    ' 1 Email
    regex.Pattern = "[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}"
    If regex.Test(rawText) Then info(1) = regex.Execute(rawText)(0).Value

    ' 2 Nomor HP
    regex.Pattern = "(\+?62|0)[\s\-\.\(\)]?[0-9]{2,4}[\s\-\.\(\)]?[0-9]{3,4}[\s\-\.\(\)]?[0-9]{3,5}"
    If regex.Test(rawText) Then info(2) = CleanPhone(regex.Execute(rawText)(0).Value)

    ' 3 Universitas
    info(3) = NormalizeUniversity(rawText)

    ' 4 Jurusan
    info(4) = NormalizeMajor(rawText)

    ' 5 IPK
    regex.Pattern = "(IPK|GPA|Grade Point Average)[^0-9]{0,40}([0-4][\.,]\d{1,2})"
    If regex.Test(oneLine) Then info(5) = Replace(regex.Execute(oneLine)(0).SubMatches(1), ",", ".")
    If Len(info(5)) = 0 Then
        regex.Pattern = "\b([0-4][\.,]\d{1,2})\s*/\s*4"
        If regex.Test(oneLine) Then info(5) = Replace(regex.Execute(oneLine)(0).SubMatches(0), ",", ".")
    End If
    If Len(info(5)) = 0 Then
        regex.Pattern = "\b([0-4][\.,]\d{1,2})\b"
        If regex.Test(oneLine) Then info(5) = Replace(regex.Execute(oneLine)(0).SubMatches(0), ",", ".")
    End If

    ' 6 Tahun Lulus
    regex.Pattern = "\b(20[0-9]{2})\s*[-–—]\s*([A-Za-z]{3,9}\s*)?(20[0-9]{2})\b"
    If regex.Test(oneLine) Then info(6) = regex.Execute(oneLine)(0).SubMatches(2)
    If Len(info(6)) = 0 Then
        regex.Pattern = "(Lulus|Graduated|Graduation|Tahun Lulus|Class of)[^0-9]*(20[0-9]{2})"
        If regex.Test(oneLine) Then info(6) = regex.Execute(oneLine)(0).SubMatches(1)
    End If
    If Len(info(6)) = 0 Then
        regex.Pattern = "\b(20[0-9]{2})\b"
        If regex.Test(oneLine) Then
            Dim ms As Object
            Set ms = regex.Execute(oneLine)
            info(6) = ms(ms.Count - 1).SubMatches(0)
        End If
    End If

    ' 7 Skor Bahasa Inggris
    regex.Pattern = "(TOEFL|IELTS|TOEIC|English|Bahasa Inggris)[^0-9]*(\d{2,3}(\.\d)?)"
    If regex.Test(rawText) Then info(7) = regex.Execute(rawText)(0).SubMatches(1)

    ' 8 Domisili
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If i < UBound(lines) Then nl = Trim(lines(i + 1)) Else nl = ""
        low = LCase(line)
        If StartsWithAny(low, Array("domisili", "location", "lokasi", "alamat", "address", "city", "kota")) Then
            info(8) = CleanTextValue(ExtractAfterSep(line))
            If Len(info(8)) = 0 Then info(8) = CleanTextValue(nl)
            If Len(info(8)) > 0 Then Exit For
        End If
    Next i
    If Len(info(8)) = 0 Then
        Dim cities As Variant
        cities = Array("Jakarta", "Bogor", "Depok", "Tangerang", "Bekasi", "Bandung", "Semarang", "Surabaya", "Yogyakarta", "Jogja", "Malang", "Medan", "Makassar", "Denpasar", "Bali")
        For j = 0 To UBound(cities)
            If InStr(1, rawText, cities(j), vbTextCompare) > 0 Then
                info(8) = cities(j)
                Exit For
            End If
        Next j
    End If

    ' 9 Jenjang Pendidikan
    info(9) = ExtractEducationLevel(rawText)

    ' 10 University Tier
    info(10) = GetUniversityTierCV(info(3))

    ' 11 IPK Tier
    info(11) = GetIPKTierCV(info(5))

    ParseCVTextSmart = info
End Function

Private Function NormalizeUniversity(ByVal text As String) As String
    Dim t As String, i As Long
    t = LCase(Replace(Replace(Replace(text, vbCr, " "), vbLf, " "), vbTab, " "))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = " " & Trim(t) & " "
    If InStr(1, t, " universitas indonesia ") > 0 Or InStr(1, t, " university of indonesia ") > 0 Or InStr(1, t, " indonesia university ") > 0 Then NormalizeUniversity = "Universitas Indonesia": Exit Function
    If InStr(1, t, " universitas gadjah mada ") > 0 Or InStr(1, t, " gadjah mada university ") > 0 Or InStr(1, t, " ugm ") > 0 Then NormalizeUniversity = "Universitas Gadjah Mada": Exit Function
    If InStr(1, t, " institut teknologi bandung ") > 0 Or InStr(1, t, " bandung institute of technology ") > 0 Or InStr(1, t, " itb ") > 0 Then NormalizeUniversity = "Institut Teknologi Bandung": Exit Function
    If InStr(1, t, " universitas airlangga ") > 0 Or InStr(1, t, " airlangga university ") > 0 Or InStr(1, t, " unair ") > 0 Then NormalizeUniversity = "Universitas Airlangga": Exit Function
    If InStr(1, t, " universitas padjadjaran ") > 0 Or InStr(1, t, " padjadjaran university ") > 0 Or InStr(1, t, " unpad ") > 0 Then NormalizeUniversity = "Universitas Padjadjaran": Exit Function
    If InStr(1, t, " universitas diponegoro ") > 0 Or InStr(1, t, " diponegoro university ") > 0 Or InStr(1, t, " diponegro university ") > 0 Or InStr(1, t, " undip ") > 0 Then NormalizeUniversity = "Universitas Diponegoro": Exit Function
    If InStr(1, t, " universitas brawijaya ") > 0 Or InStr(1, t, " brawijaya university ") > 0 Then NormalizeUniversity = "Universitas Brawijaya": Exit Function
    If InStr(1, t, " institut pertanian bogor ") > 0 Or InStr(1, t, " bogor agricultural university ") > 0 Or InStr(1, t, " ipb ") > 0 Then NormalizeUniversity = "Institut Pertanian Bogor": Exit Function
    If InStr(1, t, " universitas sebelas maret ") > 0 Or InStr(1, t, " sebelas maret university ") > 0 Or InStr(1, t, " uns ") > 0 Then NormalizeUniversity = "Universitas Sebelas Maret": Exit Function
    If InStr(1, t, " institut teknologi sepuluh nopember ") > 0 Or InStr(1, t, " sepuluh nopember institute ") > 0 Or InStr(1, t, " its ") > 0 Then NormalizeUniversity = "Institut Teknologi Sepuluh Nopember": Exit Function
    If InStr(1, t, " telkom university ") > 0 Or InStr(1, t, " universitas telkom ") > 0 Then NormalizeUniversity = "Telkom University": Exit Function
    If InStr(1, t, " universitas bina nusantara ") > 0 Or InStr(1, t, " bina nusantara university ") > 0 Or InStr(1, t, " binus ") > 0 Then NormalizeUniversity = "Universitas Bina Nusantara": Exit Function
    If InStr(1, t, " universitas gunadarma ") > 0 Or InStr(1, t, " gunadarma university ") > 0 Then NormalizeUniversity = "Universitas Gunadarma": Exit Function
    Dim lines2() As String, ln As String
    lines2 = Split(text, vbLf)
    For i = 0 To UBound(lines2)
        ln = Trim(lines2(i))
        If StartsWithAny(LCase(ln), Array("universitas", "institut", "politeknik", "sekolah tinggi", "smk ", "sma ")) Then
            NormalizeUniversity = ln
            Exit Function
        End If
    Next i
End Function

Private Function NormalizeMajor(ByVal text As String) As String
    Dim t As String
    t = LCase(Replace(Replace(Replace(text, vbCr, " "), vbLf, " "), vbTab, " "))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    If InStr(1, t, "teknik industri") > 0 Or InStr(1, t, "industrial engineering") > 0 Then NormalizeMajor = "Teknik Industri": Exit Function
    If InStr(1, t, "sistem informasi") > 0 Or InStr(1, t, "information system") > 0 Then NormalizeMajor = "Sistem Informasi": Exit Function
    If InStr(1, t, "teknik informatika") > 0 Or InStr(1, t, "informatics engineering") > 0 Or InStr(1, t, "computer science") > 0 Then NormalizeMajor = "Teknik Informatika": Exit Function
    If InStr(1, t, "teknik mesin") > 0 Or InStr(1, t, "mechanical engineering") > 0 Then NormalizeMajor = "Teknik Mesin": Exit Function
    If InStr(1, t, "teknik elektro") > 0 Or InStr(1, t, "electrical engineering") > 0 Then NormalizeMajor = "Teknik Elektro": Exit Function
    If InStr(1, t, "teknik kimia") > 0 Or InStr(1, t, "chemical engineering") > 0 Then NormalizeMajor = "Teknik Kimia": Exit Function
    If InStr(1, t, "teknik sipil") > 0 Or InStr(1, t, "civil engineering") > 0 Then NormalizeMajor = "Teknik Sipil": Exit Function
    If InStr(1, t, "akuntansi") > 0 Or InStr(1, t, "accounting") > 0 Then NormalizeMajor = "Akuntansi": Exit Function
    If InStr(1, t, "manajemen") > 0 Or InStr(1, t, "management") > 0 Then NormalizeMajor = "Manajemen": Exit Function
    If InStr(1, t, "ekonomi") > 0 Or InStr(1, t, "economics") > 0 Then NormalizeMajor = "Ekonomi": Exit Function
    If InStr(1, t, "psikologi") > 0 Or InStr(1, t, "psychology") > 0 Then NormalizeMajor = "Psikologi": Exit Function
    If InStr(1, t, "hukum") > 0 Or InStr(1, t, "law") > 0 Then NormalizeMajor = "Hukum": Exit Function
    If InStr(1, t, "ilmu komunikasi") > 0 Or InStr(1, t, "communication") > 0 Then NormalizeMajor = "Ilmu Komunikasi": Exit Function
    If InStr(1, t, "administrasi bisnis") > 0 Or InStr(1, t, "business administration") > 0 Then NormalizeMajor = "Administrasi Bisnis": Exit Function
    If InStr(1, t, "statistika") > 0 Or InStr(1, t, "statistics") > 0 Then NormalizeMajor = "Statistika": Exit Function
End Function

Private Function ExtractEducationLevel(ByVal text As String) As String
    Dim t As String
    t = LCase(text)
    If InStr(t, "s1") > 0 Or InStr(t, "bachelor") > 0 Or InStr(t, "sarjana") > 0 Then ExtractEducationLevel = "S1": Exit Function
    If InStr(t, "s2") > 0 Or InStr(t, "master") > 0 Or InStr(t, "magister") > 0 Then ExtractEducationLevel = "S2": Exit Function
    If InStr(t, "d4") > 0 Or InStr(t, "diploma 4") > 0 Or InStr(t, "diploma iv") > 0 Then ExtractEducationLevel = "D4": Exit Function
    If InStr(t, "d3") > 0 Or InStr(t, "diploma 3") > 0 Then ExtractEducationLevel = "D3": Exit Function
    If InStr(t, "smk") > 0 Or InStr(t, "vocational") > 0 Then ExtractEducationLevel = "SMK": Exit Function
    If InStr(t, "sma") > 0 Or InStr(t, "high school") > 0 Then ExtractEducationLevel = "SMA": Exit Function
End Function

Public Function StartsWithAny(ByVal text As String, ByVal arr As Variant) As Boolean
    Dim v As Variant
    For Each v In arr
        If Left$(Trim(text), Len(CStr(v))) = LCase(CStr(v)) Then
            StartsWithAny = True
            Exit Function
        End If
    Next v
End Function

Public Function ExtractAfterSep(ByVal text As String) As String
    Dim p As Long
    p = InStr(1, text, ":")
    If p = 0 Then p = InStr(1, text, "-")
    If p > 0 Then
        ExtractAfterSep = Trim(Mid$(text, p + 1))
    Else
        ExtractAfterSep = ""
    End If
End Function

Public Function CleanTextValue(ByVal text As String) As String
    text = Trim(text)
    text = Replace(text, Chr(34), "")
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    CleanTextValue = text
End Function

Public Function LooksLikeName(ByVal text As String) As Boolean
    Dim words As Variant, w As Variant
    text = CleanTextValue(text)
    If Len(text) < 5 Or Len(text) > 60 Then Exit Function
    If InStr(1, text, "@") > 0 Then Exit Function
    If text Like "*[0-9]*" Then Exit Function
    If IsLabelLine(text) Then Exit Function
    words = Split(text, " ")
    If UBound(words) < 1 Or UBound(words) > 5 Then Exit Function
    For Each w In words
        If Len(CStr(w)) < 2 Then Exit Function
    Next w
    LooksLikeName = True
End Function

Private Function IsLabelLine(ByVal text As String) As Boolean
    Dim labels As Variant, l As Variant
    labels = Array("CV", "CURRICULUM", "RESUME", "DATA", "PENDIDIKAN", "EDUCATION", "PENGALAMAN", "EXPERIENCE", "SKILL", "KEMAMPUAN", "BAHASA", "HOBI", "PROFIL", "PROFILE", "TENTANG", "ABOUT")
    For Each l In labels
        If InStr(1, UCase(text), UCase(l), vbTextCompare) = 1 Then
            IsLabelLine = True
            Exit Function
        End If
    Next l
End Function

Public Function CleanPhone(ByVal text As String) As String
    text = Replace(text, " ", "")
    text = Replace(text, "-", "")
    text = Replace(text, ".", "")
    text = Replace(text, "(", "")
    text = Replace(text, ")", "")
    text = Replace(text, "+", "")
    CleanPhone = text
End Function

Public Sub TestClipboardCVText()
    Dim t As String
    t = GetClipboardTextSafe()
    If Len(t) = 0 Then
        MsgBox "Clipboard kosong.", vbExclamation
    Else
        SaveCVRawText "Manual Clipboard Test", t
        MsgBox "Clipboard terbaca: " & Len(t) & " karakter.", vbInformation
    End If
End Sub



