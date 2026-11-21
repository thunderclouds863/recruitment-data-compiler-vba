Attribute VB_Name = "Module2"
Option Explicit

Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_MSForm As Long = 3
Private gStep As String

Public Sub InstallMonitoringSourcing()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    StepNow "Checking VBA project access"
    If Not CanAccessVBProject() Then
        MsgBox "VBA project access masih blocked." & vbCrLf & vbCrLf & _
               "Aktifkan: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model.", vbCritical
        GoTo SafeExit
    End If

    StepNow "Cleaning old components"
    RemoveIfExists "modMonitoring"
    RemoveIfExists "frmUploadEvidence"

    StepNow "Creating module"
    CreateModuleWithCode "modMonitoring", Code_modMonitoring()

    StepNow "Creating form"
    CreateBlankUserForm "frmUploadEvidence", Code_frmUploadEvidence()

    MsgBox "Install Monitoring Sourcing berhasil!" & vbCrLf & vbCrLf & _
           "Module: modMonitoring" & vbCrLf & _
           "Form: frmUploadEvidence" & vbCrLf & vbCrLf & _
           "Gunakan macro: CreateMonitoringSheet, OpenUploadEvidenceForm", vbInformation

SafeExit:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Install failed at step: " & gStep & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume SafeExit
End Sub

Private Sub StepNow(ByVal text As String)
    gStep = text
    Application.StatusBar = text
End Sub

Private Function CanAccessVBProject() As Boolean
    On Error Resume Next
    Dim n As Long
    n = ThisWorkbook.VBProject.VBComponents.Count
    CanAccessVBProject = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub RemoveIfExists(ByVal componentName As String)
    On Error Resume Next
    Dim comp As Object
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If LCase(comp.Name) = LCase(componentName) Then
            ThisWorkbook.VBProject.VBComponents.Remove comp
            Exit For
        End If
    Next comp
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub CreateModuleWithCode(ByVal moduleName As String, ByVal codeText As String)
    Dim comp As Object
    Set comp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    comp.Name = moduleName
    comp.CodeModule.AddFromString codeText
End Sub

Private Sub CreateBlankUserForm(ByVal formName As String, ByVal codeText As String)
    Dim comp As Object
    Set comp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    comp.Name = formName
    comp.CodeModule.AddFromString codeText
End Sub


' ==================== MODULE: modMonitoring ====================

Private Function Code_modMonitoring() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public SelectedEvidenceFolder As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ==================== MONITORING SOURCING ====================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub CreateMonitoringSheet()" & vbCrLf
    s = s & "    Dim t As Single" & vbCrLf
    s = s & "    t = Timer" & vbCrLf
    s = s & "    Dim wsMon As Worksheet" & vbCrLf
    s = s & "    Set wsMon = GetOrCreateSheet(""Monitoring Sourcing"")" & vbCrLf
    s = s & "    BuildMonitoringHeaders wsMon" & vbCrLf
    s = s & "    FastFillMonitoringData wsMon" & vbCrLf
    s = s & "    wsMon.Columns(""A:O"").AutoFit" & vbCrLf
    s = s & "    Application.StatusBar = False" & vbCrLf
    s = s & "    Application.ScreenUpdating = True" & vbCrLf
    s = s & "    Application.Calculation = xlCalculationAutomatic" & vbCrLf
    s = s & "    MsgBox ""Selesai! "" & Format(Timer - t, ""0.0"") & "" detik"", vbInformation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub OpenUploadEvidenceForm()" & vbCrLf
    s = s & "    If Not WSExists(""Monitoring Sourcing"") Then" & vbCrLf
    s = s & "        MsgBox ""Sheet Monitoring Sourcing belum ada. Jalankan CreateMonitoringSheet dulu."", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    frmUploadEvidence.Show vbModal" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildMonitoringHeaders(ByVal ws As Worksheet)" & vbCrLf
    s = s & "    ws.Cells.Clear" & vbCrLf
    s = s & "    ws.Cells(1, 1).Value = ""Kode Unik""" & vbCrLf
    s = s & "    ws.Cells(1, 2).Value = ""Posisi""" & vbCrLf
    s = s & "    ws.Cells(1, 3).Value = ""PIC Recruiter""" & vbCrLf
    s = s & "    ws.Cells(1, 4).Value = ""Status FPTK""" & vbCrLf
    s = s & "    Dim startDate As Date" & vbCrLf
    s = s & "    startDate = GetStartOfWeek(Date)" & vbCrLf
    s = s & "    Dim weekNum As Long" & vbCrLf
    s = s & "    weekNum = IsoWeekNumber(Date)" & vbCrLf
    s = s & "    Dim col As Long" & vbCrLf
    s = s & "    col = 5" & vbCrLf
    s = s & "    Dim d As Long" & vbCrLf
    s = s & "    For d = 0 To 4" & vbCrLf
    s = s & "        ws.Cells(1, col).Value = Format(startDate + d, ""dd-mmm-yy"")" & vbCrLf
    s = s & "        ws.Cells(1, col).ColumnWidth = 11" & vbCrLf
    s = s & "        col = col + 1" & vbCrLf
    s = s & "        ws.Cells(1, col).Value = ""Evidence "" & (d + 1)" & vbCrLf
    s = s & "        ws.Cells(1, col).ColumnWidth = 35" & vbCrLf
    s = s & "        col = col + 1" & vbCrLf
    s = s & "    Next d" & vbCrLf
    s = s & "    ws.Cells(1, col).Value = ""Total Week "" & weekNum" & vbCrLf
    s = s & "    ws.Cells(1, col).ColumnWidth = 16" & vbCrLf
    s = s & "    With ws.Range(""A1:O1"")" & vbCrLf
    s = s & "        .Font.Bold = True" & vbCrLf
    s = s & "        .Interior.Color = RGB(44, 62, 80)" & vbCrLf
    s = s & "        .Font.Color = vbWhite" & vbCrLf
    s = s & "        .WrapText = True" & vbCrLf
    s = s & "        .HorizontalAlignment = xlCenter" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "    ws.Rows(1).RowHeight = 36" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FastFillMonitoringData(ByVal ws As Worksheet)" & vbCrLf
    s = s & "    Application.ScreenUpdating = False" & vbCrLf
    s = s & "    Application.Calculation = xlCalculationManual" & vbCrLf
    s = s & "    Application.StatusBar = ""Membaca data...""" & vbCrLf
    s = s & "    Dim wsFPTK As Worksheet" & vbCrLf
    s = s & "    Set wsFPTK = GetFPTKWorksheet()" & vbCrLf
    s = s & "    If wsFPTK Is Nothing Then GoTo CleanExit" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    Dim wsDB As Worksheet" & vbCrLf
    s = s & "    Set wsDB = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    Dim colKodeFP As Long" & vbCrLf
    s = s & "    colKodeFP = FindHeaderColumn(wsFPTK, ""Kode Unik"")" & vbCrLf
    s = s & "    Dim colPosFP As Long" & vbCrLf
    s = s & "    colPosFP = FindHeaderColumn(wsFPTK, ""Posisi - Kebutuhan TA"")" & vbCrLf
    s = s & "    Dim colPIC As Long" & vbCrLf
    s = s & "    colPIC = FindHeaderColumn(wsFPTK, ""PIC Recruiter"")" & vbCrLf
    s = s & "    Dim colStatus As Long" & vbCrLf
    s = s & "    colStatus = FindHeaderColumn(wsFPTK, ""Status"")" & vbCrLf
    s = s & "    If colKodeFP = 0 Then GoTo CleanExit" & vbCrLf
    s = s & "    Dim colKodeDB As Long" & vbCrLf
    s = s & "    colKodeDB = FindHeaderColumn(wsDB, ""Kode Unik (copy value dari FPTK)"")" & vbCrLf
    s = s & "    Dim colTglDB As Long" & vbCrLf
    s = s & "    colTglDB = FindHeaderColumn(wsDB, ""Tanggal Sourcing"")" & vbCrLf
    s = s & "    If colKodeDB = 0 Or colTglDB = 0 Then GoTo CleanExit" & vbCrLf
    s = s & "    Application.StatusBar = ""Cache DB Sourcing...""" & vbCrLf
    s = s & "    Dim cache As Object" & vbCrLf
    s = s & "    Set cache = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    s = s & "    Dim lastDB As Long" & vbCrLf
    s = s & "    lastDB = wsDB.Cells(wsDB.Rows.Count, colKodeDB).End(xlUp).Row" & vbCrLf
    s = s & "    Dim arrDB As Variant" & vbCrLf
    s = s & "    arrDB = wsDB.Range(wsDB.Cells(2, colKodeDB), wsDB.Cells(lastDB, colTglDB)).Value" & vbCrLf
    s = s & "    Dim rr As Long" & vbCrLf
    s = s & "    For rr = 1 To UBound(arrDB, 1)" & vbCrLf
    s = s & "        Dim kd As String" & vbCrLf
    s = s & "        kd = Trim(CStr(arrDB(rr, 1)))" & vbCrLf
    s = s & "        If Len(kd) > 0 Then" & vbCrLf
    s = s & "            Dim rawTgl As String" & vbCrLf
    s = s & "            rawTgl = Trim(CStr(arrDB(rr, 2)))" & vbCrLf
    s = s & "            If Len(rawTgl) > 0 Then" & vbCrLf
    s = s & "                rawTgl = Replace(rawTgl, ""-"", ""/"")" & vbCrLf
    s = s & "                rawTgl = Replace(rawTgl, ""."", ""/"")" & vbCrLf
    s = s & "                rawTgl = Replace(rawTgl, "" "", ""/"")" & vbCrLf
    s = s & "                If IsDate(rawTgl) Then" & vbCrLf
    s = s & "                    Dim tglKey As String" & vbCrLf
    s = s & "                    tglKey = Format(CDate(rawTgl), ""yyyy-mm-dd"")" & vbCrLf
    s = s & "                    Dim ck As String" & vbCrLf
    s = s & "                    ck = kd & ""|"" & tglKey" & vbCrLf
    s = s & "                    If cache.Exists(ck) Then" & vbCrLf
    s = s & "                        cache(ck) = CLng(cache(ck)) + 1" & vbCrLf
    s = s & "                    Else" & vbCrLf
    s = s & "                        cache.Add ck, 1" & vbCrLf
    s = s & "                    End If" & vbCrLf
    s = s & "                End If" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next rr" & vbCrLf
    s = s & "    Application.StatusBar = ""Proses data...""" & vbCrLf
    s = s & "    Dim lastFP As Long" & vbCrLf
    s = s & "    lastFP = wsFPTK.Cells(wsFPTK.Rows.Count, colKodeFP).End(xlUp).Row" & vbCrLf
    s = s & "    Dim maxCol As Long" & vbCrLf
    s = s & "    maxCol = colKodeFP" & vbCrLf
    s = s & "    If colPosFP > maxCol Then maxCol = colPosFP" & vbCrLf
    s = s & "    If colPIC > maxCol Then maxCol = colPIC" & vbCrLf
    s = s & "    If colStatus > maxCol Then maxCol = colStatus" & vbCrLf
    s = s & "    Dim arrFP As Variant" & vbCrLf
    s = s & "    arrFP = wsFPTK.Range(wsFPTK.Cells(2, 1), wsFPTK.Cells(lastFP, maxCol)).Value" & vbCrLf
    s = s & "    Dim outArr() As Variant" & vbCrLf
    s = s & "    ReDim outArr(1 To lastFP - 1, 1 To 15)" & vbCrLf
    s = s & "    Dim startDate As Date" & vbCrLf
    s = s & "    startDate = GetStartOfWeek(Date)" & vbCrLf
    s = s & "    Dim dateKeys(0 To 4) As String" & vbCrLf
    s = s & "    Dim d As Long" & vbCrLf
    s = s & "    For d = 0 To 4" & vbCrLf
    s = s & "        dateKeys(d) = Format(startDate + d, ""yyyy-mm-dd"")" & vbCrLf
    s = s & "    Next d" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim rw As Long" & vbCrLf
    s = s & "    rw = 0" & vbCrLf
    s = s & "    For r = 1 To UBound(arrFP, 1)" & vbCrLf
    s = s & "        Dim kodeUnik As String" & vbCrLf
    s = s & "        kodeUnik = Trim(CStr(arrFP(r, colKodeFP)))" & vbCrLf
    s = s & "        If Len(kodeUnik) = 0 Then GoTo NextRow" & vbCrLf
    s = s & "        rw = rw + 1" & vbCrLf
    s = s & "        outArr(rw, 1) = kodeUnik" & vbCrLf
    s = s & "        If colPosFP > 0 Then outArr(rw, 2) = CStr(arrFP(r, colPosFP))" & vbCrLf
    s = s & "        If colPIC > 0 Then outArr(rw, 3) = CStr(arrFP(r, colPIC))" & vbCrLf
    s = s & "        If colStatus > 0 Then outArr(rw, 4) = CStr(arrFP(r, colStatus))" & vbCrLf
    s = s & "        Dim weekTotal As Long" & vbCrLf
    s = s & "        weekTotal = 0" & vbCrLf
    s = s & "        Dim col As Long" & vbCrLf
    s = s & "        col = 5" & vbCrLf
    s = s & "        For d = 0 To 4" & vbCrLf
    s = s & "            Dim cKey As String" & vbCrLf
    s = s & "            cKey = kodeUnik & ""|"" & dateKeys(d)" & vbCrLf
    s = s & "            Dim cvCount As Long" & vbCrLf
    s = s & "            cvCount = 0" & vbCrLf
    s = s & "            If cache.Exists(cKey) Then" & vbCrLf
    s = s & "                cvCount = CLng(cache(cKey))" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            weekTotal = weekTotal + cvCount" & vbCrLf
    s = s & "            outArr(rw, col) = cvCount" & vbCrLf
    s = s & "            col = col + 1" & vbCrLf
    s = s & "            outArr(rw, col) = """"" & vbCrLf
    s = s & "            col = col + 1" & vbCrLf
    s = s & "        Next d" & vbCrLf
    s = s & "        outArr(rw, 15) = weekTotal" & vbCrLf
    s = s & "NextRow:" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    Application.StatusBar = ""Tulis ke sheet...""" & vbCrLf
    s = s & "    If rw > 0 Then" & vbCrLf
    s = s & "        ws.Range(""A2"").Resize(rw, 15).Value = outArr" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "CleanExit:" & vbCrLf
    s = s & "    Application.StatusBar = False" & vbCrLf
    s = s & "    Application.Calculation = xlCalculationAutomatic" & vbCrLf
    s = s & "    Application.ScreenUpdating = True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub WriteEvidenceLink(ByVal posisi As String, ByVal tglEvidence As String, ByVal linkPath As String)" & vbCrLf
    s = s & "    Dim wsMon As Worksheet" & vbCrLf
    s = s & "    Set wsMon = ThisWorkbook.Worksheets(""Monitoring Sourcing"")" & vbCrLf
    s = s & "    Dim lastRow As Long" & vbCrLf
    s = s & "    lastRow = wsMon.Cells(wsMon.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    s = s & "    Dim startDate As Date" & vbCrLf
    s = s & "    startDate = GetStartOfWeek(Date)" & vbCrLf
    s = s & "    Dim targetDate As Date" & vbCrLf
    s = s & "    If IsDate(tglEvidence) Then targetDate = CDate(tglEvidence)" & vbCrLf
    s = s & "    Dim colIdx As Long" & vbCrLf
    s = s & "    colIdx = -1" & vbCrLf
    s = s & "    Dim d As Long" & vbCrLf
    s = s & "    For d = 0 To 4" & vbCrLf
    s = s & "        If startDate + d = targetDate Then" & vbCrLf
    s = s & "            colIdx = 6 + (d * 2)" & vbCrLf
    s = s & "            Exit For" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next d" & vbCrLf
    s = s & "    If colIdx = -1 Then Exit Sub" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    For r = 2 To lastRow" & vbCrLf
    s = s & "        If InStr(1, CStr(wsMon.Cells(r, 2).Value), posisi, vbTextCompare) > 0 Then" & vbCrLf
    s = s & "            wsMon.Cells(r, colIdx).Value = linkPath" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function GetStartOfWeek(ByVal dt As Date) As Date" & vbCrLf
    s = s & "    Dim wd As Long" & vbCrLf
    s = s & "    wd = Weekday(dt, vbMonday)" & vbCrLf
    s = s & "    GetStartOfWeek = dt - wd + 1" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsoWeekNumber(ByVal dt As Date) As Long" & vbCrLf
    s = s & "    Dim jan4 As Date" & vbCrLf
    s = s & "    jan4 = DateSerial(Year(dt), 1, 4)" & vbCrLf
    s = s & "    Dim startOfYear As Date" & vbCrLf
    s = s & "    startOfYear = jan4 - Weekday(jan4, vbMonday) + 1" & vbCrLf
    s = s & "    IsoWeekNumber = Int((dt - startOfYear) / 7) + 1" & vbCrLf
    s = s & "End Function" & vbCrLf
    Code_modMonitoring = s
End Function


' ==================== FORM: frmUploadEvidence ====================

Private Function Code_frmUploadEvidence() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mBuilt As Boolean" & vbCrLf
    s = s & "Private cboPosisi As MSForms.ComboBox" & vbCrLf
    s = s & "Private txtTanggal As MSForms.TextBox" & vbCrLf
    s = s & "Private lblTotalCV As MSForms.Label" & vbCrLf
    s = s & "Private lblFilePath As MSForms.Label" & vbCrLf
    s = s & "Private mFilePath As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents btnPilihFile As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnUpload As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnPilihFolder As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnCancel As MSForms.CommandButton" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "    LoadPosisi" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildUI()" & vbCrLf
    s = s & "    If mBuilt Then Exit Sub" & vbCrLf
    s = s & "    mBuilt = True" & vbCrLf
    s = s & "    Me.Caption = ""UPLOAD EVIDENCE SOURCING""" & vbCrLf
    s = s & "    Me.Width = 500" & vbCrLf
    s = s & "    Me.Height = 330" & vbCrLf
    s = s & "    Me.BackColor = RGB(245, 245, 245)" & vbCrLf
    s = s & "    Dim y As Single" & vbCrLf
    s = s & "    y = 12" & vbCrLf
    s = s & "    AddL Me, ""lblTitle"", ""UPLOAD EVIDENCE SOURCING"", 20, y, 440, 20, True, 11" & vbCrLf
    s = s & "    y = y + 32" & vbCrLf
    s = s & "    AddL Me, ""lblPosisi"", ""Pilih Posisi:"", 20, y, 100, 16, True, 8" & vbCrLf
    s = s & "    Set cboPosisi = AddC(Me, ""cboPosisi"", 130, y - 2, 330, 22)" & vbCrLf
    s = s & "    cboPosisi.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    y = y + 32" & vbCrLf
    s = s & "    AddL Me, ""lblTanggal"", ""Tanggal (dd/mm/yyyy):"", 20, y, 130, 16, True, 8" & vbCrLf
    s = s & "    Set txtTanggal = AddT(Me, ""txtTanggal"", 155, y - 2, 120, 22)" & vbCrLf
    s = s & "    txtTanggal.ControlTipText = ""Format: dd/mm/yyyy""" & vbCrLf
    s = s & "    AddL Me, ""lblTotalLabel"", ""Total CV:"", 290, y, 60, 16, True, 8" & vbCrLf
    s = s & "    Set lblTotalCV = AddL(Me, ""lblTotalCV"", ""0"", 355, y, 100, 16, True, 10)" & vbCrLf
    s = s & "    lblTotalCV.ForeColor = RGB(39, 174, 96)" & vbCrLf
    s = s & "    y = y + 36" & vbCrLf
    s = s & "    Set btnPilihFolder = AddB(Me, ""btnPilihFolder"", ""PILIH FOLDER ONEDRIVE"", 20, y, 200, 30, RGB(52, 152, 219))" & vbCrLf
    s = s & "    Set btnPilihFile = AddB(Me, ""btnPilihFile"", ""PILIH FILE BUKTI"", 240, y, 220, 30, RGB(243, 156, 18))" & vbCrLf
    s = s & "    y = y + 36" & vbCrLf
    s = s & "    Set lblFilePath = AddL(Me, ""lblFilePath"", ""Belum ada file dipilih"", 20, y, 440, 16, False, 8)" & vbCrLf
    s = s & "    lblFilePath.ForeColor = RGB(149, 165, 166)" & vbCrLf
    s = s & "    y = y + 28" & vbCrLf
    s = s & "    Set btnUpload = AddB(Me, ""btnUpload"", ""UPLOAD EVIDENCE"", 120, y, 140, 34, RGB(39, 174, 96))" & vbCrLf
    s = s & "    Set btnCancel = AddB(Me, ""btnCancel"", ""TUTUP"", 280, y, 80, 34, RGB(231, 76, 60))" & vbCrLf
    s = s & "    If Len(SelectedEvidenceFolder) > 0 Then" & vbCrLf
    s = s & "        lblFilePath.Caption = ""Folder: "" & SelectedEvidenceFolder" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadPosisi()" & vbCrLf
    s = s & "    Dim wsMon As Worksheet" & vbCrLf
    s = s & "    Set wsMon = ThisWorkbook.Worksheets(""Monitoring Sourcing"")" & vbCrLf
    s = s & "    Dim lastRow As Long" & vbCrLf
    s = s & "    lastRow = wsMon.Cells(wsMon.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim posDict As Object" & vbCrLf
    s = s & "    Set posDict = CreateObject(""Scripting.Dictionary"")" & vbCrLf
    s = s & "    For r = 2 To lastRow" & vbCrLf
    s = s & "        Dim p As String" & vbCrLf
    s = s & "        p = Trim(CStr(wsMon.Cells(r, 2).Value))" & vbCrLf
    s = s & "        If Len(p) > 0 And Not posDict.Exists(p) Then" & vbCrLf
    s = s & "            posDict.Add p, True" & vbCrLf
    s = s & "            cboPosisi.AddItem p" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    If cboPosisi.ListCount > 0 Then cboPosisi.ListIndex = 0" & vbCrLf
    s = s & "    txtTanggal.Value = Format(Date, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboPosisi_Change()" & vbCrLf
    s = s & "    UpdateTotalCV" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtTanggal_Change()" & vbCrLf
    s = s & "    UpdateTotalCV" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateTotalCV()" & vbCrLf
    s = s & "    Dim total As Long" & vbCrLf
    s = s & "    total = 0" & vbCrLf
    s = s & "    If cboPosisi.ListIndex >= 0 And Len(txtTanggal.Value) > 0 Then" & vbCrLf
    s = s & "        Dim wsMon As Worksheet" & vbCrLf
    s = s & "        Set wsMon = ThisWorkbook.Worksheets(""Monitoring Sourcing"")" & vbCrLf
    s = s & "        Dim lastRow As Long" & vbCrLf
    s = s & "        lastRow = wsMon.Cells(wsMon.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    s = s & "        Dim r As Long" & vbCrLf
    s = s & "        For r = 2 To lastRow" & vbCrLf
    s = s & "            If InStr(1, CStr(wsMon.Cells(r, 2).Value), cboPosisi.Value, vbTextCompare) > 0 Then" & vbCrLf
    s = s & "                total = total + CLng(Val(wsMon.Cells(r, 15).Value))" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next r" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lblTotalCV.Caption = CStr(total)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnPilihFolder_Click()" & vbCrLf
    s = s & "    Dim objShell As Object" & vbCrLf
    s = s & "    Set objShell = CreateObject(""Shell.Application"")" & vbCrLf
    s = s & "    Dim objFolder As Object" & vbCrLf
    s = s & "    Set objFolder = objShell.BrowseForFolder(0, ""Pilih folder Evidence di OneDrive:"", 0, 0)" & vbCrLf
    s = s & "    If Not objFolder Is Nothing Then" & vbCrLf
    s = s & "        SelectedEvidenceFolder = objFolder.Self.Path" & vbCrLf
    s = s & "        lblFilePath.Caption = ""Folder: "" & SelectedEvidenceFolder" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnPilihFile_Click()" & vbCrLf
    s = s & "    Dim fd As Object" & vbCrLf
    s = s & "    Set fd = Application.FileDialog(1)" & vbCrLf
    s = s & "    fd.Title = ""Pilih file bukti evidence""" & vbCrLf
    s = s & "    fd.Filters.Clear" & vbCrLf
    s = s & "    fd.Filters.Add ""Semua File"", ""*.*""" & vbCrLf
    s = s & "    If fd.Show = -1 Then" & vbCrLf
    s = s & "        mFilePath = CStr(fd.SelectedItems(1))" & vbCrLf
    s = s & "        lblFilePath.Caption = ""File: "" & Mid(mFilePath, InStrRev(mFilePath, ""\"") + 1)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnUpload_Click()" & vbCrLf
    s = s & "    If cboPosisi.ListIndex < 0 Then" & vbCrLf
    s = s & "        MsgBox ""Pilih posisi dulu!"", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Len(Trim(txtTanggal.Value)) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Isi tanggal dulu!"", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Len(mFilePath) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Pilih file bukti dulu!"", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Len(SelectedEvidenceFolder) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Pilih folder OneDrive dulu!"", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim posisi As String" & vbCrLf
    s = s & "    posisi = CleanName(cboPosisi.Value)" & vbCrLf
    s = s & "    Dim tglStr As String" & vbCrLf
    s = s & "    tglStr = Trim(txtTanggal.Value)" & vbCrLf
    s = s & "    tglStr = Replace(tglStr, ""-"", ""/"")" & vbCrLf
    s = s & "    tglStr = Replace(tglStr, ""."", ""/"")" & vbCrLf
    s = s & "    If Not IsDate(tglStr) Then" & vbCrLf
    s = s & "        MsgBox ""Format tanggal salah! Gunakan dd/mm/yyyy"", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim dtEvidence As Date" & vbCrLf
    s = s & "    dtEvidence = CDate(tglStr)" & vbCrLf
    s = s & "    Dim totalCV As Long" & vbCrLf
    s = s & "    totalCV = CLng(Val(lblTotalCV.Caption))" & vbCrLf
    s = s & "    Dim ext As String" & vbCrLf
    s = s & "    ext = LCase(Mid(mFilePath, InStrRev(mFilePath, ""."") + 1))" & vbCrLf
    s = s & "    Dim fileName As String" & vbCrLf
    s = s & "    fileName = Format(dtEvidence, ""yyyy-mm-dd"") & "" "" & posisi & "" - "" & totalCV & "" CV."" & ext" & vbCrLf
    s = s & "    Dim destPath As String" & vbCrLf
    s = s & "    destPath = SelectedEvidenceFolder & ""\"" & fileName" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Kill destPath" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    FileCopy mFilePath, destPath" & vbCrLf
    s = s & "    WriteEvidenceLink cboPosisi.Value, Format(dtEvidence, ""dd/mm/yyyy""), destPath" & vbCrLf
    s = s & "    MsgBox ""Evidence berhasil diupload!"" & vbCrLf & vbCrLf & destPath, vbInformation" & vbCrLf
    s = s & "    mFilePath = """"" & vbCrLf
    s = s & "    lblFilePath.Caption = ""Folder: "" & SelectedEvidenceFolder" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function CleanName(ByVal s As String) As String" & vbCrLf
    s = s & "    s = Replace(s, ""/"", ""-"")" & vbCrLf
    s = s & "    s = Replace(s, ""\"", ""-"")" & vbCrLf
    s = s & "    s = Replace(s, "":"", ""-"")" & vbCrLf
    s = s & "    s = Replace(s, ""*"", """")" & vbCrLf
    s = s & "    s = Replace(s, ""?"", """")" & vbCrLf
    s = s & "    s = Replace(s, """""", """")" & vbCrLf
    s = s & "    s = Replace(s, ""<"", """")" & vbCrLf
    s = s & "    s = Replace(s, "">"", """")" & vbCrLf
    s = s & "    s = Replace(s, ""|"", """")" & vbCrLf
    s = s & "    CleanName = s" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddL(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal b As Boolean = False, Optional ByVal s As Single = 8) As MSForms.Label" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.Label.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l" & vbCrLf
    s = s & "    o.Top = t" & vbCrLf
    s = s & "    o.Width = w" & vbCrLf
    s = s & "    o.Height = h" & vbCrLf
    s = s & "    o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = s" & vbCrLf
    s = s & "    o.Font.Bold = b" & vbCrLf
    s = s & "    o.BackStyle = 0" & vbCrLf
    s = s & "    If b Then o.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    Set AddL = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "Private Function AddT(ByVal p As Object, ByVal n As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal lk As Boolean = False) As MSForms.TextBox" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.TextBox.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l" & vbCrLf
    s = s & "    o.Top = t" & vbCrLf
    s = s & "    o.Width = w" & vbCrLf
    s = s & "    o.Height = h" & vbCrLf
    s = s & "    o.Font.Size = 8" & vbCrLf
    s = s & "    If lk Then o.Locked = True" & vbCrLf
    s = s & "    Set AddT = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "Private Function AddC(ByVal p As Object, ByVal n As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As MSForms.ComboBox" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.ComboBox.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l" & vbCrLf
    s = s & "    o.Top = t" & vbCrLf
    s = s & "    o.Width = w" & vbCrLf
    s = s & "    o.Height = h" & vbCrLf
    s = s & "    o.Font.Size = 8" & vbCrLf
    s = s & "    Set AddC = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "Private Function AddB(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, ByVal bc As Long) As MSForms.CommandButton" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.CommandButton.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l" & vbCrLf
    s = s & "    o.Top = t" & vbCrLf
    s = s & "    o.Width = w" & vbCrLf
    s = s & "    o.Height = h" & vbCrLf
    s = s & "    o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = 8" & vbCrLf
    s = s & "    o.Font.Bold = True" & vbCrLf
    s = s & "    o.BackColor = bc" & vbCrLf
    s = s & "    o.ForeColor = vbWhite" & vbCrLf
    s = s & "    Set AddB = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    Code_frmUploadEvidence = s
End Function

