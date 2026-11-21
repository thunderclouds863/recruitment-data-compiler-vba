Attribute VB_Name = "modinstallerdbfptk"
Option Explicit

Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_MSForm As Long = 3
Private gStep As String

Public Sub InstallFPTKInputForm()

    If SchemaModeEnabled Then

        MsgBox _
            "Installer legacy diblokir karena Schema Adjustment aktif." & _
            vbCrLf & _
            "Gunakan UPDATE ADJUSTMENT untuk mengubah header/form.", _
            vbExclamation, _
            "Schema Adjustment"

        Exit Sub

    End If

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
    RemoveIfExists "modFPTKInput"
    RemoveIfExists "frmInputFPTK"

    StepNow "Creating module"
    CreateModuleWithCode "modFPTKInput", Code_modFPTKInput()

    StepNow "Creating form"
    CreateBlankUserForm "frmInputFPTK", Code_frmInputFPTK()

    MsgBox "Install FPTK Input Form berhasil!" & vbCrLf & vbCrLf & _
           "Gunakan macro: OpenFPTKInputForm", vbInformation

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


' ==================== MODULE: modFPTKInput ====================

Private Function Code_modFPTKInput() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub OpenFPTKInputForm()" & vbCrLf
    s = s & "    frmInputFPTK.Show vbModal" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function LookupMappingUser(ByVal posisi As String) As Variant" & vbCrLf
    s = s & "    Dim result(2) As String" & vbCrLf
    s = s & "    result(0) = """"" & vbCrLf
    s = s & "    result(1) = """"" & vbCrLf
    s = s & "    result(2) = """"" & vbCrLf
    s = s & "    If Len(Trim(posisi)) = 0 Then LookupMappingUser = result: Exit Function" & vbCrLf
    s = s & "    Dim wsMap As Worksheet" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set wsMap = ThisWorkbook.Worksheets(""Mapping User"")" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    If wsMap Is Nothing Then LookupMappingUser = result: Exit Function" & vbCrLf
    s = s & "    Dim lastRow As Long" & vbCrLf
    s = s & "    lastRow = wsMap.Cells(wsMap.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    For r = lastRow To 2 Step -1" & vbCrLf
    s = s & "        If InStr(1, LCase(CStr(wsMap.Cells(r, 1).Value)), LCase(posisi), vbTextCompare) > 0 Then" & vbCrLf
    s = s & "            result(0) = CStr(wsMap.Cells(r, 2).Value)" & vbCrLf
    s = s & "            result(1) = CStr(wsMap.Cells(r, 3).Value)" & vbCrLf
    s = s & "            result(2) = CStr(wsMap.Cells(r, 4).Value)" & vbCrLf
    s = s & "            Exit For" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    LookupMappingUser = result" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub SaveMappingUser(ByVal posisi As String, ByVal user As String, ByVal indirectUser As String, ByVal directorate As String)" & vbCrLf
    s = s & "    If Len(Trim(posisi)) = 0 Then Exit Sub" & vbCrLf
    s = s & "    Dim wsMap As Worksheet" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set wsMap = ThisWorkbook.Worksheets(""Mapping User"")" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    If wsMap Is Nothing Then" & vbCrLf
    s = s & "        Set wsMap = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))" & vbCrLf
    s = s & "        wsMap.Name = ""Mapping User""" & vbCrLf
    s = s & "        wsMap.Cells(1, 1).Value = ""POSITION""" & vbCrLf
    s = s & "        wsMap.Cells(1, 2).Value = ""USER (MANAGER)""" & vbCrLf
    s = s & "        wsMap.Cells(1, 3).Value = ""INDIRECT USER""" & vbCrLf
    s = s & "        wsMap.Cells(1, 4).Value = ""DIRECTORATE""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim lastRow As Long" & vbCrLf
    s = s & "    lastRow = wsMap.Cells(wsMap.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim foundRow As Long" & vbCrLf
    s = s & "    foundRow = 0" & vbCrLf
    s = s & "    For r = lastRow To 2 Step -1" & vbCrLf
    s = s & "        If InStr(1, LCase(CStr(wsMap.Cells(r, 1).Value)), LCase(posisi), vbTextCompare) > 0 Then" & vbCrLf
    s = s & "            foundRow = r" & vbCrLf
    s = s & "            Exit For" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    If foundRow > 0 Then" & vbCrLf
    s = s & "        wsMap.Cells(foundRow, 2).Value = user" & vbCrLf
    s = s & "        wsMap.Cells(foundRow, 3).Value = indirectUser" & vbCrLf
    s = s & "        wsMap.Cells(foundRow, 4).Value = directorate" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Dim newRow As Long" & vbCrLf
    s = s & "        newRow = lastRow + 1" & vbCrLf
    s = s & "        wsMap.Cells(newRow, 1).Value = posisi" & vbCrLf
    s = s & "        wsMap.Cells(newRow, 2).Value = user" & vbCrLf
    s = s & "        wsMap.Cells(newRow, 3).Value = indirectUser" & vbCrLf
    s = s & "        wsMap.Cells(newRow, 4).Value = directorate" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    FormatAsTable wsMap, ""tblMappingUser""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    Code_modFPTKInput = s
End Function


' ==================== FORM: frmInputFPTK ====================

Private Function Code_frmInputFPTK() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mBuilt As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private txtEmailBody As MSForms.TextBox" & vbCrLf
    s = s & "Private lblProgress As MSForms.Label" & vbCrLf
    s = s & "Private lstSearchResults As MSForms.ListBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private txtFPTKNumber As MSForms.TextBox" & vbCrLf
    s = s & "Private txtPosisi As MSForms.TextBox" & vbCrLf
    s = s & "Private txtBusinessUnit As MSForms.TextBox" & vbCrLf
    s = s & "Private txtDivisi As MSForms.TextBox" & vbCrLf
    s = s & "Private txtDepartment As MSForms.TextBox" & vbCrLf
    s = s & "Private txtLevel As MSForms.TextBox" & vbCrLf
    s = s & "Private txtLevelNumber As MSForms.TextBox" & vbCrLf
    s = s & "Private txtAlasan As MSForms.TextBox" & vbCrLf
    s = s & "Private txtCategory As MSForms.TextBox" & vbCrLf
    s = s & "Private txtFilter As MSForms.TextBox" & vbCrLf
    s = s & "Private txtPIC As MSForms.TextBox" & vbCrLf
    s = s & "Private txtPICCode As MSForms.TextBox" & vbCrLf
    s = s & "Private txtPositionCode As MSForms.TextBox" & vbCrLf
    s = s & "Private txtBUCode As MSForms.TextBox" & vbCrLf
    s = s & "Private txtVacancy As MSForms.TextBox" & vbCrLf
    s = s & "Private txtKodeUnik As MSForms.TextBox" & vbCrLf
    s = s & "Private txtFPTKDate As MSForms.TextBox" & vbCrLf
    s = s & "Private txtDirectorate As MSForms.TextBox" & vbCrLf
    s = s & "Private txtUser As MSForms.TextBox" & vbCrLf
    s = s & "Private txtIndirectUser As MSForms.TextBox" & vbCrLf
    s = s & "Private txtLokasiKerja As MSForms.TextBox" & vbCrLf
    s = s & "Private txtHRLocation As MSForms.TextBox" & vbCrLf
    s = s & "Private txtEmploymentStatus As MSForms.TextBox" & vbCrLf
    s = s & "Private cboStatus As MSForms.ComboBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents btnProses As MSForms.CommandButton" & vbCrLf
    s = s & "Private txtSearchFPTK As MSForms.TextBox" & vbCrLf
    s = s & "Private WithEvents btnCariFPTK As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnClearForm As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnSimpan As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnBatal As MSForms.CommandButton" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildUI()" & vbCrLf
    s = s & "    If mBuilt Then Exit Sub" & vbCrLf
    s = s & "    mBuilt = True" & vbCrLf
    s = s & "    Me.Caption = ""INPUT FPTK BARU""" & vbCrLf
    s = s & "    Me.Width = 730" & vbCrLf
    s = s & "    Me.Height = 620" & vbCrLf
    s = s & "    Me.BackColor = RGB(245, 245, 245)" & vbCrLf
    s = s & "    Me.ScrollBars = fmScrollBarsVertical" & vbCrLf
    s = s & "    Me.ScrollHeight = 950" & vbCrLf
    s = s & "    Dim y As Single" & vbCrLf
    s = s & "    y = 8" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' === SECTION: EMAIL BODY ===" & vbCrLf
    s = s & "    AddL Me, ""lblSec1"", ""PASTE EMAIL BODY DI SINI"", 10, y, 700, 18, True, 9" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = Me.Controls(""lblSec1"")" & vbCrLf
    s = s & "    o.BackColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    o.ForeColor = vbWhite" & vbCrLf
    s = s & "    o.BackStyle = 1" & vbCrLf
    s = s & "    y = y + 20" & vbCrLf
    s = s & "    Set txtEmailBody = Me.Controls.Add(""Forms.TextBox.1"", ""txtEmailBody"", True)" & vbCrLf
    s = s & "    txtEmailBody.Left = 10" & vbCrLf
    s = s & "    txtEmailBody.Top = y" & vbCrLf
    s = s & "    txtEmailBody.Width = 700" & vbCrLf
    s = s & "    txtEmailBody.Height = 100" & vbCrLf
    s = s & "    txtEmailBody.MultiLine = True" & vbCrLf
    s = s & "    txtEmailBody.ScrollBars = fmScrollBarsVertical" & vbCrLf
    s = s & "    txtEmailBody.Font.Size = 8" & vbCrLf
    s = s & "    y = y + 105" & vbCrLf
    s = s & "    Set btnProses = AddB(Me, ""btnProses"", ""PROSES EMAIL"", 250, y, 180, 32, RGB(39, 174, 96))" & vbCrLf
    s = s & "    Set lblProgress = AddL(Me, ""lblProgress"", ""Belum diproses"", 440, y + 6, 250, 18, False, 8)" & vbCrLf
    s = s & "    lblProgress.ForeColor = RGB(149, 165, 166)" & vbCrLf
    s = s & "    y = y + 40" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' === SECTION: CARI FPTK EXISTING ===" & vbCrLf
    s = s & "    AddL Me, ""lblSec0"", ""EDIT FPTK YANG SUDAH ADA"", 10, y, 700, 18, True, 9" & vbCrLf
    s = s & "    Set o = Me.Controls(""lblSec0"")" & vbCrLf
    s = s & "    o.BackColor = RGB(243, 156, 18)" & vbCrLf
    s = s & "    o.ForeColor = vbWhite" & vbCrLf
    s = s & "    o.BackStyle = 1" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "    AddL Me, ""lblSearchFPTK"", ""Cari Kode Unik / Posisi:"", 10, y, 150, 14, True, 7" & vbCrLf
    s = s & "    Set txtSearchFPTK = AddT(Me, ""txtSearchFPTK"", 165, y - 1, 280, 18)" & vbCrLf
    s = s & "    Set btnCariFPTK = AddB(Me, ""btnCariFPTK"", ""CARI"", 455, y - 2, 60, 22, RGB(243, 156, 18))" & vbCrLf
    s = s & "    Set btnClearForm = AddB(Me, ""btnClearForm"", ""CLEAR"", 525, y - 2, 60, 22, RGB(149, 165, 166))" & vbCrLf
    s = s & "    y = y + 2" & vbCrLf
    s = s & "    Set lstSearchResults = Me.Controls.Add(""Forms.ListBox.1"", ""lstSearchResults"", True)" & vbCrLf
    s = s & "    lstSearchResults.Left = 10" & vbCrLf
    s = s & "    lstSearchResults.Top = y" & vbCrLf
    s = s & "    lstSearchResults.Width = 700" & vbCrLf
    s = s & "    lstSearchResults.Height = 70" & vbCrLf
    s = s & "    lstSearchResults.Font.Size = 8" & vbCrLf
    s = s & "    lstSearchResults.ColumnCount = 3" & vbCrLf
    s = s & "    lstSearchResults.ColumnWidths = ""150;350;100""" & vbCrLf
    s = s & "    y = y + 80" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' === SECTION: DATA FPTK ===" & vbCrLf
    s = s & "    AddL Me, ""lblSec2"", ""DATA FPTK (Hasil Proses / Cari / Manual)"", 10, y, 700, 18, True, 9" & vbCrLf
    s = s & "    Set o = Me.Controls(""lblSec2"")" & vbCrLf
    s = s & "    o.BackColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    o.ForeColor = vbWhite" & vbCrLf
    s = s & "    o.BackStyle = 1" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 1: FPTK Number | Kode Unik" & vbCrLf
    s = s & "    AddL Me, ""lblFPTKNumber"", ""FPTK Number"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtFPTKNumber = AddT(Me, ""txtFPTKNumber"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    AddL Me, ""lblKodeUnik"", ""Kode Unik"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtKodeUnik = AddT(Me, ""txtKodeUnik"", 445, y - 1, 260, 18, True)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 2: Posisi (full width)" & vbCrLf
    s = s & "    AddL Me, ""lblPosisi"", ""Posisi"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtPosisi = AddT(Me, ""txtPosisi"", 115, y - 1, 590, 18)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 3: FPTK Date | Status" & vbCrLf
    s = s & "    AddL Me, ""lblFPTKDate"", ""FPTK Date"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtFPTKDate = AddT(Me, ""txtFPTKDate"", 115, y - 1, 200, 18, True)" & vbCrLf
    s = s & "    AddL Me, ""lblStatus"", ""Status"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set cboStatus = AddC(Me, ""cboStatus"", 445, y - 1, 260, 18)" & vbCrLf
    s = s & "    cboStatus.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    cboStatus.AddItem ""OP""" & vbCrLf
    s = s & "    cboStatus.AddItem ""Closed""" & vbCrLf
    s = s & "    cboStatus.AddItem ""Cancel""" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 4: Business Unit | Direktorat" & vbCrLf
    s = s & "    AddL Me, ""lblBusinessUnit"", ""Business Unit"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtBusinessUnit = AddT(Me, ""txtBusinessUnit"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    AddL Me, ""lblDirectorate"", ""Direktorat"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtDirectorate = AddT(Me, ""txtDirectorate"", 445, y - 1, 260, 18)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 5: Divisi | Department" & vbCrLf
    s = s & "    AddL Me, ""lblDivisi"", ""Divisi"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtDivisi = AddT(Me, ""txtDivisi"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    AddL Me, ""lblDepartment"", ""Department"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtDepartment = AddT(Me, ""txtDepartment"", 445, y - 1, 260, 18)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 6: Level | Level Number" & vbCrLf
    s = s & "    AddL Me, ""lblLevel"", ""Level"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtLevel = AddT(Me, ""txtLevel"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    AddL Me, ""lblLevelNumber"", ""Level Number"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtLevelNumber = AddT(Me, ""txtLevelNumber"", 445, y - 1, 260, 18)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 7: Alasan | Category" & vbCrLf
    s = s & "    AddL Me, ""lblAlasan"", ""Alasan FPTK"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtAlasan = AddT(Me, ""txtAlasan"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    AddL Me, ""lblCategory"", ""Category FPTK"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtCategory = AddT(Me, ""txtCategory"", 445, y - 1, 260, 18)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 8: Filter | PIC Recruiter" & vbCrLf
    s = s & "    AddL Me, ""lblFilter"", ""Filter Kategori"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtFilter = AddT(Me, ""txtFilter"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    AddL Me, ""lblPIC"", ""PIC Recruiter"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtPIC = AddT(Me, ""txtPIC"", 445, y - 1, 260, 18)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 9: Kode PIC | Kode Posisi" & vbCrLf
    s = s & "    AddL Me, ""lblPICCode"", ""Kode PIC"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtPICCode = AddT(Me, ""txtPICCode"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    AddL Me, ""lblPositionCode"", ""Kode Posisi"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtPositionCode = AddT(Me, ""txtPositionCode"", 445, y - 1, 260, 18)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 10: Kode BU | Vacancy" & vbCrLf
    s = s & "    AddL Me, ""lblBUCode"", ""Kode BU"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtBUCode = AddT(Me, ""txtBUCode"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    AddL Me, ""lblVacancy"", ""Vacancy"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtVacancy = AddT(Me, ""txtVacancy"", 445, y - 1, 260, 18)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 11: User Manager | Indirect User" & vbCrLf
    s = s & "    AddL Me, ""lblUser"", ""User (Manager)"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtUser = AddT(Me, ""txtUser"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    AddL Me, ""lblIndirectUser"", ""Indirect User"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtIndirectUser = AddT(Me, ""txtIndirectUser"", 445, y - 1, 260, 18)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 12: Lokasi Kerja | HR Location" & vbCrLf
    s = s & "    AddL Me, ""lblLokasiKerja"", ""Lokasi Kerja"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtLokasiKerja = AddT(Me, ""txtLokasiKerja"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    AddL Me, ""lblHRLocation"", ""Lokasi HR"", 340, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtHRLocation = AddT(Me, ""txtHRLocation"", 445, y - 1, 260, 18)" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Row 13: Employment Status" & vbCrLf
    s = s & "    AddL Me, ""lblEmploymentStatus"", ""Status Karyawan"", 10, y, 100, 14, True, 7" & vbCrLf
    s = s & "    Set txtEmploymentStatus = AddT(Me, ""txtEmploymentStatus"", 115, y - 1, 200, 18)" & vbCrLf
    s = s & "    y = y + 32" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' === BUTTONS ===" & vbCrLf
    s = s & "    Set btnSimpan = AddB(Me, ""btnSimpan"", ""SIMPAN KE FPTK"", 200, y, 150, 34, RGB(39, 174, 96))" & vbCrLf
    s = s & "    Set btnBatal = AddB(Me, ""btnBatal"", ""TUTUP"", 380, y, 100, 34, RGB(231, 76, 60))" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnProses_Click()" & vbCrLf
    s = s & "    If Len(Trim(txtEmailBody.Value)) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Paste email body dulu!"", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    On Error GoTo ErrHandler" & vbCrLf
    s = s & "    lblProgress.Caption = ""Memproses...""" & vbCrLf
    s = s & "    lblProgress.ForeColor = RGB(243, 156, 18)" & vbCrLf
    s = s & "    DoEvents" & vbCrLf
    s = s & "    Init" & vbCrLf
    s = s & "    ClearGlobal" & vbCrLf
    s = s & "    gToday = Date" & vbCrLf
    s = s & "    Range(""EmailBody"").Value = txtEmailBody.Value" & vbCrLf
    s = s & "    ParseEmail" & vbCrLf
    s = s & "    If Not LookupPosition Then" & vbCrLf
    s = s & "        InsertNewPosition" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    LookupPIC" & vbCrLf
    s = s & "    LookupBU" & vbCrLf
    s = s & "    GenerateAll" & vbCrLf
    s = s & "    LoadDataDariGlobal" & vbCrLf
    s = s & "    lblProgress.Caption = ""Proses selesai! Silakan cek data di bawah.""" & vbCrLf
    s = s & "    lblProgress.ForeColor = RGB(39, 174, 96)" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrHandler:" & vbCrLf
    s = s & "    lblProgress.Caption = ""Error: "" & Err.Description" & vbCrLf
    s = s & "    lblProgress.ForeColor = RGB(231, 76, 60)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadDataDariGlobal()" & vbCrLf
    s = s & "    txtFPTKNumber.Value = gFPTKNumber" & vbCrLf
    s = s & "    txtPosisi.Value = gPosition" & vbCrLf
    s = s & "    txtBusinessUnit.Value = gBusinessUnit" & vbCrLf
    s = s & "    txtDivisi.Value = gDivisionEmail" & vbCrLf
    s = s & "    txtDepartment.Value = gDepartmentEmail" & vbCrLf
    s = s & "    txtLevel.Value = gLevelEmail" & vbCrLf
    s = s & "    txtLevelNumber.Value = CStr(gLevelNumber)" & vbCrLf
    s = s & "    txtAlasan.Value = gAlasanFPTK" & vbCrLf
    s = s & "    txtCategory.Value = gCategoryFPTK" & vbCrLf
    s = s & "    txtFilter.Value = gFilterKategori" & vbCrLf
    s = s & "    txtPIC.Value = gPIC" & vbCrLf
    s = s & "    txtPICCode.Value = gPICCode" & vbCrLf
    s = s & "    txtPositionCode.Value = gPositionCode" & vbCrLf
    s = s & "    txtBUCode.Value = gBUCode" & vbCrLf
    s = s & "    txtVacancy.Value = CStr(gVacancy)" & vbCrLf
    s = s & "    txtDirectorate.Value = gDirectorate" & vbCrLf
    s = s & "    txtKodeUnik.Value = gUniqueCode" & vbCrLf
    s = s & "    txtFPTKDate.Value = Format(gFPTKDate, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "    cboStatus.Value = ""OP""" & vbCrLf
    s = s & "    Dim mapping As Variant" & vbCrLf
    s = s & "    mapping = LookupMappingUser(gPosition)" & vbCrLf
    s = s & "    txtUser.Value = mapping(0)" & vbCrLf
    s = s & "    txtIndirectUser.Value = mapping(1)" & vbCrLf
    s = s & "    If Len(Trim(gDirectorate)) = 0 And Len(mapping(2)) > 0 Then" & vbCrLf
    s = s & "        txtDirectorate.Value = mapping(2)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    txtLokasiKerja.Value = gLocation" & vbCrLf
    s = s & "    txtHRLocation.Value = gHRLocation" & vbCrLf
    s = s & "    txtEmploymentStatus.Value = gEmploymentStatus" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSimpan_Click()" & vbCrLf
    s = s & "    If Trim(txtKodeUnik.Value) = """" Then" & vbCrLf
    s = s & "        MsgBox ""Kode Unik tidak boleh kosong!"", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    On Error GoTo ErrSave" & vbCrLf
    s = s & "    Init" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = wsFPTK" & vbCrLf
    s = s & "    Dim lastRow As Long" & vbCrLf
    s = s & "    lastRow = ws.Cells(ws.Rows.Count, colKodeUnik).End(xlUp).Row" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim targetRow As Long" & vbCrLf
    s = s & "    targetRow = 0" & vbCrLf
    s = s & "    For r = 2 To lastRow" & vbCrLf
    s = s & "        If Trim(CStr(ws.Cells(r, colKodeUnik).Value)) = Trim(txtKodeUnik.Value) Then" & vbCrLf
    s = s & "            targetRow = r" & vbCrLf
    s = s & "            Exit For" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    If targetRow = 0 Then" & vbCrLf
    s = s & "        targetRow = lastRow + 1" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    With ws" & vbCrLf
    s = s & "        .Cells(targetRow, colFPTKNumber).Value = txtFPTKNumber.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colKodeUnik).Value = txtKodeUnik.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colPosisi).Value = txtPosisi.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colFPTKDate).Value = txtFPTKDate.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colBusinessUnit).Value = txtBusinessUnit.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colDirektorat).Value = txtDirectorate.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colDivisi).Value = txtDivisi.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colDepartment).Value = txtDepartment.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colLevel).Value = txtLevel.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colLevelNumber).Value = txtLevelNumber.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colCategory).Value = txtAlasan.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colCategoryFPTK).Value = txtCategory.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colFilterKategori).Value = txtFilter.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colPICRecruiter).Value = txtPIC.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colKodePIC).Value = txtPICCode.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colKodeAngka).Value = txtPositionCode.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colVacancy).Value = txtVacancy.Value" & vbCrLf
    s = s & "        .Cells(targetRow, colStatus).Value = cboStatus.Value" & vbCrLf
    s = s & "        If colUser > 0 Then .Cells(targetRow, colUser).Value = txtUser.Value" & vbCrLf
    s = s & "        If colIndirectUser > 0 Then .Cells(targetRow, colIndirectUser).Value = txtIndirectUser.Value" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "    SaveMappingUser txtPosisi.Value, txtUser.Value, txtIndirectUser.Value, txtDirectorate.Value" & vbCrLf
    s = s & "    FormatAsTable ws, ""tblFPTK""" & vbCrLf
    s = s & "    MsgBox ""FPTK berhasil disimpan! Row: "" & targetRow, vbInformation" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrSave:" & vbCrLf
    s = s & "    MsgBox ""Error saat simpan: "" & Err.Description, vbCritical" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnBatal_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnCariFPTK_Click()" & vbCrLf
    s = s & "    Dim keyword As String" & vbCrLf
    s = s & "    keyword = Trim(txtSearchFPTK.Value)" & vbCrLf
    s = s & "    If Len(keyword) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Masukkan Kode Unik atau Posisi!"", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Init" & vbCrLf
    s = s & "    lstSearchResults.Clear" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = wsFPTK" & vbCrLf
    s = s & "    Dim lastRow As Long" & vbCrLf
    s = s & "    lastRow = ws.Cells(ws.Rows.Count, colKodeUnik).End(xlUp).Row" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    Dim idx As Long" & vbCrLf
    s = s & "    idx = 0" & vbCrLf
    s = s & "    For r = 2 To lastRow" & vbCrLf
    s = s & "        Dim kodeUnik As String" & vbCrLf
    s = s & "        kodeUnik = Trim(CStr(ws.Cells(r, colKodeUnik).Value))" & vbCrLf
    s = s & "        Dim fptkNum As String" & vbCrLf
    s = s & "        fptkNum = Trim(CStr(ws.Cells(r, colFPTKNumber).Value))" & vbCrLf
    s = s & "        Dim posisi As String" & vbCrLf
    s = s & "        posisi = Trim(CStr(ws.Cells(r, colPosisi).Value))" & vbCrLf
    s = s & "        If InStr(1, LCase(kodeUnik), LCase(keyword)) > 0 Or _" & vbCrLf
    s = s & "           InStr(1, LCase(fptkNum), LCase(keyword)) > 0 Or _" & vbCrLf
    s = s & "           InStr(1, LCase(posisi), LCase(keyword)) > 0 Then" & vbCrLf
    s = s & "            lstSearchResults.AddItem """"" & vbCrLf
    s = s & "            lstSearchResults.List(idx, 0) = kodeUnik" & vbCrLf
    s = s & "            lstSearchResults.List(idx, 1) = posisi" & vbCrLf
    s = s & "            lstSearchResults.List(idx, 2) = r" & vbCrLf
    s = s & "            idx = idx + 1" & vbCrLf
    s = s & "            If idx >= 100 Then Exit For" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    If idx = 0 Then" & vbCrLf
    s = s & "        lblProgress.Caption = ""Data tidak ditemukan!""" & vbCrLf
    s = s & "        lblProgress.ForeColor = RGB(231, 76, 60)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblProgress.Caption = ""Ditemukan "" & idx & "" data. Double-click untuk pilih.""" & vbCrLf
    s = s & "        lblProgress.ForeColor = RGB(39, 174, 96)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstSearchResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    If lstSearchResults.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "    Dim rowNum As Long" & vbCrLf
    s = s & "    rowNum = CLng(lstSearchResults.List(lstSearchResults.ListIndex, 2))" & vbCrLf
    s = s & "    LoadFPTKRow rowNum" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadFPTKRow(ByVal r As Long)" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = wsFPTK" & vbCrLf
    s = s & "    txtFPTKNumber.Value = CStr(ws.Cells(r, colFPTKNumber).Value)" & vbCrLf
    s = s & "    txtKodeUnik.Value = Trim(CStr(ws.Cells(r, colKodeUnik).Value))" & vbCrLf
    s = s & "    txtPosisi.Value = CStr(ws.Cells(r, colPosisi).Value)" & vbCrLf
    s = s & "    txtBusinessUnit.Value = CStr(ws.Cells(r, colBusinessUnit).Value)" & vbCrLf
    s = s & "    txtDirectorate.Value = CStr(ws.Cells(r, colDirektorat).Value)" & vbCrLf
    s = s & "    txtDivisi.Value = CStr(ws.Cells(r, colDivisi).Value)" & vbCrLf
    s = s & "    txtDepartment.Value = CStr(ws.Cells(r, colDepartment).Value)" & vbCrLf
    s = s & "    txtLevel.Value = CStr(ws.Cells(r, colLevel).Value)" & vbCrLf
    s = s & "    txtLevelNumber.Value = CStr(ws.Cells(r, colLevelNumber).Value)" & vbCrLf
    s = s & "    txtAlasan.Value = CStr(ws.Cells(r, colCategory).Value)" & vbCrLf
    s = s & "    txtCategory.Value = CStr(ws.Cells(r, colCategoryFPTK).Value)" & vbCrLf
    s = s & "    txtFilter.Value = CStr(ws.Cells(r, colFilterKategori).Value)" & vbCrLf
    s = s & "    txtPIC.Value = CStr(ws.Cells(r, colPICRecruiter).Value)" & vbCrLf
    s = s & "    txtPICCode.Value = CStr(ws.Cells(r, colKodePIC).Value)" & vbCrLf
    s = s & "    txtPositionCode.Value = CStr(ws.Cells(r, colKodeAngka).Value)" & vbCrLf
    s = s & "    txtBUCode.Value = """"" & vbCrLf
    s = s & "    txtVacancy.Value = CStr(ws.Cells(r, colVacancy).Value)" & vbCrLf
    s = s & "    txtFPTKDate.Value = CStr(ws.Cells(r, colFPTKDate).Value)" & vbCrLf
    s = s & "    cboStatus.Value = CStr(ws.Cells(r, colStatus).Value)" & vbCrLf
    s = s & "    Dim mapping As Variant" & vbCrLf
    s = s & "    mapping = LookupMappingUser(txtPosisi.Value)" & vbCrLf
    s = s & "    If Len(Trim(txtUser.Value)) = 0 Then txtUser.Value = mapping(0)" & vbCrLf
    s = s & "    If Len(Trim(txtIndirectUser.Value)) = 0 Then txtIndirectUser.Value = mapping(1)" & vbCrLf
    s = s & "    If Len(Trim(txtDirectorate.Value)) = 0 Then txtDirectorate.Value = mapping(2)" & vbCrLf
    s = s & "    lblProgress.Caption = ""Data Row "" & r & "" - Edit lalu klik SIMPAN""" & vbCrLf
    s = s & "    lblProgress.ForeColor = RGB(39, 174, 96)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnClearForm_Click()" & vbCrLf
    s = s & "    txtEmailBody.Value = """"" & vbCrLf
    s = s & "    txtSearchFPTK.Value = """"" & vbCrLf
    s = s & "    lstSearchResults.Clear" & vbCrLf
    s = s & "    txtFPTKNumber.Value = """"" & vbCrLf
    s = s & "    txtKodeUnik.Value = """"" & vbCrLf
    s = s & "    txtPosisi.Value = """"" & vbCrLf
    s = s & "    txtBusinessUnit.Value = """"" & vbCrLf
    s = s & "    txtDirectorate.Value = """"" & vbCrLf
    s = s & "    txtDivisi.Value = """"" & vbCrLf
    s = s & "    txtDepartment.Value = """"" & vbCrLf
    s = s & "    txtLevel.Value = """"" & vbCrLf
    s = s & "    txtLevelNumber.Value = """"" & vbCrLf
    s = s & "    txtAlasan.Value = """"" & vbCrLf
    s = s & "    txtCategory.Value = """"" & vbCrLf
    s = s & "    txtFilter.Value = """"" & vbCrLf
    s = s & "    txtPIC.Value = """"" & vbCrLf
    s = s & "    txtPICCode.Value = """"" & vbCrLf
    s = s & "    txtPositionCode.Value = """"" & vbCrLf
    s = s & "    txtBUCode.Value = """"" & vbCrLf
    s = s & "    txtVacancy.Value = """"" & vbCrLf
    s = s & "    txtFPTKDate.Value = """"" & vbCrLf
    s = s & "    txtUser.Value = """"" & vbCrLf
    s = s & "    txtIndirectUser.Value = """"" & vbCrLf
    s = s & "    txtLokasiKerja.Value = """"" & vbCrLf
    s = s & "    txtHRLocation.Value = """"" & vbCrLf
    s = s & "    txtEmploymentStatus.Value = """"" & vbCrLf
    s = s & "    cboStatus.Value = ""OP""" & vbCrLf
    s = s & "    lblProgress.Caption = ""Form direset""" & vbCrLf
    s = s & "    lblProgress.ForeColor = RGB(149, 165, 166)" & vbCrLf
    s = s & "End Sub" & vbCrLf
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
    Code_frmInputFPTK = s
End Function



