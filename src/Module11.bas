Attribute VB_Name = "Module11"
Option Explicit

Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_MSForm As Long = 3
Private gStep As String
Public Sub InstallSourcingCV_RUNTIME_ONLY()
    On Error GoTo ErrHandler
    
    If SchemaModeEnabled Then

        MsgBox _
            "Installer legacy dinonaktifkan karena Schema_Config sudah aktif." & _
            vbCrLf & _
            "Module dan form schema-driven sudah dipasang manual.", _
            vbExclamation, _
            "Schema Mode Active"

        Exit Sub

    End If
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    StepNow "Checking VBA project access"
    If Not CanAccessVBProject() Then
        MsgBox "VBA project access masih blocked." & vbCrLf & vbCrLf & "Aktifkan: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model.", vbCritical
        GoTo SafeExit
    End If

    StepNow "Cleaning old components"
    RemoveIfExists "frmSourcing"
    RemoveIfExists "frmSearchFPTK"
    RemoveIfExists "frmPreviewCV"
    RemoveIfExists "frmSourcingList"
    RemoveIfExists "frmPasteCVText"
    RemoveIfExists "frmEditSourcing"
    RemoveIfExists "modSourcingSupport"
    RemoveIfExists "modMainSourcing"
    RemoveIfExists "frmSourcingDetail"

    StepNow "Creating support modules"
    CreateModuleWithCode "modSourcingSupport", Code_ModSourcingSupport()
    CreateModuleWithCode "modMainSourcing", Code_ModMainSourcing()

    StepNow "Creating UserForms"
    CreateBlankUserForm "frmSearchFPTK", Code_frmSearchFPTK()
    CreateBlankUserForm "frmSourcing", Code_frmSourcing()
    CreateBlankUserForm "frmSourcingList", Code_frmSourcingList()
    CreateBlankUserForm "frmPasteCVText", Code_frmPasteCVText()
    CreateBlankUserForm "frmEditSourcing", Code_frmEditSourcing()
    CreateBlankUserForm "frmSourcingDetail", Code_frmSourcingDetail()
    ' modMonitoring sudah diimport sebagai module terpisah.
    ' Jangan dibuat ulang oleh installer legacy.

    StepNow "Creating database sheets"
    Application.Run "InitSourcingCV"

    MsgBox "Install berhasil!" & vbCrLf & vbCrLf & "Buka form: OpenSourcingForm" & vbCrLf & "Update: DB Sourcing sekarang bisa Edit via Form.", vbInformation

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
Private Function Code_frmSourcingDetail() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "Private mBuilt As Boolean" & vbCrLf
    s = s & "Private mRow As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private txtNama As MSForms.TextBox" & vbCrLf
    s = s & "Private txtPosisi As MSForms.TextBox" & vbCrLf
    s = s & "Private txtKode As MSForms.TextBox" & vbCrLf
    s = s & "Private txtSumber As MSForms.TextBox" & vbCrLf
    s = s & "Private txtModel As MSForms.TextBox" & vbCrLf
    s = s & "Private txtRekruter As MSForms.TextBox" & vbCrLf
    s = s & "Private txtUniv As MSForms.TextBox" & vbCrLf
    s = s & "Private txtJurusan As MSForms.TextBox" & vbCrLf
    s = s & "Private txtJenjang As MSForms.TextBox" & vbCrLf
    s = s & "Private txtIPK As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTahunLulus As MSForms.TextBox" & vbCrLf
    s = s & "Private txtEnglish As MSForms.TextBox" & vbCrLf
    s = s & "Private txtUnivTier As MSForms.TextBox" & vbCrLf
    s = s & "Private txtIPKTier As MSForms.TextBox" & vbCrLf
    s = s & "Private txtHP As MSForms.TextBox" & vbCrLf
    s = s & "Private txtEmail As MSForms.TextBox" & vbCrLf
    s = s & "Private txtDomisili As MSForms.TextBox" & vbCrLf
    s = s & "Private lblStatusBesar As MSForms.Label" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private lblSourcingHR As MSForms.Label" & vbCrLf
    s = s & "Private lblTglSourcingHR As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanSourcingHR As MSForms.Label" & vbCrLf
    s = s & "Private lblShortlist As MSForms.Label" & vbCrLf
    s = s & "Private lblTglShortlist As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanShortlist As MSForms.Label" & vbCrLf
    s = s & "Private lblPsikotes As MSForms.Label" & vbCrLf
    s = s & "Private lblTglPsikotes As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanPsikotes As MSForms.Label" & vbCrLf
    s = s & "Private lblHRInterview As MSForms.Label" & vbCrLf
    s = s & "Private lblTglHRInterview As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanHRInterview As MSForms.Label" & vbCrLf
    s = s & "Private lblTechnicalTest As MSForms.Label" & vbCrLf
    s = s & "Private lblTglTechnicalTest As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanTechnicalTest As MSForms.Label" & vbCrLf
    s = s & "Private lblMarketVisit As MSForms.Label" & vbCrLf
    s = s & "Private lblTglMarketVisit As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanMarketVisit As MSForms.Label" & vbCrLf
    s = s & "Private lblUserInterview As MSForms.Label" & vbCrLf
    s = s & "Private lblTglUserInterview As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanUserInterview As MSForms.Label" & vbCrLf
    s = s & "Private lblPanelInterview As MSForms.Label" & vbCrLf
    s = s & "Private lblTglPanelInterview As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanPanelInterview As MSForms.Label" & vbCrLf
    s = s & "Private lblReferenceCheck As MSForms.Label" & vbCrLf
    s = s & "Private lblTglReferenceCheck As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanReferenceCheck As MSForms.Label" & vbCrLf
    s = s & "Private lblMCU As MSForms.Label" & vbCrLf
    s = s & "Private lblTglMCU As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanMCU As MSForms.Label" & vbCrLf
    s = s & "Private lblOffering As MSForms.Label" & vbCrLf
    s = s & "Private lblTglOffering As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanOffering As MSForms.Label" & vbCrLf
    s = s & "Private lblDay1 As MSForms.Label" & vbCrLf
    s = s & "Private lblTglDay1 As MSForms.Label" & vbCrLf
    s = s & "Private lblAlasanDay1 As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents btnEdit As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnClose As MSForms.CommandButton" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildUI()" & vbCrLf
    s = s & "    If mBuilt Then Exit Sub" & vbCrLf
    s = s & "    mBuilt = True" & vbCrLf
    s = s & "    Me.Caption = ""DETAIL KANDIDAT""" & vbCrLf
    s = s & "    Me.Width = 720" & vbCrLf
    s = s & "    Me.Height = 600" & vbCrLf
    s = s & "    Me.BackColor = RGB(245, 245, 245)" & vbCrLf
    s = s & "    Me.ScrollBars = fmScrollBarsVertical" & vbCrLf
    s = s & "    Me.ScrollHeight = 850" & vbCrLf
    s = s & "    Dim y As Single" & vbCrLf
    s = s & "    y = 4" & vbCrLf
    s = s & "    AddSec Me, ""sec1"", ""DATA KANDIDAT"", 8, y, 695" & vbCrLf
    s = s & "    y = y + 14" & vbCrLf
    s = s & "    AddL Me, ""lblNama"", ""Nama"", 8, y, 40, 12, True, 7" & vbCrLf
    s = s & "    Set txtNama = AddT(Me, ""txtNama"", 50, y - 1, 180, 16, True)" & vbCrLf
    s = s & "    txtNama.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblKode"", ""Kode"", 240, y, 40, 12, True, 7" & vbCrLf
    s = s & "    Set txtKode = AddT(Me, ""txtKode"", 285, y - 1, 110, 16, True)" & vbCrLf
    s = s & "    txtKode.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblModel"", ""Model"", 405, y, 35, 12, True, 7" & vbCrLf
    s = s & "    Set txtModel = AddT(Me, ""txtModel"", 440, y - 1, 50, 16, True)" & vbCrLf
    s = s & "    txtModel.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblRekruter"", ""Rekr"", 500, y, 30, 12, True, 7" & vbCrLf
    s = s & "    Set txtRekruter = AddT(Me, ""txtRekruter"", 530, y - 1, 60, 16, True)" & vbCrLf
    s = s & "    txtRekruter.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblDomisili"", ""Dom"", 600, y, 30, 12, True, 7" & vbCrLf
    s = s & "    Set txtDomisili = AddT(Me, ""txtDomisili"", 630, y - 1, 75, 16, True)" & vbCrLf
    s = s & "    txtDomisili.BackColor = vbWhite" & vbCrLf
    s = s & "    y = y + 17" & vbCrLf
    s = s & "    AddL Me, ""lblPosisi"", ""Posisi"", 8, y, 40, 12, True, 7" & vbCrLf
    s = s & "    Set txtPosisi = AddT(Me, ""txtPosisi"", 50, y - 1, 280, 16, True)" & vbCrLf
    s = s & "    txtPosisi.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblSumber"", ""Sumber"", 340, y, 45, 12, True, 7" & vbCrLf
    s = s & "    Set txtSumber = AddT(Me, ""txtSumber"", 385, y - 1, 100, 16, True)" & vbCrLf
    s = s & "    txtSumber.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblEmail"", ""Email"", 495, y, 35, 12, True, 7" & vbCrLf
    s = s & "    Set txtEmail = AddT(Me, ""txtEmail"", 530, y - 1, 175, 16, True)" & vbCrLf
    s = s & "    txtEmail.BackColor = vbWhite" & vbCrLf
    s = s & "    y = y + 17" & vbCrLf
    s = s & "    AddL Me, ""lblHP"", ""HP"", 8, y, 40, 12, True, 7" & vbCrLf
    s = s & "    Set txtHP = AddT(Me, ""txtHP"", 50, y - 1, 130, 16, True)" & vbCrLf
    s = s & "    txtHP.BackColor = vbWhite" & vbCrLf
    s = s & "    y = y + 16" & vbCrLf
    s = s & "    AddSec Me, ""sec2"", ""PENDIDIKAN"", 8, y, 695" & vbCrLf
    s = s & "    y = y + 14" & vbCrLf
    s = s & "    AddL Me, ""lblUniv"", ""Universitas"", 8, y, 55, 12, True, 7" & vbCrLf
    s = s & "    Set txtUniv = AddT(Me, ""txtUniv"", 65, y - 1, 250, 16, True)" & vbCrLf
    s = s & "    txtUniv.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblJenjang"", ""Jenjang"", 325, y, 45, 12, True, 7" & vbCrLf
    s = s & "    Set txtJenjang = AddT(Me, ""txtJenjang"", 370, y - 1, 50, 16, True)" & vbCrLf
    s = s & "    txtJenjang.BackColor = vbWhite" & vbCrLf
    s = s & "    y = y + 17" & vbCrLf
    s = s & "    AddL Me, ""lblJurusan"", ""Jurusan"", 8, y, 55, 12, True, 7" & vbCrLf
    s = s & "    Set txtJurusan = AddT(Me, ""txtJurusan"", 65, y - 1, 250, 16, True)" & vbCrLf
    s = s & "    txtJurusan.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblIPK"", ""IPK"", 325, y, 25, 12, True, 7" & vbCrLf
    s = s & "    Set txtIPK = AddT(Me, ""txtIPK"", 350, y - 1, 38, 16, True)" & vbCrLf
    s = s & "    txtIPK.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblTahun"", ""Lulus"", 395, y, 35, 12, True, 7" & vbCrLf
    s = s & "    Set txtTahunLulus = AddT(Me, ""txtTahunLulus"", 430, y - 1, 42, 16, True)" & vbCrLf
    s = s & "    txtTahunLulus.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblEnglish"", ""Eng"", 480, y, 25, 12, True, 7" & vbCrLf
    s = s & "    Set txtEnglish = AddT(Me, ""txtEnglish"", 505, y - 1, 38, 16, True)" & vbCrLf
    s = s & "    txtEnglish.BackColor = vbWhite" & vbCrLf
    s = s & "    y = y + 17" & vbCrLf
    s = s & "    AddL Me, ""lblUnivTier"", ""Univ Tier"", 8, y, 55, 12, True, 7" & vbCrLf
    s = s & "    Set txtUnivTier = AddT(Me, ""txtUnivTier"", 65, y - 1, 120, 16, True)" & vbCrLf
    s = s & "    txtUnivTier.BackColor = vbWhite" & vbCrLf
    s = s & "    AddL Me, ""lblIPKTier"", ""IPK Tier"", 195, y, 55, 12, True, 7" & vbCrLf
    s = s & "    Set txtIPKTier = AddT(Me, ""txtIPKTier"", 250, y - 1, 130, 16, True)" & vbCrLf
    s = s & "    txtIPKTier.BackColor = vbWhite" & vbCrLf
    s = s & "    y = y + 18" & vbCrLf
    s = s & "    Set lblStatusBesar = AddL(Me, ""lblStatusBesar"", ""STATUS: Menunggu proses Sourcing HR"", 8, y, 695, 20, True, 9)" & vbCrLf
    s = s & "    lblStatusBesar.BackColor = RGB(236, 240, 241)" & vbCrLf
    s = s & "    lblStatusBesar.TextAlign = fmTextAlignCenter" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "    AddSec Me, ""sec3"", ""PIPELINE REKRUTMEN"", 8, y, 695" & vbCrLf
    s = s & "    y = y + 14" & vbCrLf
    s = s & "    AddL Me, ""lblHdr1"", ""TAHAPAN"", 8, y, 125, 13, True, 7" & vbCrLf
    s = s & "    AddL Me, ""lblHdr2"", ""STATUS"", 138, y, 80, 13, True, 7" & vbCrLf
    s = s & "    AddL Me, ""lblHdr3"", ""TANGGAL (dd/mm/yyyy)"", 223, y, 105, 13, True, 7" & vbCrLf
    s = s & "    AddL Me, ""lblHdr4"", ""ALASAN"", 333, y, 370, 13, True, 7" & vbCrLf
    s = s & "    y = y + 13" & vbCrLf
    s = s & "    AddSep Me, 8, y, 695" & vbCrLf
    s = s & "    y = y + 2" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""Sourcing HR"", lblSourcingHR, lblTglSourcingHR, lblAlasanSourcingHR)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""Shortlist CV"", lblShortlist, lblTglShortlist, lblAlasanShortlist)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""Psikotes"", lblPsikotes, lblTglPsikotes, lblAlasanPsikotes)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""HR Interview"", lblHRInterview, lblTglHRInterview, lblAlasanHRInterview)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""Technical Test"", lblTechnicalTest, lblTglTechnicalTest, lblAlasanTechnicalTest)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""Market Visit"", lblMarketVisit, lblTglMarketVisit, lblAlasanMarketVisit)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""User Interview"", lblUserInterview, lblTglUserInterview, lblAlasanUserInterview)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""Panel Interview"", lblPanelInterview, lblTglPanelInterview, lblAlasanPanelInterview)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""Reference Check"", lblReferenceCheck, lblTglReferenceCheck, lblAlasanReferenceCheck)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""MCU"", lblMCU, lblTglMCU, lblAlasanMCU)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""Offering"", lblOffering, lblTglOffering, lblAlasanOffering)" & vbCrLf
    s = s & "    y = AddDRow(Me, 8, y, ""Day 1"", lblDay1, lblTglDay1, lblAlasanDay1)" & vbCrLf
    s = s & "    y = y + 2" & vbCrLf
    s = s & "    AddSep Me, 8, y, 695" & vbCrLf
    s = s & "    y = y + 8" & vbCrLf
    s = s & "    AddL Me, ""lblLegend"", ""Status: Lolos | Tidak Lolos | Belum  --  Tanggal: dd/mm/yyyy"", 8, y, 695, 13, False, 7" & vbCrLf
    s = s & "    y = y + 18" & vbCrLf
    s = s & "    Set btnEdit = AddB(Me, ""btnEdit"", ""EDIT KANDIDAT"", 200, y, 130, 26, RGB(243, 156, 18))" & vbCrLf
    s = s & "    Set btnClose = AddB(Me, ""btnClose"", ""TUTUP"", 360, y, 100, 26, RGB(231, 76, 60))" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddDRow(ByVal p As Object, ByVal xL As Single, ByVal y As Single, ByVal labelText As String, ByRef lblStatus As MSForms.Label, ByRef lblTgl As MSForms.Label, ByRef lblAlasan As MSForms.Label) As Single" & vbCrLf
    s = s & "    AddL p, ""lblName"" & labelText, labelText, xL, y, 125, 14, True, 7" & vbCrLf
    s = s & "    Set lblStatus = AddL(p, ""lblStatus"" & labelText, ""Belum"", xL + 130, y, 80, 14, True, 7)" & vbCrLf
    s = s & "    lblStatus.ForeColor = RGB(149, 165, 166)" & vbCrLf
    s = s & "    Set lblTgl = AddL(p, ""lblTgl"" & labelText, ""-"", xL + 215, y, 105, 14, False, 7)" & vbCrLf
    s = s & "    Set lblAlasan = AddL(p, ""lblAlasan"" & labelText, ""-"", xL + 325, y, 375, 14, False, 7)" & vbCrLf
    s = s & "    lblAlasan.ForeColor = RGB(149, 165, 166)" & vbCrLf
    s = s & "    AddDRow = y + 15" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub LoadRecord(ByVal rowNum As Long)" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "    mRow = rowNum" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    Me.Caption = ""DETAIL - Row "" & rowNum" & vbCrLf
    s = s & "    txtKode.Value = GetByHeader(ws, rowNum, ""Kode Unik (copy value dari FPTK)"")" & vbCrLf
    s = s & "    txtPosisi.Value = GetByHeader(ws, rowNum, ""Posisi"")" & vbCrLf
    s = s & "    txtNama.Value = GetByHeader(ws, rowNum, ""Nama"")" & vbCrLf
    s = s & "    txtSumber.Value = GetByHeader(ws, rowNum, ""Sumber Sourcing"")" & vbCrLf
    s = s & "    txtModel.Value = GetByHeader(ws, rowNum, ""Model Rekrutmen (Lihat di Sheet Flow Map Model)"")" & vbCrLf
    s = s & "    txtRekruter.Value = GetByHeader(ws, rowNum, ""Rekruter"")" & vbCrLf
    s = s & "    txtEmail.Value = GetByHeader(ws, rowNum, ""Email"")" & vbCrLf
    s = s & "    txtHP.Value = GetByHeader(ws, rowNum, ""Nomor HP"")" & vbCrLf
    s = s & "    txtDomisili.Value = GetByHeader(ws, rowNum, ""Domisili"")" & vbCrLf
    s = s & "    Dim ut As String" & vbCrLf
    s = s & "    ut = CStr(GetByHeader(ws, rowNum, ""Nama Universitas/Sekolah (TOP 10)""))" & vbCrLf
    s = s & "    Dim uo As String" & vbCrLf
    s = s & "    uo = CStr(GetByHeader(ws, rowNum, ""Nama Universitas/Sekolah Lainnya""))" & vbCrLf
    s = s & "    If Len(ut) > 0 Then" & vbCrLf
    s = s & "        txtUniv.Value = ut" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        txtUniv.Value = uo" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    txtJenjang.Value = GetByHeader(ws, rowNum, ""Jenjang Pendidikan"")" & vbCrLf
    s = s & "    txtJurusan.Value = GetByHeader(ws, rowNum, ""Jurusan"")" & vbCrLf
    s = s & "    txtIPK.Value = GetByHeader(ws, rowNum, ""IPK"")" & vbCrLf
    s = s & "    txtTahunLulus.Value = GetByHeader(ws, rowNum, ""Tahun Lulus"")" & vbCrLf
    s = s & "    txtEnglish.Value = GetByHeader(ws, rowNum, ""Skor Bahasa Inggris"")" & vbCrLf
    s = s & "    txtUnivTier.Value = GetByHeader(ws, rowNum, ""University Tier"")" & vbCrLf
    s = s & "    txtIPKTier.Value = GetByHeader(ws, rowNum, ""IPK Tier"")" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""Sourcing HR"", ""Detail Keterangan Sourcing HR"", ""Tanggal Sourcing"", lblSourcingHR, lblTglSourcingHR, lblAlasanSourcingHR" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""Shortlist CV"", ""Detail Keterangan Shortlist CV"", ""Tanggal Shortlist CV"", lblShortlist, lblTglShortlist, lblAlasanShortlist" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""Psikotes"", ""Detail Keterangan Psikotes"", ""Tanggal Psikotes / Cek psikotes"", lblPsikotes, lblTglPsikotes, lblAlasanPsikotes" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""HR Interview"", ""Detail Keterangan HR Interview"", ""Tanggal HR Interview"", lblHRInterview, lblTglHRInterview, lblAlasanHRInterview" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""Technical Test/ Case Study"", ""Detail Keterangan Technical Test/ Case Study"", ""Tanggal Technical Test/ Case Study"", lblTechnicalTest, lblTglTechnicalTest, lblAlasanTechnicalTest" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""Market Visit"", ""Detail Market Visit"", ""Tanggal Market Visit"", lblMarketVisit, lblTglMarketVisit, lblAlasanMarketVisit" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""User Interview"", ""Detail Keterangan User Interview"", ""Tanggal User Interview"", lblUserInterview, lblTglUserInterview, lblAlasanUserInterview" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""Panel Interview"", ""Detail Keterangan Panel Interview"", ""Tanggal Panel Interview"", lblPanelInterview, lblTglPanelInterview, lblAlasanPanelInterview" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""Reference Check"", ""Detail Keterangan Reference Check"", ""Tanggal Reference Check"", lblReferenceCheck, lblTglReferenceCheck, lblAlasanReferenceCheck" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""MCU"", ""Detail Keterangan MCU"", ""Tanggal MCU"", lblMCU, lblTglMCU, lblAlasanMCU" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""Offering"", ""Detail Keterangan Offering"", ""Tanggal Offering"", lblOffering, lblTglOffering, lblAlasanOffering" & vbCrLf
    s = s & "    LoadDL ws, rowNum, ""Day 1"", ""Detail Keterangan Day 1"", ""Tanggal Day 1"", lblDay1, lblTglDay1, lblAlasanDay1" & vbCrLf
    s = s & "    UpdateStatus" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadDL(ByVal ws As Worksheet, ByVal r As Long, ByVal hS As String, ByVal hD As String, ByVal hT As String, ByRef lblStatus As MSForms.Label, ByRef lblTgl As MSForms.Label, ByRef lblAlasan As MSForms.Label)" & vbCrLf
    s = s & "    Dim sv As String" & vbCrLf
    s = s & "    sv = Trim(CStr(GetByHeader(ws, r, hS)))" & vbCrLf
    s = s & "    Dim dv As String" & vbCrLf
    s = s & "    dv = Trim(CStr(GetByHeader(ws, r, hD)))" & vbCrLf
    s = s & "    Dim tv As String" & vbCrLf
    s = s & "    tv = FormatTanggal(CStr(GetByHeader(ws, r, hT)))" & vbCrLf
    s = s & "    If InStr(1, sv, ""V"") > 0 Or InStr(1, LCase(sv), ""lolos"") > 0 Or InStr(1, LCase(sv), ""yes"") > 0 Then" & vbCrLf
    s = s & "        lblStatus.Caption = ""Lolos""" & vbCrLf
    s = s & "        lblStatus.ForeColor = RGB(39, 174, 96)" & vbCrLf
    s = s & "    ElseIf InStr(1, LCase(sv), ""x"") > 0 Or InStr(1, LCase(sv), ""tidak"") > 0 Or InStr(1, LCase(sv), ""no"") > 0 Then" & vbCrLf
    s = s & "        lblStatus.Caption = ""Tidak Lolos""" & vbCrLf
    s = s & "        lblStatus.ForeColor = RGB(231, 76, 60)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblStatus.Caption = ""Belum""" & vbCrLf
    s = s & "        lblStatus.ForeColor = RGB(149, 165, 166)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Len(tv) > 0 Then" & vbCrLf
    s = s & "        lblTgl.Caption = tv" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblTgl.Caption = ""-""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Len(dv) > 0 Then" & vbCrLf
    s = s & "        lblAlasan.Caption = dv" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblAlasan.Caption = ""-""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If InStr(1, lblStatus.Caption, ""Tidak Lolos"") > 0 And Len(dv) > 0 Then" & vbCrLf
    s = s & "        lblAlasan.ForeColor = RGB(231, 76, 60)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblAlasan.ForeColor = RGB(149, 165, 166)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FormatTanggal(ByVal tgl As String) As String" & vbCrLf
    s = s & "    If Len(Trim(tgl)) = 0 Then" & vbCrLf
    s = s & "        FormatTanggal = """"" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim d As String" & vbCrLf
    s = s & "    d = Replace(Trim(tgl), ""-"", ""/"")" & vbCrLf
    s = s & "    d = Replace(d, ""."", ""/"")" & vbCrLf
    s = s & "    d = Replace(d, "" "", ""/"")" & vbCrLf
    s = s & "    Dim parts() As String" & vbCrLf
    s = s & "    parts = Split(d, ""/"")" & vbCrLf
    s = s & "    If UBound(parts) >= 2 Then" & vbCrLf
    s = s & "        Dim p1 As Long" & vbCrLf
    s = s & "        p1 = Val(parts(0))" & vbCrLf
    s = s & "        Dim p2 As Long" & vbCrLf
    s = s & "        p2 = Val(parts(1))" & vbCrLf
    s = s & "        Dim yy As String" & vbCrLf
    s = s & "        yy = parts(2)" & vbCrLf
    s = s & "        If Len(yy) = 2 Then" & vbCrLf
    s = s & "            If Val(yy) > 50 Then" & vbCrLf
    s = s & "                yy = ""19"" & yy" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                yy = ""20"" & yy" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        Dim dd As String" & vbCrLf
    s = s & "        Dim mm As String" & vbCrLf
    s = s & "        If p1 > 31 Then" & vbCrLf
    s = s & "            FormatTanggal = tgl" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If p1 > 12 Then" & vbCrLf
    s = s & "            dd = Format(p1, ""00"")" & vbCrLf
    s = s & "            mm = Format(p2, ""00"")" & vbCrLf
    s = s & "        ElseIf p2 > 12 Then" & vbCrLf
    s = s & "            mm = Format(p1, ""00"")" & vbCrLf
    s = s & "            dd = Format(p2, ""00"")" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            dd = Format(p1, ""00"")" & vbCrLf
    s = s & "            mm = Format(p2, ""00"")" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If Val(mm) > 12 Or Val(mm) < 1 Then" & vbCrLf
    s = s & "            FormatTanggal = tgl" & vbCrLf
    s = s & "        ElseIf Val(dd) > 31 Or Val(dd) < 1 Then" & vbCrLf
    s = s & "            FormatTanggal = tgl" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            FormatTanggal = dd & ""/"" & mm & ""/"" & yy" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        FormatTanggal = tgl" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateStatus()" & vbCrLf
    s = s & "    Dim st As String" & vbCrLf
    s = s & "    Dim bg As Long" & vbCrLf
    s = s & "    If lblDay1.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Day 1 | "" & lblTglDay1.Caption" & vbCrLf
    s = s & "        bg = RGB(39, 174, 96)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf lblDay1.Caption = ""Tidak Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Tidak Lolos Day 1 | "" & lblAlasanDay1.Caption" & vbCrLf
    s = s & "        bg = RGB(231, 76, 60)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf lblOffering.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Offering | "" & lblTglOffering.Caption" & vbCrLf
    s = s & "        bg = RGB(39, 174, 96)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf lblOffering.Caption = ""Tidak Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Tidak Lolos Offering | "" & lblAlasanOffering.Caption" & vbCrLf
    s = s & "        bg = RGB(231, 76, 60)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf lblMCU.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos MCU | "" & lblTglMCU.Caption" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf lblMCU.Caption = ""Tidak Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Tidak Lolos MCU | "" & lblAlasanMCU.Caption" & vbCrLf
    s = s & "        bg = RGB(231, 76, 60)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf lblReferenceCheck.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Reference Check | "" & lblTglReferenceCheck.Caption" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf lblReferenceCheck.Caption = ""Tidak Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Tidak Lolos Reference Check | "" & lblAlasanReferenceCheck.Caption" & vbCrLf
    s = s & "        bg = RGB(231, 76, 60)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf lblPanelInterview.Caption = ""Lolos"" Or lblUserInterview.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos User/Panel Interview | "" & lblTglUserInterview.Caption" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf lblPanelInterview.Caption = ""Tidak Lolos"" Or lblUserInterview.Caption = ""Tidak Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Tidak Lolos User/Panel Interview | "" & lblAlasanUserInterview.Caption" & vbCrLf
    s = s & "        bg = RGB(231, 76, 60)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf lblMarketVisit.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Market Visit | "" & lblTglMarketVisit.Caption" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf lblTechnicalTest.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Technical Test | "" & lblTglTechnicalTest.Caption" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf lblHRInterview.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos HR Interview | "" & lblTglHRInterview.Caption" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf lblPsikotes.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Psikotes | "" & lblTglPsikotes.Caption" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf lblShortlist.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Shortlist | "" & lblTglShortlist.Caption" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf lblSourcingHR.Caption = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Sourcing HR | "" & lblTglSourcingHR.Caption" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        st = ""STATUS: Menunggu proses Sourcing HR""" & vbCrLf
    s = s & "        bg = RGB(236, 240, 241)" & vbCrLf
    s = s & "        lblStatusBesar.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lblStatusBesar.Caption = st" & vbCrLf
    s = s & "    lblStatusBesar.BackColor = bg" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnEdit_Click()" & vbCrLf
    s = s & "    If mRow = 0 Then Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglSourcingHR.Value) Then MsgBox ""Tanggal Sourcing HR tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglShortlist.Value) Then MsgBox ""Tanggal Shortlist tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglPsikotes.Value) Then MsgBox ""Tanggal Psikotes tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglHRInterview.Value) Then MsgBox ""Tanggal HR Interview tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglTechnicalTest.Value) Then MsgBox ""Tanggal Technical Test tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglMarketVisit.Value) Then MsgBox ""Tanggal Market Visit tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglUserInterview.Value) Then MsgBox ""Tanggal User Interview tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglPanelInterview.Value) Then MsgBox ""Tanggal Panel Interview tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglReferenceCheck.Value) Then MsgBox ""Tanggal Reference Check tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglMCU.Value) Then MsgBox ""Tanggal MCU tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglOffering.Value) Then MsgBox ""Tanggal Offering tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Not IsValidTgl(txtTglDay1.Value) Then MsgBox ""Tanggal Day 1 tidak valid! Gunakan dd/mm/yyyy"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    frmEditSourcing.LoadRecord mRow" & vbCrLf
    s = s & "    frmEditSourcing.Show vbModal" & vbCrLf
    s = s & "    LoadRecord mRow" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnClose_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
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
    s = s & "    If b Then" & vbCrLf
    s = s & "        o.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set AddL = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddT(ByVal p As Object, ByVal n As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal lk As Boolean = False) As MSForms.TextBox" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.TextBox.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l" & vbCrLf
    s = s & "    o.Top = t" & vbCrLf
    s = s & "    o.Width = w" & vbCrLf
    s = s & "    o.Height = h" & vbCrLf
    s = s & "    o.Font.Size = 8" & vbCrLf
    s = s & "    If lk Then" & vbCrLf
    s = s & "        o.Locked = True" & vbCrLf
    s = s & "        o.BackColor = RGB(250, 250, 250)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set AddT = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
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
    s = s & "" & vbCrLf
    s = s & "Private Sub AddSec(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single)" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.Label.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l" & vbCrLf
    s = s & "    o.Top = t" & vbCrLf
    s = s & "    o.Width = w" & vbCrLf
    s = s & "    o.Height = 14" & vbCrLf
    s = s & "    o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = 7" & vbCrLf
    s = s & "    o.Font.Bold = True" & vbCrLf
    s = s & "    o.BackColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    o.ForeColor = vbWhite" & vbCrLf
    s = s & "    o.BackStyle = 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub AddSep(ByVal p As Object, ByVal l As Single, ByVal t As Single, ByVal w As Single)" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.Label.1"", ""sep"" & t, True)" & vbCrLf
    s = s & "    o.Left = l" & vbCrLf
    s = s & "    o.Top = t" & vbCrLf
    s = s & "    o.Width = w" & vbCrLf
    s = s & "    o.Height = 1" & vbCrLf
    s = s & "    o.BackColor = RGB(189, 195, 199)" & vbCrLf
    s = s & "    o.BackStyle = 1" & vbCrLf
    s = s & "    o.Caption = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    Code_frmSourcingDetail = s
    s = s & "Private Function IsValidTgl(ByVal tgl As String) As Boolean" & vbCrLf
    s = s & "    If Len(Trim(tgl)) = 0 Then" & vbCrLf
    s = s & "        IsValidTgl = True" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim ft As String" & vbCrLf
    s = s & "    ft = FormatTgl(tgl)" & vbCrLf
    s = s & "    If ft = tgl Then" & vbCrLf
    s = s & "        Dim parts() As String" & vbCrLf
    s = s & "        parts = Split(Replace(Replace(Replace(Trim(tgl), ""-"", ""/""), ""."", ""/""), "" "", ""/""), ""/"")" & vbCrLf
    s = s & "        If UBound(parts) >= 2 Then" & vbCrLf
    s = s & "            If Val(parts(1)) > 12 Then" & vbCrLf
    s = s & "                IsValidTgl = False" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            If Val(parts(0)) > 31 Then" & vbCrLf
    s = s & "                IsValidTgl = False" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    IsValidTgl = True" & vbCrLf
    s = s & "End Function" & vbCrLf
End Function

Private Function Code_frmEditSourcing() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "Private mBuilt As Boolean" & vbCrLf
    s = s & "Private mRow As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private txtNama As MSForms.TextBox" & vbCrLf
    s = s & "Private txtPosisi As MSForms.TextBox" & vbCrLf
    s = s & "Private txtKode As MSForms.TextBox" & vbCrLf
    s = s & "Private cboSumber As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboModel As MSForms.ComboBox" & vbCrLf
    s = s & "Private txtRekruter As MSForms.TextBox" & vbCrLf
    s = s & "Private cboUnivTop As MSForms.ComboBox" & vbCrLf
    s = s & "Private txtUnivOther As MSForms.TextBox" & vbCrLf
    s = s & "Private cboJenjang As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboJurusan As MSForms.ComboBox" & vbCrLf
    s = s & "Private txtTahunLulus As MSForms.TextBox" & vbCrLf
    s = s & "Private txtIPK As MSForms.TextBox" & vbCrLf
    s = s & "Private txtEnglish As MSForms.TextBox" & vbCrLf
    s = s & "Private cboUnivTier As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboIPKTier As MSForms.ComboBox" & vbCrLf
    s = s & "Private txtHP As MSForms.TextBox" & vbCrLf
    s = s & "Private txtEmail As MSForms.TextBox" & vbCrLf
    s = s & "Private txtDomisili As MSForms.TextBox" & vbCrLf
    s = s & "Private lblStatusTerakhir As MSForms.Label" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private cboSourcingHR As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboShortlist As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboPsikotes As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboHRInterview As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboTechnicalTest As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboMarketVisit As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboUserInterview As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboPanelInterview As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboReferenceCheck As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboMCU As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboOffering As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboDay1 As MSForms.ComboBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private txtTglSourcingHR As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglShortlist As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglPsikotes As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglHRInterview As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglTechnicalTest As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglMarketVisit As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglUserInterview As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglPanelInterview As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglReferenceCheck As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglMCU As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglOffering As MSForms.TextBox" & vbCrLf
    s = s & "Private txtTglDay1 As MSForms.TextBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private cboAlasanSourcingHR As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanShortlist As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanPsikotes As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanHRInterview As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanTechnicalTest As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanMarketVisit As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanUserInterview As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanPanelInterview As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanReferenceCheck As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanMCU As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanOffering As MSForms.ComboBox" & vbCrLf
    s = s & "Private cboAlasanDay1 As MSForms.ComboBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private WithEvents btnSave As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnCancel As MSForms.CommandButton" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "    FillDropdowns" & vbCrLf
    s = s & "    FillAlasanDropdowns" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FillAlasanDropdowns()" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    If cboAlasanSourcingHR.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Pengalaman tidak sesuai dengan User"",""Jenjang pendidikan tidak sesuai dengan kualifikasi"",""Gender tidak sesuai dengan permintaan User"",""Usia tidak sesuai dengan permintaan User"",""Tidak bersedia dengan lokasi penempatan"",""Kandidat sudah pernah diproses sebelumnya"",""Background keluarga tidak sesuai"",""Riwayat organisasi tidak sesuai"",""Gaji tidak sesuai dengan budget FPTK"",""Kurang pengalaman organisasi"",""Tahun pengalaman tidak sesuai dengan permintaan User"",""PTN di luar Top 20"",""PTS di luar Top 10"",""IPK kurang dari 3,5"",""IPK kurang dari 3,2"",""Score bahasa inggris kurang dari standard"",""Pengalaman lebih dari 1 tahun"",""Kandidat withdraw"",""Jurusan tidak sesuai dengan permintaan User"",""Tidak melewati tahap shortlist CV dari User"")" & vbCrLf
    s = s & "            cboAlasanSourcingHR.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanShortlist.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Pengalaman tidak sesuai dengan User"",""Jenjang pendidikan tidak sesuai dengan kualifikasi"",""Gender tidak sesuai dengan permintaan User"",""Usia tidak sesuai dengan permintaan User"",""Tidak bersedia dengan lokasi penempatan"",""Kandidat sudah pernah diproses sebelumnya"",""Background keluarga tidak sesuai"",""Riwayat organisasi tidak sesuai"",""Gaji tidak sesuai dengan Budget FPTK"",""PTN di luar Top 20"",""PTS di luar Top 10"",""IPK kurang dari 3,5"",""IPK kurang dari 3,2"",""Score bahasa inggris kurang dari standard"",""Pengalaman lebih dari 1 tahun"",""Kandidat withdraw"",""Jurusan tidak sesuai dengan permintaan User"",""Tidak melewati tahap shortlist CV dari User"")" & vbCrLf
    s = s & "            cboAlasanShortlist.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanPsikotes.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Logika Tidak Lulus"",""IQ Tidak Lulus"",""RA Tidak Lulus"",""Daya Tangkap Tidak Lulus"",""DISC Tidak Lulus"",""Lebih dari 2 Aspek Tidak Lulus"",""Lulus namun usia tidak sesuai FPTK"",""Kode Expired"",""Mengerjakan dengan orang lain"",""Tidak merespon setelah di follow up berulang kali"",""Tidak melewati tahap psikotes"",""Belum mengerjakan"",""Kandidat withdraw"")" & vbCrLf
    s = s & "            cboAlasanPsikotes.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanHRInterview.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Tidak bersedia dengan lokasi penempatan"",""Data kandidat tidak sesuai dengan CV"",""Tidak hadir interview"",""Background keluarga tidak sesuai"",""Kandidat tidak menguasai technical skill yang dibutuhkan"",""Kurang memiliki drive yang tinggi"",""Kurang inisiatif dalam meningkatkan efektivitas/efisiensi kerja"",""Kurang pengalaman leadership"",""Kurang terstruktur dalam menjawab pertanyaan"",""Tidak bersedia ditempatkan di seluruh Indonesia"",""Tidak bersedia memperpanjang SIM"",""Kandidat withdraw"",""Tidak melewati tahap HR Interview"",""Tidak lancar percakapan bahasa inggris"",""Ekspektasi gaji di luar budget / non-negotiable"",""Tidak mempelajari ARC sebelum interview"",""Tidak tertarik dengan posisi setelah mendengar penjelasan dari Rekruter"",""Kurang memiliki analytical thinking dan problem solving"")" & vbCrLf
    s = s & "            cboAlasanHRInterview.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanUserInterview.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Soft skills tidak sesuai"",""Culture tidak sesuai dengan tim"",""Sudah diterima di perusahaan lain"",""Technical skills tidak sesuai"",""Kandidat tidak merespon setelah di follow up berulang kali"",""Ekspektasi gaji tidak sesuai dengan HR Interview"",""Kandidat withdraw"",""Tidak hadir interview"",""Kurang memiliki drive yang tinggi"",""Kurang inisiatif dalam meningkatkan efektivitas/efisiensi kerja"",""Kurang pengalaman leadership"",""Kurang terstruktur dalam menjawab pertanyaan"",""Kandidat cenderung menyalahkan pihak lain (blaming)"")" & vbCrLf
    s = s & "            cboAlasanUserInterview.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanPanelInterview.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Soft skills tidak sesuai"",""Culture tidak sesuai dengan tim"",""Sudah diterima di perusahaan lain"",""Technical skills tidak sesuai"",""Kandidat tidak merespon setelah di follow up berulang kali"",""Ekspektasi gaji tidak sesuai dengan HR Interview"",""Kandidat withdraw"",""Tidak hadir interview"",""Kurang memiliki drive yang tinggi"",""Kurang inisiatif dalam meningkatkan efektivitas/efisiensi kerja"",""Kurang pengalaman leadership"",""Kurang terstruktur dalam menjawab pertanyaan"",""Kandidat cenderung menyalahkan pihak lain (blaming)"")" & vbCrLf
    s = s & "            cboAlasanPanelInterview.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanTechnicalTest.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Technical skills tidak sesuai"",""Laptop & internet tidak support"",""Tidak melewati tahap technical test"",""Kandidat withdraw"",""Tidak hadir technical test"",""Kandidat tidak submit Case Study"",""Case study didominasi oleh AI"",""Analisa case study kurang tajam"")" & vbCrLf
    s = s & "            cboAlasanTechnicalTest.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanMarketVisit.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Kandidat withdraw"",""Analisa kurang tajam"",""Data pendukung kurang lengkap"",""Kurang komunikatif saat presentasi"",""Kurang memahami teknis di lapangan"",""Presentasi didominasi oleh AI"",""Tidak melewati tahap Market Visit"",""Kurang menguasai product knowledge"")" & vbCrLf
    s = s & "            cboAlasanMarketVisit.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanReferenceCheck.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Kasus Integritas/Fraud"",""Kurang perform"",""Attitude Issue"",""Kedisiplinan"",""Kandidat tidak merespon setelah di follow up berulang kali"",""Tidak melewati reference check"",""Belum reference check"",""Kandidat memalsukan kontak reference check"",""Kandidat withdraw"")" & vbCrLf
    s = s & "            cboAlasanReferenceCheck.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanMCU.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Kurang sehat"",""Perlu pengecekan/diagnosa lebih lanjut"",""Keberatan membayar MCU"",""Tidak melewati tahap MCU"",""Kandidat withdraw"",""Sudah menerima offering dari perusahaan lain"",""Ditahan oleh perusahaan saat ini"",""Kandidat tidak merespon setelah di follow up berulang kali"")" & vbCrLf
    s = s & "            cboAlasanMCU.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanOffering.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Sudah menerima offering dari perusahaan lain"",""Keluarga keberatan terkait offer"",""Tidak bersedia dengan lokasi penempatan"",""Ditahan oleh perusahaan saat ini"",""Tidak menjawab offering"",""Keberatan dengan status karyawan"",""Gaji yang ditawarkan tidak sesuai ekspektasi"",""Keberatan dengan penalty"",""Kandidat withdraw"",""Keberatan dengan ikatan dinas"")" & vbCrLf
    s = s & "            cboAlasanOffering.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboAlasanDay1.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Sudah menerima offering dari perusahaan lain"",""Ditahan oleh perusahaan saat ini"",""Kandidat tidak merespon setelah di follow up berulang kali"")" & vbCrLf
    s = s & "            cboAlasanDay1.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildUI()" & vbCrLf
    s = s & "    If mBuilt Then Exit Sub" & vbCrLf
    s = s & "    mBuilt = True" & vbCrLf
    s = s & "    Me.Caption = ""EDIT KANDIDAT & PIPELINE""" & vbCrLf
    s = s & "    Me.Width = 730" & vbCrLf
    s = s & "    Me.Height = 580" & vbCrLf
    s = s & "    Me.BackColor = RGB(245, 245, 245)" & vbCrLf
    s = s & "    Me.ScrollBars = fmScrollBarsVertical" & vbCrLf
    s = s & "    Me.ScrollHeight = 950" & vbCrLf
    s = s & "    Dim y As Single" & vbCrLf
    s = s & "    y = 4" & vbCrLf
    s = s & "    AddSection Me, ""sec1"", ""DATA KANDIDAT"", 8, y, 705" & vbCrLf
    s = s & "    y = y + 16" & vbCrLf
    s = s & "    AddL Me, ""lblNama"", ""Nama"", 8, y, 40, 12, True, 7" & vbCrLf
    s = s & "    Set txtNama = AddT(Me, ""txtNama"", 50, y - 1, 180, 17)" & vbCrLf
    s = s & "    AddL Me, ""lblPosisi"", ""Posisi"", 240, y, 40, 12, True, 7" & vbCrLf
    s = s & "    Set txtPosisi = AddT(Me, ""txtPosisi"", 285, y - 1, 430, 17, True)" & vbCrLf
    s = s & "    y = y + 19" & vbCrLf
    s = s & "    AddL Me, ""lblKode"", ""Kode"", 8, y, 40, 12, True, 7" & vbCrLf
    s = s & "    Set txtKode = AddT(Me, ""txtKode"", 50, y - 1, 110, 17, True)" & vbCrLf
    s = s & "    AddL Me, ""lblSumber"", ""Sumber"", 170, y, 40, 12, True, 7" & vbCrLf
    s = s & "    Set cboSumber = AddC(Me, ""cboSumber"", 210, y - 1, 115, 17)" & vbCrLf
    s = s & "    cboSumber.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    AddL Me, ""lblModel"", ""Model"", 335, y, 35, 12, True, 7" & vbCrLf
    s = s & "    Set cboModel = AddC(Me, ""cboModel"", 370, y - 1, 55, 17)" & vbCrLf
    s = s & "    cboModel.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    AddL Me, ""lblRekruter"", ""Rekr"", 435, y, 30, 12, True, 7" & vbCrLf
    s = s & "    Set txtRekruter = AddT(Me, ""txtRekruter"", 465, y - 1, 55, 17)" & vbCrLf
    s = s & "    AddL Me, ""lblDomisili"", ""Dom"", 530, y, 30, 12, True, 7" & vbCrLf
    s = s & "    Set txtDomisili = AddT(Me, ""txtDomisili"", 560, y - 1, 70, 17)" & vbCrLf
    s = s & "    AddL Me, ""lblHP"", ""HP"", 640, y, 25, 12, True, 7" & vbCrLf
    s = s & "    Set txtHP = AddT(Me, ""txtHP"", 665, y - 1, 50, 17)" & vbCrLf
    s = s & "    y = y + 19" & vbCrLf
    s = s & "    AddL Me, ""lblEmail"", ""Email"", 8, y, 40, 12, True, 7" & vbCrLf
    s = s & "    Set txtEmail = AddT(Me, ""txtEmail"", 50, y - 1, 200, 17)" & vbCrLf
    s = s & "    y = y + 20" & vbCrLf
    s = s & "    AddSection Me, ""sec2"", ""PENDIDIKAN"", 8, y, 705" & vbCrLf
    s = s & "    y = y + 16" & vbCrLf
    s = s & "    AddL Me, ""lblUniv"", ""Univ TOP10"", 8, y, 60, 12, True, 7" & vbCrLf
    s = s & "    Set cboUnivTop = AddC(Me, ""cboUnivTop"", 70, y - 1, 170, 17)" & vbCrLf
    s = s & "    cboUnivTop.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    AddL Me, ""lblUnivOther"", ""Lain"", 250, y, 30, 12, True, 7" & vbCrLf
    s = s & "    Set txtUnivOther = AddT(Me, ""txtUnivOther"", 280, y - 1, 150, 17)" & vbCrLf
    s = s & "    AddL Me, ""lblJenjang"", ""Jenjang"", 440, y, 45, 12, True, 7" & vbCrLf
    s = s & "    Set cboJenjang = AddC(Me, ""cboJenjang"", 485, y - 1, 60, 17)" & vbCrLf
    s = s & "    cboJenjang.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    y = y + 19" & vbCrLf
    s = s & "    AddL Me, ""lblJurusan"", ""Jurusan"", 8, y, 60, 12, True, 7" & vbCrLf
    s = s & "    Set cboJurusan = AddC(Me, ""cboJurusan"", 70, y - 1, 170, 17)" & vbCrLf
    s = s & "    cboJurusan.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    AddL Me, ""lblTahun"", ""Lulus"", 250, y, 35, 12, True, 7" & vbCrLf
    s = s & "    Set txtTahunLulus = AddT(Me, ""txtTahunLulus"", 285, y - 1, 50, 17)" & vbCrLf
    s = s & "    AddL Me, ""lblIPK"", ""IPK"", 345, y, 22, 12, True, 7" & vbCrLf
    s = s & "    Set txtIPK = AddT(Me, ""txtIPK"", 367, y - 1, 42, 17)" & vbCrLf
    s = s & "    AddL Me, ""lblEnglish"", ""Eng"", 418, y, 22, 12, True, 7" & vbCrLf
    s = s & "    Set txtEnglish = AddT(Me, ""txtEnglish"", 440, y - 1, 42, 17)" & vbCrLf
    s = s & "    AddL Me, ""lblUnivTier"", ""Univ Tier"", 492, y, 50, 12, True, 7" & vbCrLf
    s = s & "    Set cboUnivTier = AddC(Me, ""cboUnivTier"", 542, y - 1, 120, 17)" & vbCrLf
    s = s & "    cboUnivTier.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    y = y + 19" & vbCrLf
    s = s & "    AddL Me, ""lblIPKTier"", ""IPK Tier"", 8, y, 60, 12, True, 7" & vbCrLf
    s = s & "    Set cboIPKTier = AddC(Me, ""cboIPKTier"", 70, y - 1, 135, 17)" & vbCrLf
    s = s & "    cboIPKTier.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    y = y + 22" & vbCrLf
    s = s & "    Set lblStatusTerakhir = AddL(Me, ""lblStatusTerakhir"", ""STATUS: Belum ada proses"", 8, y, 705, 20, True, 9)" & vbCrLf
    s = s & "    lblStatusTerakhir.BackColor = RGB(236, 240, 241)" & vbCrLf
    s = s & "    lblStatusTerakhir.TextAlign = fmTextAlignCenter" & vbCrLf
    s = s & "    y = y + 24" & vbCrLf
    s = s & "    AddSection Me, ""sec3"", ""PIPELINE (Lolos / Tidak Lolos / Belum) | Tanggal: dd/mm/yyyy"", 8, y, 705" & vbCrLf
    s = s & "    y = y + 16" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""Sourcing HR"", cboSourcingHR, txtTglSourcingHR, cboAlasanSourcingHR)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""Shortlist CV"", cboShortlist, txtTglShortlist, cboAlasanShortlist)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""Psikotes"", cboPsikotes, txtTglPsikotes, cboAlasanPsikotes)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""HR Interview"", cboHRInterview, txtTglHRInterview, cboAlasanHRInterview)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""Technical Test"", cboTechnicalTest, txtTglTechnicalTest, cboAlasanTechnicalTest)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""Market Visit"", cboMarketVisit, txtTglMarketVisit, cboAlasanMarketVisit)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""User Interview"", cboUserInterview, txtTglUserInterview, cboAlasanUserInterview)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""Panel Interview"", cboPanelInterview, txtTglPanelInterview, cboAlasanPanelInterview)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""Reference Check"", cboReferenceCheck, txtTglReferenceCheck, cboAlasanReferenceCheck)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""MCU"", cboMCU, txtTglMCU, cboAlasanMCU)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""Offering"", cboOffering, txtTglOffering, cboAlasanOffering)" & vbCrLf
    s = s & "    y = AddPipe(Me, 8, y, ""Day 1"", cboDay1, txtTglDay1, cboAlasanDay1)" & vbCrLf
    s = s & "    y = y + 6" & vbCrLf
    s = s & "    Set btnSave = AddB(Me, ""btnSave"", ""SIMPAN"", 200, y, 130, 28, RGB(39, 174, 96))" & vbCrLf
    s = s & "    Set btnCancel = AddB(Me, ""btnCancel"", ""TUTUP"", 360, y, 100, 28, RGB(231, 76, 60))" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddPipe(ByVal p As Object, ByVal xL As Single, ByVal y As Single, ByVal labelText As String, ByRef cboStatus As MSForms.ComboBox, ByRef txtTgl As MSForms.TextBox, ByRef cboAlasan As MSForms.ComboBox) As Single" & vbCrLf
    s = s & "    AddL p, ""lbl"" & labelText, labelText, xL, y, 110, 13, True, 7" & vbCrLf
    s = s & "    Set cboStatus = AddC(p, ""cbo"" & labelText, xL + 113, y, 78, 15)" & vbCrLf
    s = s & "    cboStatus.Font.Size = 7" & vbCrLf
    s = s & "    cboStatus.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    cboStatus.AddItem ""Belum""" & vbCrLf
    s = s & "    cboStatus.AddItem ""Lolos""" & vbCrLf
    s = s & "    cboStatus.AddItem ""Tidak Lolos""" & vbCrLf
    s = s & "    cboStatus.Value = ""Belum""" & vbCrLf
    s = s & "    Set txtTgl = AddT(p, ""txtTgl"" & labelText, xL + 194, y, 78, 15)" & vbCrLf
    s = s & "    txtTgl.Font.Size = 7" & vbCrLf
    s = s & "    txtTgl.ControlTipText = ""Format: dd/mm/yyyy (contoh: 15/02/2026)""" & vbCrLf
    s = s & "    Set cboAlasan = AddC(p, ""cboAlasan"" & labelText, xL + 275, y, 440, 15)" & vbCrLf
    s = s & "    cboAlasan.Font.Size = 7" & vbCrLf
    s = s & "    cboAlasan.Style = fmStyleDropDownCombo" & vbCrLf
    s = s & "    AddPipe = y + 16" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FillDropdowns()" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    If cboSumber.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Jobstreet"",""LinkedIn"",""Google Form"",""Referensi User"",""Referensi Karyawan"",""Campus Hiring"",""Walk-in Interview"",""Database Internal"",""Freelance"")" & vbCrLf
    s = s & "            cboSumber.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboModel.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Model 1"",""Model 2"",""Model 3"",""Model 4"")" & vbCrLf
    s = s & "            cboModel.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboJenjang.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""SMA/SMK"",""D3"",""D4"",""S1"",""S2"")" & vbCrLf
    s = s & "            cboJenjang.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboUnivTop.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Universitas Indonesia"",""Universitas Gadjah Mada"",""Institut Teknologi Bandung"",""Universitas Airlangga"",""Universitas Padjadjaran"",""Universitas Diponegoro"",""Universitas Brawijaya"",""Institut Pertanian Bogor"",""Universitas Sebelas Maret"",""Telkom University"",""Lainnya"")" & vbCrLf
    s = s & "            cboUnivTop.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboJurusan.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Manajemen"",""Akuntansi"",""Teknik Industri"",""Teknik Informatika"",""Teknik Mesin"",""Ekonomi Bisnis"",""Psikologi"",""Ilmu Komunikasi"",""Sistem Informasi"",""Lainnya"")" & vbCrLf
    s = s & "            cboJurusan.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboUnivTier.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Top 3 PTN"",""Top 10 PTN"",""Top 20 PTN"",""Top 10 PTS"",""Lainnya"")" & vbCrLf
    s = s & "            cboUnivTier.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboIPKTier.ListCount = 0 Then" & vbCrLf
    s = s & "        For Each v In Array(""Lebih dari 3,5"",""Lebih dari 3,2"",""Kurang dari 3,5"",""Kurang dari 3,2"")" & vbCrLf
    s = s & "            cboIPKTier.AddItem v" & vbCrLf
    s = s & "        Next v" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function FormatTgl(ByVal tgl As String) As String" & vbCrLf
    s = s & "    If Len(Trim(tgl)) = 0 Then" & vbCrLf
    s = s & "        FormatTgl = """"" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim d As String" & vbCrLf
    s = s & "    d = Replace(Trim(tgl), ""-"", ""/"")" & vbCrLf
    s = s & "    d = Replace(d, ""."", ""/"")" & vbCrLf
    s = s & "    d = Replace(d, "" "", ""/"")" & vbCrLf
    s = s & "    Dim parts() As String" & vbCrLf
    s = s & "    parts = Split(d, ""/"")" & vbCrLf
    s = s & "    If UBound(parts) >= 2 Then" & vbCrLf
    s = s & "        Dim p1 As Long" & vbCrLf
    s = s & "        p1 = Val(parts(0))" & vbCrLf
    s = s & "        Dim p2 As Long" & vbCrLf
    s = s & "        p2 = Val(parts(1))" & vbCrLf
    s = s & "        Dim yy As String" & vbCrLf
    s = s & "        yy = parts(2)" & vbCrLf
    s = s & "        If Len(yy) = 2 Then" & vbCrLf
    s = s & "            If Val(yy) > 50 Then" & vbCrLf
    s = s & "                yy = ""19"" & yy" & vbCrLf
    s = s & "            Else" & vbCrLf
    s = s & "                yy = ""20"" & yy" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        Dim dd As String" & vbCrLf
    s = s & "        Dim mm As String" & vbCrLf
    s = s & "        If p1 > 31 Then" & vbCrLf
    s = s & "            FormatTgl = tgl" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If p1 > 12 Then" & vbCrLf
    s = s & "            dd = Format(p1, ""00"")" & vbCrLf
    s = s & "            mm = Format(p2, ""00"")" & vbCrLf
    s = s & "        ElseIf p2 > 12 Then" & vbCrLf
    s = s & "            mm = Format(p1, ""00"")" & vbCrLf
    s = s & "            dd = Format(p2, ""00"")" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            dd = Format(p1, ""00"")" & vbCrLf
    s = s & "            mm = Format(p2, ""00"")" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If Val(mm) > 12 Or Val(mm) < 1 Then" & vbCrLf
    s = s & "            FormatTgl = tgl" & vbCrLf
    s = s & "        ElseIf Val(dd) > 31 Or Val(dd) < 1 Then" & vbCrLf
    s = s & "            FormatTgl = tgl" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            FormatTgl = dd & ""/"" & mm & ""/"" & yy" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        FormatTgl = tgl" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub LoadRecord(ByVal rowNum As Long)" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "    FillDropdowns" & vbCrLf
    s = s & "    FillAlasanDropdowns" & vbCrLf
    s = s & "    mRow = rowNum" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    Me.Caption = ""EDIT - Row "" & rowNum" & vbCrLf
    s = s & "    txtKode.Value = GetByHeader(ws, rowNum, ""Kode Unik (copy value dari FPTK)"")" & vbCrLf
    s = s & "    txtPosisi.Value = GetByHeader(ws, rowNum, ""Posisi"")" & vbCrLf
    s = s & "    txtNama.Value = GetByHeader(ws, rowNum, ""Nama"")" & vbCrLf
    s = s & "    cboSumber.Value = GetByHeader(ws, rowNum, ""Sumber Sourcing"")" & vbCrLf
    s = s & "    cboModel.Value = GetByHeader(ws, rowNum, ""Model Rekrutmen (Lihat di Sheet Flow Map Model)"")" & vbCrLf
    s = s & "    txtRekruter.Value = GetByHeader(ws, rowNum, ""Rekruter"")" & vbCrLf
    s = s & "    txtHP.Value = GetByHeader(ws, rowNum, ""Nomor HP"")" & vbCrLf
    s = s & "    txtEmail.Value = GetByHeader(ws, rowNum, ""Email"")" & vbCrLf
    s = s & "    txtDomisili.Value = GetByHeader(ws, rowNum, ""Domisili"")" & vbCrLf
    s = s & "    cboUnivTop.Value = GetByHeader(ws, rowNum, ""Nama Universitas/Sekolah (TOP 10)"")" & vbCrLf
    s = s & "    txtUnivOther.Value = GetByHeader(ws, rowNum, ""Nama Universitas/Sekolah Lainnya"")" & vbCrLf
    s = s & "    cboJenjang.Value = GetByHeader(ws, rowNum, ""Jenjang Pendidikan"")" & vbCrLf
    s = s & "    cboJurusan.Value = GetByHeader(ws, rowNum, ""Jurusan"")" & vbCrLf
    s = s & "    txtTahunLulus.Value = GetByHeader(ws, rowNum, ""Tahun Lulus"")" & vbCrLf
    s = s & "    txtIPK.Value = GetByHeader(ws, rowNum, ""IPK"")" & vbCrLf
    s = s & "    txtEnglish.Value = GetByHeader(ws, rowNum, ""Skor Bahasa Inggris"")" & vbCrLf
    s = s & "    cboUnivTier.Value = GetByHeader(ws, rowNum, ""University Tier"")" & vbCrLf
    s = s & "    cboIPKTier.Value = GetByHeader(ws, rowNum, ""IPK Tier"")" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboSourcingHR, txtTglSourcingHR, cboAlasanSourcingHR, ""Sourcing HR"", ""Detail Keterangan Sourcing HR"", ""Tanggal Sourcing""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboShortlist, txtTglShortlist, cboAlasanShortlist, ""Shortlist CV"", ""Detail Keterangan Shortlist CV"", ""Tanggal Shortlist CV""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboPsikotes, txtTglPsikotes, cboAlasanPsikotes, ""Psikotes"", ""Detail Keterangan Psikotes"", ""Tanggal Psikotes / Cek psikotes""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboHRInterview, txtTglHRInterview, cboAlasanHRInterview, ""HR Interview"", ""Detail Keterangan HR Interview"", ""Tanggal HR Interview""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboTechnicalTest, txtTglTechnicalTest, cboAlasanTechnicalTest, ""Technical Test/ Case Study"", ""Detail Keterangan Technical Test/ Case Study"", ""Tanggal Technical Test/ Case Study""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboMarketVisit, txtTglMarketVisit, cboAlasanMarketVisit, ""Market Visit"", ""Detail Market Visit"", ""Tanggal Market Visit""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboUserInterview, txtTglUserInterview, cboAlasanUserInterview, ""User Interview"", ""Detail Keterangan User Interview"", ""Tanggal User Interview""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboPanelInterview, txtTglPanelInterview, cboAlasanPanelInterview, ""Panel Interview"", ""Detail Keterangan Panel Interview"", ""Tanggal Panel Interview""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboReferenceCheck, txtTglReferenceCheck, cboAlasanReferenceCheck, ""Reference Check"", ""Detail Keterangan Reference Check"", ""Tanggal Reference Check""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboMCU, txtTglMCU, cboAlasanMCU, ""MCU"", ""Detail Keterangan MCU"", ""Tanggal MCU""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboOffering, txtTglOffering, cboAlasanOffering, ""Offering"", ""Detail Keterangan Offering"", ""Tanggal Offering""" & vbCrLf
    s = s & "    LoadSC ws, rowNum, cboDay1, txtTglDay1, cboAlasanDay1, ""Day 1"", ""Detail Keterangan Day 1"", ""Tanggal Day 1""" & vbCrLf
    s = s & "    UpdateStatus" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadSC(ByVal ws As Worksheet, ByVal r As Long, ByRef cbo As MSForms.ComboBox, ByRef txtTgl As MSForms.TextBox, ByRef cboAlasan As MSForms.ComboBox, ByVal hS As String, ByVal hD As String, ByVal hT As String)" & vbCrLf
    s = s & "    Dim sv As String" & vbCrLf
    s = s & "    sv = Trim(CStr(GetByHeader(ws, r, hS)))" & vbCrLf
    s = s & "    If InStr(1, sv, ""V"") > 0 Or InStr(1, LCase(sv), ""lolos"") > 0 Then" & vbCrLf
    s = s & "        cbo.Value = ""Lolos""" & vbCrLf
    s = s & "    ElseIf InStr(1, LCase(sv), ""x"") > 0 Or InStr(1, LCase(sv), ""tidak"") > 0 Then" & vbCrLf
    s = s & "        cbo.Value = ""Tidak Lolos""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        cbo.Value = ""Belum""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    txtTgl.Value = FormatTgl(CStr(GetByHeader(ws, r, hT)))" & vbCrLf
    s = s & "    cboAlasan.Value = CStr(GetByHeader(ws, r, hD))" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateStatus()" & vbCrLf
    s = s & "    Dim st As String" & vbCrLf
    s = s & "    Dim bg As Long" & vbCrLf
    s = s & "    If cboDay1.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Day 1 | "" & txtTglDay1.Value" & vbCrLf
    s = s & "        bg = RGB(39, 174, 96)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf cboDay1.Value = ""Tidak Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Tidak Lolos Day 1 | "" & cboAlasanDay1.Value" & vbCrLf
    s = s & "        bg = RGB(231, 76, 60)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf cboOffering.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Offering | "" & txtTglOffering.Value" & vbCrLf
    s = s & "        bg = RGB(39, 174, 96)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf cboOffering.Value = ""Tidak Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Tidak Lolos Offering | "" & cboAlasanOffering.Value" & vbCrLf
    s = s & "        bg = RGB(231, 76, 60)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf cboMCU.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos MCU | "" & txtTglMCU.Value" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf cboMCU.Value = ""Tidak Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Tidak Lolos MCU | "" & cboAlasanMCU.Value" & vbCrLf
    s = s & "        bg = RGB(231, 76, 60)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf cboReferenceCheck.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Reference Check | "" & txtTglReferenceCheck.Value" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf cboReferenceCheck.Value = ""Tidak Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Tidak Lolos Reference Check | "" & cboAlasanReferenceCheck.Value" & vbCrLf
    s = s & "        bg = RGB(231, 76, 60)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf cboPanelInterview.Value = ""Lolos"" Or cboUserInterview.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos User/Panel Interview | "" & txtTglUserInterview.Value" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf cboPanelInterview.Value = ""Tidak Lolos"" Or cboUserInterview.Value = ""Tidak Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Tidak Lolos User/Panel Interview | "" & cboAlasanUserInterview.Value" & vbCrLf
    s = s & "        bg = RGB(231, 76, 60)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = vbWhite" & vbCrLf
    s = s & "    ElseIf cboMarketVisit.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Market Visit | "" & txtTglMarketVisit.Value" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf cboTechnicalTest.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Technical Test | "" & txtTglTechnicalTest.Value" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf cboHRInterview.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos HR Interview | "" & txtTglHRInterview.Value" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf cboPsikotes.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Psikotes | "" & txtTglPsikotes.Value" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf cboShortlist.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Shortlist | "" & txtTglShortlist.Value" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    ElseIf cboSourcingHR.Value = ""Lolos"" Then" & vbCrLf
    s = s & "        st = ""STATUS: Lolos Sourcing HR | "" & txtTglSourcingHR.Value" & vbCrLf
    s = s & "        bg = RGB(255, 193, 7)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        st = ""STATUS: Menunggu proses Sourcing HR""" & vbCrLf
    s = s & "        bg = RGB(236, 240, 241)" & vbCrLf
    s = s & "        lblStatusTerakhir.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lblStatusTerakhir.Caption = st" & vbCrLf
    s = s & "    lblStatusTerakhir.BackColor = bg" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    If mRow = 0 Then Exit Sub" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Nama"", txtNama.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Sumber Sourcing"", cboSumber.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Nomor HP"", txtHP.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Email"", txtEmail.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Domisili"", txtDomisili.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Nama Universitas/Sekolah (TOP 10)"", cboUnivTop.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Nama Universitas/Sekolah Lainnya"", txtUnivOther.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Jenjang Pendidikan"", cboJenjang.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Jurusan"", cboJurusan.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Tahun Lulus"", txtTahunLulus.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""IPK"", txtIPK.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""Skor Bahasa Inggris"", txtEnglish.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""University Tier"", cboUnivTier.Value" & vbCrLf
    s = s & "    SetByHeader ws, mRow, ""IPK Tier"", cboIPKTier.Value" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboSourcingHR, txtTglSourcingHR, cboAlasanSourcingHR, ""Sourcing HR"", ""Detail Keterangan Sourcing HR"", ""Tanggal Sourcing""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboShortlist, txtTglShortlist, cboAlasanShortlist, ""Shortlist CV"", ""Detail Keterangan Shortlist CV"", ""Tanggal Shortlist CV""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboPsikotes, txtTglPsikotes, cboAlasanPsikotes, ""Psikotes"", ""Detail Keterangan Psikotes"", ""Tanggal Psikotes / Cek psikotes""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboHRInterview, txtTglHRInterview, cboAlasanHRInterview, ""HR Interview"", ""Detail Keterangan HR Interview"", ""Tanggal HR Interview""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboTechnicalTest, txtTglTechnicalTest, cboAlasanTechnicalTest, ""Technical Test/ Case Study"", ""Detail Keterangan Technical Test/ Case Study"", ""Tanggal Technical Test/ Case Study""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboMarketVisit, txtTglMarketVisit, cboAlasanMarketVisit, ""Market Visit"", ""Detail Market Visit"", ""Tanggal Market Visit""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboUserInterview, txtTglUserInterview, cboAlasanUserInterview, ""User Interview"", ""Detail Keterangan User Interview"", ""Tanggal User Interview""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboPanelInterview, txtTglPanelInterview, cboAlasanPanelInterview, ""Panel Interview"", ""Detail Keterangan Panel Interview"", ""Tanggal Panel Interview""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboReferenceCheck, txtTglReferenceCheck, cboAlasanReferenceCheck, ""Reference Check"", ""Detail Keterangan Reference Check"", ""Tanggal Reference Check""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboMCU, txtTglMCU, cboAlasanMCU, ""MCU"", ""Detail Keterangan MCU"", ""Tanggal MCU""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboOffering, txtTglOffering, cboAlasanOffering, ""Offering"", ""Detail Keterangan Offering"", ""Tanggal Offering""" & vbCrLf
    s = s & "    SaveSC ws, mRow, cboDay1, txtTglDay1, cboAlasanDay1, ""Day 1"", ""Detail Keterangan Day 1"", ""Tanggal Day 1""" & vbCrLf
    s = s & "    ws.Columns(""A:BM"").AutoFit" & vbCrLf
    s = s & "    MsgBox ""Data berhasil diupdate!"", vbInformation" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SaveSC(ByVal ws As Worksheet, ByVal r As Long, ByVal cbo As MSForms.ComboBox, ByVal txtTgl As MSForms.TextBox, ByVal cboAlasan As MSForms.ComboBox, ByVal hS As String, ByVal hD As String, ByVal hT As String)" & vbCrLf
    s = s & "    Select Case cbo.Value" & vbCrLf
    s = s & "        Case ""Lolos""" & vbCrLf
    s = s & "            SetByHeader ws, r, hS, ""V""" & vbCrLf
    s = s & "        Case ""Tidak Lolos""" & vbCrLf
    s = s & "            SetByHeader ws, r, hS, ""X""" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            SetByHeader ws, r, hS, """"" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "    SetByHeader ws, r, hD, cboAlasan.Value" & vbCrLf
    s = s & "    SetByHeader ws, r, hT, FormatTgl(txtTgl.Value)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ApplyManualText(ByVal rawText As String, Optional ByVal sourceLabel As String = ""Manual Paste - Edit"")" & vbCrLf
    s = s & "    If Len(Trim(rawText)) = 0 Then Exit Sub" & vbCrLf
    s = s & "    SaveCVRawText sourceLabel, rawText" & vbCrLf
    s = s & "    Dim info As Variant" & vbCrLf
    s = s & "    info = ParseCVTextSmart(rawText)" & vbCrLf
    s = s & "    txtNama.Value = """"" & vbCrLf
    s = s & "    txtEmail.Value = """"" & vbCrLf
    s = s & "    txtHP.Value = """"" & vbCrLf
    s = s & "    cboUnivTop.Value = """"" & vbCrLf
    s = s & "    txtUnivOther.Value = """"" & vbCrLf
    s = s & "    cboJurusan.Value = """"" & vbCrLf
    s = s & "    txtIPK.Value = """"" & vbCrLf
    s = s & "    txtTahunLulus.Value = """"" & vbCrLf
    s = s & "    txtEnglish.Value = """"" & vbCrLf
    s = s & "    txtDomisili.Value = """"" & vbCrLf
    s = s & "    cboJenjang.Value = """"" & vbCrLf
    s = s & "    cboUnivTier.Value = """"" & vbCrLf
    s = s & "    cboIPKTier.Value = """"" & vbCrLf
    s = s & "    If Len(info(0)) > 0 Then txtNama.Value = info(0)" & vbCrLf
    s = s & "    If Len(info(1)) > 0 Then txtEmail.Value = info(1)" & vbCrLf
    s = s & "    If Len(info(2)) > 0 Then txtHP.Value = info(2)" & vbCrLf
    s = s & "    If Len(info(3)) > 0 Then SetUnivVal CStr(info(3))" & vbCrLf
    s = s & "    If Len(info(4)) > 0 Then SetCboVal cboJurusan, CStr(info(4))" & vbCrLf
    s = s & "    If Len(info(5)) > 0 Then txtIPK.Value = info(5)" & vbCrLf
    s = s & "    If Len(info(6)) > 0 Then txtTahunLulus.Value = info(6)" & vbCrLf
    s = s & "    If Len(info(7)) > 0 Then txtEnglish.Value = info(7)" & vbCrLf
    s = s & "    If UBound(info) >= 8 Then If Len(info(8)) > 0 Then txtDomisili.Value = info(8)" & vbCrLf
    s = s & "    If UBound(info) >= 9 Then If Len(info(9)) > 0 Then SetCboVal cboJenjang, CStr(info(9))" & vbCrLf
    s = s & "    If UBound(info) >= 10 Then If Len(info(10)) > 0 Then cboUnivTier.Value = info(10)" & vbCrLf
    s = s & "    If UBound(info) >= 11 Then If Len(info(11)) > 0 Then cboIPKTier.Value = info(11)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetUnivVal(ByVal univName As String)" & vbCrLf
    s = s & "    If CboHas(cboUnivTop, univName) Then" & vbCrLf
    s = s & "        cboUnivTop.Value = univName" & vbCrLf
    s = s & "        txtUnivOther.Value = """"" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        cboUnivTop.Value = ""Lainnya""" & vbCrLf
    s = s & "        txtUnivOther.Value = univName" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Sub SetCboVal(ByVal cbo As MSForms.ComboBox, ByVal val As String)" & vbCrLf
    s = s & "    If Len(val) = 0 Then Exit Sub" & vbCrLf
    s = s & "    If Not CboHas(cbo, val) Then cbo.AddItem val" & vbCrLf
    s = s & "    cbo.Value = val" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "Private Function CboHas(ByVal cbo As MSForms.ComboBox, ByVal val As String) As Boolean" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    For i = 0 To cbo.ListCount - 1" & vbCrLf
    s = s & "        If LCase(Trim(cbo.List(i))) = LCase(Trim(val)) Then" & vbCrLf
    s = s & "            CboHas = True" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
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
    s = s & "    If b Then" & vbCrLf
    s = s & "        o.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set AddL = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddT(ByVal p As Object, ByVal n As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal lk As Boolean = False) As MSForms.TextBox" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.TextBox.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l" & vbCrLf
    s = s & "    o.Top = t" & vbCrLf
    s = s & "    o.Width = w" & vbCrLf
    s = s & "    o.Height = h" & vbCrLf
    s = s & "    o.Font.Size = 8" & vbCrLf
    s = s & "    If lk Then" & vbCrLf
    s = s & "        o.Locked = True" & vbCrLf
    s = s & "        o.BackColor = RGB(236, 240, 241)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Set AddT = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
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
    s = s & "" & vbCrLf
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
    s = s & "" & vbCrLf
    s = s & "Private Sub AddSection(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single)" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.Label.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l" & vbCrLf
    s = s & "    o.Top = t" & vbCrLf
    s = s & "    o.Width = w" & vbCrLf
    s = s & "    o.Height = 14" & vbCrLf
    s = s & "    o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = 7" & vbCrLf
    s = s & "    o.Font.Bold = True" & vbCrLf
    s = s & "    o.BackColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    o.ForeColor = vbWhite" & vbCrLf
    s = s & "    o.BackStyle = 1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsValidTgl(ByVal tgl As String) As Boolean" & vbCrLf
    s = s & "    If Len(Trim(tgl)) = 0 Then" & vbCrLf
    s = s & "        IsValidTgl = True" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim d As String" & vbCrLf
    s = s & "    d = Replace(Trim(tgl), ""-"", ""/"")" & vbCrLf
    s = s & "    d = Replace(d, ""."", ""/"")" & vbCrLf
    s = s & "    d = Replace(d, "" "", ""/"")" & vbCrLf
    s = s & "    Dim parts() As String" & vbCrLf
    s = s & "    parts = Split(d, ""/"")" & vbCrLf
    s = s & "    If UBound(parts) >= 2 Then" & vbCrLf
    s = s & "        Dim p1 As Long" & vbCrLf
    s = s & "        p1 = Val(parts(0))" & vbCrLf
    s = s & "        Dim p2 As Long" & vbCrLf
    s = s & "        p2 = Val(parts(1))" & vbCrLf
    s = s & "        Dim mm As Long" & vbCrLf
    s = s & "        If p1 > 12 Then" & vbCrLf
    s = s & "            mm = p2" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            mm = p1" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If mm > 12 Or mm < 1 Then" & vbCrLf
    s = s & "            IsValidTgl = False" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If p1 > 31 Or p2 > 31 Then" & vbCrLf
    s = s & "            IsValidTgl = False" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    IsValidTgl = True" & vbCrLf
    s = s & "End Function" & vbCrLf
    Code_frmEditSourcing = s
End Function

Private Function Code_ModSourcingSupport() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function WSExists(ByVal sheetName As String) As Boolean" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    WSExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing" & vbCrLf
    s = s & "    Err.Clear" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet" & vbCrLf
    s = s & "    If WSExists(sheetName) Then" & vbCrLf
    s = s & "        Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))" & vbCrLf
    s = s & "        GetOrCreateSheet.Name = sheetName" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub InitSourcingCV()" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    EnsureCVRawText" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub EnsureDBSourcing()" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = GetOrCreateSheet(""DB Sourcing"")" & vbCrLf
    s = s & "    If Len(ws.Cells(1, 1).Value) = 0 Then" & vbCrLf
    s = s & "        SetHeader ws, 1, ""No""" & vbCrLf
    s = s & "        SetHeader ws, 2, ""Sourcing Date""" & vbCrLf
    s = s & "        SetHeader ws, 3, ""Kode Unik (copy value dari FPTK)""" & vbCrLf
    s = s & "        SetHeader ws, 4, ""Posisi""" & vbCrLf
    s = s & "        SetHeader ws, 5, ""Model Rekrutmen (Lihat di Sheet Flow Map Model)""" & vbCrLf
    s = s & "        SetHeader ws, 6, ""Rekruter""" & vbCrLf
    s = s & "        SetHeader ws, 7, ""Sumber Sourcing""" & vbCrLf
    s = s & "        SetHeader ws, 8, ""Nama""" & vbCrLf
    s = s & "        SetHeader ws, 9, ""Nama Universitas/Sekolah (TOP 10)""" & vbCrLf
    s = s & "        SetHeader ws, 10, ""Nama Universitas/Sekolah Lainnya""" & vbCrLf
    s = s & "        SetHeader ws, 11, ""Jenjang Pendidikan""" & vbCrLf
    s = s & "        SetHeader ws, 12, ""Jurusan""" & vbCrLf
    s = s & "        SetHeader ws, 13, ""Tahun Lulus""" & vbCrLf
    s = s & "        SetHeader ws, 14, ""IPK""" & vbCrLf
    s = s & "        SetHeader ws, 15, ""Skor Bahasa Inggris""" & vbCrLf
    s = s & "        SetHeader ws, 16, ""University Tier""" & vbCrLf
    s = s & "        SetHeader ws, 17, ""IPK Tier""" & vbCrLf
    s = s & "        SetHeader ws, 18, ""Nomor HP""" & vbCrLf
    s = s & "        SetHeader ws, 19, ""Email""" & vbCrLf
    s = s & "        SetHeader ws, 20, ""Domisili""" & vbCrLf
    s = s & "        SetHeader ws, 21, ""Sourcin Freelance""" & vbCrLf
    s = s & "        SetHeader ws, 22, ""Tanggal Sourcing Freelance""" & vbCrLf
    s = s & "        SetHeader ws, 23, ""Sourcing HR""" & vbCrLf
    s = s & "        SetHeader ws, 24, ""Detail Keterangan Sourcing HR""" & vbCrLf
    s = s & "        SetHeader ws, 25, ""Tanggal Sourcing""" & vbCrLf
    s = s & "        SetHeader ws, 26, ""Shortlist CV""" & vbCrLf
    s = s & "        SetHeader ws, 27, ""Detail Keterangan Shortlist CV""" & vbCrLf
    s = s & "        SetHeader ws, 28, ""Tanggal Shortlist CV""" & vbCrLf
    s = s & "        SetHeader ws, 29, ""Psikotes""" & vbCrLf
    s = s & "        SetHeader ws, 30, ""Kode Psikotes""" & vbCrLf
    s = s & "        SetHeader ws, 31, ""Detail Keterangan Psikotes""" & vbCrLf
    s = s & "        SetHeader ws, 32, ""Tanggal Psikotes / Cek psikotes""" & vbCrLf
    s = s & "        SetHeader ws, 33, ""Nilai Logika""" & vbCrLf
    s = s & "        SetHeader ws, 34, ""Nilai IQ""" & vbCrLf
    s = s & "        SetHeader ws, 35, ""Nilai Daya Tangkap""" & vbCrLf
    s = s & "        SetHeader ws, 36, ""Nilai RA""" & vbCrLf
    s = s & "        SetHeader ws, 37, ""DISC""" & vbCrLf
    s = s & "        SetHeader ws, 38, ""HR Interview""" & vbCrLf
    s = s & "        SetHeader ws, 39, ""Detail Keterangan HR Interview""" & vbCrLf
    s = s & "        SetHeader ws, 40, ""Tanggal HR Interview""" & vbCrLf
    s = s & "        SetHeader ws, 41, ""Technical Test/ Case Study""" & vbCrLf
    s = s & "        SetHeader ws, 42, ""Detail Keterangan Technical Test/ Case Study""" & vbCrLf
    s = s & "        SetHeader ws, 43, ""Tanggal Technical Test/ Case Study""" & vbCrLf
    s = s & "        SetHeader ws, 44, ""Market Visit""" & vbCrLf
    s = s & "        SetHeader ws, 45, ""Detail Market Visit""" & vbCrLf
    s = s & "        SetHeader ws, 46, ""Tanggal Market Visit""" & vbCrLf
    s = s & "        SetHeader ws, 47, ""User Interview""" & vbCrLf
    s = s & "        SetHeader ws, 48, ""Detail Keterangan User Interview""" & vbCrLf
    s = s & "        SetHeader ws, 49, ""Tanggal User Interview""" & vbCrLf
    s = s & "        SetHeader ws, 50, ""Panel Interview""" & vbCrLf
    s = s & "        SetHeader ws, 51, ""Detail Keterangan Panel Interview""" & vbCrLf
    s = s & "        SetHeader ws, 52, ""Tanggal Panel Interview""" & vbCrLf
    s = s & "        SetHeader ws, 53, ""Reference Check""" & vbCrLf
    s = s & "        SetHeader ws, 54, ""Detail Keterangan Reference Check""" & vbCrLf
    s = s & "        SetHeader ws, 55, ""Tanggal Reference Check""" & vbCrLf
    s = s & "        SetHeader ws, 56, ""MCU""" & vbCrLf
    s = s & "        SetHeader ws, 57, ""Detail Keterangan MCU""" & vbCrLf
    s = s & "        SetHeader ws, 58, ""Tanggal MCU""" & vbCrLf
    s = s & "        SetHeader ws, 59, ""Offering""" & vbCrLf
    s = s & "        SetHeader ws, 60, ""Detail Keterangan Offering""" & vbCrLf
    s = s & "        SetHeader ws, 61, ""Tanggal Offering""" & vbCrLf
    s = s & "        SetHeader ws, 62, ""Notes""" & vbCrLf
    s = s & "        SetHeader ws, 63, ""Day 1""" & vbCrLf
    s = s & "        SetHeader ws, 64, ""Detail Keterangan Day 1""" & vbCrLf
    s = s & "        SetHeader ws, 65, ""Tanggal Day 1""" & vbCrLf
    s = s & "        FormatHeader ws, 65" & vbCrLf
    s = s & "        ws.Columns(""A:BM"").AutoFit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub EnsureCVRawText()" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = GetOrCreateSheet(""CV Raw Text"")" & vbCrLf
    s = s & "    If Len(ws.Cells(1, 1).Value) = 0 Then" & vbCrLf
    s = s & "        SetHeader ws, 1, ""Timestamp""" & vbCrLf
    s = s & "        SetHeader ws, 2, ""CV Path""" & vbCrLf
    s = s & "        SetHeader ws, 3, ""File Name""" & vbCrLf
    s = s & "        SetHeader ws, 4, ""Raw Text Length""" & vbCrLf
    s = s & "        SetHeader ws, 5, ""Raw Text""" & vbCrLf
    s = s & "        FormatHeader ws, 5" & vbCrLf
    s = s & "        ws.Columns(""A:E"").AutoFit" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub SetHeader(ByVal ws As Worksheet, ByVal colNum As Long, ByVal headerText As String)" & vbCrLf
    s = s & "    ws.Cells(1, colNum).Value = headerText" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub FormatHeader(ByVal ws As Worksheet, ByVal colCount As Long)" & vbCrLf
    s = s & "    With ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))" & vbCrLf
    s = s & "        .Font.Bold = True" & vbCrLf
    s = s & "        .Interior.Color = RGB(44, 62, 80)" & vbCrLf
    s = s & "        .Font.Color = vbWhite" & vbCrLf
    s = s & "        .WrapText = True" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "    ws.Rows(1).RowHeight = 32" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long" & vbCrLf
    s = s & "    Dim lastCol As Long, c As Long" & vbCrLf
    s = s & "    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column" & vbCrLf
    s = s & "    For c = 1 To lastCol" & vbCrLf
    s = s & "        If Trim(LCase(CStr(ws.Cells(1, c).Value))) = Trim(LCase(headerText)) Then" & vbCrLf
    s = s & "            FindHeaderColumn = c" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub SetByHeader(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal headerText As String, ByVal valueText As Variant)" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    c = FindHeaderColumn(ws, headerText)" & vbCrLf
    s = s & "    If c > 0 Then ws.Cells(rowNum, c).Value = valueText" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function GetByHeader(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal headerText As String) As Variant" & vbCrLf
    s = s & "    Dim c As Long" & vbCrLf
    s = s & "    c = FindHeaderColumn(ws, headerText)" & vbCrLf
    s = s & "    If c > 0 Then" & vbCrLf
    s = s & "        GetByHeader = ws.Cells(rowNum, c).Value" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        GetByHeader = """"" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function GetFPTKWorksheet() As Worksheet" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    If WSExists(""FPTK"") Then" & vbCrLf
    s = s & "        Set GetFPTKWorksheet = ThisWorkbook.Worksheets(""FPTK"")" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    For Each ws In ThisWorkbook.Worksheets" & vbCrLf
    s = s & "        If FindHeaderColumn(ws, ""Kode Unik"") > 0 Then" & vbCrLf
    s = s & "            If FindHeaderColumn(ws, ""Posisi - Kebutuhan TA"") > 0 Then" & vbCrLf
    s = s & "                Set GetFPTKWorksheet = ws" & vbCrLf
    s = s & "                Exit Function" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next ws" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function SearchFPTK(ByVal keyword As String) As Collection" & vbCrLf
    s = s & "    Dim c As New Collection" & vbCrLf
    s = s & "    Dim ws As Worksheet, lastRow As Long, r As Long" & vbCrLf
    s = s & "    Dim colKode As Long, colPosisi As Long, colStatus As Long, colRecruiter As Long" & vbCrLf
    s = s & "    Set ws = GetFPTKWorksheet()" & vbCrLf
    s = s & "    If ws Is Nothing Then Set SearchFPTK = c: Exit Function" & vbCrLf
    s = s & "    colKode = FindHeaderColumn(ws, ""Kode Unik"")" & vbCrLf
    s = s & "    colPosisi = FindHeaderColumn(ws, ""Posisi - Kebutuhan TA"")" & vbCrLf
    s = s & "    colStatus = FindHeaderColumn(ws, ""Status"")" & vbCrLf
    s = s & "    colRecruiter = FindHeaderColumn(ws, ""PIC Recruiter"")" & vbCrLf
    s = s & "    If colKode = 0 Or colPosisi = 0 Then Set SearchFPTK = c: Exit Function" & vbCrLf
    s = s & "    keyword = LCase(Trim(keyword))" & vbCrLf
    s = s & "    lastRow = ws.Cells(ws.Rows.Count, colKode).End(xlUp).Row" & vbCrLf
    s = s & "    Dim k As String, p As String, st As String, rc As String" & vbCrLf
    s = s & "    Dim md As String, hs As String" & vbCrLf
    s = s & "    For r = 2 To lastRow" & vbCrLf
    s = s & "        k = CStr(ws.Cells(r, colKode).Value)" & vbCrLf
    s = s & "        p = CStr(ws.Cells(r, colPosisi).Value)" & vbCrLf
    s = s & "        st = """"" & vbCrLf
    s = s & "        If colStatus > 0 Then st = CStr(ws.Cells(r, colStatus).Value)" & vbCrLf
    s = s & "        rc = """"" & vbCrLf
    s = s & "        If colRecruiter > 0 Then rc = CStr(ws.Cells(r, colRecruiter).Value)" & vbCrLf
    s = s & "        md = DetermineModelRekrutmen(p, """")" & vbCrLf
    s = s & "        hs = LCase(k & "" "" & p & "" "" & st & "" "" & rc & "" "" & md)" & vbCrLf
    s = s & "        If keyword = """" Or InStr(1, hs, keyword) > 0 Then" & vbCrLf
    s = s & "            c.Add k & ""|"" & p & ""|"" & st & ""|"" & rc & ""|"" & md" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    Set SearchFPTK = c" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function GetFPTKDetail(ByVal kodeUnik As String) As Variant" & vbCrLf
    s = s & "    Dim arr(4) As Variant" & vbCrLf
    s = s & "    Dim ws As Worksheet, lastRow As Long, r As Long" & vbCrLf
    s = s & "    Dim colKode As Long, colPosisi As Long, colStatus As Long, colRecruiter As Long" & vbCrLf
    s = s & "    Set ws = GetFPTKWorksheet()" & vbCrLf
    s = s & "    If ws Is Nothing Then GetFPTKDetail = arr: Exit Function" & vbCrLf
    s = s & "    colKode = FindHeaderColumn(ws, ""Kode Unik"")" & vbCrLf
    s = s & "    colPosisi = FindHeaderColumn(ws, ""Posisi - Kebutuhan TA"")" & vbCrLf
    s = s & "    colStatus = FindHeaderColumn(ws, ""Status"")" & vbCrLf
    s = s & "    colRecruiter = FindHeaderColumn(ws, ""PIC Recruiter"")" & vbCrLf
    s = s & "    If colKode = 0 Or colPosisi = 0 Then GetFPTKDetail = arr: Exit Function" & vbCrLf
    s = s & "    lastRow = ws.Cells(ws.Rows.Count, colKode).End(xlUp).Row" & vbCrLf
    s = s & "    For r = 2 To lastRow" & vbCrLf
    s = s & "        If CStr(ws.Cells(r, colKode).Value) = kodeUnik Then" & vbCrLf
    s = s & "            arr(0) = ws.Cells(r, colKode).Value" & vbCrLf
    s = s & "            arr(1) = ws.Cells(r, colPosisi).Value" & vbCrLf
    s = s & "            arr(2) = """"" & vbCrLf
    s = s & "            If colStatus > 0 Then arr(2) = ws.Cells(r, colStatus).Value" & vbCrLf
    s = s & "            arr(3) = """"" & vbCrLf
    s = s & "            If colRecruiter > 0 Then arr(3) = ws.Cells(r, colRecruiter).Value" & vbCrLf
    s = s & "            arr(4) = DetermineModelRekrutmen(CStr(arr(1)), """")" & vbCrLf
    s = s & "            GetFPTKDetail = arr" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    GetFPTKDetail = arr" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function DetermineModelRekrutmen(ByVal posisi As String, Optional ByVal existingModel As Variant) As String" & vbCrLf
    s = s & "    Dim mapModel As String" & vbCrLf
    s = s & "    If Len(Trim(CStr(existingModel))) > 0 Then" & vbCrLf
    s = s & "        DetermineModelRekrutmen = CStr(existingModel)" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mapModel = GetModelFromFlowMap(posisi)" & vbCrLf
    s = s & "    If Len(mapModel) > 0 Then" & vbCrLf
    s = s & "        DetermineModelRekrutmen = mapModel" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    posisi = LCase(posisi)" & vbCrLf
    s = s & "    If InStr(posisi, ""fgdp"") > 0 Then DetermineModelRekrutmen = ""Model 1"": Exit Function" & vbCrLf
    s = s & "    If InStr(posisi, ""clap"") > 0 Then DetermineModelRekrutmen = ""Model 2"": Exit Function" & vbCrLf
    s = s & "    If InStr(posisi, ""sales taking order"") > 0 Or InStr(posisi, ""sto"") > 0 Then DetermineModelRekrutmen = ""Model 4"": Exit Function" & vbCrLf
    s = s & "    If InStr(posisi, ""area sales supervisor"") > 0 Or InStr(posisi, ""ass"") > 0 Then DetermineModelRekrutmen = ""Model 3"": Exit Function" & vbCrLf
    s = s & "    If InStr(posisi, ""manager"") > 0 Then DetermineModelRekrutmen = ""Model 2"": Exit Function" & vbCrLf
    s = s & "    DetermineModelRekrutmen = ""Model 1""" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function GetModelFromFlowMap(ByVal posisi As String) As String" & vbCrLf
    s = s & "    On Error GoTo ErrHandler" & vbCrLf
    s = s & "    Dim ws As Worksheet, r As Long, lastRow As Long" & vbCrLf
    s = s & "    Dim colPos As Long, colModel As Long" & vbCrLf
    s = s & "    If Not WSExists(""Flow Map Model"") Then Exit Function" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""Flow Map Model"")" & vbCrLf
    s = s & "    colPos = FindHeaderColumn(ws, ""Posisi"")" & vbCrLf
    s = s & "    colModel = FindHeaderColumn(ws, ""Model Rekrutmen"")" & vbCrLf
    s = s & "    If colModel = 0 Then colModel = FindHeaderColumn(ws, ""Model"")" & vbCrLf
    s = s & "    If colPos = 0 Or colModel = 0 Then Exit Function" & vbCrLf
    s = s & "    lastRow = ws.Cells(ws.Rows.Count, colPos).End(xlUp).Row" & vbCrLf
    s = s & "    For r = 2 To lastRow" & vbCrLf
    s = s & "        If InStr(1, posisi, CStr(ws.Cells(r, colPos).Value), vbTextCompare) > 0 Then" & vbCrLf
    s = s & "            GetModelFromFlowMap = CStr(ws.Cells(r, colModel).Value)" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "ErrHandler:" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function PickFileCV() As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    With Application.FileDialog(3)" & vbCrLf
    s = s & "        .title = ""Pilih CV Kandidat""" & vbCrLf
    s = s & "        .Filters.Clear" & vbCrLf
    s = s & "        .Filters.Add ""CV Files"", ""*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.pdf""" & vbCrLf
    s = s & "        .AllowMultiSelect = False" & vbCrLf
    s = s & "        If .Show = -1 Then PickFileCV = CStr(.SelectedItems(1))" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function IsImageFile(ByVal filePath As String) As Boolean" & vbCrLf
    s = s & "    Dim ext As String" & vbCrLf
    s = s & "    ext = LCase(Mid(filePath, InStrRev(filePath, ""."") + 1))" & vbCrLf
    s = s & "    IsImageFile = (ext = ""jpg"" Or ext = ""jpeg"" Or ext = ""png"" Or ext = ""bmp"" Or ext = ""gif"")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function ShortFileName(ByVal filePath As String, Optional ByVal maxLen As Long = 35) As String" & vbCrLf
    s = s & "    Dim p As Long" & vbCrLf
    s = s & "    p = InStrRev(filePath, Application.PathSeparator)" & vbCrLf
    s = s & "    If p > 0 Then" & vbCrLf
    s = s & "        ShortFileName = Mid(filePath, p + 1)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ShortFileName = filePath" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Len(ShortFileName) > maxLen Then" & vbCrLf
    s = s & "        ShortFileName = Left(ShortFileName, maxLen - 3) & ""...""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub SaveCVRawText(ByVal filePath As String, ByVal rawText As String)" & vbCrLf
    s = s & "    EnsureCVRawText" & vbCrLf
    s = s & "    Dim ws As Worksheet, r As Long" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""CV Raw Text"")" & vbCrLf
    s = s & "    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    s = s & "    ws.Cells(r, 1).Value = Now" & vbCrLf
    s = s & "    ws.Cells(r, 2).Value = filePath" & vbCrLf
    s = s & "    ws.Cells(r, 3).Value = ShortFileName(filePath, 80)" & vbCrLf
    s = s & "    ws.Cells(r, 4).Value = Len(rawText)" & vbCrLf
    s = s & "    ws.Cells(r, 5).Value = rawText" & vbCrLf
    s = s & "    ws.Cells(r, 5).WrapText = False" & vbCrLf
    s = s & "    ws.Columns(""A:E"").AutoFit" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function ReadPDFAutoCopy(ByVal pdfPath As String) As String" & vbCrLf
    s = s & "    On Error GoTo ErrHandler" & vbCrLf
    s = s & "    Dim txt As String" & vbCrLf
    s = s & "    txt = ReadPDFAutoCopyFocused(pdfPath)" & vbCrLf
    s = s & "    If Len(Trim(txt)) < 50 Then" & vbCrLf
    s = s & "        MsgBox ""PDF tidak terbaca. Coba copy manual (Ctrl+A, Ctrl+C) lalu klik OK."", vbInformation" & vbCrLf
    s = s & "        txt = GetClipboardTextSafe()" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    ReadPDFAutoCopy = txt" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrHandler:" & vbCrLf
    s = s & "    ReadPDFAutoCopy = """"" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ReadPDFAutoCopyFocused(ByVal pdfPath As String) As String" & vbCrLf
    s = s & "    On Error GoTo ErrHandler" & vbCrLf
    s = s & "    Dim i As Long, clipText As String" & vbCrLf
    s = s & "    ClearClipboardText" & vbCrLf
    s = s & "    Shell ""cmd /c start """""""" "" & Chr(34) & pdfPath & Chr(34), vbNormalFocus" & vbCrLf
    s = s & "    Application.Wait Now + TimeValue(""00:00:04"")" & vbCrLf
    s = s & "    AppActivatePDFWindow pdfPath" & vbCrLf
    s = s & "    Application.Wait Now + TimeValue(""00:00:01"")" & vbCrLf
    s = s & "    Application.SendKeys ""^a"", True" & vbCrLf
    s = s & "    Application.Wait Now + TimeValue(""00:00:01"")" & vbCrLf
    s = s & "    Application.SendKeys ""^c"", True" & vbCrLf
    s = s & "    Application.Wait Now + TimeValue(""00:00:01"")" & vbCrLf
    s = s & "    For i = 1 To 6" & vbCrLf
    s = s & "        clipText = GetClipboardTextSafe()" & vbCrLf
    s = s & "        If Len(Trim(clipText)) > 20 Then Exit For" & vbCrLf
    s = s & "        AppActivatePDFWindow pdfPath" & vbCrLf
    s = s & "        Application.SendKeys ""^c"", True" & vbCrLf
    s = s & "        Application.Wait Now + TimeValue(""00:00:01"")" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    ReadPDFAutoCopyFocused = clipText" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrHandler:" & vbCrLf
    s = s & "    ReadPDFAutoCopyFocused = """"" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AppActivatePDFWindow(ByVal pdfPath As String) As Boolean" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim wsh As Object, fileName As String" & vbCrLf
    s = s & "    Set wsh = CreateObject(""WScript.Shell"")" & vbCrLf
    s = s & "    fileName = ShortFileName(pdfPath, 200)" & vbCrLf
    s = s & "    AppActivatePDFWindow = wsh.AppActivate(fileName)" & vbCrLf
    s = s & "    If AppActivatePDFWindow Then Exit Function" & vbCrLf
    s = s & "    AppActivatePDFWindow = wsh.AppActivate(""Adobe Acrobat"")" & vbCrLf
    s = s & "    If AppActivatePDFWindow Then Exit Function" & vbCrLf
    s = s & "    AppActivatePDFWindow = wsh.AppActivate(""Adobe Reader"")" & vbCrLf
    s = s & "    If AppActivatePDFWindow Then Exit Function" & vbCrLf
    s = s & "    AppActivatePDFWindow = wsh.AppActivate(""Acrobat"")" & vbCrLf
    s = s & "    If AppActivatePDFWindow Then Exit Function" & vbCrLf
    s = s & "    AppActivatePDFWindow = wsh.AppActivate(""Microsoft Edge"")" & vbCrLf
    s = s & "    If AppActivatePDFWindow Then Exit Function" & vbCrLf
    s = s & "    AppActivatePDFWindow = wsh.AppActivate(""Google Chrome"")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ClearClipboardText()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim dataObj As Object" & vbCrLf
    s = s & "    Set dataObj = CreateObject(""New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"")" & vbCrLf
    s = s & "    dataObj.SetText """"" & vbCrLf
    s = s & "    dataObj.PutInClipboard" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function GetClipboardTextSafe() As String" & vbCrLf
    s = s & "    On Error GoTo ErrHandler" & vbCrLf
    s = s & "    Dim dataObj As Object" & vbCrLf
    s = s & "    Set dataObj = CreateObject(""New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"")" & vbCrLf
    s = s & "    dataObj.GetFromClipboard" & vbCrLf
    s = s & "    GetClipboardTextSafe = dataObj.GetText" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "ErrHandler:" & vbCrLf
    s = s & "    GetClipboardTextSafe = """"" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function GetUniversityTierCV(ByVal univName As String) As String" & vbCrLf
    s = s & "    Dim u As String" & vbCrLf
    s = s & "    u = LCase(Trim(univName))" & vbCrLf
    s = s & "    If Len(u) = 0 Then Exit Function" & vbCrLf
    s = s & "    Select Case u" & vbCrLf
    s = s & "        Case ""universitas indonesia"", ""universitas gadjah mada"", ""institut teknologi bandung""" & vbCrLf
    s = s & "            GetUniversityTierCV = ""Top 3 PTN""" & vbCrLf
    s = s & "        Case ""universitas airlangga"", ""universitas padjadjaran"", ""universitas diponegoro"", ""universitas brawijaya"", ""institut pertanian bogor"", ""universitas sebelas maret"", ""institut teknologi sepuluh nopember""" & vbCrLf
    s = s & "            GetUniversityTierCV = ""Top 10 PTN""" & vbCrLf
    s = s & "        Case ""telkom university"", ""universitas bina nusantara"", ""universitas gunadarma""" & vbCrLf
    s = s & "            GetUniversityTierCV = ""Top 10 PTS""" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            GetUniversityTierCV = ""Lainnya""" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function GetIPKTierCV(ByVal ipkText As String) As String" & vbCrLf
    s = s & "    Dim gpa As Double" & vbCrLf
    s = s & "    If Len(Trim(ipkText)) = 0 Then Exit Function" & vbCrLf
    s = s & "    gpa = Val(Replace(Trim(ipkText), "","", "".""))" & vbCrLf
    s = s & "    If gpa >= 3.5 Then" & vbCrLf
    s = s & "        GetIPKTierCV = ""Lebih dari 3,5""" & vbCrLf
    s = s & "    ElseIf gpa >= 3.2 Then" & vbCrLf
    s = s & "        GetIPKTierCV = ""Lebih dari 3,2""" & vbCrLf
    s = s & "    ElseIf gpa > 0 Then" & vbCrLf
    s = s & "        GetIPKTierCV = ""Kurang dari 3,2""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function ParseCVTextSmart(ByVal rawText As String) As Variant" & vbCrLf
    s = s & "    Dim info(11) As String" & vbCrLf
    s = s & "    Dim lines() As String, line As String, nl As String, low As String" & vbCrLf
    s = s & "    Dim i As Long, j As Long, regex As Object, oneLine As String" & vbCrLf
    s = s & "    Set regex = CreateObject(""VBScript.RegExp"")" & vbCrLf
    s = s & "    regex.Global = True" & vbCrLf
    s = s & "    regex.IgnoreCase = True" & vbCrLf
    s = s & "    rawText = Replace(rawText, vbCr, vbLf)" & vbCrLf
    s = s & "    rawText = Replace(rawText, vbTab, vbLf)" & vbCrLf
    s = s & "    Do While InStr(rawText, vbLf & vbLf & vbLf) > 0" & vbCrLf
    s = s & "        rawText = Replace(rawText, vbLf & vbLf & vbLf, vbLf & vbLf)" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    If Len(Trim(rawText)) = 0 Then ParseCVTextSmart = info: Exit Function" & vbCrLf
    s = s & "    lines = Split(rawText, vbLf)" & vbCrLf
    s = s & "    oneLine = Replace(Replace(rawText, vbCr, "" ""), vbLf, "" | "")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 0 Nama" & vbCrLf
    s = s & "    For i = 0 To UBound(lines)" & vbCrLf
    s = s & "        line = Trim(lines(i))" & vbCrLf
    s = s & "        If i < UBound(lines) Then nl = Trim(lines(i + 1)) Else nl = """"" & vbCrLf
    s = s & "        low = LCase(line)" & vbCrLf
    s = s & "        If StartsWithAny(low, Array(""nama"", ""name"", ""full name"", ""candidate name"")) Then" & vbCrLf
    s = s & "            info(0) = CleanTextValue(ExtractAfterSep(line))" & vbCrLf
    s = s & "            If Len(info(0)) = 0 And LooksLikeName(nl) Then info(0) = nl" & vbCrLf
    s = s & "            If Len(info(0)) > 0 Then Exit For" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    If Len(info(0)) = 0 Then" & vbCrLf
    s = s & "        For i = 0 To Application.Min(UBound(lines), 35)" & vbCrLf
    s = s & "            line = CleanTextValue(Trim(lines(i)))" & vbCrLf
    s = s & "            If LooksLikeName(line) Then" & vbCrLf
    s = s & "                info(0) = line" & vbCrLf
    s = s & "                Exit For" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 1 Email" & vbCrLf
    s = s & "    regex.Pattern = ""[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}""" & vbCrLf
    s = s & "    If regex.Test(rawText) Then info(1) = regex.Execute(rawText)(0).Value" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 2 Nomor HP" & vbCrLf
    s = s & "    regex.Pattern = ""(\+?62|0)[\s\-\.\(\)]?[0-9]{2,4}[\s\-\.\(\)]?[0-9]{3,4}[\s\-\.\(\)]?[0-9]{3,5}""" & vbCrLf
    s = s & "    If regex.Test(rawText) Then info(2) = CleanPhone(regex.Execute(rawText)(0).Value)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 3 Universitas" & vbCrLf
    s = s & "    info(3) = NormalizeUniversity(rawText)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 4 Jurusan" & vbCrLf
    s = s & "    info(4) = NormalizeMajor(rawText)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 5 IPK" & vbCrLf
    s = s & "    regex.Pattern = ""(IPK|GPA|Grade Point Average)[^0-9]{0,40}([0-4][\.,]\d{1,2})""" & vbCrLf
    s = s & "    If regex.Test(oneLine) Then info(5) = Replace(regex.Execute(oneLine)(0).SubMatches(1), "","", ""."")" & vbCrLf
    s = s & "    If Len(info(5)) = 0 Then" & vbCrLf
    s = s & "        regex.Pattern = ""\b([0-4][\.,]\d{1,2})\s*/\s*4""" & vbCrLf
    s = s & "        If regex.Test(oneLine) Then info(5) = Replace(regex.Execute(oneLine)(0).SubMatches(0), "","", ""."")" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Len(info(5)) = 0 Then" & vbCrLf
    s = s & "        regex.Pattern = ""\b([0-4][\.,]\d{1,2})\b""" & vbCrLf
    s = s & "        If regex.Test(oneLine) Then info(5) = Replace(regex.Execute(oneLine)(0).SubMatches(0), "","", ""."")" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 6 Tahun Lulus" & vbCrLf
    s = s & "    regex.Pattern = ""\b(20[0-9]{2})\s*[-ï¿½ï¿½]\s*([A-Za-z]{3,9}\s*)?(20[0-9]{2})\b""" & vbCrLf
    s = s & "    If regex.Test(oneLine) Then info(6) = regex.Execute(oneLine)(0).SubMatches(2)" & vbCrLf
    s = s & "    If Len(info(6)) = 0 Then" & vbCrLf
    s = s & "        regex.Pattern = ""(Lulus|Graduated|Graduation|Tahun Lulus|Class of)[^0-9]*(20[0-9]{2})""" & vbCrLf
    s = s & "        If regex.Test(oneLine) Then info(6) = regex.Execute(oneLine)(0).SubMatches(1)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Len(info(6)) = 0 Then" & vbCrLf
    s = s & "        regex.Pattern = ""\b(20[0-9]{2})\b""" & vbCrLf
    s = s & "        If regex.Test(oneLine) Then" & vbCrLf
    s = s & "            Dim ms As Object" & vbCrLf
    s = s & "            Set ms = regex.Execute(oneLine)" & vbCrLf
    s = s & "            info(6) = ms(ms.Count - 1).SubMatches(0)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 7 Skor Bahasa Inggris" & vbCrLf
    s = s & "    regex.Pattern = ""(TOEFL|IELTS|TOEIC|English|Bahasa Inggris)[^0-9]*(\d{2,3}(\.\d)?)""" & vbCrLf
    s = s & "    If regex.Test(rawText) Then info(7) = regex.Execute(rawText)(0).SubMatches(1)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 8 Domisili" & vbCrLf
    s = s & "    For i = 0 To UBound(lines)" & vbCrLf
    s = s & "        line = Trim(lines(i))" & vbCrLf
    s = s & "        If i < UBound(lines) Then nl = Trim(lines(i + 1)) Else nl = """"" & vbCrLf
    s = s & "        low = LCase(line)" & vbCrLf
    s = s & "        If StartsWithAny(low, Array(""domisili"", ""location"", ""lokasi"", ""alamat"", ""address"", ""city"", ""kota"")) Then" & vbCrLf
    s = s & "            info(8) = CleanTextValue(ExtractAfterSep(line))" & vbCrLf
    s = s & "            If Len(info(8)) = 0 Then info(8) = CleanTextValue(nl)" & vbCrLf
    s = s & "            If Len(info(8)) > 0 Then Exit For" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    If Len(info(8)) = 0 Then" & vbCrLf
    s = s & "        Dim cities As Variant" & vbCrLf
    s = s & "        cities = Array(""Jakarta"", ""Bogor"", ""Depok"", ""Tangerang"", ""Bekasi"", ""Bandung"", ""Semarang"", ""Surabaya"", ""Yogyakarta"", ""Jogja"", ""Malang"", ""Medan"", ""Makassar"", ""Denpasar"", ""Bali"")" & vbCrLf
    s = s & "        For j = 0 To UBound(cities)" & vbCrLf
    s = s & "            If InStr(1, rawText, cities(j), vbTextCompare) > 0 Then" & vbCrLf
    s = s & "                info(8) = cities(j)" & vbCrLf
    s = s & "                Exit For" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "        Next j" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 9 Jenjang Pendidikan" & vbCrLf
    s = s & "    info(9) = ExtractEducationLevel(rawText)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 10 University Tier" & vbCrLf
    s = s & "    info(10) = GetUniversityTierCV(info(3))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' 11 IPK Tier" & vbCrLf
    s = s & "    info(11) = GetIPKTierCV(info(5))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ParseCVTextSmart = info" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function NormalizeUniversity(ByVal text As String) As String" & vbCrLf
    s = s & "    Dim t As String, i As Long" & vbCrLf
    s = s & "    t = LCase(Replace(Replace(Replace(text, vbCr, "" ""), vbLf, "" ""), vbTab, "" ""))" & vbCrLf
    s = s & "    Do While InStr(t, ""  "") > 0" & vbCrLf
    s = s & "        t = Replace(t, ""  "", "" "")" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    t = "" "" & Trim(t) & "" """ & vbCrLf
    s = s & "    If InStr(1, t, "" universitas indonesia "") > 0 Or InStr(1, t, "" university of indonesia "") > 0 Or InStr(1, t, "" indonesia university "") > 0 Then NormalizeUniversity = ""Universitas Indonesia"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" universitas gadjah mada "") > 0 Or InStr(1, t, "" gadjah mada university "") > 0 Or InStr(1, t, "" ugm "") > 0 Then NormalizeUniversity = ""Universitas Gadjah Mada"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" institut teknologi bandung "") > 0 Or InStr(1, t, "" bandung institute of technology "") > 0 Or InStr(1, t, "" itb "") > 0 Then NormalizeUniversity = ""Institut Teknologi Bandung"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" universitas airlangga "") > 0 Or InStr(1, t, "" airlangga university "") > 0 Or InStr(1, t, "" unair "") > 0 Then NormalizeUniversity = ""Universitas Airlangga"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" universitas padjadjaran "") > 0 Or InStr(1, t, "" padjadjaran university "") > 0 Or InStr(1, t, "" unpad "") > 0 Then NormalizeUniversity = ""Universitas Padjadjaran"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" universitas diponegoro "") > 0 Or InStr(1, t, "" diponegoro university "") > 0 Or InStr(1, t, "" diponegro university "") > 0 Or InStr(1, t, "" undip "") > 0 Then NormalizeUniversity = ""Universitas Diponegoro"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" universitas brawijaya "") > 0 Or InStr(1, t, "" brawijaya university "") > 0 Then NormalizeUniversity = ""Universitas Brawijaya"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" institut pertanian bogor "") > 0 Or InStr(1, t, "" bogor agricultural university "") > 0 Or InStr(1, t, "" ipb "") > 0 Then NormalizeUniversity = ""Institut Pertanian Bogor"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" universitas sebelas maret "") > 0 Or InStr(1, t, "" sebelas maret university "") > 0 Or InStr(1, t, "" uns "") > 0 Then NormalizeUniversity = ""Universitas Sebelas Maret"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" institut teknologi sepuluh nopember "") > 0 Or InStr(1, t, "" sepuluh nopember institute "") > 0 Or InStr(1, t, "" its "") > 0 Then NormalizeUniversity = ""Institut Teknologi Sepuluh Nopember"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" telkom university "") > 0 Or InStr(1, t, "" universitas telkom "") > 0 Then NormalizeUniversity = ""Telkom University"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" universitas bina nusantara "") > 0 Or InStr(1, t, "" bina nusantara university "") > 0 Or InStr(1, t, "" binus "") > 0 Then NormalizeUniversity = ""Universitas Bina Nusantara"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, "" universitas gunadarma "") > 0 Or InStr(1, t, "" gunadarma university "") > 0 Then NormalizeUniversity = ""Universitas Gunadarma"": Exit Function" & vbCrLf
    s = s & "    Dim lines2() As String, ln As String" & vbCrLf
    s = s & "    lines2 = Split(text, vbLf)" & vbCrLf
    s = s & "    For i = 0 To UBound(lines2)" & vbCrLf
    s = s & "        ln = Trim(lines2(i))" & vbCrLf
    s = s & "        If StartsWithAny(LCase(ln), Array(""universitas"", ""institut"", ""politeknik"", ""sekolah tinggi"", ""smk "", ""sma "")) Then" & vbCrLf
    s = s & "            NormalizeUniversity = ln" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function NormalizeMajor(ByVal text As String) As String" & vbCrLf
    s = s & "    Dim t As String" & vbCrLf
    s = s & "    t = LCase(Replace(Replace(Replace(text, vbCr, "" ""), vbLf, "" ""), vbTab, "" ""))" & vbCrLf
    s = s & "    Do While InStr(t, ""  "") > 0" & vbCrLf
    s = s & "        t = Replace(t, ""  "", "" "")" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    If InStr(1, t, ""teknik industri"") > 0 Or InStr(1, t, ""industrial engineering"") > 0 Then NormalizeMajor = ""Teknik Industri"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""sistem informasi"") > 0 Or InStr(1, t, ""information system"") > 0 Then NormalizeMajor = ""Sistem Informasi"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""teknik informatika"") > 0 Or InStr(1, t, ""informatics engineering"") > 0 Or InStr(1, t, ""computer science"") > 0 Then NormalizeMajor = ""Teknik Informatika"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""teknik mesin"") > 0 Or InStr(1, t, ""mechanical engineering"") > 0 Then NormalizeMajor = ""Teknik Mesin"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""teknik elektro"") > 0 Or InStr(1, t, ""electrical engineering"") > 0 Then NormalizeMajor = ""Teknik Elektro"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""teknik kimia"") > 0 Or InStr(1, t, ""chemical engineering"") > 0 Then NormalizeMajor = ""Teknik Kimia"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""teknik sipil"") > 0 Or InStr(1, t, ""civil engineering"") > 0 Then NormalizeMajor = ""Teknik Sipil"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""akuntansi"") > 0 Or InStr(1, t, ""accounting"") > 0 Then NormalizeMajor = ""Akuntansi"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""manajemen"") > 0 Or InStr(1, t, ""management"") > 0 Then NormalizeMajor = ""Manajemen"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""ekonomi"") > 0 Or InStr(1, t, ""economics"") > 0 Then NormalizeMajor = ""Ekonomi"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""psikologi"") > 0 Or InStr(1, t, ""psychology"") > 0 Then NormalizeMajor = ""Psikologi"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""hukum"") > 0 Or InStr(1, t, ""law"") > 0 Then NormalizeMajor = ""Hukum"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""ilmu komunikasi"") > 0 Or InStr(1, t, ""communication"") > 0 Then NormalizeMajor = ""Ilmu Komunikasi"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""administrasi bisnis"") > 0 Or InStr(1, t, ""business administration"") > 0 Then NormalizeMajor = ""Administrasi Bisnis"": Exit Function" & vbCrLf
    s = s & "    If InStr(1, t, ""statistika"") > 0 Or InStr(1, t, ""statistics"") > 0 Then NormalizeMajor = ""Statistika"": Exit Function" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ExtractEducationLevel(ByVal text As String) As String" & vbCrLf
    s = s & "    Dim t As String" & vbCrLf
    s = s & "    t = LCase(text)" & vbCrLf
    s = s & "    If InStr(t, ""s1"") > 0 Or InStr(t, ""bachelor"") > 0 Or InStr(t, ""sarjana"") > 0 Then ExtractEducationLevel = ""S1"": Exit Function" & vbCrLf
    s = s & "    If InStr(t, ""s2"") > 0 Or InStr(t, ""master"") > 0 Or InStr(t, ""magister"") > 0 Then ExtractEducationLevel = ""S2"": Exit Function" & vbCrLf
    s = s & "    If InStr(t, ""d4"") > 0 Or InStr(t, ""diploma 4"") > 0 Or InStr(t, ""diploma iv"") > 0 Then ExtractEducationLevel = ""D4"": Exit Function" & vbCrLf
    s = s & "    If InStr(t, ""d3"") > 0 Or InStr(t, ""diploma 3"") > 0 Then ExtractEducationLevel = ""D3"": Exit Function" & vbCrLf
    s = s & "    If InStr(t, ""smk"") > 0 Or InStr(t, ""vocational"") > 0 Then ExtractEducationLevel = ""SMK"": Exit Function" & vbCrLf
    s = s & "    If InStr(t, ""sma"") > 0 Or InStr(t, ""high school"") > 0 Then ExtractEducationLevel = ""SMA"": Exit Function" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function StartsWithAny(ByVal text As String, ByVal arr As Variant) As Boolean" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    For Each v In arr" & vbCrLf
    s = s & "        If Left$(Trim(text), Len(CStr(v))) = LCase(CStr(v)) Then" & vbCrLf
    s = s & "            StartsWithAny = True" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next v" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function ExtractAfterSep(ByVal text As String) As String" & vbCrLf
    s = s & "    Dim p As Long" & vbCrLf
    s = s & "    p = InStr(1, text, "":"")" & vbCrLf
    s = s & "    If p = 0 Then p = InStr(1, text, ""-"")" & vbCrLf
    s = s & "    If p > 0 Then" & vbCrLf
    s = s & "        ExtractAfterSep = Trim(Mid$(text, p + 1))" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ExtractAfterSep = """"" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function CleanTextValue(ByVal text As String) As String" & vbCrLf
    s = s & "    text = Trim(text)" & vbCrLf
    s = s & "    text = Replace(text, Chr(34), """")" & vbCrLf
    s = s & "    Do While InStr(text, ""  "") > 0" & vbCrLf
    s = s & "        text = Replace(text, ""  "", "" "")" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    CleanTextValue = text" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function LooksLikeName(ByVal text As String) As Boolean" & vbCrLf
    s = s & "    Dim words As Variant, w As Variant" & vbCrLf
    s = s & "    text = CleanTextValue(text)" & vbCrLf
    s = s & "    If Len(text) < 5 Or Len(text) > 60 Then Exit Function" & vbCrLf
    s = s & "    If InStr(1, text, ""@"") > 0 Then Exit Function" & vbCrLf
    s = s & "    If text Like ""*[0-9]*"" Then Exit Function" & vbCrLf
    s = s & "    If IsLabelLine(text) Then Exit Function" & vbCrLf
    s = s & "    words = Split(text, "" "")" & vbCrLf
    s = s & "    If UBound(words) < 1 Or UBound(words) > 5 Then Exit Function" & vbCrLf
    s = s & "    For Each w In words" & vbCrLf
    s = s & "        If Len(CStr(w)) < 2 Then Exit Function" & vbCrLf
    s = s & "    Next w" & vbCrLf
    s = s & "    LooksLikeName = True" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function IsLabelLine(ByVal text As String) As Boolean" & vbCrLf
    s = s & "    Dim labels As Variant, l As Variant" & vbCrLf
    s = s & "    labels = Array(""CV"", ""CURRICULUM"", ""RESUME"", ""DATA"", ""PENDIDIKAN"", ""EDUCATION"", ""PENGALAMAN"", ""EXPERIENCE"", ""SKILL"", ""KEMAMPUAN"", ""BAHASA"", ""HOBI"", ""PROFIL"", ""PROFILE"", ""TENTANG"", ""ABOUT"")" & vbCrLf
    s = s & "    For Each l In labels" & vbCrLf
    s = s & "        If InStr(1, UCase(text), UCase(l), vbTextCompare) = 1 Then" & vbCrLf
    s = s & "            IsLabelLine = True" & vbCrLf
    s = s & "            Exit Function" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next l" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function CleanPhone(ByVal text As String) As String" & vbCrLf
    s = s & "    text = Replace(text, "" "", """")" & vbCrLf
    s = s & "    text = Replace(text, ""-"", """")" & vbCrLf
    s = s & "    text = Replace(text, ""."", """")" & vbCrLf
    s = s & "    text = Replace(text, ""("", """")" & vbCrLf
    s = s & "    text = Replace(text, "")"", """")" & vbCrLf
    s = s & "    text = Replace(text, ""+"", """")" & vbCrLf
    s = s & "    CleanPhone = text" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub TestClipboardCVText()" & vbCrLf
    s = s & "    Dim t As String" & vbCrLf
    s = s & "    t = GetClipboardTextSafe()" & vbCrLf
    s = s & "    If Len(t) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Clipboard kosong."", vbExclamation" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        SaveCVRawText ""Manual Clipboard Test"", t" & vbCrLf
    s = s & "        MsgBox ""Clipboard terbaca: "" & Len(t) & "" karakter."", vbInformation" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ExportDBSourcingCSV()" & vbCrLf
    s = s & "    On Error GoTo ErrHandler" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    ws.Copy" & vbCrLf
    s = s & "    Dim f As Variant" & vbCrLf
    s = s & "    f = Application.GetSaveAsFilename(""DB_Sourcing.csv"", ""CSV Files (*.csv),*.csv"")" & vbCrLf
    s = s & "    If f <> False Then ActiveWorkbook.SaveAs CStr(f), xlCSVUTF8" & vbCrLf
    s = s & "    ActiveWorkbook.Close False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrHandler:" & vbCrLf
    s = s & "    MsgBox ""Export gagal: "" & Err.Description, vbExclamation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    Code_ModSourcingSupport = s
End Function

Private Function Code_ModMainSourcing() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'==================== MAIN MENU MACROS ====================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub OpenSourcingForm()" & vbCrLf
    s = s & "    InitSourcingCV" & vbCrLf
    s = s & "    frmSourcing.Show vbModeless" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub OpenSourcingList()" & vbCrLf
    s = s & "    InitSourcingCV" & vbCrLf
    s = s & "    frmSourcingList.Show vbModeless" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub OpenSearchFPTK()" & vbCrLf
    s = s & "    InitSourcingCV" & vbCrLf
    s = s & "    frmSearchFPTK.Show vbModal" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub OpenPasteCVText()" & vbCrLf
    s = s & "    InitSourcingCV" & vbCrLf
    s = s & "    frmSourcing.Show vbModeless" & vbCrLf
    s = s & "    frmPasteCVText.Show vbModal" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub OpenEditSourcingForm()" & vbCrLf
    s = s & "    InitSourcingCV" & vbCrLf
    s = s & "    MsgBox ""Buka dari frmSourcingList: pilih kandidat lalu klik Edit/Double-click."", vbInformation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "'==================== SHORTCUT MACROS (BISA DI-ASSIGN KE BUTTON) ====================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ParseClipboardCVToDBSourcing()" & vbCrLf
    s = s & "    ' Copy text CV dulu (Ctrl+C), lalu jalankan macro ini." & vbCrLf
    s = s & "    ' Hasil akan masuk ke ROW BARU di DB Sourcing." & vbCrLf
    s = s & "    Dim rawText As String" & vbCrLf
    s = s & "    rawText = GetClipboardTextSafe()" & vbCrLf
    s = s & "    If Len(Trim(rawText)) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Clipboard kosong. Copy text CV dulu (Ctrl+A, Ctrl+C dari PDF/Jobstreet/LinkedIn)."", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    SaveCVRawText ""Clipboard CV"", rawText" & vbCrLf
    s = s & "    Dim ws As Worksheet, rowNum As Long" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    rowNum = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    s = s & "    If rowNum < 2 Then rowNum = 2" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""No"", rowNum - 1" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Sourcing Date"", Date" & vbCrLf
    s = s & "    Dim info As Variant" & vbCrLf
    s = s & "    info = ParseCVTextSmart(rawText)" & vbCrLf
    s = s & "    ' Clear old data dulu (meskipun row baru, aman)" & vbCrLf
    s = s & "    ClearParsedCVFields ws, rowNum" & vbCrLf
    s = s & "    ' Write parsed info" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Nama"", CStr(info(0))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Email"", CStr(info(1))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Nomor HP"", CStr(info(2))" & vbCrLf
    s = s & "    If Len(CStr(info(3))) > 0 Then" & vbCrLf
    s = s & "        If IsTopUniversityDropdown(CStr(info(3))) Then" & vbCrLf
    s = s & "            SetByHeader ws, rowNum, ""Nama Universitas/Sekolah (TOP 10)"", CStr(info(3))" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            SetByHeader ws, rowNum, ""Nama Universitas/Sekolah Lainnya"", CStr(info(3))" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Jurusan"", CStr(info(4))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""IPK"", CStr(info(5))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Tahun Lulus"", CStr(info(6))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Skor Bahasa Inggris"", CStr(info(7))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Domisili"", CStr(info(8))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Jenjang Pendidikan"", CStr(info(9))" & vbCrLf
    s = s & "    ' AUTO TIER" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""University Tier"", GetUniversityTierCV(CStr(info(3)))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""IPK Tier"", GetIPKTierCV(CStr(info(5)))" & vbCrLf
    s = s & "    ws.Columns(""A:BM"").AutoFit" & vbCrLf
    s = s & "    MsgBox ""CV berhasil diparse ke row "" & rowNum & ""."" & vbCrLf & vbCrLf & ""Nama: "" & info(0) & vbCrLf & ""Universitas: "" & info(3) & vbCrLf & ""IPK: "" & info(5) & vbCrLf & ""Univ Tier: "" & GetUniversityTierCV(CStr(info(3))) & vbCrLf & ""IPK Tier: "" & GetIPKTierCV(CStr(info(5))), vbInformation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ParseClipboardCVToSelectedRow()" & vbCrLf
    s = s & "    ' Copy text CV dulu, klik row tujuan di DB Sourcing, lalu jalankan macro ini." & vbCrLf
    s = s & "    ' Hasil akan MENIMPA row aktif (info CV lama di-clear dulu)." & vbCrLf
    s = s & "    Dim rawText As String" & vbCrLf
    s = s & "    rawText = GetClipboardTextSafe()" & vbCrLf
    s = s & "    If Len(Trim(rawText)) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Clipboard kosong. Copy text CV dulu."", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    Dim ws As Worksheet, rowNum As Long" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    If ActiveSheet.Name = ws.Name And ActiveCell.Row >= 2 Then" & vbCrLf
    s = s & "        rowNum = ActiveCell.Row" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        rowNum = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    s = s & "        If rowNum < 2 Then rowNum = 2" & vbCrLf
    s = s & "        SetByHeader ws, rowNum, ""No"", rowNum - 1" & vbCrLf
    s = s & "        SetByHeader ws, rowNum, ""Sourcing Date"", Date" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SaveCVRawText ""Clipboard CV"", rawText" & vbCrLf
    s = s & "    Dim info As Variant" & vbCrLf
    s = s & "    info = ParseCVTextSmart(rawText)" & vbCrLf
    s = s & "    ClearParsedCVFields ws, rowNum" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Nama"", CStr(info(0))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Email"", CStr(info(1))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Nomor HP"", CStr(info(2))" & vbCrLf
    s = s & "    If Len(CStr(info(3))) > 0 Then" & vbCrLf
    s = s & "        If IsTopUniversityDropdown(CStr(info(3))) Then" & vbCrLf
    s = s & "            SetByHeader ws, rowNum, ""Nama Universitas/Sekolah (TOP 10)"", CStr(info(3))" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            SetByHeader ws, rowNum, ""Nama Universitas/Sekolah Lainnya"", CStr(info(3))" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Jurusan"", CStr(info(4))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""IPK"", CStr(info(5))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Tahun Lulus"", CStr(info(6))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Skor Bahasa Inggris"", CStr(info(7))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Domisili"", CStr(info(8))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Jenjang Pendidikan"", CStr(info(9))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""University Tier"", GetUniversityTierCV(CStr(info(3)))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""IPK Tier"", GetIPKTierCV(CStr(info(5)))" & vbCrLf
    s = s & "    ws.Columns(""A:BM"").AutoFit" & vbCrLf
    s = s & "    MsgBox ""CV berhasil diparse ke row "" & rowNum & "" (data CV lama sudah diclear)."" & vbCrLf & vbCrLf & ""University Tier: "" & GetUniversityTierCV(CStr(info(3))) & vbCrLf & ""IPK Tier: "" & GetIPKTierCV(CStr(info(5))), vbInformation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ScanCVFileToNewRow()" & vbCrLf
    s = s & "    ' Pilih file PDF, hasil masuk row baru." & vbCrLf
    s = s & "    Dim filePath As String" & vbCrLf
    s = s & "    filePath = PickFileCV()" & vbCrLf
    s = s & "    If Len(Trim(filePath)) = 0 Then Exit Sub" & vbCrLf
    s = s & "    Dim rawText As String" & vbCrLf
    s = s & "    rawText = ReadPDFAutoCopy(filePath)" & vbCrLf
    s = s & "    If Len(Trim(rawText)) < 50 Then" & vbCrLf
    s = s & "        MsgBox ""Text CV tidak terbaca dari PDF. Coba copy manual (Ctrl+A, Ctrl+C) lalu jalankan ParseClipboardCVToDBSourcing."", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    SaveCVRawText filePath, rawText" & vbCrLf
    s = s & "    Dim ws As Worksheet, rowNum As Long" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    rowNum = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    s = s & "    If rowNum < 2 Then rowNum = 2" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""No"", rowNum - 1" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Sourcing Date"", Date" & vbCrLf
    s = s & "    Dim info As Variant" & vbCrLf
    s = s & "    info = ParseCVTextSmart(rawText)" & vbCrLf
    s = s & "    ClearParsedCVFields ws, rowNum" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Nama"", CStr(info(0))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Email"", CStr(info(1))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Nomor HP"", CStr(info(2))" & vbCrLf
    s = s & "    If Len(CStr(info(3))) > 0 Then" & vbCrLf
    s = s & "        If IsTopUniversityDropdown(CStr(info(3))) Then" & vbCrLf
    s = s & "            SetByHeader ws, rowNum, ""Nama Universitas/Sekolah (TOP 10)"", CStr(info(3))" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            SetByHeader ws, rowNum, ""Nama Universitas/Sekolah Lainnya"", CStr(info(3))" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Jurusan"", CStr(info(4))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""IPK"", CStr(info(5))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Tahun Lulus"", CStr(info(6))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Skor Bahasa Inggris"", CStr(info(7))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Domisili"", CStr(info(8))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Jenjang Pendidikan"", CStr(info(9))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""University Tier"", GetUniversityTierCV(CStr(info(3)))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""IPK Tier"", GetIPKTierCV(CStr(info(5)))" & vbCrLf
    s = s & "    ws.Columns(""A:BM"").AutoFit" & vbCrLf
    s = s & "    MsgBox ""CV PDF berhasil discan ke row baru DB Sourcing."" & vbCrLf & vbCrLf & ""Nama: "" & info(0) & vbCrLf & ""University Tier: "" & GetUniversityTierCV(CStr(info(3))) & vbCrLf & ""IPK Tier: "" & GetIPKTierCV(CStr(info(5))), vbInformation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ScanCVFileToSelectedRow()" & vbCrLf
    s = s & "    ' Pilih file PDF, hasil menimpa row aktif." & vbCrLf
    s = s & "    Dim filePath As String" & vbCrLf
    s = s & "    filePath = PickFileCV()" & vbCrLf
    s = s & "    If Len(Trim(filePath)) = 0 Then Exit Sub" & vbCrLf
    s = s & "    Dim rawText As String" & vbCrLf
    s = s & "    rawText = ReadPDFAutoCopy(filePath)" & vbCrLf
    s = s & "    If Len(Trim(rawText)) < 50 Then" & vbCrLf
    s = s & "        MsgBox ""Text CV tidak terbaca dari PDF."", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    Dim ws As Worksheet, rowNum As Long" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    If ActiveSheet.Name = ws.Name And ActiveCell.Row >= 2 Then" & vbCrLf
    s = s & "        rowNum = ActiveCell.Row" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        rowNum = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    s = s & "        If rowNum < 2 Then rowNum = 2" & vbCrLf
    s = s & "        SetByHeader ws, rowNum, ""No"", rowNum - 1" & vbCrLf
    s = s & "        SetByHeader ws, rowNum, ""Sourcing Date"", Date" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SaveCVRawText filePath, rawText" & vbCrLf
    s = s & "    Dim info As Variant" & vbCrLf
    s = s & "    info = ParseCVTextSmart(rawText)" & vbCrLf
    s = s & "    ClearParsedCVFields ws, rowNum" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Nama"", CStr(info(0))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Email"", CStr(info(1))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Nomor HP"", CStr(info(2))" & vbCrLf
    s = s & "    If Len(CStr(info(3))) > 0 Then" & vbCrLf
    s = s & "        If IsTopUniversityDropdown(CStr(info(3))) Then" & vbCrLf
    s = s & "            SetByHeader ws, rowNum, ""Nama Universitas/Sekolah (TOP 10)"", CStr(info(3))" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            SetByHeader ws, rowNum, ""Nama Universitas/Sekolah Lainnya"", CStr(info(3))" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Jurusan"", CStr(info(4))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""IPK"", CStr(info(5))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Tahun Lulus"", CStr(info(6))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Skor Bahasa Inggris"", CStr(info(7))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Domisili"", CStr(info(8))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""Jenjang Pendidikan"", CStr(info(9))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""University Tier"", GetUniversityTierCV(CStr(info(3)))" & vbCrLf
    s = s & "    SetByHeader ws, rowNum, ""IPK Tier"", GetIPKTierCV(CStr(info(5)))" & vbCrLf
    s = s & "    ws.Columns(""A:BM"").AutoFit" & vbCrLf
    s = s & "    MsgBox ""CV PDF berhasil discan ke row "" & rowNum & "" (data CV lama sudah diclear)."" & vbCrLf & vbCrLf & ""University Tier: "" & GetUniversityTierCV(CStr(info(3))) & vbCrLf & ""IPK Tier: "" & GetIPKTierCV(CStr(info(5))), vbInformation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub TestParseLastRawText()" & vbCrLf
    s = s & "    ' Ambil raw text terakhir dari sheet CV Raw Text, lalu parse ke row baru DB Sourcing." & vbCrLf
    s = s & "    If Not WSExists(""CV Raw Text"") Then" & vbCrLf
    s = s & "        MsgBox ""Sheet CV Raw Text belum ada."", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim wsRaw As Worksheet, wsDB As Worksheet" & vbCrLf
    s = s & "    Set wsRaw = ThisWorkbook.Worksheets(""CV Raw Text"")" & vbCrLf
    s = s & "    Dim rowRaw As Long" & vbCrLf
    s = s & "    rowRaw = wsRaw.Cells(wsRaw.Rows.Count, 5).End(xlUp).Row" & vbCrLf
    s = s & "    If rowRaw < 2 Then" & vbCrLf
    s = s & "        MsgBox ""Belum ada raw text CV di sheet CV Raw Text."", vbExclamation" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim rawText As String" & vbCrLf
    s = s & "    rawText = CStr(wsRaw.Cells(rowRaw, 5).Value)" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    Set wsDB = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    Dim rowDB As Long" & vbCrLf
    s = s & "    rowDB = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    s = s & "    If rowDB < 2 Then rowDB = 2" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""No"", rowDB - 1" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""Sourcing Date"", Date" & vbCrLf
    s = s & "    Dim info As Variant" & vbCrLf
    s = s & "    info = ParseCVTextSmart(rawText)" & vbCrLf
    s = s & "    ClearParsedCVFields wsDB, rowDB" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""Nama"", CStr(info(0))" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""Email"", CStr(info(1))" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""Nomor HP"", CStr(info(2))" & vbCrLf
    s = s & "    If Len(CStr(info(3))) > 0 Then" & vbCrLf
    s = s & "        If IsTopUniversityDropdown(CStr(info(3))) Then" & vbCrLf
    s = s & "            SetByHeader wsDB, rowDB, ""Nama Universitas/Sekolah (TOP 10)"", CStr(info(3))" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            SetByHeader wsDB, rowDB, ""Nama Universitas/Sekolah Lainnya"", CStr(info(3))" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""Jurusan"", CStr(info(4))" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""IPK"", CStr(info(5))" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""Tahun Lulus"", CStr(info(6))" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""Skor Bahasa Inggris"", CStr(info(7))" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""Domisili"", CStr(info(8))" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""Jenjang Pendidikan"", CStr(info(9))" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""University Tier"", GetUniversityTierCV(CStr(info(3)))" & vbCrLf
    s = s & "    SetByHeader wsDB, rowDB, ""IPK Tier"", GetIPKTierCV(CStr(info(5)))" & vbCrLf
    s = s & "    wsDB.Columns(""A:BM"").AutoFit" & vbCrLf
    s = s & "    MsgBox ""Raw text terakhir berhasil dites ke DB Sourcing row "" & rowDB & "".'' & vbCrLf & vbCrLf & ""University Tier: "" & GetUniversityTierCV(CStr(info(3))) & vbCrLf & ""IPK Tier: "" & GetIPKTierCV(CStr(info(5))), vbInformation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Function IsTopUniversityDropdown(ByVal univName As String) As Boolean" & vbCrLf
    s = s & "    Dim u As String" & vbCrLf
    s = s & "    u = LCase(Trim(univName))" & vbCrLf
    s = s & "    Select Case u" & vbCrLf
    s = s & "        Case ""universitas indonesia"", ""universitas gadjah mada"", ""institut teknologi bandung"", _" & vbCrLf
    s = s & "             ""universitas airlangga"", ""universitas padjadjaran"", ""universitas diponegoro"", _" & vbCrLf
    s = s & "             ""universitas brawijaya"", ""institut pertanian bogor"", ""universitas sebelas maret"", _" & vbCrLf
    s = s & "             ""institut teknologi sepuluh nopember"", ""telkom university"", _" & vbCrLf
    s = s & "             ""universitas bina nusantara"", ""universitas gunadarma""" & vbCrLf
    s = s & "            IsTopUniversityDropdown = True" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ClearParsedCVFields(ByVal ws As Worksheet, ByVal rowNum As Long)" & vbCrLf
    s = s & "    ' Membersihkan field hasil parse CV sebelum menulis data baru" & vbCrLf
    s = s & "    If rowNum < 2 Then Exit Sub" & vbCrLf
    s = s & "    Dim headersToClear As Variant" & vbCrLf
    s = s & "    headersToClear = Array(""Nama"", ""Nama Kandidat"", ""Candidate Name"", _" & vbCrLf
    s = s & "        ""Nomor HP"", ""No HP"", ""Phone"", ""Email"", ""Domisili"", ""Domicile"", ""Location"", _" & vbCrLf
    s = s & "        ""Nama Universitas"", ""Nama Universitas/Sekolah"", ""Nama Universitas/Sekolah (TOP 10)"", ""Nama Universitas/Sekolah Lainnya"", _" & vbCrLf
    s = s & "        ""Universitas"", ""University"", ""Jenjang Pendidikan"", ""Education Level"", ""Jurusan"", ""Major"", _" & vbCrLf
    s = s & "        ""Tahun Lulus"", ""Graduation Year"", ""IPK"", ""GPA"", ""Skor Bahasa Inggris"", ""English Score"", _" & vbCrLf
    s = s & "        ""University Tier"", ""Univ Tier"", ""Tier Universitas"", ""IPK Tier"", ""GPA Tier"", ""Tier IPK"")" & vbCrLf
    s = s & "    Dim i As Long, c As Long" & vbCrLf
    s = s & "    For i = LBound(headersToClear) To UBound(headersToClear)" & vbCrLf
    s = s & "        c = FindHeaderColumn(ws, CStr(headersToClear(i)))" & vbCrLf
    s = s & "        If c > 0 Then ws.Cells(rowNum, c).ClearContents" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ExportDBSourcingCSV()" & vbCrLf
    s = s & "    On Error GoTo ErrHandler" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    ws.Copy" & vbCrLf
    s = s & "    Dim f As Variant" & vbCrLf
    s = s & "    f = Application.GetSaveAsFilename(""DB_Sourcing.csv"", ""CSV Files (*.csv),*.csv"")" & vbCrLf
    s = s & "    If f <> False Then ActiveWorkbook.SaveAs CStr(f), xlCSVUTF8" & vbCrLf
    s = s & "    ActiveWorkbook.Close False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ErrHandler:" & vbCrLf
    s = s & "    MsgBox ""Export gagal: "" & Err.Description, vbExclamation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    Code_ModMainSourcing = s
End Function

Private Function Code_frmSearchFPTK() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "Private mBuilt As Boolean" & vbCrLf
    s = s & "Private txtKeyword As MSForms.TextBox" & vbCrLf
    s = s & "Private WithEvents btnSearch As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnSelect As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnCancel As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents lstResults As MSForms.ListBox" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildUI()" & vbCrLf
    s = s & "    If mBuilt Then Exit Sub" & vbCrLf
    s = s & "    mBuilt = True" & vbCrLf
    s = s & "    Me.Caption = ""SEARCH FPTK""" & vbCrLf
    s = s & "    Me.Width = 440" & vbCrLf
    s = s & "    Me.Height = 355" & vbCrLf
    s = s & "    Me.BackColor = RGB(245, 245, 245)" & vbCrLf
    s = s & "    AddLabel Me, ""lblTitle"", ""Cari FPTK dari sheet FPTK"", 12, 12, 410, 20, True, 10" & vbCrLf
    s = s & "    AddLabel Me, ""lblKeyword"", ""Keyword:"", 12, 42, 60, 20" & vbCrLf
    s = s & "    Set txtKeyword = AddTextBox(Me, ""txtKeyword"", 80, 40, 270, 22)" & vbCrLf
    s = s & "    Set btnSearch = AddButton(Me, ""btnSearch"", ""Cari"", 360, 38, 60, 26, RGB(52, 152, 219))" & vbCrLf
    s = s & "    Set lstResults = AddListBox(Me, ""lstResults"", 12, 75, 410, 170)" & vbCrLf
    s = s & "    lstResults.ColumnCount = 3" & vbCrLf
    s = s & "    lstResults.ColumnWidths = ""125;220;55""" & vbCrLf
    s = s & "    AddLabel Me, ""lblTip"", ""Double-click untuk memilih FPTK"", 12, 252, 410, 18, False, 8, True" & vbCrLf
    s = s & "    Set btnSelect = AddButton(Me, ""btnSelect"", ""Pilih FPTK"", 145, 282, 125, 28, RGB(39, 174, 96))" & vbCrLf
    s = s & "    Set btnCancel = AddButton(Me, ""btnCancel"", ""Cancel"", 280, 282, 125, 28, RGB(231, 76, 60))" & vbCrLf
    s = s & "    txtKeyword.SetFocus" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSearch_Click()" & vbCrLf
    s = s & "    SearchAndRender" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtKeyword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
    s = s & "    If KeyCode = 13 Then SearchAndRender" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SearchAndRender()" & vbCrLf
    s = s & "    Dim keyword As String" & vbCrLf
    s = s & "    keyword = Trim(txtKeyword.Value)" & vbCrLf
    s = s & "    If keyword = """" Then MsgBox ""Masukkan keyword!"", vbExclamation: txtKeyword.SetFocus: Exit Sub" & vbCrLf
    s = s & "    lstResults.Clear" & vbCrLf
    s = s & "    Dim results As Collection" & vbCrLf
    s = s & "    Set results = SearchFPTK(keyword)" & vbCrLf
    s = s & "    If results.Count = 0 Then" & vbCrLf
    s = s & "        lstResults.AddItem ""Tidak ditemukan""" & vbCrLf
    s = s & "        lstResults.List(0, 1) = ""Cek sheet FPTK / keyword""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim item As Variant, parts() As String, i As Long" & vbCrLf
    s = s & "    i = 0" & vbCrLf
    s = s & "    For Each item In results" & vbCrLf
    s = s & "        parts = Split(CStr(item), ""|"")" & vbCrLf
    s = s & "        lstResults.AddItem """"" & vbCrLf
    s = s & "        lstResults.List(i, 0) = parts(0)" & vbCrLf
    s = s & "        lstResults.List(i, 1) = parts(1)" & vbCrLf
    s = s & "        lstResults.List(i, 2) = parts(2)" & vbCrLf
    s = s & "        i = i + 1" & vbCrLf
    s = s & "    Next item" & vbCrLf
    s = s & "    If lstResults.ListCount > 0 Then lstResults.ListIndex = 0: lstResults.SetFocus" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSelect_Click()" & vbCrLf
    s = s & "    If lstResults.ListIndex < 0 Then MsgBox ""Pilih FPTK!"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If lstResults.List(lstResults.ListIndex, 0) = ""Tidak ditemukan"" Then Exit Sub" & vbCrLf
    s = s & "    Dim kode As String" & vbCrLf
    s = s & "    kode = CStr(lstResults.List(lstResults.ListIndex, 0))" & vbCrLf
    s = s & "    Dim fptkDetail As Variant" & vbCrLf
    s = s & "    fptkDetail = GetFPTKDetail(kode)" & vbCrLf
    s = s & "    frmSourcing.SetFPTKSelection CStr(fptkDetail(0)), CStr(fptkDetail(1)), CStr(fptkDetail(3)), CStr(fptkDetail(4))" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    btnSelect_Click" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddLabel(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal b As Boolean = False, Optional ByVal s As Single = 8, Optional ByVal it As Boolean = False) As MSForms.Label" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.Label.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = s: o.Font.Bold = b: o.Font.Italic = it: o.BackStyle = 0" & vbCrLf
    s = s & "    If b Then o.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    Set AddLabel = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddTextBox(ByVal p As Object, ByVal n As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal lk As Boolean = False) As MSForms.TextBox" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.TextBox.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Font.Size = 8" & vbCrLf
    s = s & "    If lk Then o.Locked = True: o.BackColor = RGB(236, 240, 241)" & vbCrLf
    s = s & "    Set AddTextBox = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddButton(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, ByVal bc As Long) As MSForms.CommandButton" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.CommandButton.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = 8: o.Font.Bold = True: o.BackColor = bc: o.ForeColor = vbWhite" & vbCrLf
    s = s & "    Set AddButton = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddListBox(ByVal p As Object, ByVal n As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As MSForms.ListBox" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.ListBox.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Font.Size = 8" & vbCrLf
    s = s & "    Set AddListBox = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    Code_frmSearchFPTK = s
End Function

Private Function Code_frmSourcing() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "Private mBuilt As Boolean, mCVPath As String" & vbCrLf
    s = s & "Private txtKodeUnik As MSForms.TextBox, txtPosisi As MSForms.TextBox, txtRekruter As MSForms.TextBox" & vbCrLf
    s = s & "Private cboModel As MSForms.ComboBox, cboSumber As MSForms.ComboBox" & vbCrLf
    s = s & "Private imgPreview As MSForms.Image, lblCVStatus As MSForms.Label" & vbCrLf
    s = s & "Private txtNama As MSForms.TextBox, cboUniversitas As MSForms.ComboBox, txtUnivOther As MSForms.TextBox" & vbCrLf
    s = s & "Private cboJenjang As MSForms.ComboBox, cboJurusan As MSForms.ComboBox, txtTahunLulus As MSForms.TextBox" & vbCrLf
    s = s & "Private txtIPK As MSForms.TextBox, txtEnglishScore As MSForms.TextBox, cboUnivTier As MSForms.ComboBox, cboIPKTier As MSForms.ComboBox" & vbCrLf
    s = s & "Private txtHP As MSForms.TextBox, txtEmail As MSForms.TextBox, txtDomisili As MSForms.TextBox" & vbCrLf
    s = s & "Private WithEvents btnSearchFPTK As MSForms.CommandButton, WithEvents btnUploadCV As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnSave As MSForms.CommandButton, WithEvents btnCancel As MSForms.CommandButton, WithEvents btnList As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnPasteText As MSForms.CommandButton" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "    FillDropdowns" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildUI()" & vbCrLf
    s = s & "    If mBuilt Then Exit Sub" & vbCrLf
    s = s & "    mBuilt = True" & vbCrLf
    s = s & "    Me.Caption = ""SOURCING CV - INPUT DATA KANDIDAT""" & vbCrLf
    s = s & "    Me.Width = 560" & vbCrLf
    s = s & "    Me.Height = 520" & vbCrLf
    s = s & "    Me.BackColor = RGB(245, 245, 245)" & vbCrLf
    s = s & "    Me.ScrollBars = fmScrollBarsNone" & vbCrLf
    s = s & "    Dim fra As MSForms.Frame" & vbCrLf
    s = s & "    Set fra = AddFrame(Me, ""fraFPTK"", ""Pilih FPTK"", 10, 8, 520, 70)" & vbCrLf
    s = s & "    AddLabel fra, ""lblKodeUnik"", ""Kode:"", 10, 18, 45, 16" & vbCrLf
    s = s & "    Set txtKodeUnik = AddTextBox(fra, ""txtKodeUnik"", 55, 16, 350, 20, True)" & vbCrLf
    s = s & "    Set btnSearchFPTK = AddButton(fra, ""btnSearchFPTK"", ""Cari"", 415, 15, 90, 22, RGB(52, 152, 219))" & vbCrLf
    s = s & "    AddLabel fra, ""lblPosisi"", ""Posisi:"", 10, 43, 45, 16" & vbCrLf
    s = s & "    Set txtPosisi = AddTextBox(fra, ""txtPosisi"", 55, 41, 450, 20, True)" & vbCrLf
    s = s & "    Set fra = AddFrame(Me, ""fraDetail"", ""Detail Rekrutmen"", 10, 82, 520, 60)" & vbCrLf
    s = s & "    AddLabel fra, ""lblModel"", ""Model:"", 10, 18, 45, 16" & vbCrLf
    s = s & "    Set cboModel = AddCombo(fra, ""cboModel"", 55, 16, 110, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblRekruter"", ""Rekruter:"", 175, 18, 60, 16" & vbCrLf
    s = s & "    Set txtRekruter = AddTextBox(fra, ""txtRekruter"", 235, 16, 270, 20, True)" & vbCrLf
    s = s & "    AddLabel fra, ""lblSumber"", ""Sumber:"", 10, 40, 45, 16" & vbCrLf
    s = s & "    Set cboSumber = AddCombo(fra, ""cboSumber"", 55, 38, 450, 20)" & vbCrLf
    s = s & "    Set fra = AddFrame(Me, ""fraKandidat"", ""Data Kandidat"", 10, 146, 520, 220)" & vbCrLf
    s = s & "    AddLabel fra, ""lblNama"", ""Nama:"", 10, 18, 65, 16" & vbCrLf
    s = s & "    Set txtNama = AddTextBox(fra, ""txtNama"", 80, 16, 180, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblJenjang"", ""Jenjang:"", 10, 42, 65, 16" & vbCrLf
    s = s & "    Set cboJenjang = AddCombo(fra, ""cboJenjang"", 80, 40, 180, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblUniv"", ""Univ TOP10:"", 10, 66, 65, 16" & vbCrLf
    s = s & "    Set cboUniversitas = AddCombo(fra, ""cboUniversitas"", 80, 64, 180, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblUnivOther"", ""Univ Lain:"", 10, 90, 65, 16" & vbCrLf
    s = s & "    Set txtUnivOther = AddTextBox(fra, ""txtUnivOther"", 80, 88, 180, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblJurusan"", ""Jurusan:"", 10, 114, 65, 16" & vbCrLf
    s = s & "    Set cboJurusan = AddCombo(fra, ""cboJurusan"", 80, 112, 180, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblTahun"", ""Thn Lulus:"", 10, 138, 65, 16" & vbCrLf
    s = s & "    Set txtTahunLulus = AddTextBox(fra, ""txtTahunLulus"", 80, 136, 180, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblDomisili"", ""Domisili:"", 10, 162, 65, 16" & vbCrLf
    s = s & "    Set txtDomisili = AddTextBox(fra, ""txtDomisili"", 80, 160, 180, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblIPK"", ""IPK:"", 280, 18, 70, 16" & vbCrLf
    s = s & "    Set txtIPK = AddTextBox(fra, ""txtIPK"", 355, 16, 145, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblEnglish"", ""Eng Score:"", 280, 42, 70, 16" & vbCrLf
    s = s & "    Set txtEnglishScore = AddTextBox(fra, ""txtEnglishScore"", 355, 40, 145, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblUnivTier"", ""Univ Tier:"", 280, 66, 70, 16" & vbCrLf
    s = s & "    Set cboUnivTier = AddCombo(fra, ""cboUnivTier"", 355, 64, 145, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblIPKTier"", ""IPK Tier:"", 280, 90, 70, 16" & vbCrLf
    s = s & "    Set cboIPKTier = AddCombo(fra, ""cboIPKTier"", 355, 88, 145, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblHP"", ""No HP:"", 280, 114, 70, 16" & vbCrLf
    s = s & "    Set txtHP = AddTextBox(fra, ""txtHP"", 355, 112, 145, 20)" & vbCrLf
    s = s & "    AddLabel fra, ""lblEmail"", ""Email:"", 280, 138, 70, 16" & vbCrLf
    s = s & "    Set txtEmail = AddTextBox(fra, ""txtEmail"", 355, 136, 145, 20)" & vbCrLf
    s = s & "    Set fra = AddFrame(Me, ""fraCV"", ""CV Upload + Raw Text"", 10, 370, 520, 65)" & vbCrLf
    s = s & "    Set imgPreview = fra.Controls.Add(""Forms.Image.1"", ""imgPreview"", True)" & vbCrLf
    s = s & "    imgPreview.Left = 10: imgPreview.Top = 18: imgPreview.Width = 135: imgPreview.Height = 35" & vbCrLf
    s = s & "    imgPreview.BorderStyle = 1: imgPreview.PictureAlignment = 2: imgPreview.PictureSizeMode = 1" & vbCrLf
    s = s & "    Set btnUploadCV = AddButton(fra, ""btnUploadCV"", ""Upload CV"", 155, 22, 80, 24, RGB(52, 152, 219))" & vbCrLf
    s = s & "    Set btnPasteText = AddButton(fra, ""btnPasteText"", ""Paste Text"", 245, 22, 85, 24, RGB(39, 174, 96))" & vbCrLf
    s = s & "    Set lblCVStatus = AddLabel(fra, ""lblCVStatus"", ""Belum ada CV"", 340, 14, 165, 44, False, 8)" & vbCrLf
    s = s & "    lblCVStatus.ForeColor = RGB(127, 140, 141)" & vbCrLf
    s = s & "    Set btnSave = AddButton(Me, ""btnSave"", ""SIMPAN"", 135, 448, 95, 30, RGB(39, 174, 96))" & vbCrLf
    s = s & "    Set btnCancel = AddButton(Me, ""btnCancel"", ""CANCEL"", 240, 448, 95, 30, RGB(231, 76, 60))" & vbCrLf
    s = s & "    Set btnList = AddButton(Me, ""btnList"", ""LIHAT DB"", 345, 448, 95, 30, RGB(52, 152, 219))" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FillDropdowns()" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    If cboModel.ListCount = 0 Then For Each v In Array(""Model 1"", ""Model 2"", ""Model 3"", ""Model 4""): cboModel.AddItem v: Next v" & vbCrLf
    s = s & "    If cboSumber.ListCount = 0 Then For Each v In Array(""Jobstreet"", ""LinkedIn"", ""Google Form"", ""Referensi User"", ""Referensi Karyawan"", ""Campus Hiring"", ""Walk-in Interview"", ""Database Internal"", ""Freelance""): cboSumber.AddItem v: Next v" & vbCrLf
    s = s & "    If cboJenjang.ListCount = 0 Then For Each v In Array(""SMA/SMK"", ""D3"", ""D4"", ""S1"", ""S2""): cboJenjang.AddItem v: Next v" & vbCrLf
    s = s & "    If cboUniversitas.ListCount = 0 Then For Each v In Array(""Universitas Indonesia"", ""Universitas Gadjah Mada"", ""Institut Teknologi Bandung"", ""Universitas Airlangga"", ""Universitas Padjadjaran"", ""Universitas Diponegoro"", ""Universitas Brawijaya"", ""Institut Pertanian Bogor"", ""Universitas Sebelas Maret"", ""Telkom University"", ""Lainnya""): cboUniversitas.AddItem v: Next v" & vbCrLf
    s = s & "    If cboJurusan.ListCount = 0 Then For Each v In Array(""Manajemen"", ""Akuntansi"", ""Teknik Industri"", ""Teknik Informatika"", ""Teknik Mesin"", ""Ekonomi Bisnis"", ""Psikologi"", ""Ilmu Komunikasi"", ""Manajemen Pemasaran"", ""Manajemen SDM"", ""Sistem Informasi"", ""Lainnya""): cboJurusan.AddItem v: Next v" & vbCrLf
    s = s & "    If cboUnivTier.ListCount = 0 Then For Each v In Array(""Top 3 PTN"", ""Top 10 PTN"", ""Top 20 PTN"", ""Top 10 PTS"", ""Lainnya""): cboUnivTier.AddItem v: Next v" & vbCrLf
    s = s & "    If cboIPKTier.ListCount = 0 Then For Each v In Array(""Lebih dari 3,5"", ""Lebih dari 3,2"", ""Kurang dari 3,5"", ""Kurang dari 3,2""): cboIPKTier.AddItem v: Next v" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub SetFPTKSelection(ByVal kodeUnik As String, ByVal posisi As String, ByVal rekruter As String, ByVal model As String)" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "    txtKodeUnik.Value = kodeUnik" & vbCrLf
    s = s & "    txtPosisi.Value = posisi" & vbCrLf
    s = s & "    txtRekruter.Value = rekruter" & vbCrLf
    s = s & "    cboModel.Value = DetermineModelRekrutmen(posisi, model)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSearchFPTK_Click()" & vbCrLf
    s = s & "    frmSearchFPTK.Show vbModal" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnList_Click()" & vbCrLf
    s = s & "    frmSourcingList.Show vbModeless" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnPasteText_Click()" & vbCrLf
    s = s & "    frmPasteCVText.Show vbModal" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnUploadCV_Click()" & vbCrLf
    s = s & "    Dim filePath As String" & vbCrLf
    s = s & "    filePath = PickFileCV()" & vbCrLf
    s = s & "    If Len(filePath) = 0 Then Exit Sub" & vbCrLf
    s = s & "    mCVPath = filePath" & vbCrLf
    s = s & "    If IsImageFile(filePath) Then" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        Set imgPreview.Picture = LoadPicture(filePath)" & vbCrLf
    s = s & "        If Err.Number <> 0 Then" & vbCrLf
    s = s & "            lblCVStatus.Caption = ""File: "" & ShortFileName(filePath, 24) & vbCrLf & ""No preview""" & vbCrLf
    s = s & "            Err.Clear" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            lblCVStatus.Caption = ""CV image loaded: "" & ShortFileName(filePath, 24)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        On Error GoTo 0" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lblCVStatus.Caption = ""Membaca PDF / clipboard...""" & vbCrLf
    s = s & "    Me.Caption = ""READING PDF...""" & vbCrLf
    s = s & "    DoEvents" & vbCrLf
    s = s & "    Dim extractedText As String" & vbCrLf
    s = s & "    extractedText = ReadPDFAutoCopy(filePath)" & vbCrLf
    s = s & "    If Len(extractedText) > 50 Then" & vbCrLf
    s = s & "        ApplyManualText extractedText, filePath" & vbCrLf
    s = s & "        lblCVStatus.Caption = ""PDF parsed: "" & Len(extractedText) & "" char""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblCVStatus.Caption = ""PDF kosong. Paste text manual.""" & vbCrLf
    s = s & "        If MsgBox(""PDF tidak terbaca. Mau paste text dari Jobstreet/LinkedIn/CV manual?"", vbQuestion + vbYesNo) = vbYes Then frmPasteCVText.Show vbModal" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Me.Caption = ""SOURCING CV - INPUT DATA KANDIDAT""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub ApplyManualText(ByVal rawText As String, Optional ByVal sourceLabel As String = ""Manual Paste - Edit"")" & vbCrLf
    s = s & "    If Len(Trim(rawText)) = 0 Then Exit Sub" & vbCrLf
    s = s & "    SaveCVRawText sourceLabel, rawText" & vbCrLf
    s = s & "    Dim info As Variant" & vbCrLf
    s = s & "    info = ParseCVTextSmart(rawText)" & vbCrLf
    s = s & "    ApplyParsedInfo info" & vbCrLf
    s = s & "    lblCVStatus.Caption = ""Text parsed: "" & Len(rawText) & "" char""" & vbCrLf
    s = s & "    MsgBox ""Text berhasil diproses."", vbInformation" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ApplyParsedInfo(ByVal info As Variant)" & vbCrLf
    s = s & "    ' CLEAR dulu field CV sebelum diisi ulang" & vbCrLf
    s = s & "    txtNama.Value = """"" & vbCrLf
    s = s & "    txtEmail.Value = """"" & vbCrLf
    s = s & "    txtHP.Value = """"" & vbCrLf
    s = s & "    cboUniversitas.Value = """"" & vbCrLf
    s = s & "    txtUnivOther.Value = """"" & vbCrLf
    s = s & "    cboJurusan.Value = """"" & vbCrLf
    s = s & "    txtIPK.Value = """"" & vbCrLf
    s = s & "    txtTahunLulus.Value = """"" & vbCrLf
    s = s & "    txtEnglishScore.Value = """"" & vbCrLf
    s = s & "    txtDomisili.Value = """"" & vbCrLf
    s = s & "    cboJenjang.Value = """"" & vbCrLf
    s = s & "    cboUnivTier.Value = """"" & vbCrLf
    s = s & "    cboIPKTier.Value = """"" & vbCrLf
    s = s & "    ' ISI ulang dari hasil parse" & vbCrLf
    s = s & "    If Len(info(0)) > 0 Then txtNama.Value = info(0)" & vbCrLf
    s = s & "    If Len(info(1)) > 0 Then txtEmail.Value = info(1)" & vbCrLf
    s = s & "    If Len(info(2)) > 0 Then txtHP.Value = info(2)" & vbCrLf
    s = s & "    If Len(info(3)) > 0 Then SetUniversityValue CStr(info(3))" & vbCrLf
    s = s & "    If Len(info(4)) > 0 Then SetComboValueOrAdd cboJurusan, CStr(info(4))" & vbCrLf
    s = s & "    If Len(info(5)) > 0 Then txtIPK.Value = info(5)" & vbCrLf
    s = s & "    If Len(info(6)) > 0 Then txtTahunLulus.Value = info(6)" & vbCrLf
    s = s & "    If Len(info(7)) > 0 Then txtEnglishScore.Value = info(7)" & vbCrLf
    s = s & "    If UBound(info) >= 8 Then If Len(info(8)) > 0 Then txtDomisili.Value = info(8)" & vbCrLf
    s = s & "    If UBound(info) >= 9 Then If Len(info(9)) > 0 Then SetComboValueOrAdd cboJenjang, CStr(info(9))" & vbCrLf
    s = s & "    If UBound(info) >= 10 Then If Len(info(10)) > 0 Then cboUnivTier.Value = info(10)" & vbCrLf
    s = s & "    If UBound(info) >= 11 Then If Len(info(11)) > 0 Then cboIPKTier.Value = info(11)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetUniversityValue(ByVal univName As String)" & vbCrLf
    s = s & "    If ComboHasValue(cboUniversitas, univName) Then" & vbCrLf
    s = s & "        cboUniversitas.Value = univName" & vbCrLf
    s = s & "        txtUnivOther.Value = """"" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        cboUniversitas.Value = ""Lainnya""" & vbCrLf
    s = s & "        txtUnivOther.Value = univName" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ComboHasValue(ByVal cbo As MSForms.ComboBox, ByVal val As String) As Boolean" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    For i = 0 To cbo.ListCount - 1" & vbCrLf
    s = s & "        If LCase(Trim(cbo.List(i))) = LCase(Trim(val)) Then ComboHasValue = True: Exit Function" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetComboValueOrAdd(ByVal cbo As MSForms.ComboBox, ByVal val As String)" & vbCrLf
    s = s & "    If Len(val) = 0 Then Exit Sub" & vbCrLf
    s = s & "    If Not ComboHasValue(cbo, val) Then cbo.AddItem val" & vbCrLf
    s = s & "    cbo.Value = val" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnSave_Click()" & vbCrLf
    s = s & "    If Trim(txtKodeUnik.Value) = """" Then MsgBox ""Pilih FPTK dulu!"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If Trim(txtNama.Value) = """" Then MsgBox ""Nama harus diisi!"", vbExclamation: txtNama.SetFocus: Exit Sub" & vbCrLf
    s = s & "    If Trim(cboSumber.Value) = """" Then MsgBox ""Sumber harus dipilih!"", vbExclamation: cboSumber.SetFocus: Exit Sub" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    Dim ws As Worksheet, nextRow As Long" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""No"", nextRow - 1" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Sourcing Date"", Date" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Kode Unik (copy value dari FPTK)"", txtKodeUnik.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Posisi"", txtPosisi.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Model Rekrutmen (Lihat di Sheet Flow Map Model)"", cboModel.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Rekruter"", txtRekruter.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Sumber Sourcing"", cboSumber.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Nama"", txtNama.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Nama Universitas/Sekolah (TOP 10)"", cboUniversitas.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Nama Universitas/Sekolah Lainnya"", txtUnivOther.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Jenjang Pendidikan"", cboJenjang.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Jurusan"", cboJurusan.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Tahun Lulus"", txtTahunLulus.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""IPK"", txtIPK.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Skor Bahasa Inggris"", txtEnglishScore.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""University Tier"", cboUnivTier.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""IPK Tier"", cboIPKTier.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Nomor HP"", txtHP.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Email"", txtEmail.Value" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Domisili"", txtDomisili.Value" & vbCrLf
    s = s & "    If InStr(1, LCase(cboSumber.Value), ""freelance"") > 0 Then" & vbCrLf
    s = s & "        SetByHeader ws, nextRow, ""Sourcin Freelance"", ""V""" & vbCrLf
    s = s & "        SetByHeader ws, nextRow, ""Tanggal Sourcing Freelance"", Date" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Sourcing HR"", ""V""" & vbCrLf
    s = s & "    SetByHeader ws, nextRow, ""Tanggal Sourcing"", Date" & vbCrLf
    s = s & "    ws.Columns(""A:BM"").AutoFit" & vbCrLf
    s = s & "    MsgBox ""Data tersimpan ke DB Sourcing!"" & vbCrLf & vbCrLf & txtNama.Value & "" - "" & txtPosisi.Value, vbInformation" & vbCrLf
    s = s & "    ClearFields" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ClearFields()" & vbCrLf
    s = s & "    txtNama.Value = """": cboJenjang.Value = """": cboUniversitas.Value = """": txtUnivOther.Value = """"" & vbCrLf
    s = s & "    cboJurusan.Value = """": txtTahunLulus.Value = """": txtIPK.Value = """": txtEnglishScore.Value = """"" & vbCrLf
    s = s & "    cboUnivTier.Value = """": cboIPKTier.Value = """": txtEmail.Value = """": txtHP.Value = """": txtDomisili.Value = """"" & vbCrLf
    s = s & "    mCVPath = """": lblCVStatus.Caption = ""Belum ada CV""" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set imgPreview.Picture = Nothing" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddLabel(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal b As Boolean = False, Optional ByVal s As Single = 8, Optional ByVal it As Boolean = False) As MSForms.Label" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.Label.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = s: o.Font.Bold = b: o.Font.Italic = it: o.BackStyle = 0" & vbCrLf
    s = s & "    If b Then o.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    Set AddLabel = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddTextBox(ByVal p As Object, ByVal n As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal lk As Boolean = False) As MSForms.TextBox" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.TextBox.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Font.Size = 8" & vbCrLf
    s = s & "    If lk Then o.Locked = True: o.BackColor = RGB(236, 240, 241)" & vbCrLf
    s = s & "    Set AddTextBox = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddCombo(ByVal p As Object, ByVal n As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As MSForms.ComboBox" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.ComboBox.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Font.Size = 8: o.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    Set AddCombo = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddButton(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, ByVal bc As Long) As MSForms.CommandButton" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.CommandButton.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = 8: o.Font.Bold = True: o.BackColor = bc: o.ForeColor = vbWhite" & vbCrLf
    s = s & "    Set AddButton = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddFrame(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As MSForms.Frame" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.Frame.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Caption = c" & vbCrLf
    s = s & "    o.BackColor = vbWhite: o.ForeColor = RGB(44, 62, 80): o.Font.Size = 8: o.Font.Bold = True" & vbCrLf
    s = s & "    Set AddFrame = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    Code_frmSourcing = s
End Function

Private Function Code_frmPasteCVText() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "Private mBuilt As Boolean" & vbCrLf
    s = s & "Private txtRaw As MSForms.TextBox" & vbCrLf
    s = s & "Private WithEvents btnClipboard As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnParse As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnCancel As MSForms.CommandButton" & vbCrLf
    s = s & "Private lblInfo As MSForms.Label" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildUI()" & vbCrLf
    s = s & "    If mBuilt Then Exit Sub" & vbCrLf
    s = s & "    mBuilt = True" & vbCrLf
    s = s & "    Me.Caption = ""PASTE TEXT CV / JOBSTREET / LINKEDIN""" & vbCrLf
    s = s & "    Me.Width = 560" & vbCrLf
    s = s & "    Me.Height = 420" & vbCrLf
    s = s & "    Me.BackColor = RGB(245, 245, 245)" & vbCrLf
    s = s & "    AddLabel Me, ""lblTitle"", ""Paste hasil copy dari CV / Jobstreet / LinkedIn di bawah ini"", 15, 12, 520, 18, True, 10" & vbCrLf
    s = s & "    Set txtRaw = Me.Controls.Add(""Forms.TextBox.1"", ""txtRaw"", True)" & vbCrLf
    s = s & "    txtRaw.Left = 15: txtRaw.Top = 40: txtRaw.Width = 520: txtRaw.Height = 270" & vbCrLf
    s = s & "    txtRaw.MultiLine = True" & vbCrLf
    s = s & "    txtRaw.ScrollBars = fmScrollBarsVertical" & vbCrLf
    s = s & "    txtRaw.EnterKeyBehavior = True" & vbCrLf
    s = s & "    txtRaw.WordWrap = True" & vbCrLf
    s = s & "    txtRaw.Font.Size = 8" & vbCrLf
    s = s & "    Set lblInfo = AddLabel(Me, ""lblInfo"", ""Tips: buka Jobstreet/LinkedIn/CV > Ctrl+A/Ctrl+C > klik Ambil Clipboard atau paste manual di box ini."", 15, 318, 520, 28, False, 8)" & vbCrLf
    s = s & "    lblInfo.ForeColor = RGB(127, 140, 141)" & vbCrLf
    s = s & "    Set btnClipboard = AddButton(Me, ""btnClipboard"", ""Ambil Clipboard"", 80, 355, 125, 30, RGB(52, 152, 219))" & vbCrLf
    s = s & "    Set btnParse = AddButton(Me, ""btnParse"", ""Parse ke Form"", 220, 355, 125, 30, RGB(39, 174, 96))" & vbCrLf
    s = s & "    Set btnCancel = AddButton(Me, ""btnCancel"", ""Cancel"", 360, 355, 90, 30, RGB(231, 76, 60))" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnClipboard_Click()" & vbCrLf
    s = s & "    Dim t As String" & vbCrLf
    s = s & "    t = GetClipboardTextSafe()" & vbCrLf
    s = s & "    If Len(t) = 0 Then" & vbCrLf
    s = s & "        MsgBox ""Clipboard kosong. Copy text dulu dari Jobstreet/LinkedIn/CV."", vbExclamation" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        txtRaw.Value = t" & vbCrLf
    s = s & "        lblInfo.Caption = ""Clipboard terbaca: "" & Len(t) & "" karakter.""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnParse_Click()" & vbCrLf
    s = s & "    If Len(Trim(txtRaw.Value)) = 0 Then MsgBox ""Paste text dulu."", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    If frmEditSourcing.Visible Then" & vbCrLf
    s = s & "        frmEditSourcing.ApplyManualText txtRaw.Value, ""Manual Paste - Jobstreet/LinkedIn/CV""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        frmSourcing.ApplyManualText txtRaw.Value, ""Manual Paste - Jobstreet/LinkedIn/CV""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddLabel(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal b As Boolean = False, Optional ByVal s As Single = 8) As MSForms.Label" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.Label.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = s: o.Font.Bold = b: o.BackStyle = 0" & vbCrLf
    s = s & "    If b Then o.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    Set AddLabel = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddButton(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, ByVal bc As Long) As MSForms.CommandButton" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.CommandButton.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = 8: o.Font.Bold = True: o.BackColor = bc: o.ForeColor = vbWhite" & vbCrLf
    s = s & "    Set AddButton = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    Code_frmPasteCVText = s
End Function

Private Function Code_frmSourcingList() As String
    Dim s As String
    s = ""
    s = s & "Option Explicit" & vbCrLf
    s = s & "Private mBuilt As Boolean" & vbCrLf
    s = s & "Private WithEvents cboPosisi As MSForms.ComboBox, WithEvents cboSumber As MSForms.ComboBox, WithEvents cboModel As MSForms.ComboBox" & vbCrLf
    s = s & "Private lstData As MSForms.ListBox, lblSummary As MSForms.Label" & vbCrLf
    s = s & "Private WithEvents btnDetail As MSForms.CommandButton, WithEvents btnEdit As MSForms.CommandButton, WithEvents btnDelete As MSForms.CommandButton, WithEvents btnExport As MSForms.CommandButton, WithEvents btnRefresh As MSForms.CommandButton" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    BuildUI" & vbCrLf
    s = s & "    LoadFilters" & vbCrLf
    s = s & "    LoadData" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub BuildUI()" & vbCrLf
    s = s & "    If mBuilt Then Exit Sub" & vbCrLf
    s = s & "    mBuilt = True" & vbCrLf
    s = s & "    Me.Caption = ""DATABASE SOURCING""" & vbCrLf
    s = s & "    Me.Width = 640" & vbCrLf
    s = s & "    Me.Height = 500" & vbCrLf
    s = s & "    Me.BackColor = RGB(245, 245, 245)" & vbCrLf
    s = s & "    AddLabel Me, ""lblFilter"", ""Filter:"", 20, 15, 100, 18, True, 9" & vbCrLf
    s = s & "    AddLabel Me, ""lblPosisi"", ""Posisi:"", 20, 42, 50, 18" & vbCrLf
    s = s & "    Set cboPosisi = AddCombo(Me, ""cboPosisi"", 75, 40, 210, 22)" & vbCrLf
    s = s & "    AddLabel Me, ""lblSumber"", ""Sumber:"", 310, 42, 50, 18" & vbCrLf
    s = s & "    Set cboSumber = AddCombo(Me, ""cboSumber"", 365, 40, 200, 22)" & vbCrLf
    s = s & "    AddLabel Me, ""lblModel"", ""Model:"", 20, 72, 50, 18" & vbCrLf
    s = s & "    Set cboModel = AddCombo(Me, ""cboModel"", 75, 70, 210, 22)" & vbCrLf
    s = s & "    Set lstData = Me.Controls.Add(""Forms.ListBox.1"", ""lstData"", True)" & vbCrLf
    s = s & "    lstData.Left = 20: lstData.Top = 105: lstData.Width = 590: lstData.Height = 280: lstData.Font.Size = 8" & vbCrLf
    s = s & "    lstData.ColumnCount = 6" & vbCrLf
    s = s & "    lstData.ColumnWidths = ""30;140;190;95;80;0""" & vbCrLf
    s = s & "    Set btnDetail = AddButton(Me, ""btnDetail"", ""Detail"", 20, 400, 75, 30, RGB(52, 152, 219))" & vbCrLf
    s = s & "    Set btnEdit = AddButton(Me, ""btnEdit"", ""Edit"", 105, 400, 75, 30, RGB(243, 156, 18))" & vbCrLf
    s = s & "    Set btnDelete = AddButton(Me, ""btnDelete"", ""Hapus"", 190, 400, 75, 30, RGB(231, 76, 60))" & vbCrLf
    s = s & "    Set btnExport = AddButton(Me, ""btnExport"", ""Export"", 275, 400, 75, 30, RGB(52, 152, 219))" & vbCrLf
    s = s & "    Set btnRefresh = AddButton(Me, ""btnRefresh"", ""Refresh"", 360, 400, 80, 30, RGB(39, 174, 96))" & vbCrLf
    s = s & "    Set lblSummary = AddLabel(Me, ""lblSummary"", ""Total: 0 kandidat"", 20, 445, 590, 20, True, 9)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadFilters()" & vbCrLf
    s = s & "    Dim v As Variant" & vbCrLf
    s = s & "    cboPosisi.Clear: cboSumber.Clear: cboModel.Clear" & vbCrLf
    s = s & "    cboPosisi.AddItem ""Semua Posisi""" & vbCrLf
    s = s & "    cboSumber.AddItem ""Semua Sumber""" & vbCrLf
    s = s & "    cboModel.AddItem ""Semua Model""" & vbCrLf
    s = s & "    For Each v In Array(""Jobstreet"", ""LinkedIn"", ""Google Form"", ""Referensi User"", ""Referensi Karyawan"", ""Campus Hiring"", ""Walk-in Interview"", ""Database Internal"", ""Freelance""): cboSumber.AddItem v: Next v" & vbCrLf
    s = s & "    For Each v In Array(""Model 1"", ""Model 2"", ""Model 3"", ""Model 4""): cboModel.AddItem v: Next v" & vbCrLf
    s = s & "    LoadPosisiFilter" & vbCrLf
    s = s & "    cboPosisi.Value = ""Semua Posisi""" & vbCrLf
    s = s & "    cboSumber.Value = ""Semua Sumber""" & vbCrLf
    s = s & "    cboModel.Value = ""Semua Model""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadPosisiFilter()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    Dim ws As Worksheet, lastRow As Long, r As Long, p As String" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    s = s & "    For r = 2 To lastRow" & vbCrLf
    s = s & "        p = CStr(GetByHeader(ws, r, ""Posisi""))" & vbCrLf
    s = s & "        If Len(p) > 0 Then If Not ComboHasValue(cboPosisi, p) Then cboPosisi.AddItem p" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ComboHasValue(ByVal cbo As MSForms.ComboBox, ByVal val As String) As Boolean" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    For i = 0 To cbo.ListCount - 1" & vbCrLf
    s = s & "        If LCase(Trim(cbo.List(i))) = LCase(Trim(val)) Then ComboHasValue = True: Exit Function" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub RefreshList()" & vbCrLf
    s = s & "    LoadFilters" & vbCrLf
    s = s & "    LoadData" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadData()" & vbCrLf
    s = s & "    EnsureDBSourcing" & vbCrLf
    s = s & "    Dim ws As Worksheet, lastRow As Long, r As Long, cnt As Long, idx As Long" & vbCrLf
    s = s & "    Set ws = ThisWorkbook.Worksheets(""DB Sourcing"")" & vbCrLf
    s = s & "    lstData.Clear" & vbCrLf
    s = s & "    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row" & vbCrLf
    s = s & "    For r = 2 To lastRow" & vbCrLf
    s = s & "        If RowMatches(ws, r) Then" & vbCrLf
    s = s & "            lstData.AddItem """"" & vbCrLf
    s = s & "            idx = lstData.ListCount - 1" & vbCrLf
    s = s & "            cnt = cnt + 1" & vbCrLf
    s = s & "            lstData.List(idx, 0) = cnt" & vbCrLf
    s = s & "            lstData.List(idx, 1) = CStr(GetByHeader(ws, r, ""Nama""))" & vbCrLf
    s = s & "            lstData.List(idx, 2) = CStr(GetByHeader(ws, r, ""Posisi""))" & vbCrLf
    s = s & "            lstData.List(idx, 3) = CStr(GetByHeader(ws, r, ""Sumber Sourcing""))" & vbCrLf
    s = s & "            lstData.List(idx, 4) = CStr(GetByHeader(ws, r, ""Sourcing HR""))" & vbCrLf
    s = s & "            lstData.List(idx, 5) = r" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next r" & vbCrLf
    s = s & "    lblSummary.Caption = ""Total: "" & cnt & "" kandidat""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function RowMatches(ByVal ws As Worksheet, ByVal r As Long) As Boolean" & vbCrLf
    s = s & "    RowMatches = True" & vbCrLf
    s = s & "    If cboSumber.Value <> """" And cboSumber.Value <> ""Semua Sumber"" Then If CStr(GetByHeader(ws, r, ""Sumber Sourcing"")) <> cboSumber.Value Then RowMatches = False" & vbCrLf
    s = s & "    If cboModel.Value <> """" And cboModel.Value <> ""Semua Model"" Then If CStr(GetByHeader(ws, r, ""Model Rekrutmen (Lihat di Sheet Flow Map Model)"")) <> cboModel.Value Then RowMatches = False" & vbCrLf
    s = s & "    If cboPosisi.Value <> """" And cboPosisi.Value <> ""Semua Posisi"" Then If CStr(GetByHeader(ws, r, ""Posisi"")) <> cboPosisi.Value Then RowMatches = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function SelectedRow() As Long" & vbCrLf
    s = s & "    If lstData.ListIndex >= 0 Then SelectedRow = CLng(lstData.List(lstData.ListIndex, 5)) Else SelectedRow = 0" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboSumber_Change()" & vbCrLf
    s = s & "    LoadData" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboModel_Change()" & vbCrLf
    s = s & "    LoadData" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboPosisi_Change()" & vbCrLf
    s = s & "    LoadData" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnRefresh_Click()" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnExport_Click()" & vbCrLf
    s = s & "    ExportDBSourcingCSV" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnEdit_Click()" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    r = SelectedRow()" & vbCrLf
    s = s & "    If r = 0 Then MsgBox ""Pilih kandidat yang mau diedit!"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    frmEditSourcing.LoadRecord r" & vbCrLf
    s = s & "    frmEditSourcing.Show vbModal" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    btnDetail_Click" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnDetail_Click()" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    r = SelectedRow()" & vbCrLf
    s = s & "    If r = 0 Then MsgBox ""Pilih kandidat yang mau dilihat detailnya!"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    frmSourcingDetail.LoadRecord r" & vbCrLf
    s = s & "    frmSourcingDetail.Show vbModal" & vbCrLf
    s = s & "    RefreshList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnDelete_Click()" & vbCrLf
    s = s & "    Dim r As Long" & vbCrLf
    s = s & "    r = SelectedRow()" & vbCrLf
    s = s & "    If r = 0 Then MsgBox ""Pilih kandidat!"", vbExclamation: Exit Sub" & vbCrLf
    s = s & "    If MsgBox(""Hapus data ini?"", vbQuestion + vbYesNo) = vbYes Then ThisWorkbook.Worksheets(""DB Sourcing"").Rows(r).Delete: LoadData" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddLabel(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, Optional ByVal b As Boolean = False, Optional ByVal s As Single = 8) As MSForms.Label" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.Label.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = s: o.Font.Bold = b: o.BackStyle = 0" & vbCrLf
    s = s & "    If b Then o.ForeColor = RGB(44, 62, 80)" & vbCrLf
    s = s & "    Set AddLabel = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddCombo(ByVal p As Object, ByVal n As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As MSForms.ComboBox" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.ComboBox.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Font.Size = 8: o.Style = fmStyleDropDownList" & vbCrLf
    s = s & "    Set AddCombo = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function AddButton(ByVal p As Object, ByVal n As String, ByVal c As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single, ByVal bc As Long) As MSForms.CommandButton" & vbCrLf
    s = s & "    Dim o As Object" & vbCrLf
    s = s & "    Set o = p.Controls.Add(""Forms.CommandButton.1"", n, True)" & vbCrLf
    s = s & "    o.Left = l: o.Top = t: o.Width = w: o.Height = h: o.Caption = c" & vbCrLf
    s = s & "    o.Font.Size = 8: o.Font.Bold = True: o.BackColor = bc: o.ForeColor = vbWhite" & vbCrLf
    s = s & "    Set AddButton = o" & vbCrLf
    s = s & "End Function" & vbCrLf
    Code_frmSourcingList = s
End Function

'===========================
' MASTER UNIVERSITY TIER MAP
'===========================

Public Function GetUniversityTier(ByVal univ As String) As String
    Dim u As String
    u = LCase(Trim(univ))

    Select Case True

        Case InStr(u, "universitas indonesia") > 0 _
          Or InStr(u, "ugm") > 0 _
          Or InStr(u, "institut teknologi bandung") > 0
            GetUniversityTier = "Top 3 PTN"
            Exit Function

        Case InStr(u, "universitas airlangga") > 0 _
          Or InStr(u, "universitas padjadjaran") > 0 _
          Or InStr(u, "universitas diponegoro") > 0 _
          Or InStr(u, "universitas brawijaya") > 0 _
          Or InStr(u, "ipb") > 0 _
          Or InStr(u, "universitas sebelas maret") > 0
            GetUniversityTier = "Top 10 PTN"
            Exit Function

        Case InStr(u, "telkom university") > 0 _
          Or InStr(u, "binus") > 0 _
          Or InStr(u, "universitas gunadarma") > 0 _
          Or InStr(u, "mercu buana") > 0
            GetUniversityTier = "Top 10 PTS"
            Exit Function

        Case Else
            GetUniversityTier = "Lainnya"
    End Select
End Function


'===========================
' MAJOR NORMALIZATION
'===========================

Public Function NormalizeMajor(ByVal major As String) As String
    Dim m As String
    m = LCase(Trim(major))

    Select Case True

        Case InStr(m, "account") > 0
            NormalizeMajor = "Accounting"

        Case InStr(m, "management") > 0 Or InStr(m, "manajemen") > 0
            NormalizeMajor = "Management"

        Case InStr(m, "finance") > 0 Or InStr(m, "keuangan") > 0
            NormalizeMajor = "Finance"

        Case InStr(m, "marketing") > 0 Or InStr(m, "pemasaran") > 0
            NormalizeMajor = "Marketing"

        Case InStr(m, "computer") > 0 Or InStr(m, "informatika") > 0 Or InStr(m, "information") > 0
            NormalizeMajor = "Information Technology"

        Case InStr(m, "industrial") > 0 Or InStr(m, "industri") > 0
            NormalizeMajor = "Industrial Engineering"

        Case InStr(m, "economics") > 0 Or InStr(m, "ekonomi") > 0
            NormalizeMajor = "Economics"

        Case Else
            NormalizeMajor = major
    End Select
End Function



