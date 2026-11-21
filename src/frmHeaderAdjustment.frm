Attribute VB_Name = "frmHeaderAdjustment"
Attribute VB_Base = "0{42D3D51C-B32A-4A17-BF1D-5939DDDD47E8}{7F90DB5A-A940-4F79-BB6D-2C8C45C64DD8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private mBuilt As Boolean

Private WithEvents cboDatabase As MSForms.ComboBox
Attribute cboDatabase.VB_VarHelpID = -1
Private WithEvents cboAction As MSForms.ComboBox
Attribute cboAction.VB_VarHelpID = -1
Private cboField As MSForms.ComboBox
Private txtHeader As MSForms.TextBox
Private txtPosition As MSForms.TextBox
Private cboDataType As MSForms.ComboBox
Private cboControlType As MSForms.ComboBox
Private txtDropdown As MSForms.TextBox

Private chkShowForm As MSForms.CheckBox
Private chkRequired As MSForms.CheckBox
Private chkReadOnly As MSForms.CheckBox
Private chkSearchable As MSForms.CheckBox

Private WithEvents btnApply As MSForms.CommandButton
Attribute btnApply.VB_VarHelpID = -1
Private WithEvents btnRefresh As MSForms.CommandButton
Attribute btnRefresh.VB_VarHelpID = -1
Private WithEvents btnClose As MSForms.CommandButton
Attribute btnClose.VB_VarHelpID = -1
Private Sub UserForm_Initialize()

    BuildUI
    LoadDatabases
    LoadActions
    LoadOptions

    If cboDatabase.ListCount > 0 Then
        cboDatabase.listIndex = 0
    End If

    If cboAction.ListCount > 0 Then
        cboAction.listIndex = 0
    End If

    RefreshFieldList

End Sub

Private Sub BuildUI()

    If mBuilt Then Exit Sub
    mBuilt = True

    Me.Caption = "UPDATE ADJUSTMENT HEADER - SAFE OVERLAY"
    Me.Width = 540
    Me.Height = 455
    Me.backColor = RGB(245, 245, 245)

    AddTitle _
        "HEADER & FORM ADJUSTMENT", _
        12, 10, 500, 24

    AddLabel _
        "Database", _
        18, 52, 115, 18

    Set cboDatabase = AddCombo( _
        "cboDatabase", _
        145, 50, 350, 20)

    AddLabel _
        "Tindakan", _
        18, 82, 115, 18

    Set cboAction = AddCombo( _
        "cboAction", _
        145, 80, 350, 20)

    AddLabel _
        "Field/Header", _
        18, 112, 115, 18

    Set cboField = AddCombo( _
        "cboField", _
        145, 110, 350, 20)

    AddLabel _
        "Nama Header Baru", _
        18, 142, 115, 18

    Set txtHeader = addText( _
        "txtHeader", _
        145, 140, 350, 20)

    AddLabel _
        "Posisi Kolom", _
        18, 172, 115, 18

    Set txtPosition = addText( _
        "txtPosition", _
        145, 170, 90, 20)

    AddLabel _
        "Tipe Data", _
        250, 172, 75, 18

    Set cboDataType = AddCombo( _
        "cboDataType", _
        325, 170, 170, 20)

    AddLabel _
        "Tipe Input", _
        18, 202, 115, 18

    Set cboControlType = AddCombo( _
        "cboControlType", _
        145, 200, 160, 20)

    AddLabel _
        "Pilihan Dropdown", _
        18, 232, 115, 18

    Set txtDropdown = addText( _
        "txtDropdown", _
        145, 230, 350, 20)

    txtDropdown.ControlTipText = _
        "Pisahkan pilihan dengan koma atau titik koma."

    Set chkShowForm = AddCheck( _
        "chkShowForm", _
        "Tampilkan di form", _
        145, 265, 145, 18)

    Set chkRequired = AddCheck( _
        "chkRequired", _
        "Wajib diisi", _
        310, 265, 120, 18)

    Set chkReadOnly = AddCheck( _
        "chkReadOnly", _
        "Read only", _
        145, 290, 120, 18)

    Set chkSearchable = AddCheck( _
        "chkSearchable", _
        "Masuk pencarian", _
        310, 290, 145, 18)

    AddLabel _
        "Core field: rename/move boleh, nonaktifkan tidak boleh.", _
        18, 320, 475, 18

    Set btnRefresh = AddButton( _
        "btnRefresh", _
        "REFRESH", _
        80, 355, 105, 30, _
        RGB(52, 152, 219))

    Set btnApply = AddButton( _
        "btnApply", _
        "APPLY ADJUSTMENT", _
        195, 355, 170, 30, _
        RGB(39, 174, 96))

    Set btnClose = AddButton( _
        "btnClose", _
        "TUTUP", _
        375, 355, 105, 30, _
        RGB(231, 76, 60))

End Sub

Private Sub LoadDatabases()

    Dim databaseKeys As Variant
    Dim databaseItem As Variant

    databaseKeys = SchemaGetDatabaseKeys()

    cboDatabase.Clear

    For Each databaseItem In databaseKeys
        cboDatabase.AddItem CStr(databaseItem)
    Next databaseItem

End Sub

Private Sub LoadActions()

    cboAction.Clear
    cboAction.AddItem "TAMBAH"
    cboAction.AddItem "UBAH NAMA"
    cboAction.AddItem "NONAKTIFKAN"
    cboAction.AddItem "AKTIFKAN"
    cboAction.AddItem "PINDAH URUTAN"

End Sub

Private Sub LoadOptions()

    cboDataType.Clear
    cboDataType.AddItem "Text"
    cboDataType.AddItem "Number"
    cboDataType.AddItem "Date"
    cboDataType.AddItem "Boolean"

    cboControlType.Clear
    cboControlType.AddItem "TextBox"
    cboControlType.AddItem "ComboBox"

    cboDataType.Value = "Text"
    cboControlType.Value = "TextBox"

    chkShowForm.Value = True
    chkRequired.Value = False
    chkReadOnly.Value = False
    chkSearchable.Value = True

End Sub

Private Sub cboDatabase_Change()

    RefreshFieldList

End Sub

Private Sub cboAction_Change()

    RefreshFieldList
    UpdateActionState

End Sub

Private Sub RefreshFieldList()

    Dim fieldList As Collection
    Dim item As Variant
    Dim parts() As String
    Dim displayText As String

    cboField.Clear

    If Len(Trim$(cboDatabase.Value)) = 0 Then
        Exit Sub
    End If

    Set fieldList = SchemaGetFieldList( _
        cboDatabase.Value, True)

    For Each item In fieldList

        parts = Split(CStr(item), "|")

        If UBound(parts) >= 3 Then

            displayText = _
                parts(0) & " | " & _
                parts(1) & " | " & _
                parts(2) & " | Pos " & _
                parts(3)

            cboField.AddItem displayText

        End If

    Next item

    If cboField.ListCount > 0 Then
        cboField.listIndex = 0
    End If

End Sub

Private Sub UpdateActionState()

    Dim actionText As String

    actionText = UCase$( _
        Trim$(cboAction.Value))

    txtHeader.Enabled = _
        (actionText = "TAMBAH" Or _
         actionText = "UBAH NAMA")

    txtPosition.Enabled = _
        (actionText = "TAMBAH" Or _
         actionText = "PINDAH URUTAN")

    cboDataType.Enabled = _
        (actionText = "TAMBAH")

    cboControlType.Enabled = _
        (actionText = "TAMBAH")

    txtDropdown.Enabled = _
        (actionText = "TAMBAH")

    chkShowForm.Enabled = _
        (actionText = "TAMBAH")

    chkRequired.Enabled = _
        (actionText = "TAMBAH")

    chkReadOnly.Enabled = _
        (actionText = "TAMBAH")

    chkSearchable.Enabled = _
        (actionText = "TAMBAH")

    cboField.Enabled = _
        (actionText <> "TAMBAH")

End Sub

Private Function SelectedFieldID() As String

    Dim parts() As String

    If Len(Trim$(cboField.Value)) = 0 Then
        Exit Function
    End If

    parts = Split(CStr(cboField.Value), "|")

    If UBound(parts) < 0 Then Exit Function

    SelectedFieldID = Trim$(parts(0))

End Function

Private Sub btnApply_Click()

    Dim actionText As String
    Dim fieldID As String
    Dim requestedPosition As Long
    Dim confirmation As VbMsgBoxResult

    On Error GoTo ErrorHandler

    actionText = Trim$( _
        cboAction.Value)

    fieldID = SelectedFieldID()

    requestedPosition = CLng( _
        val(txtPosition.Value))

    If Len(Trim$(cboDatabase.Value)) = 0 Then

        MsgBox _
            "Pilih database terlebih dahulu.", _
            vbExclamation

        Exit Sub

    End If

    If Len(actionText) = 0 Then

        MsgBox _
            "Pilih tindakan adjustment.", _
            vbExclamation

        Exit Sub

    End If

    If UCase$(actionText) <> "TAMBAH" And _
       Len(fieldID) = 0 Then

        MsgBox _
            "Pilih field/header yang akan diubah.", _
            vbExclamation

        Exit Sub

    End If

    If UCase$(actionText) = "TAMBAH" And _
       Len(Trim$(txtHeader.Value)) = 0 Then

        MsgBox _
            "Isi Nama Header Baru.", _
            vbExclamation

        Exit Sub

    End If

    If UCase$(actionText) = "UBAH NAMA" And _
       Len(Trim$(txtHeader.Value)) = 0 Then

        MsgBox _
            "Isi Nama Header Baru.", _
            vbExclamation

        Exit Sub

    End If

    If UCase$(actionText) = "PINDAH URUTAN" And _
       requestedPosition <= 0 Then

        MsgBox _
            "Isi posisi kolom dengan angka lebih dari 0.", _
            vbExclamation

        Exit Sub

    End If

    confirmation = MsgBox( _
        "Terapkan adjustment ini?" & _
        vbCrLf & vbCrLf & _
        "Database: " & cboDatabase.Value & _
        vbCrLf & _
        "Tindakan: " & actionText & _
        vbCrLf & _
        "Field: " & fieldID & _
        vbCrLf & _
        "Header baru: " & txtHeader.Value, _
        vbYesNo + vbQuestion, _
        "Konfirmasi Adjustment")

    If confirmation <> vbYes Then
        Exit Sub
    End If

    SchemaApplyRequest _
        cboDatabase.Value, _
        actionText, _
        fieldID, _
        Trim$(txtHeader.Value), _
        requestedPosition, _
        cboDataType.Value, _
        chkShowForm.Value, _
        chkRequired.Value, _
        chkReadOnly.Value, _
        chkSearchable.Value, _
        cboControlType.Value, _
        Trim$(txtDropdown.Value)

    MsgBox _
        "Adjustment berhasil diterapkan." & _
        vbCrLf & _
        "Tidak ada sheet backup baru yang dibuat." & _
        vbCrLf & _
        "Riwayat tersimpan di Schema_Change_Log.", _
        vbInformation, _
        "Header Adjustment"

    txtHeader.Value = ""
    txtPosition.Value = ""

    RefreshFieldList

    Exit Sub

ErrorHandler:

    MsgBox _
        "Adjustment gagal." & _
        vbCrLf & _
        "Error " & CStr(Err.Number) & _
        ": " & Err.Description, _
        vbCritical, _
        "Header Adjustment"

End Sub

Private Sub btnRefresh_Click()

    EnsureSchemaConfig
    SchemaRefreshRegistryFromSheets
    RefreshFieldList

    MsgBox _
        "Schema registry sudah direfresh tanpa rebuild database.", _
        vbInformation

End Sub

Private Sub btnClose_Click()

    Unload Me

End Sub

Private Sub AddTitle( _
    ByVal captionText As String, _
    ByVal leftPosition As Single, _
    ByVal topPosition As Single, _
    ByVal controlWidth As Single, _
    ByVal controlHeight As Single)

    Dim result As Object

    Set result = Me.Controls.Add( _
        "Forms.Label.1", _
        "lblTitle", _
        True)

    With result
        .Left = leftPosition
        .Top = topPosition
        .Width = controlWidth
        .Height = controlHeight
        .Caption = captionText
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(44, 62, 80)
        .BackStyle = 0
        .TextAlign = fmTextAlignCenter
    End With

End Sub

Private Sub AddLabel( _
    ByVal captionText As String, _
    ByVal leftPosition As Single, _
    ByVal topPosition As Single, _
    ByVal controlWidth As Single, _
    ByVal controlHeight As Single)

    Dim result As Object

    Set result = Me.Controls.Add( _
        "Forms.Label.1", _
        "lbl_" & CStr(Me.Controls.Count), _
        True)

    With result
        .Left = leftPosition
        .Top = topPosition
        .Width = controlWidth
        .Height = controlHeight
        .Caption = captionText
        .Font.Size = 8
        .Font.Bold = True
        .BackStyle = 0
    End With

End Sub

Private Function addText( _
    ByVal controlName As String, _
    ByVal leftPosition As Single, _
    ByVal topPosition As Single, _
    ByVal controlWidth As Single, _
    ByVal controlHeight As Single) As MSForms.TextBox

    Dim result As Object

    Set result = Me.Controls.Add( _
        "Forms.TextBox.1", _
        controlName, _
        True)

    With result
        .Left = leftPosition
        .Top = topPosition
        .Width = controlWidth
        .Height = controlHeight
        .Font.Size = 8
    End With

    Set addText = result

End Function

Private Function AddCombo( _
    ByVal controlName As String, _
    ByVal leftPosition As Single, _
    ByVal topPosition As Single, _
    ByVal controlWidth As Single, _
    ByVal controlHeight As Single) As MSForms.ComboBox

    Dim result As Object

    Set result = Me.Controls.Add( _
        "Forms.ComboBox.1", _
        controlName, _
        True)

    With result
        .Left = leftPosition
        .Top = topPosition
        .Width = controlWidth
        .Height = controlHeight
        .Font.Size = 8
        .Style = fmStyleDropDownList
    End With

    Set AddCombo = result

End Function

Private Function AddCheck( _
    ByVal controlName As String, _
    ByVal captionText As String, _
    ByVal leftPosition As Single, _
    ByVal topPosition As Single, _
    ByVal controlWidth As Single, _
    ByVal controlHeight As Single) As MSForms.CheckBox

    Dim result As Object

    Set result = Me.Controls.Add( _
        "Forms.CheckBox.1", _
        controlName, _
        True)

    With result
        .Left = leftPosition
        .Top = topPosition
        .Width = controlWidth
        .Height = controlHeight
        .Caption = captionText
        .Font.Size = 8
        .BackStyle = 0
    End With

    Set AddCheck = result

End Function

Private Function AddButton( _
    ByVal controlName As String, _
    ByVal captionText As String, _
    ByVal leftPosition As Single, _
    ByVal topPosition As Single, _
    ByVal controlWidth As Single, _
    ByVal controlHeight As Single, _
    ByVal backColorValue As Long) As MSForms.CommandButton

    Dim result As Object

    Set result = Me.Controls.Add( _
        "Forms.CommandButton.1", _
        controlName, _
        True)

    With result
        .Left = leftPosition
        .Top = topPosition
        .Width = controlWidth
        .Height = controlHeight
        .Caption = captionText
        .Font.Size = 8
        .Font.Bold = True
        .backColor = backColorValue
        .ForeColor = vbWhite
    End With

    Set AddButton = result

End Function


