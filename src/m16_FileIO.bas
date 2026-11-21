Attribute VB_Name = "m16_FileIO"
Option Explicit

Public Function CountInputFiles( _
    ByVal folderPath As String) As Long

    ' folderPath yang diterima adalah folder final:
    ' InputFolderPath\NamaRecruiter\
    '
    ' Jangan baca setting SubFolders lagi di function ini.
    CountInputFiles = _
        CountFilesInFolderFinal( _
            folderPath)

End Function

Private Function CountFilesInFolderFinal( _
    ByVal folderPath As String) As Long

    Dim fso As Object
    Dim folderObject As Object
    Dim fileObject As Object

    Dim extensionText As String
    Dim fileCount As Long

    On Error GoTo SafeExit

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        GoTo SafeExit
    End If

    Set folderObject = _
        fso.GetFolder(folderPath)

    For Each fileObject In _
        folderObject.files

        If Left$( _
            CStr(fileObject.Name), _
            2) <> "~$" Then

            extensionText = LCase$( _
                fso.GetExtensionName( _
                    CStr(fileObject.Name)))

            Select Case extensionText

                Case "xlsx", "xlsm"

                    fileCount = _
                        fileCount + 1

            End Select

        End If

    Next fileObject

SafeExit:

    CountFilesInFolderFinal = _
        fileCount

    Set fileObject = Nothing
    Set folderObject = Nothing
    Set fso = Nothing

End Function

Public Function GetSourceExcelApp() As Object

    If gSourceApp Is Nothing Then

        Set gSourceApp = _
            CreateObject( _
                "Excel.Application")

        gSourceApp.Visible = False
        gSourceApp.DisplayAlerts = False
        gSourceApp.EnableEvents = False
        gSourceApp.AskToUpdateLinks = False
        gSourceApp.ScreenUpdating = False

        On Error Resume Next
        gSourceApp.Calculation = -4135
        On Error GoTo 0

    End If

    Set GetSourceExcelApp = _
        gSourceApp

End Function

Public Sub CloseSourceExcelApp()

    On Error Resume Next

    If Not gSourceApp Is Nothing Then

        gSourceApp.Quit
        Set gSourceApp = Nothing

    End If

    On Error GoTo 0

End Sub

Public Function OpenSourceWorkbook( _
    ByVal fullPath As String, _
    ByRef errorText As String) As Object

    Dim srcApp As Object

    On Error GoTo ErrHandler

    Set srcApp = _
        GetSourceExcelApp()

    Set OpenSourceWorkbook = _
        srcApp.Workbooks.Open( _
            fileName:=fullPath, _
            UpdateLinks:=0, _
            ReadOnly:=True, _
            AddToMru:=False, _
            IgnoreReadOnlyRecommended:=True, _
            Notify:=False)

    Exit Function

ErrHandler:

    errorText = Err.Description

    Set OpenSourceWorkbook = Nothing

End Function

Public Sub CloseSourceWorkbook( _
    ByVal wb As Object)

    On Error Resume Next

    wb.Close SaveChanges:=False

    On Error GoTo 0

End Sub


