Attribute VB_Name = "m03_SheetManagement"
Option Explicit

Public Sub EnsureSupportSheets()
    GetOrCreateWorksheet "Settings"
    GetOrCreateWorksheet "Progress"
    GetOrCreateWorksheet "Run_Log"
    GetOrCreateWorksheet "DB Kode Posisi"
    GetOrCreateWorksheet "FPTK"
    GetOrCreateWorksheet "DB Sourcing"
    GetOrCreateWorksheet "Monitoring Sourcing"
    GetOrCreateWorksheet "Blacklist Candidate"
    GetOrCreateWorksheet "funneling"
    GetOrCreateWorksheet "funneling - Updated"
    GetOrCreateWorksheet "Schema_Config"

    ' Terapkan schema setelah seluruh sheet inti tersedia.
    SchemaB02PrepareAll

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")
    If Trim(CStr(ws.Range("A1").Value)) = "" Then
        ws.Range("A1:B1").Value = Array("Setting", "Value")
    End If

    Set ws = ThisWorkbook.Worksheets("Run_Log")
    If Trim(CStr(ws.Range("A1").Value)) = "" Then
        ws.Range("A1:D1").Value = Array("RunTime", "Status", "Notes", "User")
    End If
End Sub

Public Function GetOrCreateWorksheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateWorksheet = ws
End Function

Public Function GetSheetByPossibleNames(wb As Object, ParamArray sheetNames() As Variant) As Object
    Dim nm As Variant, ws As Object
    For Each nm In sheetNames
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then Set GetSheetByPossibleNames = ws: Exit Function
    Next nm
    For Each ws In wb.Worksheets
        For Each nm In sheetNames
            If LCase(Trim(CStr(ws.Name))) = LCase(Trim(CStr(nm))) Then
                Set GetSheetByPossibleNames = ws: Exit Function
            End If
        Next nm
    Next ws
    Set GetSheetByPossibleNames = Nothing
End Function

Public Function SheetExists(wb As Object, sheetName As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

