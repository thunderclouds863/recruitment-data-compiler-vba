Attribute VB_Name = "m14_SettingsAndLog"
Option Explicit

Public Function GetSettingValue(settingName As String) As String
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet("Settings")
    If Trim(CStr(ws.Range("A1").Value)) = "" Then
        ws.Range("A1:B1").Value = Array("Setting", "Value")
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 1 To lastRow
        If Trim(CStr(ws.Cells(i, "A").Value)) = settingName Then
            GetSettingValue = Trim(CStr(ws.Cells(i, "B").Value))
            Exit Function
        End If
    Next i
    GetSettingValue = ""
End Function

Public Sub SetSettingValue(settingName As String, settingValue As String)
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet("Settings")
    If Trim(CStr(ws.Range("A1").Value)) = "" Then
        ws.Range("A1:B1").Value = Array("Setting", "Value")
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 1 To lastRow
        If Trim(CStr(ws.Cells(i, "A").Value)) = settingName Then
            ws.Cells(i, "B").Value = settingValue
            Exit Sub
        End If
    Next i
    
    If lastRow < 1 Then lastRow = 1
    ws.Cells(lastRow + 1, "A").Value = settingName
    ws.Cells(lastRow + 1, "B").Value = settingValue
End Sub

Public Sub AddRunLog(logStatus As String, notesText As String)
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\compile_log.txt"
    
    Dim f As Integer
    f = FreeFile()
    
    On Error Resume Next
    Open logPath For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & logStatus & vbTab & notesText & vbTab & Environ("Username")
    Close #f
    On Error GoTo 0
End Sub

