Attribute VB_Name = "m13_UtilityFunctions"
Option Explicit

Public Function SafeLongFinal(ByVal v As Variant) As Long
    If IsNumeric(v) Then
        SafeLongFinal = CLng(v)
    Else
        SafeLongFinal = CLng(val(CStr(v)))
    End If
End Function

Public Function MaxLong(valueA As Long, valueB As Long) As Long
    If valueA >= valueB Then MaxLong = valueA Else MaxLong = valueB
End Function

Public Function NormalizeKey(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Trim(s)
    NormalizeKey = UCase(s)
End Function

Public Function RowHasAnyValue(ws As Object, rowNumber As Long, firstCol As Long, lastCol As Long) As Boolean
    Dim c As Long
    For c = firstCol To lastCol
        If Trim(CStr(ws.Cells(rowNumber, c).Value)) <> "" Then RowHasAnyValue = True: Exit Function
    Next c
    RowHasAnyValue = False
End Function

Public Function GetLastUsedRowByValue(ws As Object) As Long
    Dim foundCell As Object
    On Error Resume Next
    Set foundCell = ws.Cells.Find("*", ws.Cells(1, 1), xlValues, xlPart, xlByRows, xlPrevious, False)
    On Error GoTo 0
    If foundCell Is Nothing Then GetLastUsedRowByValue = 1 Else GetLastUsedRowByValue = foundCell.Row
End Function

Public Function GetLastRowInColumnByValue(ws As Object, ColumnLetter As String) As Long
    Dim searchRange As Object
    Dim foundCell As Object
    Set searchRange = ws.Columns(ColumnLetter)
    On Error Resume Next
    Set foundCell = searchRange.Find("*", searchRange.Cells(1, 1), xlValues, xlPart, xlByRows, xlPrevious, False)
    On Error GoTo 0
    If foundCell Is Nothing Then GetLastRowInColumnByValue = 1 Else GetLastRowInColumnByValue = foundCell.Row
End Function

Public Function GetNextAppendRow(ws As Worksheet, firstDataRow As Long) As Long
    Dim lastRow As Long
    lastRow = GetLastUsedRowByValue(ws)
    If lastRow < firstDataRow Then GetNextAppendRow = firstDataRow Else GetNextAppendRow = lastRow + 1
End Function

Public Function GetLastUsedColumnInRow(ws As Object, rowNumber As Long) As Long
    Dim searchRange As Object
    Dim foundCell As Object

    Set searchRange = ws.rows(rowNumber)

    On Error Resume Next
    Set foundCell = searchRange.Find( _
        "*", _
        searchRange.Cells(1, 1), _
        xlValues, _
        xlPart, _
        xlByColumns, _
        xlPrevious, _
        False)
    On Error GoTo 0

    If foundCell Is Nothing Then
        GetLastUsedColumnInRow = 1
    Else
        GetLastUsedColumnInRow = foundCell.Column
    End If
End Function

Public Function GetLastUsedColumnByValue(ws As Object) As Long
    Dim foundCell As Object

    On Error Resume Next
    Set foundCell = ws.Cells.Find( _
        "*", _
        ws.Cells(1, 1), _
        xlValues, _
        xlPart, _
        xlByColumns, _
        xlPrevious, _
        False)
    On Error GoTo 0

    If foundCell Is Nothing Then
        GetLastUsedColumnByValue = 1
    Else
        GetLastUsedColumnByValue = foundCell.Column
    End If
End Function


