Attribute VB_Name = "modFPTKCompatibilityCleanup"
Option Explicit

' ============================================================
' COMPATIBILITY CLEANUP
'
' Dipakai untuk code lama yang masih memanggil:
' CleanupDuplicateFilterKategorisasiFPTK
'
' Letakkan di STANDARD MODULE.
' Contoh nama module: modFPTKCompatibilityCleanup
' ============================================================

Public Sub CleanupDuplicateFilterKategorisasiFPTK()

    Dim ws As Worksheet
    Dim primaryColumn As Long
    Dim duplicateColumn As Long
    Dim lastRow As Long

    On Error GoTo ErrorHandler

    Set ws = Nothing

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("FPTK")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then Exit Sub

    primaryColumn = FindFilterCategoryHeaderColumn( _
        ws, False)

    If primaryColumn = 0 Then

        primaryColumn = _
            FindFilterCategoryHeaderColumn( _
                ws, True)

    End If

    If primaryColumn = 0 Then Exit Sub

    lastRow = LastUsedRowFPTKCleanup(ws)

    Do

        duplicateColumn = _
            FindDuplicateFilterCategoryColumn( _
                ws, primaryColumn)

        If duplicateColumn = 0 Then Exit Do

        MergeFilterCategoryValues _
            ws, _
            primaryColumn, _
            duplicateColumn, _
            lastRow

        ws.Columns(duplicateColumn).Delete _
            Shift:=xlToLeft

        If duplicateColumn < primaryColumn Then
            primaryColumn = primaryColumn - 1
        End If

    Loop

    ws.Cells( _
        1, _
        primaryColumn).Value = _
        "Filter Kategorisasi FPTK"

    Exit Sub

ErrorHandler:

    MsgBox _
        "Cleanup Filter Kategorisasi FPTK gagal." & _
        vbCrLf & _
        "Error " & CStr(Err.Number) & _
        ": " & Err.Description, _
        vbExclamation, _
        "FPTK Cleanup"

End Sub

Private Function FindFilterCategoryHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal acceptNumberedHeader As Boolean) As Long

    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim headerKey As String

    lastColumn = LastUsedHeaderColumnFPTKCleanup(ws)

    For columnIndex = 1 To lastColumn

        headerKey = NormalizeFPTKCleanupHeader( _
            ws.Cells(1, columnIndex).Value2)

        If headerKey = _
           "FILTERKATEGORISASIFPTK" Then

            FindFilterCategoryHeaderColumn = _
                columnIndex

            Exit Function

        End If

        If acceptNumberedHeader Then

            If headerKey Like _
               "FILTERKATEGORISASIFPTK#" Or _
               headerKey Like _
               "FILTERKATEGORISASIFPTK##" Then

                ws.Cells( _
                    1, _
                    columnIndex).Value = _
                    "Filter Kategorisasi FPTK"

                FindFilterCategoryHeaderColumn = _
                    columnIndex

                Exit Function

            End If

        End If

    Next columnIndex

End Function

Private Function FindDuplicateFilterCategoryColumn( _
    ByVal ws As Worksheet, _
    ByVal primaryColumn As Long) As Long

    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim headerKey As String

    lastColumn = LastUsedHeaderColumnFPTKCleanup(ws)

    For columnIndex = lastColumn To 1 Step -1

        If columnIndex <> primaryColumn Then

            headerKey = NormalizeFPTKCleanupHeader( _
                ws.Cells(1, columnIndex).Value2)

            If headerKey = _
               "FILTERKATEGORISASIFPTK" Or _
               headerKey Like _
               "FILTERKATEGORISASIFPTK#" Or _
               headerKey Like _
               "FILTERKATEGORISASIFPTK##" Then

                FindDuplicateFilterCategoryColumn = _
                    columnIndex

                Exit Function

            End If

        End If

    Next columnIndex

End Function

Private Sub MergeFilterCategoryValues( _
    ByVal ws As Worksheet, _
    ByVal primaryColumn As Long, _
    ByVal duplicateColumn As Long, _
    ByVal lastRow As Long)

    Dim rowIndex As Long
    Dim primaryValue As Variant
    Dim duplicateValue As Variant

    For rowIndex = 2 To lastRow

        primaryValue = ws.Cells( _
            rowIndex, _
            primaryColumn).Value2

        duplicateValue = ws.Cells( _
            rowIndex, _
            duplicateColumn).Value2

        If Len(Trim$(CStr(primaryValue))) = 0 Then

            If Len(Trim$(CStr(duplicateValue))) > 0 Then

                ws.Cells( _
                    rowIndex, _
                    primaryColumn).Value2 = _
                    duplicateValue

            End If

        End If

    Next rowIndex

End Sub

Private Function LastUsedHeaderColumnFPTKCleanup( _
    ByVal ws As Worksheet) As Long

    LastUsedHeaderColumnFPTKCleanup = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End(xlToLeft).Column

    If LastUsedHeaderColumnFPTKCleanup < 1 Then
        LastUsedHeaderColumnFPTKCleanup = 1
    End If

End Function

Private Function LastUsedRowFPTKCleanup( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    On Error Resume Next

    Set foundCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    On Error GoTo 0

    If foundCell Is Nothing Then

        LastUsedRowFPTKCleanup = 1

    Else

        LastUsedRowFPTKCleanup = _
            foundCell.Row

    End If

End Function

Private Function NormalizeFPTKCleanupHeader( _
    ByVal rawValue As Variant) As String

    Dim sourceText As String
    Dim resultText As String
    Dim characterIndex As Long
    Dim currentCharacter As String
    Dim characterCode As Long

    If IsError(rawValue) Or _
       IsNull(rawValue) Or _
       IsEmpty(rawValue) Then

        Exit Function

    End If

    sourceText = UCase$(CStr(rawValue))

    For characterIndex = 1 To Len(sourceText)

        currentCharacter = Mid$( _
            sourceText, _
            characterIndex, _
            1)

        characterCode = AscW( _
            currentCharacter)

        If characterCode >= 48 And _
           characterCode <= 57 Then

            resultText = resultText & _
                currentCharacter

        ElseIf characterCode >= 65 And _
               characterCode <= 90 Then

            resultText = resultText & _
                currentCharacter

        End If

    Next characterIndex

    NormalizeFPTKCleanupHeader = resultText

End Function





