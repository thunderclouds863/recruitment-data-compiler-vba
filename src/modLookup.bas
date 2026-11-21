Attribute VB_Name = "modLookup"
Option Explicit

' =========================================================
' MODULE-LEVEL LOOKUP STATE
' =========================================================
Private wsDB As Worksheet
Private wsDropdown As Worksheet

Private colDBKode As Long
Private colDBPosition As Long
Private colDBLocation As Long
Private colDBBusinessUnit As Long
Private colDBDivisionChris As Long
Private colDBDepartmentChris As Long
Private colDBUser As Long
Private colDBIndirectUser As Long
Private colDBDirectorate As Long
Private colDBYear As Long

Private colDropPIC As Long
Private colDropKodePIC As Long
Private colDropBU As Long
Private colDropKodeBU As Long


'=========================================================
' modLookup CLEAN
' Aman dari:
' - lastRow not defined
' - NormalizeRecruiterName not defined
' - FormatAsTable ambiguous
' - wsDB/wsDropdown belum ke-set
' - colDB / colDrop belum ke-load
'=========================================================

'=========================================================
' LOOKUP POSITION
'=========================================================

Public Function LookupPosition() As Boolean
    SchemaB02Prepare "DB Kode Posisi"

    Dim i As Long
    Dim lastDataRow As Long
    Dim keyEmail As String
    Dim keyDB As String
    Dim ws As Worksheet

    LookupPosition = False

    Set ws = LKP_GetSheet("DB Kode Posisi", False)

    If ws Is Nothing Then
        Debug.Print "LookupPosition FAILED: Sheet DB Kode Posisi tidak ditemukan."
        Exit Function
    End If

    Set wsDB = ws
    LKP_EnsureDBColumns wsDB

    If colDBPosition = 0 Or colDBBusinessUnit = 0 Then
        Debug.Print "LookupPosition FAILED: Header POSITION / BUSINESS UNIT tidak ditemukan."
        Exit Function
    End If

    keyEmail = LKP_NormalizeCompare(gPosition) & "|" & _
               LKP_NormalizeCompare(gBusinessUnit) & "|" & _
               LKP_NormalizeCompare(gLocation)
    Debug.Print "KEY EMAIL = [" & keyEmail & "]"

    lastDataRow = LKP_LastRow(wsDB, colDBPosition)

    For i = 2 To lastDataRow

        keyDB = LKP_NormalizeCompare(wsDB.Cells(i, colDBPosition).Value) & "|" & _
                LKP_NormalizeCompare(wsDB.Cells(i, colDBBusinessUnit).Value) & "|" & _
                LKP_NormalizeCompare(wsDB.Cells(i, colDBLocation).Value)

        Debug.Print "Row : " & i
        Debug.Print "DB KEY : [" & keyDB & "]"

        If keyEmail = keyDB Then

            Debug.Print "MATCH DI ROW : " & i

            If colDBKode > 0 Then gPositionCode = CStr(wsDB.Cells(i, colDBKode).Value)
            If colDBDirectorate > 0 Then gDirectorate = CStr(wsDB.Cells(i, colDBDirectorate).Value)
            If colDBDivisionChris > 0 Then gDivision = CStr(wsDB.Cells(i, colDBDivisionChris).Value)
            If colDBDepartmentChris > 0 Then gDepartment = CStr(wsDB.Cells(i, colDBDepartmentChris).Value)
            If colDBUser > 0 Then gUser = CStr(wsDB.Cells(i, colDBUser).Value)
            If colDBIndirectUser > 0 Then gIndirectUser = CStr(wsDB.Cells(i, colDBIndirectUser).Value)

            LookupPosition = True
            Exit Function

        End If

    Next i

End Function

'=========================================================
' INSERT NEW POSITION
'=========================================================

Public Function InsertNewPosition() As Boolean
    SchemaB02Prepare "DB Kode Posisi"

    Dim lastDataRow As Long
    Dim newRow As Long
    Dim ws As Worksheet

    InsertNewPosition = False

    If Len(Trim(gPosition)) = 0 Or Len(Trim(gBusinessUnit)) = 0 Then
        Debug.Print "INSERT POSITION FAILED: Position or BusinessUnit is empty"
        Exit Function
    End If

    Set ws = LKP_GetSheet("DB Kode Posisi", True)

    If ws Is Nothing Then
        Debug.Print "INSERT POSITION FAILED: Sheet DB Kode Posisi tidak bisa dibuat."
        Exit Function
    End If

    Set wsDB = ws
    LKP_EnsureDBHeaders wsDB
    LKP_EnsureDBColumns wsDB

    If LookupPosition() = True Then
        Debug.Print "INSERT POSITION CANCELLED: Position already exists in DB"
        InsertNewPosition = True
        Exit Function
    End If

    lastDataRow = LKP_LastRow(wsDB, 1)
    newRow = lastDataRow + 1

    If Len(Trim(gPositionCode)) = 0 Then
        gPositionCode = CStr(GenerateNewPositionCode())
    End If

    With wsDB

        If colDBKode > 0 Then .Cells(newRow, colDBKode).Value = gPositionCode
        If colDBPosition > 0 Then .Cells(newRow, colDBPosition).Value = Trim(gPosition)
        If colDBLocation > 0 Then .Cells(newRow, colDBLocation).Value = Trim(gLocation)
        If colDBBusinessUnit > 0 Then .Cells(newRow, colDBBusinessUnit).Value = Trim(gBusinessUnit)
        If colDBDivisionChris > 0 Then .Cells(newRow, colDBDivisionChris).Value = Trim(gDivisionEmail)
        If colDBDepartmentChris > 0 Then .Cells(newRow, colDBDepartmentChris).Value = Trim(gDepartmentEmail)
        If colDBUser > 0 Then .Cells(newRow, colDBUser).Value = ""
        If colDBIndirectUser > 0 Then .Cells(newRow, colDBIndirectUser).Value = ""
        If colDBDirectorate > 0 Then .Cells(newRow, colDBDirectorate).Value = ""
        If colDBYear > 0 Then .Cells(newRow, colDBYear).Value = Year(Date)

    End With

    LKP_FormatAsTableSafe wsDB, "tblDBPosisi"

    Debug.Print "========================================"
    Debug.Print "NEW POSITION INSERTED at row " & newRow
    Debug.Print "  Position Code : " & gPositionCode
    Debug.Print "  Position      : " & gPosition
    Debug.Print "  Business Unit : " & gBusinessUnit
    Debug.Print "  Location      : " & gLocation
    Debug.Print "  Division      : " & gDivisionEmail
    Debug.Print "  Department    : " & gDepartmentEmail
    Debug.Print "  Directorate   : (kosong - isi manual)"
    Debug.Print "  User          : (kosong - isi manual)"
    Debug.Print "  Indirect User : (kosong - isi manual)"
    Debug.Print "========================================"

    InsertNewPosition = True

End Function

'=========================================================
' GENERATE NEW POSITION CODE
'=========================================================

Private Function GenerateNewPositionCode() As Long
    SchemaB02Prepare "DB Kode Posisi"

    Dim lastDataRow As Long
    Dim i As Long
    Dim rawCode As Variant
    Dim maxNumber As Long
    Dim currentNumber As Long
    Dim ws As Worksheet

    Set ws = LKP_GetSheet("DB Kode Posisi", True)

    If ws Is Nothing Then
        GenerateNewPositionCode = 1
        Exit Function
    End If

    Set wsDB = ws
    LKP_EnsureDBHeaders wsDB
    LKP_EnsureDBColumns wsDB

    If colDBKode = 0 Then
        GenerateNewPositionCode = 1
        Exit Function
    End If

    lastDataRow = LKP_LastRow(wsDB, colDBKode)
    maxNumber = 0

    For i = 2 To lastDataRow

        rawCode = wsDB.Cells(i, colDBKode).Value

        If Not IsEmpty(rawCode) Then
            If IsNumeric(rawCode) Then

                currentNumber = CLng(rawCode)

                If currentNumber > maxNumber Then
                    maxNumber = currentNumber
                End If

            End If
        End If

    Next i

    maxNumber = maxNumber + 1

    If maxNumber <= 0 Then maxNumber = 1

    GenerateNewPositionCode = maxNumber

    Debug.Print "Generated New Position Code: " & GenerateNewPositionCode

End Function

'=========================================================
' LOOKUP PIC
'=========================================================

Public Function LookupPIC() As Boolean

    Dim i As Long
    Dim lastDataRow As Long
    Dim emailName As String
    Dim recruiterName As String
    Dim ws As Worksheet

    LookupPIC = False

    Set ws = LKP_GetSheet("Kode Dropdown", False)

    If Not ws Is Nothing Then
        Set wsDropdown = ws
        LKP_EnsureDropdownColumns wsDropdown
    End If

    If Len(Trim(gPICEmail)) > 0 Then

        If InStr(1, gPICEmail, "@", vbTextCompare) > 0 Then
            emailName = Left(gPICEmail, InStr(1, gPICEmail, "@", vbTextCompare) - 1)
        Else
            emailName = gPICEmail
        End If

        emailName = LKP_NormalizeRecruiterName(emailName)

        If Not wsDropdown Is Nothing Then
            If colDropPIC > 0 And colDropKodePIC > 0 Then

                lastDataRow = LKP_LastRow(wsDropdown, colDropPIC)

                For i = 2 To lastDataRow

                    recruiterName = LKP_NormalizeRecruiterName(wsDropdown.Cells(i, colDropPIC).Value)

                    If Len(recruiterName) > 0 Then

                        If emailName Like "*" & recruiterName & "*" Or recruiterName Like "*" & emailName & "*" Then

                            gPIC = CStr(wsDropdown.Cells(i, colDropPIC).Value)
                            gPICCode = CStr(wsDropdown.Cells(i, colDropKodePIC).Value)

                            LookupPIC = True
                            Exit Function

                        End If

                    End If

                Next i

            End If
        End If

        ' Fallback tanpa sheet dropdown
        gPIC = LKP_PICNameFromAnyText(gPICEmail)
        gPICCode = LKP_PICCodeFromPICName(gPIC)

        If Len(Trim(gPIC)) > 0 And Len(Trim(gPICCode)) > 0 Then
            LookupPIC = True
            Exit Function
        End If

    End If

    ' Kalau gPIC sudah ada dari form/parser, isi kode PIC
    If Len(Trim(gPIC)) > 0 And Len(Trim(gPICCode)) = 0 Then
        gPICCode = LKP_PICCodeFromPICName(gPIC)
    End If

    If Len(Trim(gPIC)) > 0 And Len(Trim(gPICCode)) > 0 Then
        LookupPIC = True
    End If

End Function

'=========================================================
' LOOKUP BU
'=========================================================

Public Function LookupBU() As Boolean

    Dim i As Long
    Dim lastDataRow As Long
    Dim ws As Worksheet

    LookupBU = False

    If Len(Trim(gBusinessUnit)) = 0 Then Exit Function

    Set ws = LKP_GetSheet("Kode Dropdown", False)

    If Not ws Is Nothing Then

        Set wsDropdown = ws
        LKP_EnsureDropdownColumns wsDropdown

        If colDropBU > 0 And colDropKodeBU > 0 Then

            lastDataRow = LKP_LastRow(wsDropdown, colDropBU)

            For i = 2 To lastDataRow

                If LKP_NormalizeCompare(wsDropdown.Cells(i, colDropBU).Value) = _
                   LKP_NormalizeCompare(gBusinessUnit) Then

                    gBUCode = CStr(wsDropdown.Cells(i, colDropKodeBU).Value)
                    LookupBU = True
                    Exit Function

                End If

            Next i

        End If

    End If

    ' Fallback tanpa sheet dropdown
    gBUCode = LKP_BUCodeFromBUName(gBusinessUnit)

    If Len(Trim(gBUCode)) > 0 Then
        LookupBU = True
    End If

End Function

'=========================================================
' DETERMINE FILTER KATEGORISASI
'=========================================================

Public Function DetermineFilterKategori( _
    ByVal position As String, _
    Optional ByVal levelNum As Long = 0) As String

    Dim normalizedPosition As String

    normalizedPosition = LCase$( _
        Trim$(position))

    If Left$(normalizedPosition, 6) = "cimory" Or _
       Left$(normalizedPosition, 5) = "fresh" Then

        DetermineFilterKategori = "CLAP FGDP"

    Else

        Select Case levelNum

            Case 1, 2
                DetermineFilterKategori = "Level 1-2"

            Case 3
                DetermineFilterKategori = "Level 3"

            Case 4
                DetermineFilterKategori = "Level 4"

            Case Else
                DetermineFilterKategori = ""

        End Select

    End If

End Function

'=========================================================
' CALCULATE SLA
'=========================================================

Public Function CalculateSLA(ByVal levelNum As Long) As Long

    Select Case levelNum
        Case 1 To 3
            CalculateSLA = 30
        Case 4
            CalculateSLA = 45
        Case 5
            CalculateSLA = 60
        Case Else
            CalculateSLA = 30
    End Select

    Debug.Print "SLA untuk Level " & levelNum & " = " & CalculateSLA & " hari"

End Function

'=========================================================
' CALCULATE DEADLINE SLA
'=========================================================

Public Function CalculateDeadlineSLA(ByVal fptkDate As Date, ByVal slaDays As Long) As Date
    CalculateDeadlineSLA = DateAdd("d", slaDays, fptkDate)
End Function

'=========================================================
' WEEK / MONTH
'=========================================================

Public Function GetWeekNumber(ByVal dt As Date) As Long

    If dt = 0 Then
        GetWeekNumber = 0
        Exit Function
    End If

    GetWeekNumber = CLng(Application.WorksheetFunction.WeekNum(dt, 1))

End Function

Public Function GetMonthName(ByVal dt As Date) As String

    If dt = 0 Then
        GetMonthName = ""
        Exit Function
    End If

    GetMonthName = Format(dt, "mmmm")

End Function

'=========================================================
' PRIVATE HELPERS - SHEET
'=========================================================

Private Function LKP_GetSheet(ByVal sheetName As String, _
                              Optional ByVal createIfMissing As Boolean = False) As Worksheet

    On Error Resume Next
    Set LKP_GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If LKP_GetSheet Is Nothing And createIfMissing Then

        Set LKP_GetSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        LKP_GetSheet.Name = sheetName

    End If

End Function

Private Function LKP_LastRow(ByVal ws As Worksheet, Optional ByVal colNum As Long = 1) As Long

    If ws Is Nothing Then
        LKP_LastRow = 1
        Exit Function
    End If

    If colNum <= 0 Then colNum = 1

    LKP_LastRow = ws.Cells(ws.rows.Count, colNum).End(xlUp).Row

    If LKP_LastRow < 1 Then LKP_LastRow = 1

End Function

Private Function LKP_FindHeader(ByVal ws As Worksheet, ByVal headerText As String) As Long

    Dim schemaColumn As Long
    Dim c As Range

    If ws Is Nothing Then
        LKP_FindHeader = 0
        Exit Function
    End If

    schemaColumn = SchemaB02FindColumn( _
        ws, headerText, False)

    If schemaColumn > 0 Then
        LKP_FindHeader = schemaColumn
        Exit Function
    End If

    Set c = ws.rows(1).Find(What:=headerText, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=False)

    If c Is Nothing Then
        Set c = ws.rows(1).Find(What:=headerText, LookAt:=xlPart, LookIn:=xlValues, MatchCase:=False)
    End If

    If c Is Nothing Then
        LKP_FindHeader = 0
    Else
        LKP_FindHeader = c.Column
    End If

End Function

'=========================================================
' PRIVATE HELPERS - HEADERS
'=========================================================

Private Sub LKP_EnsureDBHeaders(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    SchemaB02Prepare "DB Kode Posisi"
    SchemaB02UnlistManagedRange ws
    SchemaB02StyleHeader ws

End Sub

Private Sub LKP_EnsureDBColumns(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    LKP_EnsureDBHeaders ws

    colDBKode = LKP_FindHeader(ws, "KODE")
    colDBPosition = LKP_FindHeader(ws, "POSITION")
    colDBLocation = LKP_FindHeader(ws, "LOCATION")
    colDBBusinessUnit = LKP_FindHeader(ws, "BUSINESS UNIT")
    colDBDivisionChris = LKP_FindHeader(ws, "DIVISION CHRIS")
    colDBDepartmentChris = LKP_FindHeader(ws, "DEPARTMENT CHRIS")
    colDBUser = LKP_FindHeader(ws, "USER (MANAGER)")
    colDBIndirectUser = LKP_FindHeader(ws, "INDIRECT USER")
    colDBDirectorate = LKP_FindHeader(ws, "DIRECTORATE")
    colDBYear = LKP_FindHeader(ws, "YEAR")

    ' Tidak memakai fallback nomor kolom.
    ' Schema_Config menjadi sumber posisi kolom.

End Sub

Private Sub LKP_EnsureDropdownColumns(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    colDropPIC = LKP_FindHeader(ws, "PIC Recruiter")
    colDropKodePIC = LKP_FindHeader(ws, "Kode PIC")
    colDropBU = LKP_FindHeader(ws, "BU")
    colDropKodeBU = LKP_FindHeader(ws, "Kode BU")

    If colDropPIC = 0 Then colDropPIC = 1
    If colDropKodePIC = 0 Then colDropKodePIC = 2
    If colDropBU = 0 Then colDropBU = 3
    If colDropKodeBU = 0 Then colDropKodeBU = 4

End Sub

'=========================================================
' PRIVATE HELPERS - NORMALIZE
'=========================================================

Private Function LKP_NormalizeCompare(ByVal v As Variant) As String

    Dim s As String

    s = UCase(Trim(CStr(v)))
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    LKP_NormalizeCompare = s

End Function

Private Function LKP_NormalizeKey(ByVal v As Variant) As String

    Dim s As String

    s = LCase(Trim(CStr(v)))
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ".", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")
    s = Replace(s, " ", "")

    LKP_NormalizeKey = s

End Function

Private Function LKP_NormalizeRecruiterName(ByVal txt As String) As String

    txt = LKP_NormalizeKey(txt)

    Select Case txt
        Case "recruiter15", "recruiter15", "recruiter15"
            txt = "recruiter15"
        Case "recruiter04", "recruiter04", "recruiter04"
            txt = "recruiter04"
        Case "desiarmanda", "desiarmandasari", "desi"
            txt = "desi"
        Case "recruiter01", "recruiter01", "rec01"
            txt = "recruiter01"
        Case "recruiter14", "rec14"
            txt = "recruiter14"
        Case "recruiter09", "rec09"
            txt = "recruiter09"
        Case "recruiter08", "rec08"
            txt = "recruiter08"
        Case "recruiter13"
            txt = "recruiter13"
        Case "recruiter05", "rec05"
            txt = "recruiter05"
        Case "recruiter16", "rec16"
            txt = "recruiter16"
        Case "recruiter17", "rec17"
            txt = "recruiter17"
        Case "recruiter18", "rec18"
            txt = "recruiter18"
        Case "recruiter19", "rec19"
            txt = "recruiter19"
        Case "recruiter20", "rec20"
            txt = "recruiter20"
        Case "recruiter02"
            txt = "recruiter02"
        Case "recruiter21", "rec21", "recruiter21"
            txt = "recruiter21"
        Case "recruiter22"
            txt = "recruiter22"
        Case "recruiter23", "recruiter23", "rec23"
            txt = "recruiter23"
    End Select

    LKP_NormalizeRecruiterName = txt

End Function

'=========================================================
' PRIVATE HELPERS - PIC / BU FALLBACK
'=========================================================

Private Function LKP_PICNameFromAnyText(ByVal sourceText As String) As String

    Dim s As String
    s = LCase(CStr(sourceText))

    If InStr(s, "desi.armanda") > 0 Or InStr(s, "desi") > 0 Then
        LKP_PICNameFromAnyText = "Desi"
    ElseIf InStr(s, "recruiter04") > 0 Or InStr(s, "recruiter04") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter04"
    ElseIf InStr(s, "recruiter09") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter09"
    ElseIf InStr(s, "recruiter14") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter14"
    ElseIf InStr(s, "recruiter15") > 0 Or InStr(s, "recruiter15") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter15"
    ElseIf InStr(s, "recruiter07") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter07"
    ElseIf InStr(s, "recruiter01") > 0 Or InStr(s, "recruiter01") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter01"
    ElseIf InStr(s, "recruiter08") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter08"
    ElseIf InStr(s, "recruiter13") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter13"
    ElseIf InStr(s, "recruiter05") > 0 Or InStr(s, "recruiter05") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter05"
    ElseIf InStr(s, "recruiter16") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter16"
    ElseIf InStr(s, "recruiter17") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter17"
    ElseIf InStr(s, "recruiter18") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter18"
    ElseIf InStr(s, "recruiter19") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter19"
    ElseIf InStr(s, "recruiter20") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter20"
    ElseIf InStr(s, "recruiter02") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter02"
    ElseIf InStr(s, "recruiter21") > 0 Or InStr(s, "recruiter21") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter21"
    ElseIf InStr(s, "recruiter22") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter22"
    ElseIf InStr(s, "recruiter23") > 0 Or InStr(s, "recruiter23") > 0 Then
        LKP_PICNameFromAnyText = "Recruiter23"
    Else
        LKP_PICNameFromAnyText = ""
    End If

End Function

Private Function LKP_PICCodeFromPICName(ByVal picName As String) As String

    Select Case LKP_NormalizeRecruiterName(picName)

        Case "recruiter09"
            LKP_PICCodeFromPICName = "CORPRec09"
        Case "recruiter04"
            LKP_PICCodeFromPICName = "CORPKar"
        Case "recruiter14"
            LKP_PICCodeFromPICName = "CORPRec14"
        Case "recruiter15"
            LKP_PICCodeFromPICName = "CORPRec15"
        Case "recruiter07"
            LKP_PICCodeFromPICName = "CORPRec07"
        Case "recruiter01"
            LKP_PICCodeFromPICName = "CORPRec01"
        Case "recruiter08"
            LKP_PICCodeFromPICName = "CORPRec08"
        Case "recruiter13"
            LKP_PICCodeFromPICName = "CORPRecruiter13"
        Case "recruiter05"
            LKP_PICCodeFromPICName = "CORPRec05"
        Case "desi"
            LKP_PICCodeFromPICName = "CORPDesi"
        Case "recruiter16"
            LKP_PICCodeFromPICName = "CMDRec16"
        Case "recruiter17"
            LKP_PICCodeFromPICName = "CMDRec17"
        Case "recruiter18"
            LKP_PICCodeFromPICName = "CMDRec18"
        Case "recruiter19"
            LKP_PICCodeFromPICName = "MPRec19"
        Case "recruiter20"
            LKP_PICCodeFromPICName = "MPRec20"
        Case "recruiter02"
            LKP_PICCodeFromPICName = "MPRecruiter02"
        Case "recruiter21"
            LKP_PICCodeFromPICName = "MPRec21"
        Case "recruiter22"
            LKP_PICCodeFromPICName = "MSRecruiter22"
        Case "recruiter23"
            LKP_PICCodeFromPICName = "JESSRec23"
        Case Else
            LKP_PICCodeFromPICName = ""

    End Select

End Function

Private Function LKP_BUCodeFromBUName(ByVal buName As String) As String

    Dim b As String
    b = LCase(Trim(CStr(buName)))

    If InStr(b, "cisarua mountain dairy") > 0 Then
        LKP_BUCodeFromBUName = "CMD"
    ElseIf InStr(b, "macrosentra niagaboga") > 0 Then
        LKP_BUCodeFromBUName = "MS"
    ElseIf InStr(b, "java egg") > 0 Then
        LKP_BUCodeFromBUName = "JESS"
    ElseIf InStr(b, "macroprima panganutama") > 0 Then
        LKP_BUCodeFromBUName = "MP"
    ElseIf InStr(b, "macrotama binasantika") > 0 Then
        LKP_BUCodeFromBUName = "MB"
    ElseIf InStr(b, "bavarian culinary haus") > 0 Then
        LKP_BUCodeFromBUName = "BHC"
    ElseIf InStr(b, "artha rasa cimory") > 0 Then
        LKP_BUCodeFromBUName = "ARC"
    ElseIf b = "ho" Then
        LKP_BUCodeFromBUName = "HO"
    Else
        LKP_BUCodeFromBUName = ""
    End If

End Function

'=========================================================
' PRIVATE HELPER - TABLE FORMAT
'=========================================================

Private Sub LKP_FormatAsTableSafe( _
    ByVal ws As Worksheet, _
    ByVal tableName As String)

    Dim lastDataRow As Long
    Dim lastColumn As Long
    Dim usedRange As Range

    If ws Is Nothing Then Exit Sub

    SchemaB02UnlistManagedRange ws

    lastDataRow = SchemaB02LastUsedRow(ws)
    lastColumn = SchemaB02LastUsedColumn(ws)

    If lastDataRow < 1 Or lastColumn < 1 Then Exit Sub

    Set usedRange = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(lastDataRow, lastColumn))

    With usedRange
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With

    SchemaB02StyleHeader ws

End Sub



