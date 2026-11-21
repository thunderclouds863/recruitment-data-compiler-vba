Attribute VB_Name = "modRecruiterMapping"
Option Explicit

' ============================================================
' RECRUITER MAPPING REPORT V2
'
' READ ONLY terhadap sheet FPTK.
'
' OUTPUT SHEET:
'   Mapping FPTK Recruiter
'
' DASHBOARD:
'   KPI Cards + Summary Recruiter
'
' BAGIAN 2:
'   Mapping per Bagian
'   Nama | Scope Group | FPTK | STO | Total | % Recruiter
'
' BAGIAN 3:
'   Mapping Detail
'   Nama | Scope Group | Scope Detail | BU | Level Scope |
'   FPTK | STO | Total | % Recruiter
'   Nama | FPTK | STO | Total | %Total | Scope FPTK
'
' DEFINISI:
'   TOTAL OPEN = semua Status OP / OPEN
'   OPEN TEAM  = semua Status OP / OPEN milik PIC terpilih, termasuk STO
'   FPTK per PIC = Status OP / OPEN dan bukan STO
'   STO per PIC  = Status OP / OPEN dan kategori STO
'   BELUM ADA PIC = semua Status OP / OPEN dengan PIC Recruiter kosong,
'                   baik FPTK maupun STO
'
' EXCLUDED DEFAULT:
'   Recruiter09, Recruiter14
'
' CLAP & FGDP:
'   Selalu disatukan menjadi:
'   Scope Group  = CLAP & FGDP
'   Scope Detail = All Program
'   Level Scope  = All Level
'   sehingga tidak lagi muncul duplikat:
'   "CLAP & FGDP (ALL)" dan
'   "CLAP & FGDP (ALL) (Level 3)"
' ============================================================

Private Const RMR_FPTK_SHEET As String = "FPTK"
Private Const RMR_MASTER_SHEET As String = "Recruiter Master"
Private Const RMR_OUTPUT_SHEET As String = "Mapping FPTK Recruiter"

Private Const RMR_HEADER_ROW As Long = 1
Private Const RMR_FIRST_DATA_ROW As Long = 2
Private Const RMR_SEP As String = "||~||"

' ============================================================
' ENTRY POINT
' ============================================================
Public Sub BuildRecruiterMappingReport()

    Dim wsF As Worksheet
    Dim wsMaster As Worksheet
    Dim wsOut As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataArr As Variant

    Dim colPIC As Long
    Dim colKodePIC As Long
    Dim colKodeUnik As Long
    Dim colStatus As Long
    Dim colFilter As Long
    Dim colBU As Long
    Dim colDirectorate As Long
    Dim colDivision As Long
    Dim colDepartment As Long
    Dim colLevel As Long
    Dim colLevelNumber As Long
    Dim colPosition As Long

    Dim recruiterRegular As Object
    Dim recruiterSTO As Object

    Dim groupRegular As Object
    Dim groupSTO As Object

    Dim detailRegular As Object
    Dim detailSTO As Object

    Dim summaryScopeBU As Object
    Dim summaryScopeLevel As Object

    Dim r As Long

    Dim pic As String
    Dim rawPIC As String
    Dim rawKodePIC As String
    Dim rawKodeUnik As String
    Dim statusText As String
    Dim filterText As String
    Dim bu As String
    Dim directorate As String
    Dim division As String
    Dim department As String
    Dim levelDetail As String
    Dim levelNumber As String
    Dim positionText As String

    Dim scopeGroup As String
    Dim scopeDetail As String
    Dim levelScope As String

    Dim groupKey As String
    Dim detailKey As String
    Dim summaryKey As String

    Dim allOpenCount As Long
    Dim selectedOpenCount As Long
    Dim selectedSTOCount As Long
    Dim noPICOpenCount As Long
    Dim noPICSTOCount As Long
    Dim totalSTOMappingCount As Long

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = _
        "Building recruiter mapping report..."

    Set wsF = RMR_GetSheet(RMR_FPTK_SHEET)

    If wsF Is Nothing Then

        Err.Raise _
            vbObjectError + 8101, _
            "BuildRecruiterMappingReport", _
            "Sheet FPTK tidak ditemukan."

    End If

    Set wsMaster = RMR_GetOrCreateRecruiterMaster()
    Set wsOut = RMR_GetOrCreateSheet(RMR_OUTPUT_SHEET)

    colPIC = RMR_FindHeader(wsF, "PIC Recruiter")
    colKodePIC = RMR_FindHeader(wsF, "Kode PIC")
    colKodeUnik = RMR_FindHeader(wsF, "Kode Unik")
    colStatus = RMR_FindHeader(wsF, "Status")
    ' Schema-aware lookup supaya rename / alias tetap terbaca.
    colFilter = SchemaFindColumnAuto( _
        wsF, _
        "Filter Kategorisasi FPTK", _
        False)

    If colFilter = 0 Then
        colFilter = RMR_FindHeader( _
            wsF, _
            "Filter Kategorisasi FPTK")
    End If

    If colFilter = 0 Then
        colFilter = RMR_FindHeader( _
            wsF, _
            "Filter Kategorisasi FPTK2")
    End If
    colBU = RMR_FindHeader(wsF, "Business Unit")
    colDirectorate = RMR_FindHeader(wsF, "Direktorat")
    colDivision = RMR_FindHeader(wsF, "Divisi")
    colDepartment = RMR_FindHeader(wsF, "Department")
    colLevel = RMR_FindHeader(wsF, "Level")
    colLevelNumber = RMR_FindHeader(wsF, "Level Number")
    colPosition = RMR_FindHeader(wsF, "Posisi")

    If colPIC = 0 Or _
       colStatus = 0 Or _
       colFilter = 0 Or _
       colPosition = 0 Then

        Err.Raise _
            vbObjectError + 8102, _
            "BuildRecruiterMappingReport", _
            "Header wajib tidak lengkap: PIC Recruiter / Status / Filter Kategorisasi FPTK / Posisi."

    End If

    lastRow = RMR_LastDataRow(wsF)

    lastCol = _
        wsF.Cells( _
            RMR_HEADER_ROW, _
            wsF.Columns.Count).End( _
                xlToLeft).Column

    If lastRow < RMR_FIRST_DATA_ROW Then

        Err.Raise _
            vbObjectError + 8103, _
            "BuildRecruiterMappingReport", _
            "Sheet FPTK belum memiliki data."

    End If

    dataArr = _
        wsF.Range( _
            wsF.Cells(1, 1), _
            wsF.Cells(lastRow, lastCol)).Value2

    Set recruiterRegular = RMR_NewDict()
    Set recruiterSTO = RMR_NewDict()

    Set groupRegular = RMR_NewDict()
    Set groupSTO = RMR_NewDict()

    Set detailRegular = RMR_NewDict()
    Set detailSTO = RMR_NewDict()

    Set summaryScopeBU = RMR_NewDict()
    Set summaryScopeLevel = RMR_NewDict()

    For r = 2 To UBound(dataArr, 1)

        ' =====================================================
        ' STATUS OPEN
        ' =====================================================
        statusText = _
            RMR_ArrayText( _
                dataArr, _
                r, _
                colStatus)

        If Not RMR_IsOpenStatus(statusText) Then _
            GoTo NextRow

        ' =====================================================
        ' KPI 1 - TOTAL FPTK OPEN
        ' Semua Status OP / OPEN.
        ' Termasuk STO, PIC di luar team, dan PIC kosong.
        ' =====================================================
        allOpenCount = allOpenCount + 1

        ' =====================================================
        ' FILTER KATEGORISASI
        ' =====================================================
        filterText = _
            RMR_ArrayText( _
                dataArr, _
                r, _
                colFilter)

        ' =====================================================
        ' KPI 3 - STO OPEN
        ' Semua OPEN yang kategorisasinya STO.
        ' Tidak tergantung PIC.
        ' =====================================================
        If RMR_IsSTO(filterText) Then
            totalSTOMappingCount = _
                totalSTOMappingCount + 1
        End If

        ' =====================================================
        ' PIC RECRUITER
        ' Workload per PIC dibaca dari kolom PIC Recruiter.
        ' =====================================================
        rawPIC = _
            RMR_ArrayText( _
                dataArr, _
                r, _
                colPIC)

        pic = _
            RMR_CanonicalRecruiterName( _
                rawPIC)

        ' =====================================================
        ' OWNER / PIC RECRUITER
        '
        ' Selama Status OP:
        ' - PIC recruiter terpilih -> masuk nama PIC tersebut.
        ' - PIC recruiter kosong / placeholder -> Belum Ada PIC.
        ' - PIC lain di luar team -> tetap masuk TOTAL OPEN,
        '   tetapi tidak masuk workload 12 recruiter.
        ' =====================================================
        If RMR_IsNoPIC(pic) Then

            pic = "Belum Ada PIC"

            noPICOpenCount = _
                noPICOpenCount + 1

        End If

        ' =====================================================
        ' KPI 2 - FPTK OPEN TEAM
        '
        ' Semua OPEN milik 12 PIC recruiter terpilih.
        ' STO tetap termasuk.
        ' Belum Ada PIC tidak termasuk OPEN TEAM.
        ' =====================================================
        If RMR_IsSelectedRecruiter(pic) Then

            selectedOpenCount = _
                selectedOpenCount + 1

        End If

        ' =====================================================
        ' SUMMARY PER PIC
        '
        ' Recruiter terpilih:
        '   FPTK = OP non-STO
        '   STO  = OP STO
        '
        ' PIC kosong:
        '   tetap masuk Belum Ada PIC
        '   dan tetap dipisah FPTK / STO.
        ' =====================================================
        If RMR_IsSelectedRecruiter(pic) Then

            If RMR_IsSTO(filterText) Then

                selectedSTOCount = _
                    selectedSTOCount + 1

                RMR_IncrementCount _
                    recruiterSTO, _
                    pic

            Else

                RMR_IncrementCount _
                    recruiterRegular, _
                    pic

            End If

        ElseIf StrComp( _
            pic, _
            "Belum Ada PIC", _
            vbTextCompare) = 0 Then

            If RMR_IsSTO(filterText) Then

                noPICSTOCount = _
                    noPICSTOCount + 1

                RMR_IncrementCount _
                    recruiterSTO, _
                    pic

            Else

                RMR_IncrementCount _
                    recruiterRegular, _
                    pic

            End If

        Else

            ' Ada nama PIC tetapi bukan bagian dari 12 recruiter.
            ' Row tetap sudah masuk TOTAL OPEN / STO OPEN,
            ' tetapi tidak masuk workload team terpilih.
            GoTo NextRow

        End If

        ' =====================================================
        ' DETAIL MAPPING
        ' Posisi tidak memengaruhi KPI / total PIC.
        ' Posisi kosong hanya tidak masuk breakdown detail.
        ' =====================================================
        positionText = _
            RMR_ArrayText( _
                dataArr, _
                r, _
                colPosition)

        If Len(positionText) = 0 Then _
            GoTo NextRow

        bu = _
            RMR_DisplayBU( _
                RMR_ArrayText( _
                    dataArr, _
                    r, _
                    colBU))

        directorate = _
            RMR_ArrayText( _
                dataArr, _
                r, _
                colDirectorate)

        division = _
            RMR_ArrayText( _
                dataArr, _
                r, _
                colDivision)

        department = _
            RMR_ArrayText( _
                dataArr, _
                r, _
                colDepartment)

        levelDetail = _
            RMR_ArrayText( _
                dataArr, _
                r, _
                colLevel)

        levelNumber = _
            RMR_ArrayText( _
                dataArr, _
                r, _
                colLevelNumber)

        RMR_ClassifyScope _
            filterText, _
            directorate, _
            division, _
            department, _
            positionText, _
            scopeGroup, _
            scopeDetail

        levelScope = _
            RMR_LevelScope( _
                filterText, _
                levelDetail, _
                levelNumber)

        groupKey = _
            RMR_Key( _
                Array( _
                    pic, _
                    scopeGroup))

        detailKey = _
            RMR_Key( _
                Array( _
                    pic, _
                    scopeGroup, _
                    scopeDetail, _
                    bu, _
                    levelScope))

        If RMR_IsSTO(filterText) Then

            RMR_IncrementCount _
                groupSTO, _
                groupKey

            RMR_IncrementCount _
                detailSTO, _
                detailKey

        Else

            RMR_IncrementCount _
                groupRegular, _
                groupKey

            RMR_IncrementCount _
                detailRegular, _
                detailKey

            summaryKey = _
                RMR_Key( _
                    Array( _
                        pic, _
                        scopeGroup))

            RMR_AddUniqueValue _
                summaryScopeBU, _
                summaryKey, _
                bu

            RMR_AddUniqueValue _
                summaryScopeLevel, _
                summaryKey, _
                levelScope

        End If

NextRow:

    Next r

    RMR_WriteReport _
        wsMaster, _
        wsOut, _
        recruiterRegular, _
        recruiterSTO, _
        groupRegular, _
        groupSTO, _
        allOpenCount, _
        selectedOpenCount, _
        selectedSTOCount, _
        noPICOpenCount, _
        noPICSTOCount, _
        totalSTOMappingCount, _
        detailRegular, _
        detailSTO, _
        summaryScopeBU, _
        summaryScopeLevel

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Recruiter Mapping Report selesai." & _
        vbCrLf & vbCrLf & _
        "Output: " & _
        RMR_OUTPUT_SHEET & _
        vbCrLf & vbCrLf & _
        "Atas = Dashboard + Summary" & _
        vbCrLf & _
        "Bawah = Mapping per Bagian + Detail", _
        vbInformation, _
        "Recruiter Mapping"

    Exit Sub

ErrorHandler:

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "Recruiter Mapping Report gagal." & _
        vbCrLf & _
        "Error " & _
        CStr(Err.Number) & _
        ": " & _
        Err.Description, _
        vbCritical, _
        "Recruiter Mapping"

End Sub

' ============================================================
' RECRUITER MASTER
' Recruiter09 & Recruiter14 sengaja tidak dimasukkan.
' Kalau Recruiter Master sudah terlanjur ada, mereka tetap
' hanya whitelist yang ditampilkan.
' ============================================================
Private Function RMR_GetOrCreateRecruiterMaster() As Worksheet

    Dim ws As Worksheet
    Dim names As Variant
    Dim i As Long

    Set ws = RMR_GetSheet(RMR_MASTER_SHEET)

    If ws Is Nothing Then
        Set ws = RMR_GetOrCreateSheet(RMR_MASTER_SHEET)
    End If

    ws.Cells.Clear

    ws.Range("A1:B1").Value = _
        Array( _
            "Nama Recruiter", _
            "Active")

    names = Array( _
        "Recruiter15", _
        "Recruiter01", _
        "Desi", _
        "Recruiter02", _
        "Recruiter05", _
        "Recruiter08", _
        "Recruiter12", _
        "YEREMIA", _
        "Adista", _
        "Recruiter07", _
        "Recruiter04", _
        "Recruiter13", _
        "Belum Ada PIC")

    For i = LBound(names) To UBound(names)
        ws.Cells(i + 2, 1).Value2 = names(i)
        ws.Cells(i + 2, 2).Value2 = "Y"
    Next i

    With ws.Range("A1:B1")
        .Font.Bold = True
        .AutoFilter
    End With

    ws.Columns("A:B").AutoFit

    Set RMR_GetOrCreateRecruiterMaster = ws

End Function

' ============================================================
' REPORT WRITER
' ============================================================
Private Sub RMR_WriteReport( _
    ByVal wsMaster As Worksheet, _
    ByVal wsOut As Worksheet, _
    ByVal recruiterRegular As Object, _
    ByVal recruiterSTO As Object, _
    ByVal groupRegular As Object, _
    ByVal groupSTO As Object, _
    ByVal allOpenCount As Long, _
    ByVal selectedOpenCount As Long, _
    ByVal selectedSTOCount As Long, _
    ByVal noPICOpenCount As Long, _
    ByVal noPICSTOCount As Long, _
    ByVal totalSTOMappingCount As Long, _
    ByVal detailRegular As Object, _
    ByVal detailSTO As Object, _
    ByVal summaryScopeBU As Object, _
    ByVal summaryScopeLevel As Object)

    Dim recruiterSet As Object
    Dim groupSet As Object
    Dim detailSet As Object

    Dim masterLastRow As Long
    Dim masterRow As Long
    Dim activeText As String
    Dim recruiterName As String

    Dim k As Variant

    Dim summaryHeaderRow As Long
    Dim summaryOutRow As Long
    Dim summaryEndRow As Long

    Dim groupTitleRow As Long
    Dim groupHeaderRow As Long
    Dim groupOutRow As Long
    Dim groupEndRow As Long

    Dim detailTitleRow As Long
    Dim detailHeaderRow As Long
    Dim detailOutRow As Long
    Dim detailEndRow As Long

    Dim headers As Variant

    Dim regularValue As Long
    Dim stoValue As Long
    Dim totalValue As Long

    Dim grandRegular As Long
    Dim grandSTO As Long
    Dim grandTotal As Long

    Dim scopeText As String

    Set recruiterSet = RMR_NewDict()
    Set groupSet = RMR_NewDict()
    Set detailSet = RMR_NewDict()

    masterLastRow = _
        wsMaster.Cells( _
            wsMaster.rows.Count, _
            1).End( _
                xlUp).Row

    For masterRow = 2 To masterLastRow

        recruiterName = _
            RMR_CanonicalRecruiterName( _
                CStr( _
                    wsMaster.Cells( _
                        masterRow, _
                        1).Value2))

        activeText = _
            UCase$( _
                Trim$( _
                    CStr( _
                        wsMaster.Cells( _
                            masterRow, _
                            2).Value2)))

        If Len(recruiterName) > 0 And _
           activeText <> "N" And _
           (RMR_IsSelectedRecruiter(recruiterName) Or _
            StrComp( _
                recruiterName, _
                "Belum Ada PIC", _
                vbTextCompare) = 0) Then

            If Not recruiterSet.Exists(recruiterName) Then
                recruiterSet.Add recruiterName, True
            End If

        End If

    Next masterRow

    For Each k In groupRegular.keys
        If Not groupSet.Exists(CStr(k)) Then
            groupSet.Add CStr(k), True
        End If
    Next k

    For Each k In groupSTO.keys
        If Not groupSet.Exists(CStr(k)) Then
            groupSet.Add CStr(k), True
        End If
    Next k

    For Each k In detailRegular.keys
        If Not detailSet.Exists(CStr(k)) Then
            detailSet.Add CStr(k), True
        End If
    Next k

    For Each k In detailSTO.keys
        If Not detailSet.Exists(CStr(k)) Then
            detailSet.Add CStr(k), True
        End If
    Next k

    For Each k In recruiterSet.keys

        recruiterName = CStr(k)

        regularValue = _
            RMR_GetCount( _
                recruiterRegular, _
                recruiterName)

        stoValue = _
            RMR_GetCount( _
                recruiterSTO, _
                recruiterName)

        totalValue = regularValue + stoValue

        grandRegular = grandRegular + regularValue
        grandSTO = grandSTO + stoValue
        grandTotal = grandTotal + totalValue


    Next k

    ' Bersihkan output lama termasuk chart/table.
    RMR_ResetOutputSheet wsOut

    ' ========================================================
    ' DASHBOARD TITLE
    ' ========================================================
    wsOut.Range("A1:I1").Merge
    wsOut.Range("A1").Value2 = "RECRUITER WORKLOAD & SCOPE MAPPING"
    wsOut.Range("A2:I2").Merge
    wsOut.Range("A2").Value2 = _
        "Dashboard current workload | Generated " & _
        Format$(Now, "dd-mmm-yyyy hh:mm")

    RMR_StyleDashboardTitle wsOut

    ' ========================================================
    ' KPI CARDS
    ' ========================================================
    RMR_WriteKPICard _
        wsOut, _
        "A4:C6", _
        "TOTAL FPTK OPEN", _
        allOpenCount, _
        "Semua Status OP"

    RMR_WriteKPICard _
        wsOut, _
        "D4:F6", _
        "FPTK OPEN TEAM", _
        selectedOpenCount, _
        "12 PIC recruiter"

    RMR_WriteKPICard _
        wsOut, _
        "G4:I6", _
        "STO OPEN", _
        totalSTOMappingCount, _
        CStr(noPICSTOCount) & " belum ada PIC"

    ' ========================================================
    ' DASHBOARD RULE NOTE
    ' ========================================================
    wsOut.Range("A7:J7").Merge
    wsOut.Range("A7").Value2 = _
        "Team: Recruiter15, Recruiter01, Desi, Recruiter02, Recruiter05, Recruiter08, Recruiter12, Yeremia, Adista, Recruiter07, Recruiter04, Recruiter13 | PIC Recruiter kosong = Belum Ada PIC"
    With wsOut.Range("A7:J7")
        .Interior.Color = RGB(255, 242, 204)
        .Font.Color = RGB(127, 96, 0)
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' ========================================================
    ' SUMMARY RECRUITER - PALING ATAS
    ' ========================================================
    wsOut.Range("A8:F8").Merge
    wsOut.Range("A8").Value2 = "SUMMARY RECRUITER"
    RMR_StyleSectionTitle wsOut.Range("A8:F8")

    summaryHeaderRow = 10

    headers = Array( _
        "Nama", _
        "FPTK", _
        "STO", _
        "Total", _
        "%Total", _
        "Scope FPTK")

    wsOut.Range( _
        wsOut.Cells(summaryHeaderRow, 1), _
        wsOut.Cells(summaryHeaderRow, 6)).Value = _
            headers

    summaryOutRow = summaryHeaderRow + 1

    For Each k In recruiterSet.keys

        recruiterName = CStr(k)

        regularValue = _
            RMR_GetCount( _
                recruiterRegular, _
                recruiterName)

        stoValue = _
            RMR_GetCount( _
                recruiterSTO, _
                recruiterName)

        totalValue = regularValue + stoValue

        If StrComp( _
            recruiterName, _
            "Belum Ada PIC", _
            vbTextCompare) = 0 Then

            If regularValue > 0 And _
               stoValue > 0 Then

                scopeText = _
                    "FPTK & STO - Belum Ada PIC"

            ElseIf stoValue > 0 Then

                scopeText = _
                    "STO - Belum Ada PIC"

            ElseIf regularValue > 0 Then

                scopeText = _
                    "FPTK - Belum Ada PIC"

            Else

                scopeText = _
                    "Belum Ada PIC"

            End If

        Else

            scopeText = _
                RMR_BuildSummaryScopeText( _
                    recruiterName, _
                    summaryScopeBU, _
                    summaryScopeLevel)

        End If

        wsOut.Cells(summaryOutRow, 1).Value2 = recruiterName
        wsOut.Cells(summaryOutRow, 2).Value2 = regularValue
        wsOut.Cells(summaryOutRow, 3).Value2 = stoValue
        wsOut.Cells(summaryOutRow, 4).Value2 = totalValue

        If selectedOpenCount + noPICOpenCount > 0 Then

            wsOut.Cells(summaryOutRow, 5).Value2 = _
                totalValue / _
                (selectedOpenCount + noPICOpenCount)

        Else

            wsOut.Cells(summaryOutRow, 5).Value2 = 0

        End If

        wsOut.Cells(summaryOutRow, 6).Value2 = scopeText

        summaryOutRow = summaryOutRow + 1

    Next k

    summaryEndRow = summaryOutRow - 1

    RMR_CreateVisualTable _
        wsOut, _
        "tblRecruiterSummary", _
        summaryHeaderRow, _
        1, _
        summaryEndRow, _
        6, _
        "TableStyleLight9"

    If summaryEndRow >= summaryHeaderRow + 1 Then

        wsOut.Range( _
            wsOut.Cells(summaryHeaderRow + 1, 5), _
            wsOut.Cells(summaryEndRow, 5)).numberFormat = _
                "0.0%"

        wsOut.Range( _
            wsOut.Cells(summaryHeaderRow + 1, 6), _
            wsOut.Cells(summaryEndRow, 6)).WrapText = True

        RMR_AddDataBar _
            wsOut.Range( _
                wsOut.Cells(summaryHeaderRow + 1, 4), _
                wsOut.Cells(summaryEndRow, 4))

        RMR_AddDataBar _
            wsOut.Range( _
                wsOut.Cells(summaryHeaderRow + 1, 5), _
                wsOut.Cells(summaryEndRow, 5))

        RMR_ApplyRecruiterColors _
            wsOut, _
            summaryHeaderRow + 1, _
            summaryEndRow, _
            6

    End If

    ' ========================================================
    ' MAPPING PER BAGIAN
    ' ========================================================
    groupTitleRow = summaryEndRow + 5

    wsOut.Range( _
        wsOut.Cells(groupTitleRow, 1), _
        wsOut.Cells(groupTitleRow, 6)).Merge

    wsOut.Cells(groupTitleRow, 1).Value2 = _
        "MAPPING PER BAGIAN / SCOPE GROUP"

    RMR_StyleSectionTitle _
        wsOut.Range( _
            wsOut.Cells(groupTitleRow, 1), _
            wsOut.Cells(groupTitleRow, 6))

    wsOut.Cells(groupTitleRow + 1, 1).Value2 = _
        "Gunakan bagian ini untuk menjawab: Marketing Desi ada berapa? Sales GT Recruiter08 ada berapa?"

    wsOut.Cells(groupTitleRow + 1, 1).Font.Italic = True

    groupHeaderRow = groupTitleRow + 3

    headers = Array( _
        "Nama", _
        "Bagian / Scope Group", _
        "FPTK", _
        "STO", _
        "Total", _
        "% Recruiter")

    wsOut.Range( _
        wsOut.Cells(groupHeaderRow, 1), _
        wsOut.Cells(groupHeaderRow, 6)).Value = _
            headers

    groupOutRow = groupHeaderRow + 1

    For Each k In recruiterSet.keys

        RMR_WriteRecruiterGroupRows _
            wsOut, _
            groupOutRow, _
            CStr(k), _
            RMR_GetCount( _
                recruiterRegular, _
                CStr(k)) + _
            RMR_GetCount( _
                recruiterSTO, _
                CStr(k)), _
            groupSet, _
            groupRegular, _
            groupSTO

    Next k

    groupEndRow = groupOutRow - 1

    If groupEndRow >= groupHeaderRow + 1 Then

        RMR_CreateVisualTable _
            wsOut, _
            "tblRecruiterScopeGroup", _
            groupHeaderRow, _
            1, _
            groupEndRow, _
            6, _
            "TableStyleLight9"

        wsOut.Range( _
            wsOut.Cells(groupHeaderRow + 1, 6), _
            wsOut.Cells(groupEndRow, 6)).numberFormat = _
                "0.0%"

        RMR_AddDataBar _
            wsOut.Range( _
                wsOut.Cells(groupHeaderRow + 1, 5), _
                wsOut.Cells(groupEndRow, 5))

        RMR_ApplyRecruiterColors _
            wsOut, _
            groupHeaderRow + 1, _
            groupEndRow, _
            6

    End If

    ' ========================================================
    ' DETAIL
    ' ========================================================
    detailTitleRow = _
        Application.Max( _
            groupEndRow, _
            groupHeaderRow) + 4

    wsOut.Range( _
        wsOut.Cells(detailTitleRow, 1), _
        wsOut.Cells(detailTitleRow, 9)).Merge

    wsOut.Cells(detailTitleRow, 1).Value2 = _
        "DETAIL SCOPE RECRUITER"

    RMR_StyleSectionTitle _
        wsOut.Range( _
            wsOut.Cells(detailTitleRow, 1), _
            wsOut.Cells(detailTitleRow, 9))

    wsOut.Cells(detailTitleRow + 1, 1).Value2 = _
        "Breakdown sampai Scope Detail, Business Unit, dan Level."

    wsOut.Cells(detailTitleRow + 1, 1).Font.Italic = True

    detailHeaderRow = detailTitleRow + 3

    headers = Array( _
        "Nama", _
        "Scope Group", _
        "Scope Detail", _
        "Business Unit", _
        "Level Scope", _
        "FPTK", _
        "STO", _
        "Total", _
        "% Recruiter")

    wsOut.Range( _
        wsOut.Cells(detailHeaderRow, 1), _
        wsOut.Cells(detailHeaderRow, 9)).Value = _
            headers

    detailOutRow = detailHeaderRow + 1

    For Each k In recruiterSet.keys

        RMR_WriteRecruiterDetailRows _
            wsOut, _
            detailOutRow, _
            CStr(k), _
            RMR_GetCount( _
                recruiterRegular, _
                CStr(k)) + _
            RMR_GetCount( _
                recruiterSTO, _
                CStr(k)), _
            detailSet, _
            detailRegular, _
            detailSTO

    Next k

    detailEndRow = detailOutRow - 1

    If detailEndRow >= detailHeaderRow + 1 Then

        RMR_CreateVisualTable _
            wsOut, _
            "tblRecruiterDetail", _
            detailHeaderRow, _
            1, _
            detailEndRow, _
            9, _
            "TableStyleLight9"

        wsOut.Range( _
            wsOut.Cells(detailHeaderRow + 1, 9), _
            wsOut.Cells(detailEndRow, 9)).numberFormat = _
                "0.0%"

        RMR_AddDataBar _
            wsOut.Range( _
                wsOut.Cells(detailHeaderRow + 1, 8), _
                wsOut.Cells(detailEndRow, 8))

        RMR_ApplyRecruiterColors _
            wsOut, _
            detailHeaderRow + 1, _
            detailEndRow, _
            9

        RMR_HighlightUnmapped _
            wsOut, _
            detailHeaderRow + 1, _
            detailEndRow

    End If

    ' ========================================================
    ' FINAL FORMAT
    ' ========================================================
    wsOut.Cells.Font.Name = "Calibri"
    wsOut.Cells.Font.Size = 10

    wsOut.Columns("A:A").columnWidth = 16
    wsOut.Columns("B:B").columnWidth = 26
    wsOut.Columns("C:C").columnWidth = 28
    wsOut.Columns("D:E").columnWidth = 18
    wsOut.Columns("F:F").columnWidth = 52
    wsOut.Columns("G:I").columnWidth = 13

    wsOut.rows("1:2").RowHeight = 24

    wsOut.Range( _
        wsOut.Cells(1, 1), _
        wsOut.Cells( _
            Application.Max(detailEndRow, summaryEndRow), _
            9)).VerticalAlignment = _
                xlTop

    wsOut.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    wsOut.Range("A11").Select
    ActiveWindow.FreezePanes = True

End Sub

Private Sub RMR_WriteRecruiterGroupRows( _
    ByVal wsOut As Worksheet, _
    ByRef outRow As Long, _
    ByVal recruiterName As String, _
    ByVal recruiterTotal As Long, _
    ByVal groupSet As Object, _
    ByVal groupRegular As Object, _
    ByVal groupSTO As Object)

    Dim k As Variant
    Dim parts As Variant

    Dim rowRegular As Long
    Dim rowSTO As Long
    Dim rowTotal As Long

    For Each k In groupSet.keys

        parts = Split(CStr(k), RMR_SEP)

        If UBound(parts) >= 1 Then

            If StrComp( _
                RMR_DisplayKeyPart(CStr(parts(0))), _
                recruiterName, _
                vbTextCompare) = 0 Then

                rowRegular = _
                    RMR_GetCount( _
                        groupRegular, _
                        CStr(k))

                rowSTO = _
                    RMR_GetCount( _
                        groupSTO, _
                        CStr(k))

                rowTotal = rowRegular + rowSTO

                wsOut.Cells(outRow, 1).Value2 = recruiterName
                wsOut.Cells(outRow, 2).Value2 = _
                    RMR_DisplayKeyPart(CStr(parts(1)))
                wsOut.Cells(outRow, 3).Value2 = rowRegular
                wsOut.Cells(outRow, 4).Value2 = rowSTO
                wsOut.Cells(outRow, 5).Value2 = rowTotal

                If recruiterTotal > 0 Then

                    wsOut.Cells(outRow, 6).Value2 = _
                        rowTotal / recruiterTotal

                Else

                    wsOut.Cells(outRow, 6).Value2 = 0

                End If

                outRow = outRow + 1

            End If

        End If

    Next k

End Sub

Private Sub RMR_WriteRecruiterDetailRows( _
    ByVal wsOut As Worksheet, _
    ByRef outRow As Long, _
    ByVal recruiterName As String, _
    ByVal recruiterTotal As Long, _
    ByVal detailSet As Object, _
    ByVal detailRegular As Object, _
    ByVal detailSTO As Object)

    Dim keysToWrite As Object
    Dim k As Variant
    Dim parts As Variant

    Dim rowRegular As Long
    Dim rowSTO As Long
    Dim rowTotal As Long

    Set keysToWrite = RMR_NewDict()

    For Each k In detailSet.keys

        parts = Split(CStr(k), RMR_SEP)

        If UBound(parts) >= 4 Then

            If StrComp( _
                RMR_DisplayKeyPart(CStr(parts(0))), _
                recruiterName, _
                vbTextCompare) = 0 Then

                If Not keysToWrite.Exists(CStr(k)) Then
                    keysToWrite.Add CStr(k), True
                End If

            End If

        End If

    Next k

    If keysToWrite.Count = 0 Then Exit Sub

    For Each k In keysToWrite.keys

        parts = Split(CStr(k), RMR_SEP)

        rowRegular = RMR_GetCount(detailRegular, CStr(k))
        rowSTO = RMR_GetCount(detailSTO, CStr(k))
        rowTotal = rowRegular + rowSTO

        wsOut.Cells(outRow, 1).Value2 = recruiterName
        wsOut.Cells(outRow, 2).Value2 = RMR_DisplayKeyPart(CStr(parts(1)))
        wsOut.Cells(outRow, 3).Value2 = RMR_DisplayKeyPart(CStr(parts(2)))
        wsOut.Cells(outRow, 4).Value2 = RMR_DisplayKeyPart(CStr(parts(3)))
        wsOut.Cells(outRow, 5).Value2 = RMR_DisplayKeyPart(CStr(parts(4)))
        wsOut.Cells(outRow, 6).Value2 = rowRegular
        wsOut.Cells(outRow, 7).Value2 = rowSTO
        wsOut.Cells(outRow, 8).Value2 = rowTotal

        If recruiterTotal > 0 Then

            wsOut.Cells(outRow, 9).Value2 = _
                rowTotal / recruiterTotal

        Else

            wsOut.Cells(outRow, 9).Value2 = 0

        End If

        outRow = outRow + 1

    Next k

End Sub

' ============================================================
' SUMMARY SCOPE
'
' Hanya broad scope, supaya summary tidak terlalu ramai.
' Detail spesifik ada di bagian Mapping Detail.
' ============================================================
Private Function RMR_BuildSummaryScopeText( _
    ByVal recruiterName As String, _
    ByVal scopeBUDict As Object, _
    ByVal scopeLevelDict As Object) As String

    Dim keys As Object
    Dim k As Variant
    Dim parts As Variant

    Dim recruiterKey As String
    Dim groupName As String
    Dim buText As String
    Dim levelText As String
    Dim lineText As String

    Dim buDict As Object
    Dim levelDict As Object

    Dim arr As Variant
    Dim i As Long

    Set keys = RMR_NewDict()

    recruiterKey = RMR_KeyPart(recruiterName)

    For Each k In scopeBUDict.keys

        parts = Split(CStr(k), RMR_SEP)

        If UBound(parts) >= 1 Then

            If StrComp( _
                CStr(parts(0)), _
                recruiterKey, _
                vbTextCompare) = 0 Then

                groupName = RMR_DisplayKeyPart(CStr(parts(1)))

                Set buDict = scopeBUDict(k)

                buText = RMR_JoinSortedDictKeys(buDict, ", ")

                levelText = ""

                If scopeLevelDict.Exists(CStr(k)) Then

                    Set levelDict = scopeLevelDict(CStr(k))

                    levelText = _
                        RMR_SummaryLevelText( _
                            levelDict)

                End If

                lineText = groupName

                If UCase$(groupName) = "CLAP & FGDP" Then

                    lineText = "CLAP & FGDP - All Program / All Level"

                Else

                    If Len(buText) > 0 Then

                        lineText = _
                            lineText & _
                            " - " & _
                            buText

                    End If

                    If Len(levelText) > 0 Then

                        lineText = _
                            lineText & _
                            " [" & _
                            levelText & _
                            "]"

                    End If

                End If

                If Not keys.Exists(lineText) Then
                    keys.Add lineText, True
                End If

            End If

        End If

    Next k

    If keys.Count = 0 Then Exit Function

    arr = keys.keys
    RMR_SortStringArray arr

    For i = LBound(arr) To UBound(arr)

        If Len(RMR_BuildSummaryScopeText) > 0 Then

            RMR_BuildSummaryScopeText = _
                RMR_BuildSummaryScopeText & _
                vbLf

        End If

        RMR_BuildSummaryScopeText = _
            RMR_BuildSummaryScopeText & _
            CStr(arr(i))

    Next i

End Function

Private Function RMR_SummaryLevelText( _
    ByVal levelDict As Object) As String

    Dim hasAll As Boolean
    Dim hasL12 As Boolean
    Dim hasL3 As Boolean
    Dim hasL4 As Boolean

    Dim k As Variant
    Dim t As String

    If levelDict Is Nothing Then Exit Function

    For Each k In levelDict.keys

        t = UCase$(CStr(k))

        If t = "ALL LEVEL" Then
            hasAll = True
        ElseIf InStr(1, t, "LEVEL 1", vbTextCompare) > 0 Or _
               InStr(1, t, "LEVEL 2", vbTextCompare) > 0 Then
            hasL12 = True
        ElseIf InStr(1, t, "LEVEL 3", vbTextCompare) > 0 Or _
               Left$(t, 2) = "3A" Or _
               Left$(t, 2) = "3B" Or _
               Left$(t, 2) = "3C" Then
            hasL3 = True
        ElseIf InStr(1, t, "LEVEL 4", vbTextCompare) > 0 Or _
               Left$(t, 1) = "4" Then
            hasL4 = True
        End If

    Next k

    If hasAll Then

        RMR_SummaryLevelText = "All Level"
        Exit Function

    End If

    If hasL12 Then
        RMR_AppendText RMR_SummaryLevelText, "Level 1-2", ", "
    End If

    If hasL3 Then
        RMR_AppendText RMR_SummaryLevelText, "Level 3", ", "
    End If

    If hasL4 Then
        RMR_AppendText RMR_SummaryLevelText, "Managerial", ", "
    End If

End Function

' ============================================================
' CLASSIFICATION
'
' Scope Group = broad bucket untuk summary.
' Scope Detail = detail spesifik untuk mapping detail.
'
' Contoh:
' Marketing | Brand Marketing
' Marketing | Trade Marketing
' Marketing | Marketing
'
' FBS / Finance | Financial Controller
' FBS / Finance | Internal Audit
' ============================================================
Private Sub RMR_ClassifyScope( _
    ByVal filterText As String, _
    ByVal directorate As String, _
    ByVal division As String, _
    ByVal department As String, _
    ByVal positionText As String, _
    ByRef scopeGroup As String, _
    ByRef scopeDetail As String)

    Dim combined As String

    scopeGroup = ""
    scopeDetail = ""

    combined = _
        UCase$( _
            directorate & _
            " | " & _
            division & _
            " | " & _
            department & _
            " | " & _
            positionText)

    ' --------------------------------------------------------
    ' PROGRAM
    ' --------------------------------------------------------
    If RMR_IsCLAPFGDP(filterText) Then

        scopeGroup = "CLAP & FGDP"
        scopeDetail = "All Program"
        Exit Sub

    End If

    ' --------------------------------------------------------
    ' SALES
    ' --------------------------------------------------------
    If RMR_ContainsAny( _
        combined, _
        Array( _
            "SALES GENERAL TRADE", _
            "SALES GT", _
            "GENERAL TRADE")) Then

        scopeGroup = "Sales GT"
        scopeDetail = "Sales General Trade"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("KEY ACCOUNT")) Then

        scopeGroup = "Commercial"
        scopeDetail = "Key Account"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("FOOD SERVICE")) Then

        scopeGroup = "Commercial"
        scopeDetail = "Food Service"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("COMMERCIAL")) Then

        scopeGroup = "Commercial"
        scopeDetail = "Commercial"
        Exit Sub

    End If

    ' --------------------------------------------------------
    ' MARKETING
    ' --------------------------------------------------------
    If RMR_ContainsAny( _
        combined, _
        Array("BRAND MARKETING")) Then

        scopeGroup = "Marketing"
        scopeDetail = "Brand Marketing"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("TRADE MARKETING")) Then

        scopeGroup = "Marketing"
        scopeDetail = "Trade Marketing"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array( _
            "MARCOMM", _
            "PUBLIC RELATION", _
            " PR ")) Then

        scopeGroup = "Marketing"
        scopeDetail = "Marcomm & PR"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("INTERNATIONAL MARKETING", "INTERNATIONAL MARKET")) Then

        scopeGroup = "Marketing"
        scopeDetail = "International Marketing"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("MARKETING")) Then

        scopeGroup = "Marketing"
        scopeDetail = "Marketing"
        Exit Sub

    End If

    ' --------------------------------------------------------
    ' BUSINESS / INNOVATION
    ' --------------------------------------------------------
    If RMR_ContainsAny( _
        combined, _
        Array("BUSINESS DEVELOPMENT")) Then

        scopeGroup = "Business Development"
        scopeDetail = "Business Development"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array( _
            "RESEARCH & INNOVATION", _
            "RESEARCH AND INNOVATION", _
            "R&I")) Then

        scopeGroup = "Research & Innovation"
        scopeDetail = "Research & Innovation"
        Exit Sub

    End If

    ' --------------------------------------------------------
    ' FINANCE / FBS
    ' --------------------------------------------------------
    If RMR_ContainsAny( _
        combined, _
        Array( _
            "FINANCIAL CONTROLLER", _
            "FINANCE CONTROLLER")) Then

        scopeGroup = "FBS / Finance"
        scopeDetail = "Financial Controller"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("INTERNAL AUDIT")) Then

        scopeGroup = "FBS / Finance"
        scopeDetail = "Internal Audit"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("PROCUREMENT")) Then

        scopeGroup = "FBS / Finance"
        scopeDetail = "Procurement"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("FINANCE")) Then

        scopeGroup = "FBS / Finance"
        scopeDetail = "Finance"
        Exit Sub

    End If

    ' --------------------------------------------------------
    ' PEOPLE / IT
    ' --------------------------------------------------------
    If RMR_ContainsAny( _
        combined, _
        Array( _
            "HRGA OPERATION", _
            "HRGA")) Then

        scopeGroup = "HRGA"
        scopeDetail = "HRGA Operation"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array( _
            "INFORMATION TECHNOLOGY", _
            "TECHNOLOGY", _
            " IT ")) Then

        scopeGroup = "IT"
        scopeDetail = "IT"
        Exit Sub

    End If

    ' --------------------------------------------------------
    ' MANUFACTURE / OPERATION
    ' --------------------------------------------------------
    If RMR_ContainsAny( _
        combined, _
        Array( _
            "MANUFACTURE", _
            "MANUFACTURING")) Then

        scopeGroup = "Manufacture"
        scopeDetail = "Manufacture"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array( _
            "SUPPLY CHAIN", _
            "LOGISTIC", _
            "LOGISTICS", _
            "DISTRIBUTION")) Then

        scopeGroup = "Supply Chain & Logistics"

        If RMR_ContainsAny( _
            combined, _
            Array("LOGISTIC", "LOGISTICS", "DISTRIBUTION")) Then

            scopeDetail = "Logistic & Distribution"

        Else

            scopeDetail = "Supply Chain & Operation"

        End If

        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("OPERATION CITEUREUP")) Then

        scopeGroup = "Operation"
        scopeDetail = "Operation Citeureup"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("OPERATION SEMARANG")) Then

        scopeGroup = "Operation"
        scopeDetail = "Operation Semarang"
        Exit Sub

    End If

    If RMR_ContainsAny( _
        combined, _
        Array("OPERATION")) Then

        scopeGroup = "Operation"
        scopeDetail = "Operation"
        Exit Sub

    End If

    ' --------------------------------------------------------
    ' FALLBACK
    ' Jangan pakai raw numeric/date serial sebagai scope.
    ' --------------------------------------------------------
    scopeGroup = "Other / Review"

    scopeDetail = _
        RMR_FirstUsefulScopeText( _
            division, _
            department, _
            directorate, _
            positionText)

    If Len(scopeDetail) = 0 Then
        scopeDetail = "Perlu Review"
    End If

End Sub

' ============================================================
' LEVEL SCOPE
' ============================================================
Private Function RMR_LevelScope( _
    ByVal filterText As String, _
    ByVal levelDetail As String, _
    ByVal levelNumber As String) As String

    Dim numText As String
    Dim detailText As String

    If RMR_IsCLAPFGDP(filterText) Then

        RMR_LevelScope = "All Level"
        Exit Function

    End If

    numText = RMR_ExtractLevelNumber(levelNumber)

    If Len(numText) = 0 Then
        numText = RMR_ExtractLevelNumber(levelDetail)
    End If

    detailText = RMR_CleanLevelDetail(levelDetail)

    Select Case numText

        Case "1", "2"

            RMR_LevelScope = "Level 1-2"

        Case "3"

            If Len(detailText) > 0 Then

                RMR_LevelScope = detailText

            Else

                RMR_LevelScope = "Level 3"

            End If

        Case "4"

            If Len(detailText) > 0 Then

                RMR_LevelScope = _
                    "Managerial (" & _
                    detailText & _
                    ")"

            Else

                RMR_LevelScope = "Managerial"

            End If

        Case Else

            RMR_LevelScope = "Unmapped Level"

    End Select

End Function

' ============================================================
' HELPERS - CLASSIFICATION
' ============================================================
Private Function RMR_FirstUsefulScopeText( _
    ParamArray choices() As Variant) As String

    Dim item As Variant
    Dim t As String

    For Each item In choices

        t = RMR_CleanDisplayText(CStr(item))

        If RMR_IsUsefulText(t) Then

            RMR_FirstUsefulScopeText = t
            Exit Function

        End If

    Next item

End Function

Private Function RMR_IsUsefulText( _
    ByVal rawText As String) As Boolean

    Dim t As String

    t = Trim$(rawText)

    If Len(t) = 0 Then Exit Function

    If IsNumeric(t) Then Exit Function

    If Len(t) <= 3 And _
       RMR_IsMostlyNumeric(t) Then

        Exit Function

    End If

    RMR_IsUsefulText = True

End Function

Private Function RMR_IsMostlyNumeric( _
    ByVal rawText As String) As Boolean

    Dim i As Long
    Dim numericCount As Long
    Dim otherCount As Long
    Dim ch As String

    For i = 1 To Len(rawText)

        ch = Mid$(rawText, i, 1)

        If ch >= "0" And ch <= "9" Then
            numericCount = numericCount + 1
        ElseIf ch <> " " And _
               ch <> "-" And _
               ch <> "/" Then
            otherCount = otherCount + 1
        End If

    Next i

    If numericCount > 0 And otherCount = 0 Then
        RMR_IsMostlyNumeric = True
    End If

End Function

' ============================================================
' UNIQUE VALUE COLLECTION
' ============================================================
Private Sub RMR_AddUniqueValue( _
    ByVal parentDict As Object, _
    ByVal parentKey As String, _
    ByVal valueText As String)

    Dim childDict As Object
    Dim cleanValue As String

    cleanValue = Trim$(valueText)

    If Len(cleanValue) = 0 Then Exit Sub

    If parentDict.Exists(parentKey) Then

        Set childDict = parentDict(parentKey)

    Else

        Set childDict = RMR_NewDict()

        parentDict.Add _
            parentKey, _
            childDict

    End If

    If Not childDict.Exists(cleanValue) Then
        childDict.Add cleanValue, True
    End If

End Sub

' ============================================================
' COUNTS
' ============================================================
Private Sub RMR_IncrementCount( _
    ByVal dict As Object, _
    ByVal keyText As String)

    If dict.Exists(keyText) Then

        dict(keyText) = _
            CLng(dict(keyText)) + 1

    Else

        dict.Add keyText, 1&

    End If

End Sub

Private Function RMR_GetCount( _
    ByVal dict As Object, _
    ByVal keyText As String) As Long

    If dict Is Nothing Then Exit Function

    If dict.Exists(keyText) Then
        RMR_GetCount = CLng(dict(keyText))
    End If

End Function

' ============================================================
' RECRUITER
' ============================================================
Private Function RMR_ResolveRecruiter( _
    ByVal rawPIC As String, _
    ByVal rawKodePIC As String, _
    ByVal rawKodeUnik As String) As String

    Dim picName As String
    Dim codeText As String

    ' 1. PIC Recruiter punya prioritas tertinggi.
    picName = _
        RMR_CanonicalRecruiterName( _
            rawPIC)

    If RMR_IsSelectedRecruiter(picName) Then

        RMR_ResolveRecruiter = picName
        Exit Function

    End If

    ' 2. Fallback Kode PIC.
    codeText = _
        UCase$( _
            RMR_CleanDisplayText( _
                rawKodePIC))

    picName = _
        RMR_RecruiterFromCode( _
            codeText)

    If Len(picName) > 0 Then

        RMR_ResolveRecruiter = picName
        Exit Function

    End If

    ' 3. Fallback Kode Unik.
    ' Ini menangkap STO insert yang PIC Recruiter / Kode PIC kosong.
    codeText = _
        UCase$( _
            RMR_CleanDisplayText( _
                rawKodeUnik))

    picName = _
        RMR_RecruiterFromCode( _
            codeText)

    If Len(picName) > 0 Then
        RMR_ResolveRecruiter = picName
    End If

End Function

Private Function RMR_RecruiterFromCode( _
    ByVal rawCode As String) As String

    Dim c As String

    c = UCase$(Trim$(rawCode))

    If Len(c) = 0 Then Exit Function

    ' Recruiter15
    If Left$(c, 7) = "CORPREC15" Then
        RMR_RecruiterFromCode = "Recruiter15"
        Exit Function
    End If

    ' Recruiter01
    If Left$(c, 8) = "CORPREC01" Or _
       Left$(c, 11) = "CORPRECRUITER01" Then
        RMR_RecruiterFromCode = "Recruiter01"
        Exit Function
    End If

    ' Desi
    If Left$(c, 8) = "CORPDESI" Then
        RMR_RecruiterFromCode = "Desi"
        Exit Function
    End If

    ' Recruiter02
    If Left$(c, 7) = "CORPRECRUITER02" Then
        RMR_RecruiterFromCode = "Recruiter02"
        Exit Function
    End If

    ' Recruiter05
    If Left$(c, 7) = "CORPREC05" Or _
       Left$(c, 13) = "CORPRECRUITER05" Then
        RMR_RecruiterFromCode = "Recruiter05"
        Exit Function
    End If

    ' Recruiter08 / Ome
    If Left$(c, 7) = "CORPREC08" Or _
       Left$(c, 9) = "CORPRECRUITER08" Then
        RMR_RecruiterFromCode = "Recruiter08"
        Exit Function
    End If

    ' Recruiter12
    If Left$(c, 7) = "CORPVIC" Or _
       Left$(c, 10) = "CORPRECRUITER12" Then
        RMR_RecruiterFromCode = "Recruiter12"
        Exit Function
    End If

    ' Yeremia / Yeremiaa
    If Left$(c, 7) = "CORPYER" Or _
       Left$(c, 10) = "CORPYeremia" Or _
       Left$(c, 11) = "CORPYeremiaA" Then
        RMR_RecruiterFromCode = "Yeremia"
        Exit Function
    End If

    ' Adista
    If Left$(c, 7) = "CORPADI" Or _
       Left$(c, 10) = "CORPADISTA" Then
        RMR_RecruiterFromCode = "Adista"
        Exit Function
    End If

    ' Recruiter07
    If Left$(c, 7) = "CORPREC07" Or _
       Left$(c, 9) = "CORPRECRUITER07" Then
        RMR_RecruiterFromCode = "Recruiter07"
        Exit Function
    End If

    ' Recruiter04
    If Left$(c, 7) = "CORPKAR" Or _
       Left$(c, 9) = "CORPRECRUITER04" Then
        RMR_RecruiterFromCode = "Recruiter04"
        Exit Function
    End If
    
    ' Recruiter13
    If Left$(c, 7) = "CORPRECRUITER13" Or _
       Left$(c, 9) = "CORPRECRUITER13" Then
        RMR_RecruiterFromCode = "Recruiter13"
        Exit Function
    End If

End Function

Private Function RMR_IsSelectedRecruiter( _
    ByVal recruiterName As String) As Boolean

    Select Case UCase$(Trim$(recruiterName))

        Case "RECRUITER15", _
             "RECRUITER01", _
             "DESI", _
             "RECRUITER02", _
             "RECRUITER05", _
             "RECRUITER08", _
             "RECRUITER12", _
             "YEREMIA", _
             "ADISTA", _
             "RECRUITER07", _
             "RECRUITER13", _
             "RECRUITER04"

            RMR_IsSelectedRecruiter = True

    End Select

End Function

Private Function RMR_CanonicalRecruiterName( _
    ByVal rawText As String) As String

    Dim t As String

    t = RMR_CleanDisplayText(rawText)

    If Len(t) = 0 Then Exit Function

    Select Case UCase$(t)

        Case "ADISTA": RMR_CanonicalRecruiterName = "Adista"
        Case "RECRUITER15": RMR_CanonicalRecruiterName = "Recruiter15"
        Case "RECRUITER01": RMR_CanonicalRecruiterName = "Recruiter01"
        Case "DESI": RMR_CanonicalRecruiterName = "Desi"
        Case "REC08", "RECRUITER08", "RECRUITER08 ": RMR_CanonicalRecruiterName = "Recruiter08"
        Case "RECRUITER02": RMR_CanonicalRecruiterName = "Recruiter02"
        Case "RECRUITER03": RMR_CanonicalRecruiterName = "Recruiter03"
        Case "RECRUITER04": RMR_CanonicalRecruiterName = "Recruiter04"
        Case "RECRUITER13": RMR_CanonicalRecruiterName = "Recruiter13"
        Case "BELUM ADA PIC": RMR_CanonicalRecruiterName = "Belum Ada PIC"
        Case "RECRUITER05": RMR_CanonicalRecruiterName = "Recruiter05"
        Case "RECRUITER07": RMR_CanonicalRecruiterName = "Recruiter07"
        Case "RECRUITER09": RMR_CanonicalRecruiterName = "Recruiter09"
        Case "RECRUITER14": RMR_CanonicalRecruiterName = "Recruiter14"
        Case "RECRUITER12": RMR_CanonicalRecruiterName = "Recruiter12"
        Case "YEREMIA", "YEREMIAA": RMR_CanonicalRecruiterName = "Yeremia"
        Case "RECRUITER06": RMR_CanonicalRecruiterName = "Recruiter06"
        Case "RECRUITER10": RMR_CanonicalRecruiterName = "Recruiter10"
        Case "KOSONG": RMR_CanonicalRecruiterName = "Kosong"

        Case Else
            RMR_CanonicalRecruiterName = t

    End Select

End Function

' ============================================================
' BU
' ============================================================
Private Function RMR_DisplayBU( _
    ByVal rawText As String) As String

    Dim t As String

    t = _
        UCase$( _
            RMR_CleanDisplayText( _
                rawText))

    Select Case True

        Case t = "CMD" Or _
             InStr(1, t, "CISARUA MOUNTAIN DAIRY", vbTextCompare) > 0

            RMR_DisplayBU = "CMD"

        Case t = "MP" Or _
             InStr(1, t, "MACROPRIMA PANGANUTAMA", vbTextCompare) > 0

            RMR_DisplayBU = "MP"

        Case t = "JES" Or _
             t = "JESS" Or _
             InStr(1, t, "JAVA EGG", vbTextCompare) > 0

            RMR_DisplayBU = "JES"

        Case t = "MS" Or _
             InStr(1, t, "MACROSENTRA", vbTextCompare) > 0

            RMR_DisplayBU = "MS"

        Case t = "MB" Or _
             InStr(1, t, "MACROTAMA", vbTextCompare) > 0

            RMR_DisplayBU = "MB"

        Case t = "BHC" Or _
             t = "BCH" Or _
             InStr(1, t, "BAVARIAN", vbTextCompare) > 0

            RMR_DisplayBU = "BHC"

        Case t = "ARC" Or _
             InStr(1, t, "ARTHA RASA CIMORY", vbTextCompare) > 0

            RMR_DisplayBU = "ARC"

        Case t = "HO" Or _
             t = "CORP" Or _
             t = "CORPORATE"

            RMR_DisplayBU = "HO"

        Case Else

            ' Jangan tampilkan serial date seperti 46234/46237.
            If IsNumeric(t) Then
                RMR_DisplayBU = "Unmapped BU"
            ElseIf Len(t) = 0 Then
                RMR_DisplayBU = "Unmapped BU"
            Else
                RMR_DisplayBU = RMR_CleanDisplayText(rawText)
            End If

    End Select

End Function

 ' ============================================================
' STATUS / FILTER
' ============================================================
Private Function RMR_IsOpenStatus( _
    ByVal statusText As String) As Boolean

    Dim t As String

    t = _
        UCase$( _
            RMR_CleanDisplayText( _
                statusText))

    RMR_IsOpenStatus = _
        (t = "OP" Or _
         t = "OPEN")

End Function

Private Function RMR_IsSTO( _
    ByVal filterText As String) As Boolean

    Dim t As String
    Dim compactText As String
    Dim tokens As Variant
    Dim item As Variant

    t = _
        UCase$( _
            RMR_CleanDisplayText( _
                filterText))

    If Len(t) = 0 Then Exit Function

    t = Replace$(t, "-", " ")
    t = Replace$(t, "/", " ")
    t = Replace$(t, "\", " ")
    t = Replace$(t, "_", " ")
    t = Replace$(t, ".", " ")
    t = Replace$(t, ",", " ")
    t = Replace$(t, ";", " ")
    t = Replace$(t, ":", " ")
    t = Replace$(t, "(", " ")
    t = Replace$(t, ")", " ")
    t = Replace$(t, "[", " ")
    t = Replace$(t, "]", " ")

    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace$(t, "  ", " ")
    Loop

    t = Trim$(t)
    compactText = Replace$(t, " ", "")

    If InStr(1, compactText, "NONSTO", vbTextCompare) > 0 Or _
       InStr(1, t, "BUKAN STO", vbTextCompare) > 0 Or _
       InStr(1, t, "NOT STO", vbTextCompare) > 0 Then
        Exit Function
    End If

    tokens = Split(t, " ")

    For Each item In tokens
        If UCase$(Trim$(CStr(item))) = "STO" Then
            RMR_IsSTO = True
            Exit Function
        End If
    Next item

End Function

Private Function RMR_IsNoPIC( _
    ByVal recruiterName As String) As Boolean

    Dim t As String

    t = _
        UCase$( _
            RMR_CleanDisplayText( _
                recruiterName))

    RMR_IsNoPIC = _
        (Len(t) = 0 Or _
         t = "KOSONG" Or _
         t = "BELUM ADA PIC" Or _
         t = "-" Or _
         t = "N/A")

End Function

Private Function RMR_IsCLAPFGDP( _
    ByVal filterText As String) As Boolean

    Dim t As String

    t = UCase$(filterText)

    RMR_IsCLAPFGDP = _
        InStr(1, t, "CLAP", vbTextCompare) > 0 Or _
        InStr(1, t, "FGDP", vbTextCompare) > 0

End Function

' ============================================================
' LEVEL
' ============================================================
Private Function RMR_ExtractLevelNumber( _
    ByVal rawText As String) As String

    Dim i As Long
    Dim ch As String

    For i = 1 To Len(rawText)

        ch = Mid$(rawText, i, 1)

        If ch >= "1" And ch <= "9" Then

            RMR_ExtractLevelNumber = ch
            Exit Function

        End If

    Next i

End Function

Private Function RMR_CleanLevelDetail( _
    ByVal rawText As String) As String

    Dim t As String
    Dim i As Long
    Dim ch As String
    Dim resultText As String

    t = UCase$(Trim$(rawText))

    For i = 1 To Len(t)

        ch = Mid$(t, i, 1)

        If (ch >= "0" And ch <= "9") Or _
           (ch >= "A" And ch <= "Z") Then

            resultText = resultText & ch

        End If

    Next i

    Select Case resultText

        Case "3A", "3B", "3C", _
             "4A", "4B", "4C"

            RMR_CleanLevelDetail = resultText

    End Select

End Function


' ============================================================
' DASHBOARD / VISUAL HELPERS
' ============================================================
Private Sub RMR_ResetOutputSheet( _
    ByVal ws As Worksheet)

    Dim i As Long

    On Error Resume Next

    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i

    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Delete
    Next i

    On Error GoTo 0

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.FormatConditions.Delete

End Sub

Private Sub RMR_StyleDashboardTitle( _
    ByVal ws As Worksheet)

    With ws.Range("A1:J1")

        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

    With ws.Range("A2:J2")

        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(68, 68, 68)
        .Font.Size = 10
        .HorizontalAlignment = xlLeft

    End With

End Sub

Private Sub RMR_WriteKPICard( _
    ByVal ws As Worksheet, _
    ByVal rangeAddress As String, _
    ByVal titleText As String, _
    ByVal valueNumber As Long, _
    ByVal subtitleText As String)

    Dim rg As Range
    Dim firstCell As Range
    Dim secondRow As Range
    Dim thirdRow As Range

    Set rg = ws.Range(rangeAddress)

    rg.UnMerge

    Set firstCell = _
        ws.Range( _
            ws.Cells(rg.Row, rg.Column), _
            ws.Cells(rg.Row, rg.Column + rg.Columns.Count - 1))

    Set secondRow = _
        ws.Range( _
            ws.Cells(rg.Row + 1, rg.Column), _
            ws.Cells(rg.Row + 1, rg.Column + rg.Columns.Count - 1))

    Set thirdRow = _
        ws.Range( _
            ws.Cells(rg.Row + 2, rg.Column), _
            ws.Cells(rg.Row + 2, rg.Column + rg.Columns.Count - 1))

    firstCell.Merge
    secondRow.Merge
    thirdRow.Merge

    firstCell.Value2 = titleText
    secondRow.Value2 = valueNumber
    thirdRow.Value2 = subtitleText

    With rg

        .Interior.Color = RGB(242, 246, 252)

        With .Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 210, 225)
            .Weight = xlThin
        End With

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    With firstCell
        .Font.Bold = True
        .Font.Color = RGB(68, 84, 106)
        .Font.Size = 9
    End With

    With secondRow
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .Font.Size = 20
    End With

    With thirdRow
        .Font.Color = RGB(110, 110, 110)
        .Font.Size = 8
    End With

End Sub

Private Sub RMR_StyleSectionTitle( _
    ByVal rg As Range)

    With rg

        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

End Sub

Private Sub RMR_CreateVisualTable( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal headerRow As Long, _
    ByVal firstCol As Long, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long, _
    ByVal styleName As String)

    Dim rg As Range
    Dim lo As ListObject

    If lastRow <= headerRow Then Exit Sub

    Set rg = _
        ws.Range( _
            ws.Cells(headerRow, firstCol), _
            ws.Cells(lastRow, lastCol))

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    If Not lo Is Nothing Then lo.Delete
    Set lo = Nothing
    On Error GoTo 0

    Set lo = _
        ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=rg, _
            XlListObjectHasHeaders:=xlYes)

    lo.Name = tableName
    lo.tableStyle = styleName
    lo.ShowAutoFilter = True
    lo.ShowTableStyleRowStripes = True

End Sub

Private Sub RMR_AddDataBar( _
    ByVal rg As Range)

    Dim db As Databar

    If rg Is Nothing Then Exit Sub
    If rg.rows.Count = 0 Then Exit Sub

    rg.FormatConditions.Delete

    Set db = _
        rg.FormatConditions.AddDatabar

    db.BarColor.Color = RGB(91, 155, 213)
    db.ShowValue = True

End Sub


Private Sub RMR_HighlightUnmapped( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long)

    Dim r As Long
    Dim buText As String
    Dim scopeText As String
    Dim levelText As String

    For r = firstRow To lastRow

        scopeText = _
            UCase$( _
                CStr( _
                    ws.Cells(r, 2).Value2))

        buText = _
            UCase$( _
                CStr( _
                    ws.Cells(r, 4).Value2))

        levelText = _
            UCase$( _
                CStr( _
                    ws.Cells(r, 5).Value2))

        If InStr(1, scopeText, "OTHER / REVIEW", vbTextCompare) > 0 Or _
           InStr(1, buText, "UNMAPPED", vbTextCompare) > 0 Or _
           InStr(1, levelText, "UNMAPPED", vbTextCompare) > 0 Then

            ws.Range( _
                ws.Cells(r, 1), _
                ws.Cells(r, 9)).Interior.Color = _
                    RGB(255, 242, 204)

        End If

    Next r

End Sub




Private Sub RMR_ApplyRecruiterColors( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long)

    Dim r As Long
    Dim recruiterName As String
    Dim fillColor As Long

    If lastRow < firstRow Then Exit Sub

    For r = firstRow To lastRow

        recruiterName = _
            RMR_CanonicalRecruiterName( _
                CStr( _
                    ws.Cells(r, 1).Value2))

        fillColor = _
            RMR_RecruiterColor( _
                recruiterName)

        If fillColor <> 0 Then

            With ws.Range( _
                ws.Cells(r, 1), _
                ws.Cells(r, lastCol))

                .Interior.Color = fillColor
                .Font.Color = RGB(45, 45, 45)

            End With

            ws.Cells(r, 1).Font.Bold = True

        End If

    Next r

End Sub

Private Function RMR_RecruiterColor( _
    ByVal recruiterName As String) As Long

    Select Case UCase$(Trim$(recruiterName))

        Case "RECRUITER15"
            RMR_RecruiterColor = RGB(221, 235, 247)

        Case "RECRUITER01"
            RMR_RecruiterColor = RGB(252, 228, 214)

        Case "DESI"
            RMR_RecruiterColor = RGB(226, 239, 218)

        Case "RECRUITER02"
            RMR_RecruiterColor = RGB(255, 242, 204)

        Case "RECRUITER05"
            RMR_RecruiterColor = RGB(228, 223, 236)

        Case "RECRUITER08"
            RMR_RecruiterColor = RGB(218, 238, 243)

        Case "RECRUITER12"
            RMR_RecruiterColor = RGB(244, 204, 204)

        Case "YEREMIA"
            RMR_RecruiterColor = RGB(217, 217, 217)

        Case "ADISTA"
            RMR_RecruiterColor = RGB(234, 209, 220)

        Case "RECRUITER07"
            RMR_RecruiterColor = RGB(255, 229, 238)

        Case "RECRUITER04"
            RMR_RecruiterColor = RGB(208, 224, 227)
            
        Case "BELUM ADA PIC"
            RMR_RecruiterColor = RGB(255, 199, 206)

    End Select

End Function

' ============================================================
' GENERAL STRING HELPERS
' ============================================================
Private Function RMR_NewDict() As Object

    Set RMR_NewDict = _
        CreateObject( _
            "Scripting.Dictionary")

    RMR_NewDict.CompareMode = vbTextCompare

End Function

Private Function RMR_Key( _
    ByVal values As Variant) As String

    Dim i As Long

    For i = LBound(values) To UBound(values)

        If Len(RMR_Key) > 0 Then
            RMR_Key = RMR_Key & RMR_SEP
        End If

        RMR_Key = _
            RMR_Key & _
            RMR_KeyPart( _
                CStr(values(i)))

    Next i

End Function

Private Function RMR_KeyPart( _
    ByVal rawText As String) As String

    Dim t As String

    t = _
        UCase$( _
            RMR_CleanDisplayText( _
                rawText))

    If Len(t) = 0 Then t = "(BLANK)"

    RMR_KeyPart = t

End Function

Private Function RMR_DisplayKeyPart( _
    ByVal keyText As String) As String

    If UCase$(keyText) = "(BLANK)" Then
        RMR_DisplayKeyPart = ""
    Else
        RMR_DisplayKeyPart = RMR_ToTitleSmart(keyText)
    End If

End Function

Private Function RMR_CleanDisplayText( _
    ByVal rawText As String) As String

    Dim t As String

    t = rawText

    t = Replace$(t, Chr$(160), " ")
    t = Replace$(t, vbCr, " ")
    t = Replace$(t, vbLf, " ")

    t = Trim$(t)

    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace$(t, "  ", " ")
    Loop

    RMR_CleanDisplayText = t

End Function

Private Function RMR_ToTitleSmart( _
    ByVal rawText As String) As String

    Dim t As String

    t = _
        StrConv( _
            LCase$( _
                RMR_CleanDisplayText( _
                    rawText)), _
            vbProperCase)

    t = Replace$(t, "Fbs", "FBS")
    t = Replace$(t, "Hrga", "HRGA")
    t = Replace$(t, " It", " IT")
    t = Replace$(t, "It ", "IT ")
    t = Replace$(t, "Cmd", "CMD")
    t = Replace$(t, "Mp", "MP")
    t = Replace$(t, "Jes", "JES")
    t = Replace$(t, "Ms", "MS")
    t = Replace$(t, "Sto", "STO")
    t = Replace$(t, "Clap", "CLAP")
    t = Replace$(t, "Fgdp", "FGDP")
    t = Replace$(t, "Pr", "PR")

    RMR_ToTitleSmart = t

End Function

Private Function RMR_ContainsAny( _
    ByVal haystack As String, _
    ByVal needles As Variant) As Boolean

    Dim item As Variant

    For Each item In needles

        If InStr( _
            1, _
            haystack, _
            CStr(item), _
            vbTextCompare) > 0 Then

            RMR_ContainsAny = True
            Exit Function

        End If

    Next item

End Function

Private Function RMR_JoinSortedDictKeys( _
    ByVal dict As Object, _
    ByVal delimiter As String) As String

    Dim arr As Variant
    Dim i As Long

    If dict Is Nothing Then Exit Function
    If dict.Count = 0 Then Exit Function

    arr = dict.keys

    RMR_SortStringArray arr

    For i = LBound(arr) To UBound(arr)

        RMR_AppendText _
            RMR_JoinSortedDictKeys, _
            CStr(arr(i)), _
            delimiter

    Next i

End Function

Private Sub RMR_AppendText( _
    ByRef targetText As String, _
    ByVal addText As String, _
    ByVal delimiter As String)

    If Len(addText) = 0 Then Exit Sub

    If Len(targetText) > 0 Then
        targetText = targetText & delimiter
    End If

    targetText = targetText & addText

End Sub

Private Sub RMR_SortStringArray( _
    ByRef arr As Variant)

    Dim i As Long
    Dim j As Long
    Dim temp As Variant

    If Not IsArray(arr) Then Exit Sub

    For i = LBound(arr) To UBound(arr) - 1

        For j = i + 1 To UBound(arr)

            If StrComp( _
                CStr(arr(i)), _
                CStr(arr(j)), _
                vbTextCompare) > 0 Then

                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp

            End If

        Next j

    Next i

End Sub

' ============================================================
' ARRAY HELPERS
' ============================================================
Private Function RMR_ArrayText( _
    ByRef dataArr As Variant, _
    ByVal rowIndex As Long, _
    ByVal colIndex As Long) As String

    If colIndex <= 0 Then Exit Function

    If IsError(dataArr(rowIndex, colIndex)) Then Exit Function

    RMR_ArrayText = _
        RMR_CleanDisplayText( _
            CStr( _
                dataArr( _
                    rowIndex, _
                    colIndex)))

End Function

' ============================================================
' WORKSHEET HELPERS
' ============================================================
Private Function RMR_GetSheet( _
    ByVal sheetName As String) As Worksheet

    On Error Resume Next

    Set RMR_GetSheet = _
        ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

End Function

Private Function RMR_GetOrCreateSheet( _
    ByVal sheetName As String) As Worksheet

    Set RMR_GetOrCreateSheet = _
        RMR_GetSheet(sheetName)

    If RMR_GetOrCreateSheet Is Nothing Then

        Set RMR_GetOrCreateSheet = _
            ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        RMR_GetOrCreateSheet.Name = sheetName

    End If

End Function

Private Function RMR_FindHeader( _
    ByVal ws As Worksheet, _
    ByVal headerName As String) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim wantedKey As String
    Dim currentKey As String

    wantedKey = RMR_HeaderKey(headerName)

    lastCol = _
        ws.Cells( _
            RMR_HEADER_ROW, _
            ws.Columns.Count).End( _
                xlToLeft).Column

    For c = 1 To lastCol

        currentKey = _
            RMR_HeaderKey( _
                CStr( _
                    ws.Cells( _
                        RMR_HEADER_ROW, _
                        c).Value2))

        If StrComp( _
            currentKey, _
            wantedKey, _
            vbTextCompare) = 0 Then

            RMR_FindHeader = c
            Exit Function

        End If

    Next c

End Function

Private Function RMR_HeaderKey( _
    ByVal rawText As String) As String

    Dim sourceText As String
    Dim resultText As String
    Dim i As Long
    Dim ch As String

    sourceText = UCase$(rawText)

    For i = 1 To Len(sourceText)

        ch = Mid$(sourceText, i, 1)

        If (ch >= "A" And ch <= "Z") Or _
           (ch >= "0" And ch <= "9") Then

            resultText = resultText & ch

        End If

    Next i

    RMR_HeaderKey = resultText

End Function

Private Function RMR_LastDataRow( _
    ByVal ws As Worksheet) As Long

    Dim colPosisi As Long
    Dim colKodeUnik As Long
    Dim lastPosisi As Long
    Dim lastKode As Long

    If ws Is Nothing Then
        RMR_LastDataRow = RMR_HEADER_ROW
        Exit Function
    End If

    ' Jangan pakai UsedRange / Find("*") karena FPTK bisa
    ' punya helper/dropdown di bawah data utama.
    colPosisi = RMR_FindHeader(ws, "Posisi")
    colKodeUnik = RMR_FindHeader(ws, "Kode Unik")

    If colPosisi > 0 Then

        lastPosisi = _
            ws.Cells( _
                ws.rows.Count, _
                colPosisi).End( _
                    xlUp).Row

    End If

    If colKodeUnik > 0 Then

        lastKode = _
            ws.Cells( _
                ws.rows.Count, _
                colKodeUnik).End( _
                    xlUp).Row

    End If

    RMR_LastDataRow = _
        Application.Max( _
            RMR_HEADER_ROW, _
            lastPosisi, _
            lastKode)

End Function


