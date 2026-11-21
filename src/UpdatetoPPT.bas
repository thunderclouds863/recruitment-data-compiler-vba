Attribute VB_Name = "UpdatetoPPT"
Option Explicit

' ============================================================
' POWERPOINT CONSTANTS
' Late binding: tidak perlu centang PowerPoint Object Library
' ============================================================
Private Const PP_PASTE_DEFAULT As Long = 0
Private Const PP_PASTE_ENHANCED_METAFILE As Long = 2
Private Const PP_AUTOSIZE_SHAPE_TO_FIT_TEXT As Long = 1

Private Const MSO_TRUE As Long = -1
Private Const MSO_FALSE As Long = 0


' ============================================================
' UPDATE POWERPOINT FROM MAPPING
'
' SHEET OUTPUT_PPT:
' A = Shape Name di PowerPoint
' B = Nomor Slide
' C = Source Value / Range / Nama Chart Excel
' D = Source Sheet
' E = Object Type
' H = Update Flag, isi YES
' I = AutoFit, isi FIT
' K = Status
' L = Waktu Update
'
' Lokasi PowerPoint:
' OUTPUT_PPT!B2
'
' Object Type:
' TEXTBOX
' PPT_CHART_DATA
' CHART_MPP_LINK
' CHART_MPP_PICTURE
' TABLE_RANGE
' ============================================================
Public Sub UpdatePowerPointFromMapping()

    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim pastedShapes As Object

    Dim wsMapping As Worksheet
    Dim wsSource As Worksheet
    Dim rngSource As Range

    Dim chartObj As ChartObject
    Dim chartObjPic As ChartObject

    Dim chtWorkbook As Object
    Dim chtSheet As Object

    Dim lastRow As Long
    Dim i As Long
    Dim slideNum As Long

    Dim shapeName As String
    Dim sourceSheet As String
    Dim sourceAddress As String
    Dim chartSourceName As String
    Dim objType As String
    Dim updateFlag As String
    Dim autoFitFlag As String
    Dim pptPath As String
    Dim textValue As String

    Dim sourceValue As Variant

    Dim hasText As Boolean
    Dim hasChart As Boolean

    Dim oldLeft As Single
    Dim oldTop As Single
    Dim oldWidth As Single
    Dim oldHeight As Single
    Dim oldRotation As Single
    Dim oldLockAspect As Long

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    Dim fatalNumber As Long
    Dim fatalDescription As String

    On Error GoTo FatalError

    Set wsMapping = ThisWorkbook.Worksheets("OUTPUT_PPT")

    ' Ambil alamat PPT dari OUTPUT_PPT!B2
    pptPath = GetPowerPointPath(wsMapping.Range("B2"))

    If Len(pptPath) = 0 Then

        MsgBox _
            "Alamat file PowerPoint pada OUTPUT_PPT!B2 masih kosong.", _
            vbExclamation, _
            "Update PowerPoint"

        Exit Sub

    End If

    If Not FileExists(pptPath) Then

        MsgBox _
            "File PowerPoint tidak ditemukan:" & vbCrLf & vbCrLf & _
            pptPath, _
            vbExclamation, _
            "Update PowerPoint"

        Exit Sub

    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Membuka PowerPoint..."

    ' Hitung ulang formula pada sheet mapping
    wsMapping.Calculate

    ' ========================================================
    ' BUKA ATAU AMBIL POWERPOINT YANG SUDAH TERBUKA
    ' ========================================================
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    On Error GoTo FatalError

    If pptApp Is Nothing Then

        Set pptApp = CreateObject("PowerPoint.Application")

    End If

    pptApp.Visible = True

    Set pptPres = GetOpenPresentationByPath( _
        pptApp, _
        pptPath _
    )

    If pptPres Is Nothing Then

        Set pptPres = pptApp.Presentations.Open(pptPath)

    End If

    ' ========================================================
    ' LOOP MAPPING
    ' ========================================================
    lastRow = wsMapping.Cells( _
        wsMapping.rows.Count, _
        "A" _
    ).End(xlUp).Row

    For i = 4 To lastRow

        On Error GoTo RowError

        Set pptSlide = Nothing
        Set pptShape = Nothing
        Set pastedShapes = Nothing

        Set wsSource = Nothing
        Set rngSource = Nothing

        Set chartObj = Nothing
        Set chartObjPic = Nothing

        Set chtWorkbook = Nothing
        Set chtSheet = Nothing

        shapeName = Trim$(CStr( _
            wsMapping.Cells(i, "A").Value2 _
        ))

        slideNum = val( _
            wsMapping.Cells(i, "B").Value2 _
        )

        sourceValue = _
            wsMapping.Cells(i, "C").Value2

        sourceSheet = Trim$(CStr( _
            wsMapping.Cells(i, "D").Value2 _
        ))

        objType = UCase$(Trim$(CStr( _
            wsMapping.Cells(i, "E").Value2 _
        )))

        updateFlag = UCase$(Trim$(CStr( _
            wsMapping.Cells(i, "H").Value2 _
        )))

        autoFitFlag = UCase$(Trim$(CStr( _
            wsMapping.Cells(i, "I").Value2 _
        )))

        ' Lewati mapping yang tidak aktif
        If updateFlag <> "YES" Then GoTo NextRow
        If Len(shapeName) = 0 Then GoTo NextRow
        If slideNum <= 0 Then GoTo NextRow
        If Len(objType) = 0 Then GoTo NextRow

        wsMapping.Cells(i, "K").Value = "PROCESSING"
        wsMapping.Cells(i, "L").Value = Now

        ' Cari slide target
        On Error Resume Next

        Set pptSlide = _
            pptPres.Slides(slideNum)

        On Error GoTo RowError

        If pptSlide Is Nothing Then

            Err.Raise _
                vbObjectError + 1001, _
                , _
                "Slide " & slideNum & " tidak ditemukan."

        End If

        Application.StatusBar = _
            "Update slide " & slideNum & _
            " - " & shapeName & "..."

        ' ====================================================
        ' PROSES OBJECT TYPE
        ' ====================================================
        Select Case objType

            ' =================================================
            ' TEXTBOX
            ' =================================================
            Case "TEXTBOX"

                Set pptShape = FindShape( _
                    pptSlide, _
                    shapeName _
                )

                If pptShape Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1101, _
                        , _
                        "Textbox '" & shapeName & _
                        "' tidak ditemukan."

                End If

                hasText = False

                On Error Resume Next
                hasText = CBool(pptShape.HasTextFrame)
                On Error GoTo RowError

                If Not hasText Then

                    Err.Raise _
                        vbObjectError + 1102, _
                        , _
                        "Shape '" & shapeName & _
                        "' bukan textbox."

                End If

                ' Ambil sesuai tampilan Excel
                textValue = GetExcelDisplayText( _
                    wsMapping.Cells(i, "C") _
                )

                pptShape.TextFrame.textRange.text = _
                    textValue

                ' AutoFit jika kolom I berisi FIT
                If autoFitFlag = "FIT" Then

                    On Error Resume Next

                    pptShape.TextFrame.WordWrap = _
                        MSO_TRUE

                    pptShape.TextFrame.AutoSize = _
                        PP_AUTOSIZE_SHAPE_TO_FIT_TEXT

                    On Error GoTo RowError

                End If

                wsMapping.Cells(i, "K").Value = _
                    "OK - Text updated"

                wsMapping.Cells(i, "L").Value = Now


            ' =================================================
            ' UPDATE DATA CHART NATIVE POWERPOINT
            ' =================================================
            Case "PPT_CHART_DATA"

                sourceAddress = SafeText(sourceValue)

                If Len(sourceAddress) = 0 Or _
                   UCase$(sourceAddress) = "AUTO" Then

                    GoTo NextRow

                End If

                Set rngSource = GetSourceRange( _
                    ThisWorkbook, _
                    sourceSheet, _
                    sourceAddress _
                )

                If rngSource Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1201, _
                        , _
                        "Range '" & sourceAddress & _
                        "' pada sheet '" & sourceSheet & _
                        "' tidak ditemukan."

                End If

                Set pptShape = FindShape( _
                    pptSlide, _
                    shapeName _
                )

                If pptShape Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1202, _
                        , _
                        "Chart '" & shapeName & _
                        "' tidak ditemukan."

                End If

                hasChart = False

                On Error Resume Next
                hasChart = CBool(pptShape.hasChart)
                On Error GoTo RowError

                If Not hasChart Then

                    Err.Raise _
                        vbObjectError + 1203, _
                        , _
                        "Shape '" & shapeName & _
                        "' bukan chart PowerPoint."

                End If

                pptShape.Chart.ChartData.Activate

                Set chtWorkbook = _
                    pptShape.Chart.ChartData.Workbook

                Set chtSheet = _
                    chtWorkbook.Worksheets(1)

                chtSheet.Cells.Clear

                rngSource.Copy

                chtSheet.Range("A1").PasteSpecial

                Application.CutCopyMode = False

                pptShape.Chart.Refresh

                wsMapping.Cells(i, "K").Value = _
                    "OK - Chart data updated"

                wsMapping.Cells(i, "L").Value = Now


            ' =================================================
            ' CHART EXCEL KE PPT SEBAGAI LINK
            ' =================================================
            Case "CHART_MPP_LINK"

                If Len(sourceSheet) = 0 Then

                    Err.Raise _
                        vbObjectError + 1301, _
                        , _
                        "Source Sheet masih kosong."

                End If

                Set wsSource = GetWorksheetSafe( _
                    ThisWorkbook, _
                    sourceSheet _
                )

                If wsSource Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1302, _
                        , _
                        "Sheet '" & sourceSheet & _
                        "' tidak ditemukan."

                End If

                sourceAddress = SafeText(sourceValue)

                If Len(sourceAddress) = 0 Or _
                   UCase$(sourceAddress) = "AUTO" Then

                    chartSourceName = shapeName

                Else

                    chartSourceName = sourceAddress

                End If

                Set chartObj = FindExcelChart( _
                    wsSource, _
                    chartSourceName _
                )

                If chartObj Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1303, _
                        , _
                        "Chart Excel '" & chartSourceName & _
                        "' tidak ditemukan pada sheet '" & _
                        sourceSheet & "'."

                End If

                Set pptShape = FindShape( _
                    pptSlide, _
                    shapeName _
                )

                SaveShapeGeometry _
                    pptShape, _
                    oldLeft, _
                    oldTop, _
                    oldWidth, _
                    oldHeight, _
                    oldRotation, _
                    oldLockAspect

                chartObj.Copy

                If Not pptShape Is Nothing Then
                    pptShape.Delete
                End If

                Set pastedShapes = Nothing

                On Error Resume Next
                Err.Clear

                ' Paste sebagai link
                Set pastedShapes = _
                    pptSlide.Shapes.PasteSpecial( _
                        PP_PASTE_DEFAULT, _
                        MSO_FALSE, _
                        "", _
                        0, _
                        "", _
                        MSO_TRUE _
                    )

                ' Fallback jika paste link tidak tersedia
                If pastedShapes Is Nothing Then

                    Err.Clear

                    Set pastedShapes = _
                        pptSlide.Shapes.PasteSpecial( _
                            PP_PASTE_DEFAULT _
                        )

                End If

                On Error GoTo RowError

                If pastedShapes Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1304, _
                        , _
                        "Chart gagal dipaste ke PowerPoint."

                End If

                Set pptShape = pastedShapes.item(1)

                RestoreShapeGeometry _
                    pptShape, _
                    shapeName, _
                    oldLeft, _
                    oldTop, _
                    oldWidth, _
                    oldHeight, _
                    oldRotation, _
                    oldLockAspect

                Application.CutCopyMode = False

                wsMapping.Cells(i, "K").Value = _
                    "OK - Chart linked"

                wsMapping.Cells(i, "L").Value = Now


            ' =================================================
            ' CHART EXCEL KE PPT SEBAGAI PICTURE
            '
            ' Picture lama di PPT dipakai sebagai template:
            ' - posisi tetap
            ' - ukuran tetap
            ' - rotation tetap
            ' - crop tetap jika tersedia
            ' - border/effect dicoba dipertahankan
            ' =================================================
            Case "CHART_MPP_PICTURE"

                If Len(sourceSheet) = 0 Then

                    Err.Raise _
                        vbObjectError + 1401, _
                        , _
                        "Source Sheet masih kosong."

                End If

                Set wsSource = GetWorksheetSafe( _
                    ThisWorkbook, _
                    sourceSheet _
                )

                If wsSource Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1402, _
                        , _
                        "Sheet '" & sourceSheet & _
                        "' tidak ditemukan."

                End If

                sourceAddress = SafeText(sourceValue)

                ' Kolom C boleh diisi nama chart Excel.
                ' Jika AUTO/kosong, gunakan nama shape PPT.
                If Len(sourceAddress) = 0 Or _
                   UCase$(sourceAddress) = "AUTO" Then

                    chartSourceName = shapeName

                Else

                    chartSourceName = sourceAddress

                End If

                Set chartObjPic = FindExcelChart( _
                    wsSource, _
                    chartSourceName _
                )

                If chartObjPic Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1403, _
                        , _
                        "Chart Excel '" & chartSourceName & _
                        "' tidak ditemukan pada sheet '" & _
                        sourceSheet & "'."

                End If

                Set pptShape = FindShape( _
                    pptSlide, _
                    shapeName _
                )

                If pptShape Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1404, _
                        , _
                        "Picture '" & shapeName & _
                        "' tidak ditemukan pada slide " & _
                        slideNum & "."

                End If

                ReplaceChartWithPicturePreserveFormat _
                    pptSlide, _
                    pptShape, _
                    chartObjPic, _
                    shapeName

                Application.CutCopyMode = False

                wsMapping.Cells(i, "K").Value = _
                    "OK - Picture replaced"

                wsMapping.Cells(i, "L").Value = Now


            ' =================================================
            ' UPDATE TABEL TANPA MENGHAPUS TABEL PPT
            '
            ' Hanya isi cell yang diganti.
            ' Font, ukuran, warna, border, fill,
            ' posisi, lebar kolom, dan tinggi baris tetap.
            ' =================================================
            Case "TABLE_RANGE"

                sourceAddress = SafeText(sourceValue)

                If Len(sourceAddress) = 0 Or _
                   UCase$(sourceAddress) = "AUTO" Then

                    GoTo NextRow

                End If

                Set rngSource = GetSourceRange( _
                    ThisWorkbook, _
                    sourceSheet, _
                    sourceAddress _
                )

                If rngSource Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1501, _
                        , _
                        "Range '" & sourceAddress & _
                        "' pada sheet '" & sourceSheet & _
                        "' tidak ditemukan."

                End If

                Set pptShape = FindShape( _
                    pptSlide, _
                    shapeName _
                )

                If pptShape Is Nothing Then

                    Err.Raise _
                        vbObjectError + 1502, _
                        , _
                        "Tabel '" & shapeName & _
                        "' tidak ditemukan."

                End If

                If Not ShapeHasTable(pptShape) Then

                    Err.Raise _
                        vbObjectError + 1503, _
                        , _
                        "Shape '" & shapeName & _
                        "' bukan tabel native PowerPoint."

                End If

                UpdatePowerPointTablePreserveFormat _
                    pptShape, _
                    rngSource

                wsMapping.Cells(i, "K").Value = _
                    "OK - Table content updated"

                wsMapping.Cells(i, "L").Value = Now


            Case Else

                Err.Raise _
                    vbObjectError + 1601, _
                    , _
                    "Object Type '" & objType & _
                    "' tidak dikenali."

        End Select

        GoTo NextRow


RowError:

        wsMapping.Cells(i, "K").Value = _
            "FAILED - " & Err.Number & _
            " - " & Err.Description

        wsMapping.Cells(i, "L").Value = Now

        Debug.Print _
            "Row " & i & _
            " | Slide " & slideNum & _
            " | Shape " & shapeName & _
            " | Error " & Err.Number & _
            " | " & Err.Description

        Err.Clear
        Resume NextRow


NextRow:

        On Error GoTo FatalError

        Set pptSlide = Nothing
        Set pptShape = Nothing
        Set pastedShapes = Nothing

        Set wsSource = Nothing
        Set rngSource = Nothing

        Set chartObj = Nothing
        Set chartObjPic = Nothing

        Set chtWorkbook = Nothing
        Set chtSheet = Nothing

    Next i

    ' ========================================================
    ' SIMPAN POWERPOINT
    ' ========================================================
    Application.StatusBar = _
        "Menyimpan PowerPoint..."

    pptPres.Save
    pptApp.Activate

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.StatusBar = False
    Application.CutCopyMode = False

    MsgBox _
        "PowerPoint berhasil di-update dari mapping.", _
        vbInformation, _
        "Update PowerPoint Selesai"

    Set pptSlide = Nothing
    Set pptShape = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing

    Exit Sub


FatalError:

    fatalNumber = Err.Number
    fatalDescription = Err.Description

    On Error Resume Next

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.StatusBar = False
    Application.CutCopyMode = False

    On Error GoTo 0

    MsgBox _
        "Update PowerPoint gagal." & vbCrLf & vbCrLf & _
        "Error " & fatalNumber & ": " & _
        fatalDescription, _
        vbCritical, _
        "Update PowerPoint"

End Sub


' ============================================================
' GANTI CHART MENJADI PICTURE
' DENGAN MEMPERTAHANKAN SETTING PICTURE LAMA
' ============================================================
Private Sub ReplaceChartWithPicturePreserveFormat( _
    ByVal pptSlide As Object, _
    ByVal oldPicture As Object, _
    ByVal chartObj As ChartObject, _
    ByVal requestedShapeName As String)

    Dim pastedShapes As Object
    Dim newPicture As Object

    Dim savedLeft As Single
    Dim savedTop As Single
    Dim savedWidth As Single
    Dim savedHeight As Single
    Dim savedRotation As Single
    Dim savedLockAspect As Long

    Dim savedCropLeft As Single
    Dim savedCropRight As Single
    Dim savedCropTop As Single
    Dim savedCropBottom As Single

    Dim savedAlternativeText As String
    Dim savedTitle As String

    Dim cropCanBePreserved As Boolean
    Dim formatWasPickedUp As Boolean

    If pptSlide Is Nothing Then

        Err.Raise _
            vbObjectError + 3001, _
            , _
            "Slide PowerPoint belum ditentukan."

    End If

    If oldPicture Is Nothing Then

        Err.Raise _
            vbObjectError + 3002, _
            , _
            "Picture lama PowerPoint tidak ditemukan."

    End If

    If chartObj Is Nothing Then

        Err.Raise _
            vbObjectError + 3003, _
            , _
            "Chart Excel belum ditentukan."

    End If

    ' ========================================================
    ' SIMPAN GEOMETRY PICTURE LAMA
    ' ========================================================
    savedLeft = oldPicture.Left
    savedTop = oldPicture.Top
    savedWidth = oldPicture.Width
    savedHeight = oldPicture.Height
    savedRotation = oldPicture.Rotation
    savedLockAspect = oldPicture.LockAspectRatio

    ' Simpan metadata
    On Error Resume Next

    savedAlternativeText = _
        oldPicture.AlternativeText

    savedTitle = _
        oldPicture.Title

    On Error GoTo 0

    ' ========================================================
    ' SIMPAN FORMAT PICTURE LAMA
    ' PickUp/Apply mempertahankan line, shadow, effect, dll.
    ' ========================================================
    formatWasPickedUp = False

    On Error Resume Next
    Err.Clear

    oldPicture.PickUp

    If Err.Number = 0 Then
        formatWasPickedUp = True
    End If

    Err.Clear
    On Error GoTo 0

    ' ========================================================
    ' SIMPAN SETTING CROP
    ' ========================================================
    cropCanBePreserved = False

    On Error Resume Next
    Err.Clear

    savedCropLeft = _
        oldPicture.PictureFormat.CropLeft

    savedCropRight = _
        oldPicture.PictureFormat.CropRight

    savedCropTop = _
        oldPicture.PictureFormat.CropTop

    savedCropBottom = _
        oldPicture.PictureFormat.CropBottom

    If Err.Number = 0 Then
        cropCanBePreserved = True
    End If

    Err.Clear
    On Error GoTo 0

    ' ========================================================
    ' COPY CHART EXCEL SEBAGAI GAMBAR
    ' ========================================================
    chartObj.Chart.CopyPicture _
        Appearance:=xlScreen, _
        Format:=xlPicture

    DoEvents

    Set pastedShapes = Nothing

    ' Paste sebagai Enhanced Metafile agar lebih tajam
    On Error Resume Next
    Err.Clear

    Set pastedShapes = _
        pptSlide.Shapes.PasteSpecial( _
            PP_PASTE_ENHANCED_METAFILE _
        )

    ' Fallback paste biasa
    If pastedShapes Is Nothing Then

        Err.Clear

        Set pastedShapes = _
            pptSlide.Shapes.Paste

    End If

    On Error GoTo 0

    If pastedShapes Is Nothing Then

        Err.Raise _
            vbObjectError + 3004, _
            , _
            "Chart gagal dipaste sebagai picture."

    End If

    Set newPicture = pastedShapes.item(1)

    ' Hapus picture lama setelah picture baru berhasil dibuat
    oldPicture.Delete

    ' Nama picture baru dikembalikan ke nama mapping
    On Error Resume Next
    newPicture.Name = requestedShapeName
    On Error GoTo 0

    ' Terapkan formatting picture lama
    If formatWasPickedUp Then

        On Error Resume Next
        newPicture.Apply
        On Error GoTo 0

    End If

    ' ========================================================
    ' KEMBALIKAN POSISI, UKURAN, DAN ROTATION
    ' ========================================================
    On Error Resume Next

    newPicture.LockAspectRatio = MSO_FALSE

    newPicture.Left = savedLeft
    newPicture.Top = savedTop
    newPicture.Width = savedWidth
    newPicture.Height = savedHeight
    newPicture.Rotation = savedRotation

    On Error GoTo 0

    ' ========================================================
    ' KEMBALIKAN CROP
    ' ========================================================
    If cropCanBePreserved Then

        On Error Resume Next

        newPicture.PictureFormat.CropLeft = _
            savedCropLeft

        newPicture.PictureFormat.CropRight = _
            savedCropRight

        newPicture.PictureFormat.CropTop = _
            savedCropTop

        newPicture.PictureFormat.CropBottom = _
            savedCropBottom

        On Error GoTo 0

    End If

    ' Kembalikan Lock Aspect Ratio
    On Error Resume Next

    newPicture.LockAspectRatio = _
        savedLockAspect

    newPicture.AlternativeText = _
        savedAlternativeText

    newPicture.Title = _
        savedTitle

    On Error GoTo 0

End Sub


' ============================================================
' UPDATE ISI TABEL TANPA MENGHAPUS SHAPE TABEL
'
' Jumlah baris dan kolom Excel harus sama dengan PPT.
' Hal ini menjaga ukuran dan format tabel tetap persis.
' ============================================================
Private Sub UpdatePowerPointTablePreserveFormat( _
    ByVal pptShape As Object, _
    ByVal rngSource As Range)

    Dim pptTable As Object
    Dim cellShape As Object
    Dim textRange As Object

    Dim tableRow As Long
    Dim tableColumn As Long

    Dim pptRowCount As Long
    Dim pptColumnCount As Long
    Dim sourceRowCount As Long
    Dim sourceColumnCount As Long

    Dim cellText As String

    Dim fontName As String
    Dim fontSize As Single
    Dim fontBold As Long
    Dim fontItalic As Long
    Dim fontUnderline As Long
    Dim fontColor As Long

    Dim paragraphAlignment As Long
    Dim verticalAnchor As Long

    Dim marginLeft As Single
    Dim marginRight As Single
    Dim marginTop As Single
    Dim marginBottom As Single

    If pptShape Is Nothing Then

        Err.Raise _
            vbObjectError + 4001, _
            , _
            "Shape tabel belum ditentukan."

    End If

    If rngSource Is Nothing Then

        Err.Raise _
            vbObjectError + 4002, _
            , _
            "Range sumber tabel belum ditentukan."

    End If

    If Not ShapeHasTable(pptShape) Then

        Err.Raise _
            vbObjectError + 4003, _
            , _
            "Shape bukan tabel native PowerPoint."

    End If

    Set pptTable = pptShape.Table

    pptRowCount = pptTable.rows.Count
    pptColumnCount = pptTable.Columns.Count

    sourceRowCount = rngSource.rows.Count
    sourceColumnCount = rngSource.Columns.Count

    ' Untuk menjaga ukuran tabel tetap persis,
    ' jumlah baris dan kolom harus sama.
    If pptRowCount <> sourceRowCount Then

        Err.Raise _
            vbObjectError + 4004, _
            , _
            "Jumlah baris berbeda. PPT: " & _
            pptRowCount & _
            ", Excel: " & sourceRowCount

    End If

    If pptColumnCount <> sourceColumnCount Then

        Err.Raise _
            vbObjectError + 4005, _
            , _
            "Jumlah kolom berbeda. PPT: " & _
            pptColumnCount & _
            ", Excel: " & sourceColumnCount

    End If

    For tableRow = 1 To sourceRowCount

        For tableColumn = 1 To sourceColumnCount

            Set cellShape = _
                pptTable.cell( _
                    tableRow, _
                    tableColumn _
                ).Shape

            Set textRange = _
                cellShape.TextFrame.textRange

            ' =================================================
            ' SIMPAN FORMAT TEXT CELL
            ' =================================================
            On Error Resume Next

            fontName = textRange.Font.Name
            fontSize = textRange.Font.Size
            fontBold = textRange.Font.Bold
            fontItalic = textRange.Font.Italic
            fontUnderline = textRange.Font.Underline
            fontColor = textRange.Font.Color.RGB

            paragraphAlignment = _
                textRange.ParagraphFormat.Alignment

            verticalAnchor = _
                cellShape.TextFrame.verticalAnchor

            marginLeft = _
                cellShape.TextFrame.marginLeft

            marginRight = _
                cellShape.TextFrame.marginRight

            marginTop = _
                cellShape.TextFrame.marginTop

            marginBottom = _
                cellShape.TextFrame.marginBottom

            On Error GoTo 0

            ' Ambil tampilan cell Excel
            cellText = GetExcelDisplayText( _
                rngSource.Cells( _
                    tableRow, _
                    tableColumn _
                ) _
            )

            ' Hanya ganti isi cell
            textRange.text = cellText

            ' =================================================
            ' KEMBALIKAN FORMAT TEXT CELL
            ' =================================================
            On Error Resume Next

            With textRange.Font

                .Name = fontName
                .Size = fontSize
                .Bold = fontBold
                .Italic = fontItalic
                .Underline = fontUnderline
                .Color.RGB = fontColor

            End With

            textRange.ParagraphFormat.Alignment = _
                paragraphAlignment

            With cellShape.TextFrame

                .verticalAnchor = verticalAnchor

                .marginLeft = marginLeft
                .marginRight = marginRight
                .marginTop = marginTop
                .marginBottom = marginBottom

            End With

            On Error GoTo 0

        Next tableColumn

    Next tableRow

End Sub


' ============================================================
' SIMPAN GEOMETRY SHAPE
' ============================================================
Private Sub SaveShapeGeometry( _
    ByVal pptShape As Object, _
    ByRef oldLeft As Single, _
    ByRef oldTop As Single, _
    ByRef oldWidth As Single, _
    ByRef oldHeight As Single, _
    ByRef oldRotation As Single, _
    ByRef oldLockAspect As Long)

    If pptShape Is Nothing Then

        oldLeft = 50
        oldTop = 100
        oldWidth = 400
        oldHeight = 300
        oldRotation = 0
        oldLockAspect = MSO_FALSE

    Else

        oldLeft = pptShape.Left
        oldTop = pptShape.Top
        oldWidth = pptShape.Width
        oldHeight = pptShape.Height
        oldRotation = pptShape.Rotation
        oldLockAspect = pptShape.LockAspectRatio

    End If

End Sub


' ============================================================
' KEMBALIKAN GEOMETRY SHAPE
' ============================================================
Private Sub RestoreShapeGeometry( _
    ByVal pptShape As Object, _
    ByVal requestedShapeName As String, _
    ByVal oldLeft As Single, _
    ByVal oldTop As Single, _
    ByVal oldWidth As Single, _
    ByVal oldHeight As Single, _
    ByVal oldRotation As Single, _
    ByVal oldLockAspect As Long)

    If pptShape Is Nothing Then Exit Sub

    On Error Resume Next

    pptShape.Name = requestedShapeName

    pptShape.LockAspectRatio = MSO_FALSE

    pptShape.Left = oldLeft
    pptShape.Top = oldTop
    pptShape.Width = oldWidth
    pptShape.Height = oldHeight
    pptShape.Rotation = oldRotation

    pptShape.LockAspectRatio = oldLockAspect

    On Error GoTo 0

End Sub


' ============================================================
' AMBIL TEKS SESUAI TAMPILAN EXCEL
'
' Contoh:
' 0,5119 dengan format 0,00% menjadi 51,19%
' ============================================================
Private Function GetExcelDisplayText( _
    ByVal sourceCell As Range) As String

    Dim rawValue As Variant
    Dim displayedText As String
    Dim numberFormat As String
    Dim decimalPlaces As Long

    On Error GoTo ErrorHandler

    sourceCell.Calculate

    rawValue = sourceCell.Value
    displayedText = sourceCell.text

    If IsError(rawValue) Then

        GetExcelDisplayText = displayedText
        Exit Function

    End If

    If IsEmpty(rawValue) Then

        GetExcelDisplayText = vbNullString
        Exit Function

    End If

    ' Gunakan tampilan Excel jika bukan ####
    If Not IsOnlyHashes(displayedText) Then

        GetExcelDisplayText = displayedText
        Exit Function

    End If

    ' Fallback jika lebar kolom Excel terlalu sempit
    numberFormat = _
        CStr(sourceCell.NumberFormatLocal)

    If IsNumeric(rawValue) And _
       InStr(1, numberFormat, "%", vbTextCompare) > 0 Then

        decimalPlaces = _
            GetPercentageDecimalPlaces(numberFormat)

        If decimalPlaces > 0 Then

            GetExcelDisplayText = Format$( _
                CDbl(rawValue), _
                "0." & String$(decimalPlaces, "0") & "%" _
            )

        Else

            GetExcelDisplayText = Format$( _
                CDbl(rawValue), _
                "0%" _
            )

        End If

        Exit Function

    End If

    If IsDate(rawValue) Then

        GetExcelDisplayText = _
            Format$(CDate(rawValue), "dd/mm/yyyy")

        Exit Function

    End If

    GetExcelDisplayText = CStr(rawValue)
    Exit Function

ErrorHandler:

    GetExcelDisplayText = vbNullString

End Function


' ============================================================
' HITUNG DESIMAL FORMAT PERSENTASE
' ============================================================
Private Function GetPercentageDecimalPlaces( _
    ByVal numberFormat As String) As Long

    Dim percentPosition As Long
    Dim decimalPosition As Long
    Dim characterPosition As Long
    Dim currentCharacter As String
    Dim result As Long

    percentPosition = InStr( _
        1, _
        numberFormat, _
        "%", _
        vbTextCompare _
    )

    decimalPosition = InStr( _
        1, _
        numberFormat, _
        Application.International( _
            xlDecimalSeparator _
        ) _
    )

    If decimalPosition = 0 Then

        decimalPosition = InStr( _
            1, _
            numberFormat, _
            "," _
        )

    End If

    If decimalPosition = 0 Then

        decimalPosition = InStr( _
            1, _
            numberFormat, _
            "." _
        )

    End If

    If decimalPosition > 0 And _
       percentPosition > decimalPosition Then

        For characterPosition = _
            decimalPosition + 1 _
            To percentPosition - 1

            currentCharacter = Mid$( _
                numberFormat, _
                characterPosition, _
                1 _
            )

            If currentCharacter = "0" Or _
               currentCharacter = "#" Then

                result = result + 1

            End If

        Next characterPosition

    End If

    GetPercentageDecimalPlaces = result

End Function


' ============================================================
' CEK DISPLAY CELL HANYA BERISI ####
' ============================================================
Private Function IsOnlyHashes( _
    ByVal textValue As String) As Boolean

    Dim trimmedText As String

    trimmedText = Trim$(textValue)

    If Len(trimmedText) = 0 Then

        IsOnlyHashes = False

    Else

        IsOnlyHashes = _
            (Len(Replace$( _
                trimmedText, _
                "#", _
                vbNullString _
            )) = 0)

    End If

End Function


' ============================================================
' AMBIL PATH POWERPOINT DARI CELL
' Mendukung path biasa dan hyperlink
' ============================================================
Private Function GetPowerPointPath( _
    ByVal targetCell As Range) As String

    Dim result As String

    On Error Resume Next

    If targetCell.Hyperlinks.Count > 0 Then

        result = _
            targetCell.Hyperlinks(1).Address

    End If

    On Error GoTo 0

    If Len(Trim$(result)) = 0 Then

        result = CStr(targetCell.Value2)

    End If

    result = Trim$(result)

    If Len(result) >= 2 Then

        If Left$(result, 1) = """" And _
           Right$(result, 1) = """" Then

            result = Mid$( _
                result, _
                2, _
                Len(result) - 2 _
            )

        End If

    End If

    result = Replace( _
        result, _
        "file:///", _
        "", _
        1, _
        -1, _
        vbTextCompare _
    )

    result = Replace( _
        result, _
        "file://", _
        "", _
        1, _
        -1, _
        vbTextCompare _
    )

    result = Replace( _
        result, _
        "%20", _
        " ", _
        1, _
        -1, _
        vbTextCompare _
    )

    If InStr( _
        1, _
        result, _
        "http", _
        vbTextCompare _
    ) = 0 Then

        result = Replace( _
            result, _
            "/", _
            "\" _
        )

    End If

    GetPowerPointPath = result

End Function


' ============================================================
' CEK FILE ADA
' ============================================================
Private Function FileExists( _
    ByVal filePath As String) As Boolean

    On Error Resume Next

    FileExists = _
        (Len(Dir$( _
            filePath, _
            vbNormal Or vbReadOnly Or vbHidden _
        )) > 0)

    On Error GoTo 0

End Function


' ============================================================
' AMBIL WORKSHEET DENGAN AMAN
' ============================================================
Private Function GetWorksheetSafe( _
    ByVal wb As Workbook, _
    ByVal worksheetName As String) As Worksheet

    On Error Resume Next

    Set GetWorksheetSafe = _
        wb.Worksheets(worksheetName)

    On Error GoTo 0

End Function


' ============================================================
' AMBIL RANGE SUMBER
' ============================================================
Private Function GetSourceRange( _
    ByVal wb As Workbook, _
    ByVal sourceSheet As String, _
    ByVal sourceAddress As String) As Range

    Dim ws As Worksheet

    If Len(Trim$(sourceSheet)) = 0 Then Exit Function
    If Len(Trim$(sourceAddress)) = 0 Then Exit Function

    Set ws = GetWorksheetSafe( _
        wb, _
        sourceSheet _
    )

    If ws Is Nothing Then Exit Function

    On Error Resume Next

    Set GetSourceRange = _
        ws.Range(sourceAddress)

    On Error GoTo 0

End Function


' ============================================================
' CARI SHAPE POWERPOINT BERDASARKAN NAMA
' ============================================================
Public Function FindShape( _
    ByVal sld As Object, _
    ByVal shapeName As String) As Object

    Dim shp As Object

    If sld Is Nothing Then Exit Function

    For Each shp In sld.Shapes

        If StrComp( _
            Trim$(CStr(shp.Name)), _
            Trim$(shapeName), _
            vbTextCompare _
        ) = 0 Then

            Set FindShape = shp
            Exit Function

        End If

    Next shp

End Function


' ============================================================
' CEK SHAPE ADALAH TABEL
' ============================================================
Private Function ShapeHasTable( _
    ByVal pptShape As Object) As Boolean

    If pptShape Is Nothing Then Exit Function

    On Error Resume Next

    ShapeHasTable = _
        CBool(pptShape.HasTable)

    On Error GoTo 0

End Function


' ============================================================
' CARI CHART EXCEL
'
' Prioritas:
' 1. Nama dari kolom C
' 2. Chart 6
' 3. Chart 3
' 4. Chart 1
' 5. Chart pertama
' ============================================================
Private Function FindExcelChart( _
    ByVal ws As Worksheet, _
    ByVal preferredName As String) As ChartObject

    Dim fallbackNames As Variant
    Dim fallbackName As Variant

    If ws Is Nothing Then Exit Function

    On Error Resume Next

    Set FindExcelChart = _
        ws.ChartObjects(preferredName)

    On Error GoTo 0

    If Not FindExcelChart Is Nothing Then Exit Function

    fallbackNames = Array( _
        "Chart 6", _
        "Chart 3", _
        "Chart 1" _
    )

    For Each fallbackName In fallbackNames

        On Error Resume Next

        Set FindExcelChart = _
            ws.ChartObjects(CStr(fallbackName))

        On Error GoTo 0

        If Not FindExcelChart Is Nothing Then
            Exit Function
        End If

    Next fallbackName

    If ws.ChartObjects.Count > 0 Then

        Set FindExcelChart = _
            ws.ChartObjects(1)

    End If

End Function


' ============================================================
' CARI PRESENTASI YANG SUDAH TERBUKA
' ============================================================
Private Function GetOpenPresentationByPath( _
    ByVal pptApp As Object, _
    ByVal pptPath As String) As Object

    Dim pptPres As Object

    If pptApp Is Nothing Then Exit Function

    For Each pptPres In pptApp.Presentations

        If StrComp( _
            CStr(pptPres.FullName), _
            pptPath, _
            vbTextCompare _
        ) = 0 Then

            Set GetOpenPresentationByPath = pptPres
            Exit Function

        End If

    Next pptPres

End Function


' ============================================================
' KONVERSI VALUE MENJADI TEXT AMAN
' ============================================================
Private Function SafeText( _
    ByVal valueToConvert As Variant) As String

    If IsError(valueToConvert) Then

        SafeText = vbNullString

    ElseIf IsNull(valueToConvert) Or _
           IsEmpty(valueToConvert) Then

        SafeText = vbNullString

    Else

        SafeText = Trim$(CStr(valueToConvert))

    End If

End Function


' ============================================================
' FUNGSI KOMPATIBILITAS
' Tidak mem-parsing formula dengan Mid/InStr
' ============================================================
Public Function EvaluateFormula( _
    ByVal cell As Range) As Variant

    If cell Is Nothing Then

        EvaluateFormula = Empty
        Exit Function

    End If

    cell.Calculate
    EvaluateFormula = cell.Value

End Function


' ============================================================
' KONVERSI NOMOR KOLOM MENJADI HURUF
' ============================================================
Public Function ConvertToLetter( _
    ByVal colNum As Long) As String

    Dim dividend As Long
    Dim remainder As Long
    Dim result As String

    dividend = colNum
    result = vbNullString

    Do While dividend > 0

        remainder = _
            (dividend - 1) Mod 26

        result = _
            Chr$(65 + remainder) & result

        dividend = _
            (dividend - remainder - 1) \ 26

    Loop

    ConvertToLetter = result

End Function


' ============================================================
' DEBUG SHAPE PADA SATU SLIDE
'
' Jalankan dari Immediate Window:
' DebugShapesOnSlide 24
' ============================================================
Public Sub DebugShapesOnSlide( _
    ByVal slideNum As Long)

    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim shp As Object

    Dim pptPath As String

    pptPath = GetPowerPointPath( _
        ThisWorkbook.Worksheets( _
            "OUTPUT_PPT" _
        ).Range("B2") _
    )

    If Len(pptPath) = 0 Then

        MsgBox _
            "OUTPUT_PPT!B2 masih kosong.", _
            vbExclamation

        Exit Sub

    End If

    On Error Resume Next

    Set pptApp = _
        GetObject(, "PowerPoint.Application")

    On Error GoTo 0

    If pptApp Is Nothing Then

        Set pptApp = _
            CreateObject("PowerPoint.Application")

    End If

    pptApp.Visible = True

    Set pptPres = GetOpenPresentationByPath( _
        pptApp, _
        pptPath _
    )

    If pptPres Is Nothing Then

        Set pptPres = _
            pptApp.Presentations.Open(pptPath)

    End If

    On Error Resume Next

    Set pptSlide = _
        pptPres.Slides(slideNum)

    On Error GoTo 0

    If pptSlide Is Nothing Then

        MsgBox _
            "Slide " & slideNum & _
            " tidak ditemukan.", _
            vbExclamation

        Exit Sub

    End If

    Debug.Print _
        "===== SHAPE DI SLIDE " & _
        slideNum & " ====="

    For Each shp In pptSlide.Shapes

        Debug.Print _
            "Nama: " & shp.Name & _
            " | Type: " & shp.Type

        On Error Resume Next

        If shp.HasTextFrame Then

            If shp.TextFrame.hasText Then

                Debug.Print _
                    "  Text: " & _
                    shp.TextFrame.textRange.text

            End If

        End If

        If shp.HasTable Then

            Debug.Print _
                "  Table: " & _
                shp.Table.rows.Count & _
                " baris x " & _
                shp.Table.Columns.Count & _
                " kolom"

        End If

        On Error GoTo 0

    Next shp

    MsgBox _
        "Hasil debug tersedia di Immediate Window (Ctrl+G).", _
        vbInformation

End Sub


' ============================================================
' DEBUG BEBERAPA SLIDE
' ============================================================
Public Sub DebugAllSlides()

    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim shp As Object

    Dim pptPath As String
    Dim slideNums As Variant
    Dim slideItem As Variant

    pptPath = GetPowerPointPath( _
        ThisWorkbook.Worksheets( _
            "OUTPUT_PPT" _
        ).Range("B2") _
    )

    If Len(pptPath) = 0 Then

        MsgBox _
            "OUTPUT_PPT!B2 masih kosong.", _
            vbExclamation

        Exit Sub

    End If

    On Error Resume Next

    Set pptApp = _
        GetObject(, "PowerPoint.Application")

    On Error GoTo 0

    If pptApp Is Nothing Then

        Set pptApp = _
            CreateObject("PowerPoint.Application")

    End If

    pptApp.Visible = True

    Set pptPres = GetOpenPresentationByPath( _
        pptApp, _
        pptPath _
    )

    If pptPres Is Nothing Then

        Set pptPres = _
            pptApp.Presentations.Open(pptPath)

    End If

    slideNums = Array( _
        24, _
        25, _
        26, _
        32, _
        35, _
        42, _
        43, _
        48 _
    )

    For Each slideItem In slideNums

        Set pptSlide = Nothing

        On Error Resume Next

        Set pptSlide = _
            pptPres.Slides(CLng(slideItem))

        On Error GoTo 0

        If pptSlide Is Nothing Then

            Debug.Print _
                "Slide " & slideItem & _
                " TIDAK DITEMUKAN."

        Else

            Debug.Print _
                "===== SHAPE DI SLIDE " & _
                slideItem & " ====="

            For Each shp In pptSlide.Shapes

                Debug.Print _
                    "Nama: " & shp.Name & _
                    " | Type: " & shp.Type

            Next shp

        End If

    Next slideItem

    MsgBox _
        "Hasil debug tersedia di Immediate Window (Ctrl+G).", _
        vbInformation

End Sub

