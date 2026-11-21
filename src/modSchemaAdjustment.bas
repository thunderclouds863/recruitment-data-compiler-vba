Attribute VB_Name = "modSchemaAdjustment"
Option Explicit
Private Const SCHEMA_SHEET As String = "Schema_Config"
Private Const SCHEMA_LOG_SHEET As String = "Schema_Change_Log"
Private Const SCHEMA_HEADER_ROW As Long = 1
Private Const SCHEMA_FIRST_ROW As Long = 2

' WAJIB sama dengan layout sheet FPTK:
' row 1 = header
' row 2 = data pertama
Private Const FPTK_SCHEMA_HEADER_ROW As Long = 1
Private Const FPTK_SCHEMA_FIRST_DATA_ROW As Long = 2

' Existing layout dipertahankan agar form/module lama tetap kompatibel.
Private Const C_DATABASE As Long = 1
Private Const C_FIELD_ID As Long = 2
Private Const C_HEADER_NAME As Long = 3
Private Const C_POSITION As Long = 4
Private Const C_ACTIVE As Long = 5
Private Const C_DATA_TYPE As Long = 6
Private Const C_SHOW_FORM As Long = 7
Private Const C_REQUIRED As Long = 8
Private Const C_READ_ONLY As Long = 9
Private Const C_SEARCHABLE As Long = 10
Private Const C_CONTROL_TYPE As Long = 11
Private Const C_DROPDOWN As Long = 12
Private Const C_RULE As Long = 13
Private Const C_ALIASES As Long = 14
Private Const C_KEY_FIELD As Long = 15
Private Const C_CUSTOM As Long = 16
Private Const C_UPDATED_AT As Long = 17
Private Const C_CANONICAL_HEADER As Long = 18

Private mSchemaReady As Boolean
Private mSchemaCanonicalCache As Object

' ============================================================
' INSTALL / OPEN
' ============================================================
Public Sub InstallSchemaAdjustmentSystem()

    EnsureSchemaConfig
    SchemaRefreshRegistryFromSheets
    InstallHeaderAdjustmentButton

    MsgBox _
        "Safe Schema Overlay sudah aktif." & vbCrLf & _
        "Button UPDATE ADJUSTMENT tersedia di Settings.", _
        vbInformation, _
        "Schema Adjustment"

End Sub

Public Sub InstallHeaderAdjustmentButton()

    Dim ws As Worksheet
    Dim buttonShape As Shape

    Set ws = SchemaGetOrCreateSheet("Settings")

    On Error Resume Next
    ws.Shapes("btnHeaderAdjustment").Delete
    On Error GoTo 0

    Set buttonShape = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, 20, 20, 190, 36)

    With buttonShape
        .Name = "btnHeaderAdjustment"
        .TextFrame.Characters.text = "UPDATE ADJUSTMENT"
        .OnAction = "OpenHeaderAdjustment"
        .Fill.ForeColor.RGB = RGB(52, 152, 219)
        .line.ForeColor.RGB = RGB(41, 128, 185)
        .TextFrame.Characters.Font.Color = vbWhite
        .TextFrame.Characters.Font.Bold = True
    End With

End Sub

Public Sub OpenHeaderAdjustment()

    EnsureSchemaConfig
    SchemaRefreshRegistryFromSheets
    frmHeaderAdjustment.Show

End Sub

Public Sub InitializeSchemaDrivenHeaders()

    EnsureSchemaConfig
    SchemaEnsureAll

    MsgBox _
        "Schema registry dan header aktif sudah disiapkan.", _
        vbInformation, _
        "Schema Header"

End Sub

Public Function SchemaModeEnabled() As Boolean
    SchemaModeEnabled = True
End Function

' ============================================================
' CONFIG INITIALIZATION / MIGRATION
' ============================================================
Public Sub EnsureSchemaConfig()

    Dim ws As Worksheet

    If mSchemaReady Then Exit Sub

    On Error GoTo ErrorHandler

    mSchemaReady = True

    Set ws = SchemaGetOrCreateSheet(SCHEMA_SHEET)

    SchemaWriteConfigHeaders ws
    SchemaSeedFPTK ws
    SchemaSeedDBKodePosisi ws
    SchemaSeedDBSourcing ws
    SchemaMigrateCanonicalHeaders ws

    SchemaCleanupKnownFPTKLegacyFields _
        ws

    ' Satukan registry FMCG yang sebelumnya terdaftar
    ' sebagai custom field terpisah.
    SchemaMergeKnownDBSourcingFMCGAliases _
        ws

    SchemaRegisterCurrentHeadersInternal ws, "FPTK"
    SchemaRegisterCurrentHeadersInternal ws, "DB Kode Posisi"
    SchemaRegisterCurrentHeadersInternal ws, "DB Sourcing"

    SchemaFormatConfig ws
    Exit Sub

ErrorHandler:
    mSchemaReady = False
    Err.Raise Err.Number, "EnsureSchemaConfig", Err.Description

End Sub

Public Sub SchemaResetCache()

    mSchemaReady = False
    Set mSchemaCanonicalCache = Nothing

End Sub

Public Sub SchemaRefreshRegistryFromSheets()

    Dim wsConfig As Worksheet

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)

    SchemaRegisterCurrentHeadersInternal wsConfig, "FPTK"
    SchemaRegisterCurrentHeadersInternal wsConfig, "DB Kode Posisi"
    SchemaRegisterCurrentHeadersInternal wsConfig, "DB Sourcing"

    SchemaApplyVisibilityAll
    SchemaFormatConfig wsConfig

End Sub

Public Sub SchemaRegisterCurrentHeadersSafe( _
    ByVal databaseKey As String)

    Dim wsConfig As Worksheet

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)

    SchemaRegisterCurrentHeadersInternal _
        wsConfig, _
        databaseKey

    ' Tidak AutoFit Schema_Config setiap compile.
    ' AutoFit hanya diperlukan saat membuka/refresh adjustment.

End Sub

' ============================================================
' DATABASE LAYOUT
' ============================================================
Public Function SchemaHeaderRow( _
    ByVal databaseKey As String) As Long

    If StrComp( _
        SchemaNormalizeDatabaseKey(databaseKey), _
        "FPTK", _
        vbTextCompare) = 0 Then

        SchemaHeaderRow = _
            FPTK_SCHEMA_HEADER_ROW

    Else

        SchemaHeaderRow = 1

    End If

End Function

Public Function SchemaFirstDataRow( _
    ByVal databaseKey As String) As Long

    If StrComp( _
        SchemaNormalizeDatabaseKey(databaseKey), _
        "FPTK", _
        vbTextCompare) = 0 Then

        SchemaFirstDataRow = _
            FPTK_SCHEMA_FIRST_DATA_ROW

    Else

        SchemaFirstDataRow = 2

    End If

End Function

' ============================================================
' SAFE ENSURE
' Tidak menghapus data, tidak rebuild, tidak reorder otomatis.
' ============================================================
Public Sub SchemaB02PrepareAll()

    ' Compatibility wrapper untuk module utama lama.
    ' SchemaB02PrepareAll sebelumnya dipanggil oleh
    ' CompileFPTK_Core, sedangkan schema baru memakai
    ' SchemaEnsureAll sebagai prosedur utamanya.
    SchemaEnsureAll

End Sub

Public Sub SchemaEnsureAll()

    EnsureSchemaConfig
    SchemaEnsureSheet "FPTK"
    SchemaEnsureSheet "DB Kode Posisi"
    SchemaEnsureSheet "DB Sourcing"

End Sub

Public Sub SchemaEnsureSheet( _
    ByVal databaseKey As String)

    Dim wsConfig As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim fieldID As String
    Dim currentHeader As String
    Dim canonicalHeader As String
    Dim aliases As String
    Dim columnNumber As Long
    Dim isActive As Boolean
    Dim headerRow As Long

    databaseKey = SchemaNormalizeDatabaseKey(databaseKey)
    headerRow = SchemaHeaderRow(databaseKey)

    EnsureSchemaConfig

    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)
    Set wsData = SchemaGetRequiredDataSheet(databaseKey)

    SchemaUnlistAllTables wsData
    SchemaRegisterCurrentHeadersInternal wsConfig, databaseKey

    lastRow = SchemaLastUsedRow(wsConfig)

    For rowIndex = SCHEMA_FIRST_ROW To lastRow

        If StrComp( _
            Trim$(CStr(wsConfig.Cells(rowIndex, C_DATABASE).Value2)), _
            databaseKey, vbTextCompare) = 0 Then

            fieldID = Trim$(CStr( _
                wsConfig.Cells(rowIndex, C_FIELD_ID).Value2))

            currentHeader = Trim$(CStr( _
                wsConfig.Cells(rowIndex, C_HEADER_NAME).Value2))

            canonicalHeader = Trim$(CStr( _
                wsConfig.Cells(rowIndex, C_CANONICAL_HEADER).Value2))

            aliases = CStr( _
                wsConfig.Cells(rowIndex, C_ALIASES).Value2)

            isActive = SchemaYes( _
                wsConfig.Cells(rowIndex, C_ACTIVE).Value2)

            columnNumber = SchemaFindPhysicalColumn( _
                wsData, currentHeader, canonicalHeader, aliases)

            If isActive Then

                If columnNumber = 0 Then
                    columnNumber = SchemaLastHeaderColumn(wsData) + 1
                    If SchemaLastHeaderColumn(wsData) = 1 And _
                       Len(Trim$(CStr(wsData.Cells(headerRow, 1).Value2))) = 0 Then
                        columnNumber = 1
                    End If
                    wsData.Cells(headerRow, columnNumber).Value = currentHeader
                ElseIf Len(currentHeader) > 0 Then
                    wsData.Cells(headerRow, columnNumber).Value = currentHeader
                End If

            End If

            If columnNumber > 0 Then
                SchemaSetColumnHiddenSafe _
                    wsData, columnNumber, Not isActive
            End If

        End If

    Next rowIndex

    SchemaStyleHeader wsData

End Sub

' ============================================================
' HEADER RESOLUTION
' ============================================================
Public Function SchemaFindColumnAuto( _
    ByVal ws As Worksheet, _
    ByVal fieldReference As String, _
    Optional ByVal createIfMissing As Boolean = False) As Long

    If ws Is Nothing Then Exit Function

    SchemaFindColumnAuto = SchemaFindColumn( _
        ws, ws.Name, fieldReference, createIfMissing)

End Function

Public Function SchemaFindColumn( _
    ByVal ws As Worksheet, _
    ByVal databaseKey As String, _
    ByVal fieldReference As String, _
    Optional ByVal createIfMissing As Boolean = False) As Long

    Dim wsConfig As Worksheet
    Dim configRow As Long
    Dim currentHeader As String
    Dim canonicalHeader As String
    Dim aliases As String
    Dim columnNumber As Long
    Dim headerRow As Long

    If ws Is Nothing Then Exit Function

    databaseKey = SchemaNormalizeDatabaseKey(databaseKey)
    headerRow = SchemaHeaderRow(databaseKey)

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)

    configRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, fieldReference)

    If configRow > 0 Then

        currentHeader = CStr( _
            wsConfig.Cells(configRow, C_HEADER_NAME).Value2)

        canonicalHeader = CStr( _
            wsConfig.Cells(configRow, C_CANONICAL_HEADER).Value2)

        aliases = CStr( _
            wsConfig.Cells(configRow, C_ALIASES).Value2)

        columnNumber = SchemaFindPhysicalColumn( _
            ws, currentHeader, canonicalHeader, aliases)

        If columnNumber = 0 And createIfMissing Then
            columnNumber = SchemaLastHeaderColumn(ws) + 1
            If SchemaLastHeaderColumn(ws) = 1 And _
               Len(Trim$(CStr(ws.Cells(headerRow, 1).Value2))) = 0 Then
                columnNumber = 1
            End If
            ws.Cells(headerRow, columnNumber).Value = currentHeader
        End If

    Else

        columnNumber = SchemaFindPhysicalColumn( _
            ws, fieldReference, fieldReference, vbNullString)

        If columnNumber = 0 And createIfMissing Then

            SchemaResolveOrRegisterFieldID _
                databaseKey, fieldReference

            columnNumber = SchemaLastHeaderColumn(ws) + 1
            If SchemaLastHeaderColumn(ws) = 1 And _
               Len(Trim$(CStr(ws.Cells(headerRow, 1).Value2))) = 0 Then
                columnNumber = 1
            End If
            ws.Cells(headerRow, columnNumber).Value = fieldReference

        End If

    End If

    SchemaFindColumn = columnNumber

End Function

Public Function SchemaResolveFieldID( _
    ByVal databaseKey As String, _
    ByVal fieldReference As String) As String

    Dim wsConfig As Worksheet
    Dim configRow As Long

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)

    configRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, fieldReference)

    If configRow > 0 Then
        SchemaResolveFieldID = Trim$(CStr( _
            wsConfig.Cells(configRow, C_FIELD_ID).Value2))
    End If

End Function

Public Function SchemaResolveOrRegisterFieldID( _
    ByVal databaseKey As String, _
    ByVal headerName As String, _
    Optional ByVal positionNumber As Long = 0) As String

    Dim wsConfig As Worksheet
    Dim configRow As Long

    If Len(Trim$(headerName)) = 0 Then Exit Function

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)

    configRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, headerName)

    If configRow = 0 Then

        If positionNumber <= 0 Then
            positionNumber = SchemaMaxPosition( _
                wsConfig, databaseKey) + 1
        End If

        configRow = SchemaAddFieldInternal( _
            wsConfig, databaseKey, headerName, headerName, _
            positionNumber, SchemaInferDataType(headerName), _
            True, False, False, True, _
            SchemaInferControlType(headerName), vbNullString, _
            vbNullString, False, True)

    End If

    SchemaResolveOrRegisterFieldID = Trim$(CStr( _
        wsConfig.Cells(configRow, C_FIELD_ID).Value2))

End Function

Public Function SchemaCanonicalHeaderKey( _
    ByVal databaseKey As String, _
    ByVal rawHeader As Variant) As String

    Dim wsConfig As Worksheet
    Dim configRow As Long
    Dim canonicalHeader As String

    Dim normalizedDatabase As String
    Dim normalizedHeader As String
    Dim cacheKey As String

    If IsError(rawHeader) Or _
       IsNull(rawHeader) Then

        Exit Function

    End If

    normalizedDatabase = _
        SchemaNormalizeDatabaseKey( _
            databaseKey)

    normalizedHeader = _
        SchemaNormalize( _
            CStr(rawHeader))

    If Len(normalizedDatabase) = 0 Or _
       Len(normalizedHeader) = 0 Then

        Exit Function

    End If

    If mSchemaCanonicalCache Is Nothing Then

        Set mSchemaCanonicalCache = _
            CreateObject( _
                "Scripting.Dictionary")

        mSchemaCanonicalCache.CompareMode = _
            vbTextCompare

    End If

    cacheKey = _
        normalizedDatabase & _
        "||" & _
        normalizedHeader

    If mSchemaCanonicalCache.Exists( _
        cacheKey) Then

        SchemaCanonicalHeaderKey = _
            CStr( _
                mSchemaCanonicalCache( _
                    cacheKey))

        Exit Function

    End If

    EnsureSchemaConfig

    Set wsConfig = _
        ThisWorkbook.Worksheets( _
            SCHEMA_SHEET)

    configRow = _
        SchemaFindConfigRow( _
            wsConfig, _
            normalizedDatabase, _
            CStr(rawHeader))

    If configRow > 0 Then

        canonicalHeader = CStr( _
            wsConfig.Cells( _
                configRow, _
                C_CANONICAL_HEADER).Value2)

        If Len(Trim$( _
            canonicalHeader)) = 0 Then

            canonicalHeader = CStr( _
                wsConfig.Cells( _
                    configRow, _
                    C_HEADER_NAME).Value2)

        End If

    Else

        ' Walaupun registry belum sempat diperbarui,
        ' alias yang dapat dikenali tetap diarahkan ke
        ' canonical field dan tidak dianggap header baru.
        canonicalHeader = _
            SchemaGuessCanonicalHeader( _
                normalizedDatabase, _
                CStr(rawHeader))

    End If

    If Len(Trim$( _
        canonicalHeader)) > 0 Then

        canonicalHeader = _
            SchemaNormalizeWithSpaces( _
                canonicalHeader)

    End If

    mSchemaCanonicalCache.Add _
        cacheKey, _
        canonicalHeader

    SchemaCanonicalHeaderKey = _
        canonicalHeader

End Function

Public Function SchemaCurrentHeader( _
    ByVal databaseKey As String, _
    ByVal fieldReference As String) As String

    Dim wsConfig As Worksheet
    Dim configRow As Long

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)

    configRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, fieldReference)

    If configRow > 0 Then
        SchemaCurrentHeader = CStr( _
            wsConfig.Cells(configRow, C_HEADER_NAME).Value2)
    Else
        SchemaCurrentHeader = fieldReference
    End If

End Function

Public Function SchemaFieldIsActive( _
    ByVal databaseKey As String, _
    ByVal fieldReference As String) As Boolean

    Dim wsConfig As Worksheet
    Dim configRow As Long

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)

    configRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, fieldReference)

    If configRow > 0 Then
        SchemaFieldIsActive = SchemaYes( _
            wsConfig.Cells(configRow, C_ACTIVE).Value2)
    End If

End Function

Public Function SchemaFieldIsCustom( _
    ByVal databaseKey As String, _
    ByVal fieldReference As String) As Boolean

    Dim wsConfig As Worksheet
    Dim configRow As Long

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)

    configRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, fieldReference)

    If configRow > 0 Then
        SchemaFieldIsCustom = SchemaYes( _
            wsConfig.Cells(configRow, C_CUSTOM).Value2)
    End If

End Function


' ============================================================
' FIND EXISTING PHYSICAL COLUMN BY FIELD / ALIAS / CANONICAL
'
' Compiler dapat memanggil fungsi ini sebelum append header.
' Bila hasil > 0, source header sudah memiliki kolom tujuan dan
' tidak boleh dibuat sebagai header baru.
' ============================================================
Public Function SchemaFindExistingColumnByAlias( _
    ByVal ws As Worksheet, _
    ByVal databaseKey As String, _
    ByVal sourceHeader As String) As Long

    Dim wsConfig As Worksheet
    Dim configRow As Long

    Dim currentHeader As String
    Dim canonicalHeader As String
    Dim aliases As String

    Dim headerRow As Long
    Dim lastColumn As Long
    Dim columnIndex As Long

    Dim sourceSemanticKey As String
    Dim targetSemanticKey As String

    If ws Is Nothing Then
        Exit Function
    End If

    databaseKey = _
        SchemaNormalizeDatabaseKey( _
            databaseKey)

    If Len(databaseKey) = 0 Then
        Exit Function
    End If

    EnsureSchemaConfig

    Set wsConfig = _
        ThisWorkbook.Worksheets( _
            SCHEMA_SHEET)

    configRow = _
        SchemaFindConfigRow( _
            wsConfig, _
            databaseKey, _
            sourceHeader)

    If configRow > 0 Then

        currentHeader = CStr( _
            wsConfig.Cells( _
                configRow, _
                C_HEADER_NAME).Value2)

        canonicalHeader = CStr( _
            wsConfig.Cells( _
                configRow, _
                C_CANONICAL_HEADER).Value2)

        aliases = CStr( _
            wsConfig.Cells( _
                configRow, _
                C_ALIASES).Value2)

        SchemaFindExistingColumnByAlias = _
            SchemaFindPhysicalColumn( _
                ws, _
                currentHeader, _
                canonicalHeader, _
                aliases)

        If SchemaFindExistingColumnByAlias > 0 Then
            Exit Function
        End If

    End If

    ' Fallback hanya untuk header dengan catatan instruksi,
    ' bukan fuzzy prefix bebas.
    sourceSemanticKey = _
        SchemaSemanticHeaderKey( _
            sourceHeader)

    If Len(sourceSemanticKey) = 0 Then
        Exit Function
    End If

    headerRow = _
        SchemaHeaderRow( _
            databaseKey)

    lastColumn = _
        SchemaLastHeaderColumn( _
            ws)

    For columnIndex = 1 _
        To lastColumn

        targetSemanticKey = _
            SchemaSemanticHeaderKey( _
                CStr(ws.Cells( _
                    headerRow, _
                    columnIndex).Value2))

        If Len(targetSemanticKey) > 0 Then

            If targetSemanticKey = _
               sourceSemanticKey Then

                SchemaFindExistingColumnByAlias = _
                    columnIndex

                Exit Function

            End If

        End If

    Next columnIndex

End Function

' ============================================================
' ACTIVE HEADER ARRAYS
' ============================================================
Public Function SchemaGetActiveHeadersArray( _
    ByVal databaseKey As String, _
    Optional ByVal fallbackHeaders As Variant) As Variant

    Dim wsConfig As Worksheet
    Dim sortedRows As Collection
    Dim result() As Variant
    Dim item As Variant
    Dim resultIndex As Long

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)

    Set sortedRows = SchemaRowsSorted( _
        wsConfig, databaseKey, False)

    If sortedRows.Count = 0 Then
        If IsArray(fallbackHeaders) Then
            SchemaGetActiveHeadersArray = fallbackHeaders
        Else
            SchemaGetActiveHeadersArray = Array()
        End If
        Exit Function
    End If

    ReDim result(0 To sortedRows.Count - 1)

    For Each item In sortedRows
        result(resultIndex) = CStr( _
            wsConfig.Cells(CLng(item), C_HEADER_NAME).Value2)
        resultIndex = resultIndex + 1
    Next item

    SchemaGetActiveHeadersArray = result

End Function

Public Function SchemaGetDatabaseKeys() As Variant
    SchemaGetDatabaseKeys = Array( _
        "FPTK", "DB Kode Posisi", "DB Sourcing")
End Function

Public Function SchemaGetFieldList( _
    ByVal databaseKey As String, _
    Optional ByVal includeInactive As Boolean = True) As Collection

    Dim wsConfig As Worksheet
    Dim sortedRows As Collection
    Dim result As New Collection
    Dim item As Variant
    Dim rowIndex As Long
    Dim activeText As String
    Dim displayItem As String

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)

    Set sortedRows = SchemaRowsSorted( _
        wsConfig, databaseKey, True)

    For Each item In sortedRows

        rowIndex = CLng(item)
        activeText = UCase$(Trim$(CStr( _
            wsConfig.Cells(rowIndex, C_ACTIVE).Value2)))

        If includeInactive Or activeText = "YES" Then

            displayItem = _
                CStr(wsConfig.Cells(rowIndex, C_FIELD_ID).Value2) & "|" & _
                CStr(wsConfig.Cells(rowIndex, C_HEADER_NAME).Value2) & "|" & _
                activeText & "|" & _
                CStr(wsConfig.Cells(rowIndex, C_POSITION).Value2) & "|" & _
                CStr(wsConfig.Cells(rowIndex, C_CUSTOM).Value2) & "|" & _
                CStr(wsConfig.Cells(rowIndex, C_SHOW_FORM).Value2) & "|" & _
                CStr(wsConfig.Cells(rowIndex, C_READ_ONLY).Value2) & "|" & _
                CStr(wsConfig.Cells(rowIndex, C_CONTROL_TYPE).Value2) & "|" & _
                Replace$(CStr(wsConfig.Cells(rowIndex, C_DROPDOWN).Value2), _
                    "|", "^")

            result.Add displayItem

        End If

    Next item

    Set SchemaGetFieldList = result

End Function

' ============================================================
' VALUE ACCESS
' ============================================================
Public Function SchemaGetValue( _
    ByVal databaseKey As String, _
    ByVal fieldReference As String, _
    ByVal rowNumber As Long) As Variant

    Dim ws As Worksheet
    Dim columnNumber As Long

    Set ws = SchemaGetRequiredDataSheet(databaseKey)
    columnNumber = SchemaFindColumn( _
        ws, databaseKey, fieldReference, False)

    If columnNumber > 0 And rowNumber > 0 Then
        SchemaGetValue = ws.Cells(rowNumber, columnNumber).Value
    End If

End Function

Public Sub SchemaSetValue( _
    ByVal databaseKey As String, _
    ByVal fieldReference As String, _
    ByVal rowNumber As Long, _
    ByVal newValue As Variant)

    Dim ws As Worksheet
    Dim columnNumber As Long

    Set ws = SchemaGetRequiredDataSheet(databaseKey)
    columnNumber = SchemaFindColumn( _
        ws, databaseKey, fieldReference, True)

    If columnNumber > 0 And rowNumber > 0 Then
        ws.Cells(rowNumber, columnNumber).Value = newValue
    End If

End Sub

' ============================================================
' ADJUSTMENT REQUEST
' ============================================================
Public Sub SchemaApplyRequest( _
    ByVal databaseKey As String, _
    ByVal actionName As String, _
    ByVal fieldReference As String, _
    ByVal newHeader As String, _
    ByVal requestedPosition As Long, _
    ByVal dataType As String, _
    ByVal showInForm As Boolean, _
    ByVal isRequired As Boolean, _
    ByVal isReadOnly As Boolean, _
    ByVal isSearchable As Boolean, _
    ByVal controlType As String, _
    ByVal dropdownValues As String)

    databaseKey = SchemaNormalizeDatabaseKey(databaseKey)

    If Len(databaseKey) = 0 Then
        Err.Raise vbObjectError + 8100, _
            "SchemaApplyRequest", _
            "Database tidak dikenali. Pilih FPTK, DB Kode Posisi, atau DB Sourcing."
    End If

    Select Case UCase$(Trim$(actionName))

        Case "TAMBAH", "ADD"
            SchemaAddField databaseKey, newHeader, requestedPosition, _
                dataType, showInForm, isRequired, isReadOnly, _
                isSearchable, controlType, dropdownValues

        Case "UBAH NAMA", "RENAME"
            SchemaRenameField databaseKey, fieldReference, newHeader

        Case "HAPUS", "NONAKTIFKAN", "DELETE", "DEACTIVATE"
            SchemaDeactivateField databaseKey, fieldReference

        Case "AKTIFKAN", "ACTIVATE"
            SchemaActivateField databaseKey, fieldReference

        Case "PINDAH URUTAN", "MOVE"
            SchemaMoveField databaseKey, fieldReference, requestedPosition

        Case Else
            Err.Raise vbObjectError + 8101, _
                "SchemaApplyRequest", _
                "Tindakan adjustment tidak dikenali."

    End Select

End Sub

Public Sub SchemaAddField( _
    ByVal databaseKey As String, _
    ByVal headerName As String, _
    ByVal requestedPosition As Long, _
    Optional ByVal dataType As String = "Text", _
    Optional ByVal showInForm As Boolean = True, _
    Optional ByVal isRequired As Boolean = False, _
    Optional ByVal isReadOnly As Boolean = False, _
    Optional ByVal isSearchable As Boolean = True, _
    Optional ByVal controlType As String = "TextBox", _
    Optional ByVal dropdownValues As String = "")

    Dim wsConfig As Worksheet
    Dim wsData As Worksheet
    Dim configRow As Long
    Dim physicalPosition As Long

    databaseKey = SchemaNormalizeDatabaseKey(databaseKey)

    If Len(Trim$(headerName)) = 0 Then
        Err.Raise vbObjectError + 8102, _
            "SchemaAddField", _
            "Nama header baru masih kosong."
    End If

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)
    Set wsData = SchemaGetRequiredDataSheet(databaseKey)

    SchemaRegisterCurrentHeadersInternal _
        wsConfig, databaseKey

    If SchemaFindConfigRow( _
        wsConfig, databaseKey, headerName) > 0 Then

        Err.Raise vbObjectError + 8103, _
            "SchemaAddField", _
            "Header atau alias sudah terdaftar."

    End If

    requestedPosition = SchemaClampAddPosition( _
        wsConfig, databaseKey, requestedPosition)

    physicalPosition = SchemaClampPhysicalInsertPosition( _
        wsData, requestedPosition)

    SchemaShiftPositions _
        wsConfig, databaseKey, requestedPosition

    configRow = SchemaAddFieldInternal( _
        wsConfig, databaseKey, headerName, headerName, _
        requestedPosition, dataType, showInForm, isRequired, _
        isReadOnly, isSearchable, controlType, dropdownValues, _
        vbNullString, False, True)

    SchemaInsertBlankColumnSafe _
        wsData, physicalPosition

    wsData.Cells(SchemaHeaderRow(databaseKey), physicalPosition).Value = _
        Trim$(headerName)

    SchemaNormalizePositionsFromSheet _
        wsConfig, databaseKey, wsData

    SchemaApplyVisibilityOnly databaseKey
    SchemaStyleHeader wsData

    SchemaWriteChangeLog _
        databaseKey, _
        "ADD", _
        vbNullString, _
        headerName, _
        CStr(wsConfig.Cells( _
            configRow, C_FIELD_ID).Value2)

End Sub

Public Sub SchemaRenameField( _
    ByVal databaseKey As String, _
    ByVal fieldReference As String, _
    ByVal newHeader As String)

    Dim wsConfig As Worksheet
    Dim wsData As Worksheet
    Dim configRow As Long
    Dim duplicateRow As Long
    Dim oldHeader As String
    Dim aliases As String
    Dim columnNumber As Long

    databaseKey = SchemaNormalizeDatabaseKey(databaseKey)

    If Len(Trim$(newHeader)) = 0 Then
        Err.Raise vbObjectError + 8104, _
            "SchemaRenameField", _
            "Nama header baru masih kosong."
    End If

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)
    Set wsData = SchemaGetRequiredDataSheet(databaseKey)

    configRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, fieldReference)

    If configRow = 0 Then
        Err.Raise vbObjectError + 8105, _
            "SchemaRenameField", _
            "Field tidak ditemukan."
    End If

    duplicateRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, newHeader)

    If duplicateRow > 0 And duplicateRow <> configRow Then
        Err.Raise vbObjectError + 8106, _
            "SchemaRenameField", _
            "Nama header baru sudah digunakan."
    End If

    oldHeader = CStr( _
        wsConfig.Cells(configRow, C_HEADER_NAME).Value2)

    aliases = SchemaAppendAlias( _
        CStr(wsConfig.Cells(configRow, C_ALIASES).Value2), _
        oldHeader)

    columnNumber = SchemaFindPhysicalColumn( _
        wsData, _
        oldHeader, _
        CStr(wsConfig.Cells( _
            configRow, C_CANONICAL_HEADER).Value2), _
        aliases)

    If columnNumber = 0 Then
        Err.Raise vbObjectError + 8121, _
            "SchemaRenameField", _
            "Kolom fisik untuk field tersebut tidak ditemukan di sheet " & _
            databaseKey & "."
    End If

    wsData.Cells(SchemaHeaderRow(databaseKey), columnNumber).Value = _
        Trim$(newHeader)

    wsConfig.Cells(configRow, C_HEADER_NAME).Value = _
        Trim$(newHeader)

    wsConfig.Cells(configRow, C_ALIASES).Value = aliases
    wsConfig.Cells(configRow, C_UPDATED_AT).Value = Now

    SchemaRegisterCurrentHeadersInternal _
        wsConfig, databaseKey

    SchemaStyleHeader wsData

    SchemaWriteChangeLog _
        databaseKey, _
        "RENAME", _
        oldHeader, _
        newHeader, _
        fieldReference

End Sub

Public Sub SchemaDeactivateField( _
    ByVal databaseKey As String, _
    ByVal fieldReference As String)

    Dim wsConfig As Worksheet
    Dim wsData As Worksheet
    Dim configRow As Long
    Dim columnNumber As Long

    databaseKey = SchemaNormalizeDatabaseKey(databaseKey)

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)
    Set wsData = SchemaGetRequiredDataSheet(databaseKey)

    configRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, fieldReference)

    If configRow = 0 Then
        Err.Raise vbObjectError + 8107, _
            "SchemaDeactivateField", _
            "Field tidak ditemukan."
    End If

    If Not SchemaYes( _
        wsConfig.Cells(configRow, C_CUSTOM).Value2) Then

        Err.Raise vbObjectError + 8108, _
            "SchemaDeactivateField", _
            "Core field tidak boleh dinonaktifkan. " & _
            "Core field tetap boleh di-rename atau dipindahkan."

    End If

    wsConfig.Cells(configRow, C_ACTIVE).Value = "NO"
    wsConfig.Cells(configRow, C_SHOW_FORM).Value = "NO"
    wsConfig.Cells(configRow, C_UPDATED_AT).Value = Now

    columnNumber = SchemaFindColumn( _
        wsData, databaseKey, fieldReference, False)

    If columnNumber > 0 Then
        SchemaSetColumnHiddenSafe _
            wsData, columnNumber, True
    End If

    SchemaWriteChangeLog _
        databaseKey, _
        "DEACTIVATE", _
        CStr(wsConfig.Cells( _
            configRow, C_HEADER_NAME).Value2), _
        vbNullString, _
        fieldReference

End Sub

Public Sub SchemaActivateField( _
    ByVal databaseKey As String, _
    ByVal fieldReference As String)

    Dim wsConfig As Worksheet
    Dim wsData As Worksheet
    Dim configRow As Long
    Dim columnNumber As Long
    Dim requestedPosition As Long

    databaseKey = SchemaNormalizeDatabaseKey(databaseKey)

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)
    Set wsData = SchemaGetRequiredDataSheet(databaseKey)

    configRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, fieldReference)

    If configRow = 0 Then
        Err.Raise vbObjectError + 8109, _
            "SchemaActivateField", _
            "Field tidak ditemukan."
    End If

    wsConfig.Cells(configRow, C_ACTIVE).Value = "YES"
    wsConfig.Cells(configRow, C_SHOW_FORM).Value = "YES"
    wsConfig.Cells(configRow, C_UPDATED_AT).Value = Now

    columnNumber = SchemaFindColumn( _
        wsData, databaseKey, fieldReference, False)

    If columnNumber = 0 Then

        requestedPosition = CLng(val( _
            wsConfig.Cells(configRow, C_POSITION).Value2))

        requestedPosition = _
            SchemaClampPhysicalInsertPosition( _
                wsData, requestedPosition)

        SchemaInsertBlankColumnSafe _
            wsData, requestedPosition

        wsData.Cells(SchemaHeaderRow(databaseKey), requestedPosition).Value = _
            CStr(wsConfig.Cells( _
                configRow, C_HEADER_NAME).Value2)

        columnNumber = requestedPosition

    End If

    SchemaSetColumnHiddenSafe _
        wsData, columnNumber, False

    SchemaNormalizePositionsFromSheet _
        wsConfig, databaseKey, wsData

    SchemaWriteChangeLog _
        databaseKey, _
        "ACTIVATE", _
        vbNullString, _
        CStr(wsConfig.Cells( _
            configRow, C_HEADER_NAME).Value2), _
        fieldReference

End Sub

Public Sub SchemaMoveField( _
    ByVal databaseKey As String, _
    ByVal fieldReference As String, _
    ByVal newPosition As Long)

    Dim wsConfig As Worksheet
    Dim wsData As Worksheet
    Dim configRow As Long
    Dim oldPosition As Long
    Dim sourceColumn As Long
    Dim destinationColumn As Long

    databaseKey = SchemaNormalizeDatabaseKey(databaseKey)

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)
    Set wsData = SchemaGetRequiredDataSheet(databaseKey)

    configRow = SchemaFindConfigRow( _
        wsConfig, databaseKey, fieldReference)

    If configRow = 0 Then
        Err.Raise vbObjectError + 8111, _
            "SchemaMoveField", _
            "Field tidak ditemukan."
    End If

    newPosition = SchemaClampMovePosition( _
        wsConfig, databaseKey, newPosition)

    oldPosition = CLng(val( _
        wsConfig.Cells(configRow, C_POSITION).Value2))

    sourceColumn = SchemaFindColumn( _
        wsData, databaseKey, fieldReference, False)

    If sourceColumn = 0 Then
        Err.Raise vbObjectError + 8122, _
            "SchemaMoveField", _
            "Kolom fisik tidak ditemukan di sheet " & databaseKey & "."
    End If

    destinationColumn = newPosition

    If destinationColumn > SchemaLastHeaderColumn(wsData) Then
        destinationColumn = SchemaLastHeaderColumn(wsData)
    End If

    If destinationColumn < 1 Then destinationColumn = 1

    SchemaMoveColumnSafe _
        wsData, sourceColumn, destinationColumn

    SchemaNormalizePositionsFromSheet _
        wsConfig, databaseKey, wsData

    SchemaApplyVisibilityOnly databaseKey
    SchemaStyleHeader wsData

    SchemaWriteChangeLog _
        databaseKey, _
        "MOVE", _
        CStr(oldPosition), _
        CStr(destinationColumn), _
        fieldReference

End Sub

Public Sub SchemaReorderSheet( _
    ByVal databaseKey As String)

    Dim wsConfig As Worksheet
    Dim wsData As Worksheet
    Dim sortedRows As Collection
    Dim item As Variant
    Dim desiredColumn As Long
    Dim sourceColumn As Long
    Dim fieldID As String

    databaseKey = SchemaNormalizeDatabaseKey(databaseKey)

    EnsureSchemaConfig
    Set wsConfig = ThisWorkbook.Worksheets(SCHEMA_SHEET)
    Set wsData = SchemaGetRequiredDataSheet(databaseKey)

    Set sortedRows = SchemaRowsSorted( _
        wsConfig, databaseKey, True)

    desiredColumn = 1

    For Each item In sortedRows

        fieldID = CStr( _
            wsConfig.Cells(CLng(item), C_FIELD_ID).Value2)

        sourceColumn = SchemaFindColumn( _
            wsData, databaseKey, fieldID, False)

        If sourceColumn > 0 Then

            If sourceColumn <> desiredColumn Then
                SchemaMoveColumnSafe _
                    wsData, sourceColumn, desiredColumn
            End If

            desiredColumn = desiredColumn + 1

        End If

    Next item

    SchemaNormalizePositionsFromSheet _
        wsConfig, databaseKey, wsData

    SchemaApplyVisibilityOnly databaseKey
    SchemaStyleHeader wsData

End Sub

' ============================================================
' VISIBILITY
' ============================================================
Public Sub SchemaApplyVisibilityAll()

    SchemaApplyVisibilityOnly "FPTK"
    SchemaApplyVisibilityOnly "DB Kode Posisi"
    SchemaApplyVisibilityOnly "DB Sourcing"

End Sub

Public Sub SchemaApplyVisibilityOnly( _
    ByVal databaseKey As String)

    Dim wsConfig As Worksheet
    Dim wsData As Worksheet
    Dim physicalMap As Object

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim columnNumber As Long

    Dim fieldID As String
    Dim currentHeader As String
    Dim canonicalHeader As String
    Dim aliases As String

    databaseKey = _
        SchemaNormalizeDatabaseKey(databaseKey)

    EnsureSchemaConfig

    If Not SchemaSheetExists(databaseKey) Then
        Exit Sub
    End If

    Set wsConfig = _
        ThisWorkbook.Worksheets(SCHEMA_SHEET)

    Set wsData = _
        ThisWorkbook.Worksheets(databaseKey)

    Set physicalMap = _
        SchemaBuildPhysicalHeaderMap( _
            wsData, _
            databaseKey)

    lastRow = SchemaLastUsedRow(wsConfig)

    For rowIndex = SCHEMA_FIRST_ROW _
        To lastRow

        If StrComp( _
            CStr(wsConfig.Cells( _
                rowIndex, _
                C_DATABASE).Value2), _
            databaseKey, _
            vbTextCompare) = 0 Then

            fieldID = CStr( _
                wsConfig.Cells( _
                    rowIndex, _
                    C_FIELD_ID).Value2)

            currentHeader = CStr( _
                wsConfig.Cells( _
                    rowIndex, _
                    C_HEADER_NAME).Value2)

            canonicalHeader = CStr( _
                wsConfig.Cells( _
                    rowIndex, _
                    C_CANONICAL_HEADER).Value2)

            aliases = CStr( _
                wsConfig.Cells( _
                    rowIndex, _
                    C_ALIASES).Value2)

            columnNumber = _
                SchemaColumnFromPhysicalMap( _
                    physicalMap, _
                    fieldID, _
                    currentHeader, _
                    canonicalHeader, _
                    aliases)

            If columnNumber > 0 Then

                SchemaSetColumnHiddenSafe _
                    wsData, _
                    columnNumber, _
                    Not SchemaYes( _
                        wsConfig.Cells( _
                            rowIndex, _
                            C_ACTIVE).Value2)

            End If

        End If

    Next rowIndex

    Set physicalMap = Nothing
    Set wsData = Nothing
    Set wsConfig = Nothing

End Sub

Private Function SchemaBuildPhysicalHeaderMap( _
    ByVal ws As Worksheet, _
    ByVal databaseKey As String) As Object

    Dim result As Object
    Dim headerRow As Long
    Dim lastColumn As Long
    Dim columnIndex As Long

    Dim headerName As String
    Dim normalizedKey As String
    Dim guessedCanonical As String

    Set result = CreateObject( _
        "Scripting.Dictionary")

    result.CompareMode = vbTextCompare

    headerRow = SchemaHeaderRow(databaseKey)
    lastColumn = SchemaLastHeaderColumn(ws)

    For columnIndex = 1 To lastColumn

        headerName = Trim$(CStr( _
            ws.Cells( _
                headerRow, _
                columnIndex).Value2))

        normalizedKey = _
            SchemaNormalize(headerName)

        If Len(normalizedKey) > 0 Then

            If Not result.Exists(normalizedKey) Then

                result.Add _
                    normalizedKey, _
                    columnIndex

            End If

        End If

        guessedCanonical = _
            SchemaGuessCanonicalHeader( _
                databaseKey, _
                headerName)

        normalizedKey = _
            SchemaNormalize(guessedCanonical)

        If Len(normalizedKey) > 0 Then

            If Not result.Exists(normalizedKey) Then

                result.Add _
                    normalizedKey, _
                    columnIndex

            End If

        End If

    Next columnIndex

    Set SchemaBuildPhysicalHeaderMap = result

End Function

Private Function SchemaColumnFromPhysicalMap( _
    ByVal physicalMap As Object, _
    ByVal fieldID As String, _
    ByVal currentHeader As String, _
    ByVal canonicalHeader As String, _
    ByVal aliasesText As String) As Long

    Dim requestedItems As Variant
    Dim item As Variant
    Dim normalizedKey As String

    If physicalMap Is Nothing Then
        Exit Function
    End If

    requestedItems = Array( _
        fieldID, _
        currentHeader, _
        canonicalHeader)

    For Each item In requestedItems

        normalizedKey = _
            SchemaNormalize(CStr(item))

        If Len(normalizedKey) > 0 Then

            If physicalMap.Exists(normalizedKey) Then

                SchemaColumnFromPhysicalMap = _
                    CLng(physicalMap(normalizedKey))

                Exit Function

            End If

        End If

    Next item

    For Each item In Split( _
        aliasesText, _
        "|")

        normalizedKey = _
            SchemaNormalize(CStr(item))

        If Len(normalizedKey) > 0 Then

            If physicalMap.Exists(normalizedKey) Then

                SchemaColumnFromPhysicalMap = _
                    CLng(physicalMap(normalizedKey))

                Exit Function

            End If

        End If

    Next item

End Function

' ============================================================
' INTERNAL CONFIG / SEED
' ============================================================
Private Sub SchemaWriteConfigHeaders( _
    ByVal ws As Worksheet)

    Dim headers As Variant
    Dim columnIndex As Long

    headers = Array( _
        "Database", "Field ID", "Header Name", "Position", _
        "Active", "Data Type", "Show In Form", "Required", _
        "Read Only", "Searchable", "Control Type", _
        "Dropdown Values", "Rule", "Aliases", "Key Field", _
        "Custom Field", "Updated At", "Canonical Header")

    For columnIndex = LBound(headers) To UBound(headers)
        ws.Cells(1, columnIndex + 1).Value = headers(columnIndex)
    Next columnIndex

End Sub
Private Sub SchemaSeedFPTK( _
    ByVal ws As Worksheet)

    Dim headerText As String

    headerText = _
        "Kode PIC|FPTK Date (Real)|Kode Angka|" & _
        "FPTK Date (Kode)|Kode Unik|Posisi|" & _
        "Business Unit|Direktorat|Divisi|Department|" & _
        "Level FPTK|Level Number|Alasan Permintaan FPTK|" & _
        "Category FPTK|PIC Recruiter|" & _
        "Filter Kategorisasi FPTK|Vacancy|Status"

    headerText = headerText & _
        "|Week FPTK Date (Kode)|Month FPTK Date|" & _
        "FPTK Cancel Date|Week Cancel Date|" & _
        "Month Cancel Date|Offering Date|" & _
        "Week Offering Date|Month Offering|" & _
        "Jumlah SLA|Deadline pemenuhan SLA|" & _
        "Detail SLA|Keterangan Lulus SLA|" & _
        "Keterangan Tidak Lulus SLA|" & _
        "Keterangan Cancel"

    headerText = headerText & _
        "|Nama Kandidat|Estimasi Join|" & _
        "Kebutuhan Laptop|Lokasi Onboarding|" & _
        "Tanggal Upload ke Website|" & _
        "User (Manager)|Indirect User|" & _
        "Lokasi Kerja|Lokasi HR|Status Karyawan|" & _
        "Kode BU|FPTK Availability|Created At|" & _
        "Last Updated At|Last Compile Action|" & _
        "Source File|Remark"

    SchemaSeedFromPipe _
        ws, _
        "FPTK", _
        headerText

End Sub

Private Sub SchemaSeedDBKodePosisi( _
    ByVal ws As Worksheet)

    SchemaSeedFromPipe ws, "DB Kode Posisi", _
        "KODE|POSITION|LOCATION|BUSINESS UNIT|DIVISION CHRIS|" & _
        "DEPARTMENT CHRIS|USER (MANAGER)|INDIRECT USER|" & _
        "DIRECTORATE|YEAR"

End Sub

Private Sub SchemaSeedDBSourcing( _
    ByVal ws As Worksheet)

    Dim headerText As String

    headerText = _
        "No|Sourcing Date|Kode Unik (copy value dari FPTK)|Posisi|" & _
        "Model Rekrutmen (Lihat di Sheet Flow Map Model)|Rekruter|" & _
        "Sumber Sourcing|Nama|Nama Universitas/Sekolah (TOP 10)|" & _
        "Nama Universitas/Sekolah Lainnya|Jenjang Pendidikan|Jurusan|" & _
        "Tahun Lulus|IPK|Skor Bahasa Inggris|University Tier|IPK Tier|" & _
        "Nomor HP|Email|Domisili|Last Position|Last Tenure|Last Company|" & _
        "Total Tenure|Pernah di FMCG?|Sourcing Freelance|" & _
        "Tanggal Sourcing Freelance|Sourcing HR|" & _
        "Detail Keterangan Sourcing HR|Tanggal Sourcing|Shortlist CV|"

    headerText = headerText & _
        "Detail Keterangan Shortlist CV|Tanggal Shortlist CV|Psikotes|" & _
        "Kode Psikotes|Detail Keterangan Psikotes|" & _
        "Tanggal Psikotes / Cek psikotes|Nilai Logika|Nilai IQ|" & _
        "Nilai Daya Tangkap|Nilai RA|DISC|HR Interview|" & _
        "Detail Keterangan HR Interview|Tanggal HR Interview|" & _
        "Technical Test/ Case Study|" & _
        "Detail Keterangan Technical Test/ Case Study|" & _
        "Tanggal Technical Test/ Case Study|Market Visit|" & _
        "Detail Market Visit|Tanggal Market Visit|User Interview|"

    headerText = headerText & _
        "Detail Keterangan User Interview|Tanggal User Interview|" & _
        "Panel Interview|Detail Keterangan Panel Interview|" & _
        "Tanggal Panel Interview|Reference Check|" & _
        "Detail Keterangan Reference Check|Tanggal Reference Check|" & _
        "MCU|Detail Keterangan MCU|Tanggal MCU|Offering|" & _
        "Detail Keterangan Offering|Tanggal Offering|Notes|Day 1|" & _
        "Detail Keterangan Day 1|Tanggal Day 1"

    SchemaSeedFromPipe ws, "DB Sourcing", headerText

End Sub

Private Sub SchemaSeedFromPipe( _
    ByVal ws As Worksheet, _
    ByVal databaseKey As String, _
    ByVal headerText As String)

    Dim headers As Variant
    Dim item As Variant
    Dim positionNumber As Long
    Dim configRow As Long

    headers = Split(headerText, "|")
    positionNumber = 1

    For Each item In headers

        configRow = SchemaFindConfigRow( _
            ws, databaseKey, CStr(item))

        If configRow = 0 Then

            SchemaAddFieldInternal ws, databaseKey, _
                CStr(item), CStr(item), positionNumber, _
                SchemaInferDataType(CStr(item)), True, _
                SchemaDefaultRequired(databaseKey, CStr(item)), _
                SchemaDefaultReadOnly(databaseKey, CStr(item)), _
                True, SchemaInferControlType(CStr(item)), _
                SchemaDefaultDropdown(databaseKey, CStr(item)), _
                SchemaDefaultAliases(databaseKey, CStr(item)), _
                SchemaDefaultKeyField(databaseKey, CStr(item)), False

        Else

            If Len(Trim$(CStr( _
                ws.Cells(configRow, C_CANONICAL_HEADER).Value2))) = 0 Then
                ws.Cells(configRow, C_CANONICAL_HEADER).Value = CStr(item)
            End If

            ws.Cells(configRow, C_CUSTOM).Value = "NO"

            If SchemaDefaultKeyField(databaseKey, CStr(item)) Then
                ws.Cells(configRow, C_KEY_FIELD).Value = "YES"
            End If

            ws.Cells(configRow, C_ALIASES).Value = _
                SchemaMergeAliases( _
                    CStr(ws.Cells(configRow, C_ALIASES).Value2), _
                    SchemaDefaultAliases(databaseKey, CStr(item)))

        End If

        positionNumber = positionNumber + 1

    Next item

End Sub

Private Function SchemaAddFieldInternal( _
    ByVal ws As Worksheet, _
    ByVal databaseKey As String, _
    ByVal currentHeader As String, _
    ByVal canonicalHeader As String, _
    ByVal positionNumber As Long, _
    ByVal dataType As String, _
    ByVal showInForm As Boolean, _
    ByVal isRequired As Boolean, _
    ByVal isReadOnly As Boolean, _
    ByVal isSearchable As Boolean, _
    ByVal controlType As String, _
    ByVal dropdownValues As String, _
    ByVal aliases As String, _
    ByVal isKeyField As Boolean, _
    ByVal isCustom As Boolean) As Long

    Dim targetRow As Long
    Dim fieldID As String

    targetRow = SchemaLastUsedRow(ws) + 1
    If targetRow < SCHEMA_FIRST_ROW Then targetRow = SCHEMA_FIRST_ROW

    fieldID = SchemaUniqueFieldID( _
        ws, databaseKey, SchemaMakeFieldID(canonicalHeader))

    ws.Cells(targetRow, C_DATABASE).Value = databaseKey
    ws.Cells(targetRow, C_FIELD_ID).Value = fieldID
    ws.Cells(targetRow, C_HEADER_NAME).Value = currentHeader
    ws.Cells(targetRow, C_POSITION).Value = positionNumber
    ws.Cells(targetRow, C_ACTIVE).Value = "YES"
    ws.Cells(targetRow, C_DATA_TYPE).Value = dataType
    ws.Cells(targetRow, C_SHOW_FORM).Value = SchemaYesNo(showInForm)
    ws.Cells(targetRow, C_REQUIRED).Value = SchemaYesNo(isRequired)
    ws.Cells(targetRow, C_READ_ONLY).Value = SchemaYesNo(isReadOnly)
    ws.Cells(targetRow, C_SEARCHABLE).Value = SchemaYesNo(isSearchable)
    ws.Cells(targetRow, C_CONTROL_TYPE).Value = controlType
    ws.Cells(targetRow, C_DROPDOWN).Value = dropdownValues
    ws.Cells(targetRow, C_RULE).Value = vbNullString
    ws.Cells(targetRow, C_ALIASES).Value = aliases
    ws.Cells(targetRow, C_KEY_FIELD).Value = SchemaYesNo(isKeyField)
    ws.Cells(targetRow, C_CUSTOM).Value = SchemaYesNo(isCustom)
    ws.Cells(targetRow, C_UPDATED_AT).Value = Now
    ws.Cells(targetRow, C_CANONICAL_HEADER).Value = canonicalHeader

    SchemaAddFieldInternal = targetRow

End Function
Private Sub SchemaMergeKnownDBSourcingFMCGAliases( _
    ByVal wsConfig As Worksheet)

    Dim coreRow As Long
    Dim lastRow As Long
    Dim rowIndex As Long

    Dim headerName As String
    Dim canonicalHeader As String
    Dim aliasesText As String
    Dim guessedHeader As String

    If wsConfig Is Nothing Then
        Exit Sub
    End If

    coreRow = _
        SchemaFindConfigRow( _
            wsConfig, _
            "DB Sourcing", _
            "Pernah di FMCG?")

    If coreRow = 0 Then
        Exit Sub
    End If

    aliasesText = _
        CStr( _
            wsConfig.Cells( _
                coreRow, _
                C_ALIASES).Value2)

    aliasesText = _
        SchemaMergeAliases( _
            aliasesText, _
            SchemaDefaultAliases( _
                "DB Sourcing", _
                "Pernah di FMCG?"))

    lastRow = _
        SchemaLastUsedRow( _
            wsConfig)

    ' Loop dari bawah karena duplicate registry dihapus.
    For rowIndex = lastRow _
        To SCHEMA_FIRST_ROW Step -1

        If rowIndex <> coreRow Then

            If StrComp( _
                Trim$( _
                    CStr( _
                        wsConfig.Cells( _
                            rowIndex, _
                            C_DATABASE).Value2)), _
                "DB Sourcing", _
                vbTextCompare) = 0 Then

                headerName = _
                    Trim$( _
                        CStr( _
                            wsConfig.Cells( _
                                rowIndex, _
                                C_HEADER_NAME).Value2))

                canonicalHeader = _
                    Trim$( _
                        CStr( _
                            wsConfig.Cells( _
                                rowIndex, _
                                C_CANONICAL_HEADER).Value2))

                guessedHeader = _
                    SchemaGuessCanonicalHeader( _
                        "DB Sourcing", _
                        headerName)

                If Len(guessedHeader) = 0 Then

                    guessedHeader = _
                        SchemaGuessCanonicalHeader( _
                            "DB Sourcing", _
                            canonicalHeader)

                End If

                If StrComp( _
                    guessedHeader, _
                    "Pernah di FMCG?", _
                    vbTextCompare) = 0 Then

                    aliasesText = _
                        SchemaAppendAlias( _
                            aliasesText, _
                            headerName)

                    aliasesText = _
                        SchemaAppendAlias( _
                            aliasesText, _
                            canonicalHeader)

                    aliasesText = _
                        SchemaMergeAliases( _
                            aliasesText, _
                            CStr( _
                                wsConfig.Cells( _
                                    rowIndex, _
                                    C_ALIASES).Value2))

                    wsConfig.rows( _
                        rowIndex).Delete

                    If rowIndex < coreRow Then

                        coreRow = _
                            coreRow - 1

                    End If

                End If

            End If

        End If

    Next rowIndex

    wsConfig.Cells( _
        coreRow, _
        C_HEADER_NAME).Value = _
            "Pernah di FMCG?"

    wsConfig.Cells( _
        coreRow, _
        C_CANONICAL_HEADER).Value = _
            "Pernah di FMCG?"

    wsConfig.Cells( _
        coreRow, _
        C_ALIASES).Value = _
            aliasesText

    wsConfig.Cells( _
        coreRow, _
        C_ACTIVE).Value = _
            "YES"

    wsConfig.Cells( _
        coreRow, _
        C_CUSTOM).Value = _
            "NO"

    wsConfig.Cells( _
        coreRow, _
        C_UPDATED_AT).Value = _
            Now

    Set mSchemaCanonicalCache = Nothing

End Sub

Public Sub SchemaRepairDBSourcingFMCGAliasNow()

    Dim wsConfig As Worksheet

    SchemaResetCache
    EnsureSchemaConfig

    Set wsConfig = _
        ThisWorkbook.Worksheets( _
            SCHEMA_SHEET)

    SchemaMergeKnownDBSourcingFMCGAliases _
        wsConfig

    SchemaRegisterCurrentHeadersInternal _
        wsConfig, _
        "DB Sourcing"

    SchemaFormatConfig _
        wsConfig

    SchemaResetCache

    MsgBox _
        "Registry FMCG DB Sourcing sudah digabung ke " & _
        "'Pernah di FMCG?'.", _
        vbInformation, _
        "Repair Schema FMCG"

End Sub
Private Sub SchemaCleanupKnownFPTKLegacyFields( _
    ByVal wsConfig As Worksheet)

    Dim targetHeaders As Variant
    Dim targetHeader As Variant
    Dim coreRow As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim headerName As String
    Dim canonicalHeader As String
    Dim guessedHeader As String
    Dim aliasesText As String

    If wsConfig Is Nothing Then
        Exit Sub
    End If

    targetHeaders = Array( _
        "Alasan Permintaan FPTK", _
        "Divisi", _
        "Lokasi Kerja", _
        "Filter Kategorisasi FPTK", _
        "Jumlah SLA", _
        "Keterangan Lulus SLA", _
        "Keterangan Tidak Lulus SLA", _
        "Keterangan Cancel", _
        "Kebutuhan Laptop", _
        "Lokasi HR", _
        "Status Karyawan", _
        "Kode BU", _
        "FPTK Availability")

    For Each targetHeader In targetHeaders

        coreRow = _
            SchemaFindConfigRow( _
                wsConfig, _
                "FPTK", _
                CStr(targetHeader))

        If coreRow > 0 Then

            aliasesText = _
                CStr( _
                    wsConfig.Cells( _
                        coreRow, _
                        C_ALIASES).Value2)

            lastRow = _
                SchemaLastUsedRow( _
                    wsConfig)

            For rowIndex = lastRow _
                To SCHEMA_FIRST_ROW Step -1

                If rowIndex <> coreRow Then

                    If StrComp( _
                        Trim$( _
                            CStr( _
                                wsConfig.Cells( _
                                    rowIndex, _
                                    C_DATABASE).Value2)), _
                        "FPTK", _
                        vbTextCompare) = 0 Then

                        headerName = _
                            Trim$( _
                                CStr( _
                                    wsConfig.Cells( _
                                        rowIndex, _
                                        C_HEADER_NAME).Value2))

                        canonicalHeader = _
                            Trim$( _
                                CStr( _
                                    wsConfig.Cells( _
                                        rowIndex, _
                                        C_CANONICAL_HEADER).Value2))

                        guessedHeader = _
                            SchemaGuessCanonicalHeader( _
                                "FPTK", _
                                headerName)

                        If Len(guessedHeader) = 0 Then

                            guessedHeader = _
                                SchemaGuessCanonicalHeader( _
                                    "FPTK", _
                                    canonicalHeader)

                        End If

                        If StrComp( _
                            guessedHeader, _
                            CStr(targetHeader), _
                            vbTextCompare) = 0 Then

                            aliasesText = _
                                SchemaAppendAlias( _
                                    aliasesText, _
                                    headerName)

                            aliasesText = _
                                SchemaAppendAlias( _
                                    aliasesText, _
                                    canonicalHeader)

                            aliasesText = _
                                SchemaMergeAliases( _
                                    aliasesText, _
                                    CStr( _
                                        wsConfig.Cells( _
                                            rowIndex, _
                                            C_ALIASES).Value2))

                            wsConfig.rows( _
                                rowIndex).Delete

                            If rowIndex < coreRow Then

                                coreRow = _
                                    coreRow - 1

                            End If

                        End If

                    End If

                End If

            Next rowIndex

            wsConfig.Cells( _
                coreRow, _
                C_HEADER_NAME).Value = _
                    CStr(targetHeader)

            wsConfig.Cells( _
                coreRow, _
                C_CANONICAL_HEADER).Value = _
                    CStr(targetHeader)

            wsConfig.Cells( _
                coreRow, _
                C_ALIASES).Value = _
                    aliasesText

            wsConfig.Cells( _
                coreRow, _
                C_ACTIVE).Value = _
                    "YES"

            wsConfig.Cells( _
                coreRow, _
                C_CUSTOM).Value = _
                    "NO"

        End If

    Next targetHeader

    lastRow = _
        SchemaLastUsedRow( _
            wsConfig)

    For rowIndex = lastRow _
        To SCHEMA_FIRST_ROW Step -1

        If StrComp( _
            Trim$( _
                CStr( _
                    wsConfig.Cells( _
                        rowIndex, _
                        C_DATABASE).Value2)), _
            "FPTK", _
            vbTextCompare) = 0 Then

            headerName = _
                CStr( _
                    wsConfig.Cells( _
                        rowIndex, _
                        C_HEADER_NAME).Value2)

            If SchemaIsDeprecatedFPTKHeader( _
                "FPTK", _
                headerName) Then

                wsConfig.rows( _
                    rowIndex).Delete

            End If

        End If

    Next rowIndex

    Set mSchemaCanonicalCache = Nothing

End Sub

Private Function SchemaIsDeprecatedFPTKHeader( _
    ByVal databaseKey As String, _
    ByVal rawHeader As String) As Boolean

    Dim keyText As String

    If StrComp( _
        SchemaNormalizeDatabaseKey( _
            databaseKey), _
        "FPTK", _
        vbTextCompare) <> 0 Then

        Exit Function

    End If

    keyText = _
        SchemaNormalize( _
            rawHeader)

    Select Case keyText

        Case "MONTHFPTKDATEREAL", _
             "YEARFPTK", _
             "YEARFPTKDATE", _
             "FPTKYEAR"

            SchemaIsDeprecatedFPTKHeader = _
                True

    End Select

End Function

Public Sub SchemaRepairFPTKLegacyFieldsNow()

    Dim wsConfig As Worksheet

    SchemaResetCache
    EnsureSchemaConfig

    Set wsConfig = _
        ThisWorkbook.Worksheets( _
            SCHEMA_SHEET)

    SchemaCleanupKnownFPTKLegacyFields _
        wsConfig

    SchemaRegisterCurrentHeadersInternal _
        wsConfig, _
        "FPTK"

    SchemaFormatConfig _
        wsConfig

    SchemaResetCache

    MsgBox _
        "Registry FPTK legacy sudah dibersihkan.", _
        vbInformation, _
        "Repair Schema FPTK"

End Sub



Private Sub SchemaMigrateCanonicalHeaders( _
    ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = SchemaLastUsedRow(ws)

    For rowIndex = SCHEMA_FIRST_ROW To lastRow
        If Len(Trim$(CStr( _
            ws.Cells(rowIndex, C_CANONICAL_HEADER).Value2))) = 0 Then
            ws.Cells(rowIndex, C_CANONICAL_HEADER).Value = _
                CStr(ws.Cells(rowIndex, C_HEADER_NAME).Value2)
        End If
    Next rowIndex

End Sub

Private Sub SchemaRegisterCurrentHeadersInternal( _
    ByVal wsConfig As Worksheet, _
    ByVal databaseKey As String)

    Dim wsData As Worksheet
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim headerRow As Long
    Dim headerName As String
    Dim configRow As Long

    databaseKey = SchemaNormalizeDatabaseKey(databaseKey)

    If Not SchemaSheetExists(databaseKey) Then Exit Sub

    Set wsData = ThisWorkbook.Worksheets(databaseKey)
    headerRow = SchemaHeaderRow(databaseKey)
    lastColumn = SchemaLastHeaderColumn(wsData)

    For columnIndex = 1 To lastColumn

        headerName = Trim$(CStr( _
            wsData.Cells(headerRow, columnIndex).Value2))

        If Len(headerName) > 0 And _
           Not SchemaIsDeprecatedFPTKHeader( _
                databaseKey, _
                headerName) Then

            configRow = SchemaFindConfigRow( _
                wsConfig, databaseKey, headerName)

            If configRow = 0 Then

                configRow = SchemaAddFieldInternal( _
                    wsConfig, databaseKey, headerName, headerName, _
                    columnIndex, SchemaInferDataType(headerName), _
                    True, False, False, True, _
                    SchemaInferControlType(headerName), vbNullString, _
                    vbNullString, False, True)

            Else

                ' Header fisik yang berhasil diarahkan ke field existing
                ' otomatis disimpan sebagai alias. Dengan demikian,
                ' compile berikutnya tidak akan membuat field baru lagi.
                wsConfig.Cells( _
                    configRow, _
                    C_ALIASES).Value = _
                        SchemaAppendAlias( _
                            CStr(wsConfig.Cells( _
                                configRow, _
                                C_ALIASES).Value2), _
                            headerName)

                wsConfig.Cells(configRow, C_POSITION).Value = columnIndex

                If Len(Trim$(CStr( _
                    wsConfig.Cells(configRow, C_HEADER_NAME).Value2))) = 0 Then

                    wsConfig.Cells(configRow, C_HEADER_NAME).Value = headerName

                ElseIf Not SchemaYes(wsConfig.Cells( _
                    configRow, C_CUSTOM).Value2) And _
                   SchemaNormalize(CStr(wsConfig.Cells( _
                    configRow, C_HEADER_NAME).Value2)) = _
                   SchemaNormalize(CStr(wsConfig.Cells( _
                    configRow, C_CANONICAL_HEADER).Value2)) And _
                   SchemaNormalize(headerName) <> _
                   SchemaNormalize(CStr(wsConfig.Cells( _
                    configRow, C_HEADER_NAME).Value2)) Then

                    ' First migration: pertahankan label fisik lama.
                    ' Setelah user rename, current header sudah berbeda
                    ' dari canonical sehingga tidak akan ditimpa source.
                    wsConfig.Cells(configRow, C_ALIASES).Value = _
                        SchemaAppendAlias( _
                            CStr(wsConfig.Cells( _
                                configRow, C_ALIASES).Value2), _
                            CStr(wsConfig.Cells( _
                                configRow, C_HEADER_NAME).Value2))

                    wsConfig.Cells(configRow, C_HEADER_NAME).Value = _
                        headerName

                End If

            End If

        End If

    Next columnIndex

    ' Header registry mungkin berubah.
    Set mSchemaCanonicalCache = Nothing

End Sub

' ============================================================
' CONFIG LOOKUP / SORT
' ============================================================
Private Function SchemaFindConfigRow( _
    ByVal wsConfig As Worksheet, _
    ByVal databaseKey As String, _
    ByVal fieldReference As String) As Long

    Dim requestedKey As String
    Dim requestedSemanticKey As String
    Dim guessedCanonical As String
    Dim guessedKey As String

    Dim lastRow As Long
    Dim rowIndex As Long

    Dim aliases As Variant
    Dim aliasItem As Variant
    Dim candidateText As String

    requestedKey = _
        SchemaNormalize( _
            fieldReference)

    If Len(requestedKey) = 0 Then
        Exit Function
    End If

    lastRow = _
        SchemaLastUsedRow( _
            wsConfig)

    ' --------------------------------------------------------
    ' 1. Exact match:
    ' Field ID / Header Name / Canonical / Aliases.
    ' --------------------------------------------------------
    For rowIndex = SCHEMA_FIRST_ROW _
        To lastRow

        If StrComp( _
            Trim$(CStr(wsConfig.Cells( _
                rowIndex, _
                C_DATABASE).Value2)), _
            databaseKey, _
            vbTextCompare) = 0 Then

            If SchemaNormalize(CStr( _
                wsConfig.Cells( _
                    rowIndex, _
                    C_FIELD_ID).Value2)) = _
               requestedKey Or _
               SchemaNormalize(CStr( _
                wsConfig.Cells( _
                    rowIndex, _
                    C_HEADER_NAME).Value2)) = _
               requestedKey Or _
               SchemaNormalize(CStr( _
                wsConfig.Cells( _
                    rowIndex, _
                    C_CANONICAL_HEADER).Value2)) = _
               requestedKey Then

                SchemaFindConfigRow = _
                    rowIndex

                Exit Function

            End If

            aliases = Split( _
                CStr(wsConfig.Cells( _
                    rowIndex, _
                    C_ALIASES).Value2), _
                "|")

            For Each aliasItem In aliases

                If SchemaNormalize( _
                    CStr(aliasItem)) = _
                   requestedKey Then

                    SchemaFindConfigRow = _
                        rowIndex

                    Exit Function

                End If

            Next aliasItem

        End If

    Next rowIndex

    ' --------------------------------------------------------
    ' 2. Known canonical alias.
    ' Contoh MODEL / MODEL (LIHAT DROPDOWN)
    ' diarahkan ke MODEL REKRUTMEN.
    ' --------------------------------------------------------
    guessedCanonical = _
        SchemaGuessCanonicalHeader( _
            databaseKey, _
            fieldReference)

    guessedKey = _
        SchemaNormalize( _
            guessedCanonical)

    If Len(guessedKey) > 0 Then

        For rowIndex = SCHEMA_FIRST_ROW _
            To lastRow

            If StrComp( _
                Trim$(CStr(wsConfig.Cells( _
                    rowIndex, _
                    C_DATABASE).Value2)), _
                databaseKey, _
                vbTextCompare) = 0 Then

                If SchemaNormalize(CStr( _
                    wsConfig.Cells( _
                        rowIndex, _
                        C_CANONICAL_HEADER).Value2)) = _
                   guessedKey Or _
                   SchemaNormalize(CStr( _
                    wsConfig.Cells( _
                        rowIndex, _
                        C_HEADER_NAME).Value2)) = _
                   guessedKey Then

                    SchemaFindConfigRow = _
                        rowIndex

                    Exit Function

                End If

                aliases = Split( _
                    CStr(wsConfig.Cells( _
                        rowIndex, _
                        C_ALIASES).Value2), _
                    "|")

                For Each aliasItem In aliases

                    If SchemaNormalize( _
                        CStr(aliasItem)) = _
                       guessedKey Then

                        SchemaFindConfigRow = _
                            rowIndex

                        Exit Function

                    End If

                Next aliasItem

            End If

        Next rowIndex

    End If

    ' --------------------------------------------------------
    ' 3. Semantic match untuk catatan instruksi.
    ' Contoh:
    ' MODEL
    ' MODEL (LIHAT DROPDOWN)
    '
    ' Hanya bagian instruksi yang dibuang. Field yang berbeda,
    ' seperti STATUS dan STATUS KARYAWAN, tetap tidak disamakan.
    ' --------------------------------------------------------
    requestedSemanticKey = _
        SchemaSemanticHeaderKey( _
            fieldReference)

    If Len(requestedSemanticKey) = 0 Then
        Exit Function
    End If

    For rowIndex = SCHEMA_FIRST_ROW _
        To lastRow

        If StrComp( _
            Trim$(CStr(wsConfig.Cells( _
                rowIndex, _
                C_DATABASE).Value2)), _
            databaseKey, _
            vbTextCompare) = 0 Then

            candidateText = CStr( _
                wsConfig.Cells( _
                    rowIndex, _
                    C_HEADER_NAME).Value2)

            If SchemaSemanticHeaderKey( _
                candidateText) = _
               requestedSemanticKey Then

                SchemaFindConfigRow = _
                    rowIndex

                Exit Function

            End If

            candidateText = CStr( _
                wsConfig.Cells( _
                    rowIndex, _
                    C_CANONICAL_HEADER).Value2)

            If SchemaSemanticHeaderKey( _
                candidateText) = _
               requestedSemanticKey Then

                SchemaFindConfigRow = _
                    rowIndex

                Exit Function

            End If

            aliases = Split( _
                CStr(wsConfig.Cells( _
                    rowIndex, _
                    C_ALIASES).Value2), _
                "|")

            For Each aliasItem In aliases

                If SchemaSemanticHeaderKey( _
                    CStr(aliasItem)) = _
                   requestedSemanticKey Then

                    SchemaFindConfigRow = _
                        rowIndex

                    Exit Function

                End If

            Next aliasItem

        End If

    Next rowIndex

End Function

Private Function SchemaFindPhysicalColumn( _
    ByVal ws As Worksheet, _
    ByVal currentHeader As String, _
    ByVal canonicalHeader As String, _
    ByVal aliasesText As String) As Long

    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim currentKey As String
    Dim headerRow As Long
    Dim requestedKeys As Object
    Dim aliases As Variant
    Dim item As Variant

    Set requestedKeys = CreateObject("Scripting.Dictionary")
    requestedKeys.CompareMode = vbTextCompare

    SchemaAddRequestedKey requestedKeys, currentHeader
    SchemaAddRequestedKey requestedKeys, canonicalHeader

    aliases = Split(aliasesText, "|")
    For Each item In aliases
        SchemaAddRequestedKey requestedKeys, CStr(item)
    Next item

    headerRow = SchemaHeaderRow(ws.Name)
    lastColumn = SchemaLastHeaderColumn(ws)

    For columnIndex = 1 To lastColumn
        currentKey = SchemaNormalize(CStr( _
            ws.Cells(headerRow, columnIndex).Value2))
        If requestedKeys.Exists(currentKey) Then
            SchemaFindPhysicalColumn = columnIndex
            Exit Function
        End If

        If SchemaPhysicalHeaderMatchesCanonical( _
            CStr(ws.Cells(headerRow, columnIndex).Value2), _
            canonicalHeader) Then

            SchemaFindPhysicalColumn = columnIndex
            Exit Function

        End If
    Next columnIndex

End Function

Private Sub SchemaAddRequestedKey( _
    ByVal dict As Object, _
    ByVal rawText As String)

    Dim keyText As String
    keyText = SchemaNormalize(rawText)

    If Len(keyText) > 0 Then
        If Not dict.Exists(keyText) Then dict.Add keyText, True
    End If

End Sub

Private Function SchemaPhysicalHeaderMatchesCanonical( _
    ByVal physicalHeader As String, _
    ByVal canonicalHeader As String) As Boolean

    Dim canonicalKey As String
    Dim physicalGuess As String

    canonicalKey = SchemaNormalize(canonicalHeader)
    If Len(canonicalKey) = 0 Then Exit Function

    physicalGuess = SchemaGuessCanonicalHeader( _
        vbNullString, physicalHeader)

    SchemaPhysicalHeaderMatchesCanonical = _
        (SchemaNormalize(physicalGuess) = canonicalKey)

End Function
Private Function SchemaGuessCanonicalHeader( _
    ByVal databaseKey As String, _
    ByVal rawHeader As String) As String

    Dim keyText As String

    keyText = SchemaNormalize(rawHeader)
    If Len(keyText) = 0 Then Exit Function

    ' DB Kode Posisi core fields. Prefix matching diperlukan
    ' untuk header lama yang memiliki catatan panjang.
    If StrComp(databaseKey, "DB Kode Posisi", vbTextCompare) = 0 Or _
       Len(databaseKey) = 0 Then

        If keyText = "KODE" Or _
           keyText Like "KODEANGKA*" Or _
           keyText Like "KODEPOSISI*" Or _
           keyText Like "POSITIONCODE*" Then

            SchemaGuessCanonicalHeader = "KODE"
            Exit Function

        ElseIf keyText Like "POSITION*" Or _
               keyText Like "POSISI*" Then

            SchemaGuessCanonicalHeader = "POSITION"
            Exit Function

        ElseIf keyText Like "LOCATION*" Or _
               keyText Like "LOKASI*" Then

            SchemaGuessCanonicalHeader = "LOCATION"
            Exit Function

        ElseIf keyText Like "BUSINESSUNIT*" Or _
               keyText = "BU" Or _
               keyText Like "PTBUSINESSUNIT*" Then

            SchemaGuessCanonicalHeader = "BUSINESS UNIT"
            Exit Function

        ElseIf keyText Like "DIVISIONCHRIS*" Or _
               keyText Like "DIVISION*" Or _
               keyText Like "DIVISI*" Then

            SchemaGuessCanonicalHeader = "DIVISION CHRIS"
            Exit Function

        ElseIf keyText Like "DEPARTMENTCHRIS*" Or _
               keyText Like "DEPARTMENT*" Or _
               keyText Like "DEPARTEMEN*" Then

            SchemaGuessCanonicalHeader = "DEPARTMENT CHRIS"
            Exit Function

        ElseIf keyText Like "INDIRECTUSER*" Then

            SchemaGuessCanonicalHeader = "INDIRECT USER"
            Exit Function

        ElseIf keyText Like "USERMANAGER*" Or _
               keyText = "USER" Then

            SchemaGuessCanonicalHeader = "USER (MANAGER)"
            Exit Function

        ElseIf keyText Like "DIRECTORATE*" Or _
               keyText Like "DIREKTORAT*" Then

            SchemaGuessCanonicalHeader = "DIRECTORATE"
            Exit Function

        ElseIf keyText Like "YEAR*" Or _
               keyText Like "TAHUN*" Then

            SchemaGuessCanonicalHeader = "YEAR"
            Exit Function

        End If

    End If

    ' FPTK aliases yang paling sering berubah.
    If StrComp(databaseKey, "FPTK", vbTextCompare) = 0 Then

        If keyText = "STATUS" Or _
           keyText = "STATUSREKRUTMEN" Or _
           keyText = "STATUSVACANCY" Then

            SchemaGuessCanonicalHeader = _
                "Status"

        ElseIf keyText = "STATUSFPTK" Or _
               keyText = "CATEGORY" Or _
               keyText Like "ALASANFPTK*" Or _
               keyText Like "ALASANPERMINTAANFPTK*" Then

            SchemaGuessCanonicalHeader = _
                "Alasan Permintaan FPTK"

        ElseIf keyText = "KODE" Or _
           keyText Like "KODEANGKA*" Or _
           keyText Like "POSITIONCODE*" Then
            SchemaGuessCanonicalHeader = "Kode Angka"
        ElseIf keyText Like "POSISI*" Or _
               keyText Like "POSITION*" Then
            SchemaGuessCanonicalHeader = "Posisi"
        ElseIf keyText Like "BUSINESSUNIT*" Or _
               keyText Like "PTBUSINESSUNIT*" Then
            SchemaGuessCanonicalHeader = "Business Unit"
        ElseIf keyText Like "DIVISI*" Or _
               keyText Like "DIVISION*" Then
            SchemaGuessCanonicalHeader = "Divisi"
        ElseIf keyText Like "LOKASIKERJA*" Or _
               keyText = "LOCATION" Then
            SchemaGuessCanonicalHeader = "Lokasi Kerja"
        ElseIf keyText = "FPTKDATE" Or _
               keyText Like "FPTKDATEKODE*" Or _
               keyText Like "FPTKDATEPAKEINI*" Then
            SchemaGuessCanonicalHeader = "FPTK Date (Kode)"
        ElseIf keyText Like "WEEKFPTKDATE*" Then
            SchemaGuessCanonicalHeader = "Week FPTK Date (Kode)"
        ElseIf keyText = "SLADAYS" Or _
               keyText = "JUMLAHSLA" Then
            SchemaGuessCanonicalHeader = "Jumlah SLA"
        ElseIf keyText = "SLADEADLINE" Or _
               keyText Like "DEADLINEPEMENUHANSLA*" Then
            SchemaGuessCanonicalHeader = "Deadline pemenuhan SLA"
        ElseIf keyText = "SLASTATUS" Or _
               keyText = "DETAILSLA" Then
            SchemaGuessCanonicalHeader = "Detail SLA"
        ElseIf keyText Like "LEVELFPTK*" Then
            SchemaGuessCanonicalHeader = "Level FPTK"

        ElseIf keyText = "PENEMPATAN" Then
            SchemaGuessCanonicalHeader = "Lokasi Kerja"

        ElseIf keyText = "SLA" Then
            SchemaGuessCanonicalHeader = "Jumlah SLA"

        ElseIf keyText = "KETERANGAN1" Then
            SchemaGuessCanonicalHeader = _
                "Keterangan Lulus SLA"

        ElseIf keyText = "KETERANGAN0" Then
            SchemaGuessCanonicalHeader = _
                "Keterangan Tidak Lulus SLA"

        ElseIf keyText = _
               "KETERANGANCANCELKOSONG" Then
            SchemaGuessCanonicalHeader = _
                "Keterangan Cancel"

        ElseIf keyText = _
               "KEBUTUHANLAPTOPV" Then
            SchemaGuessCanonicalHeader = _
                "Kebutuhan Laptop"

        ElseIf keyText = _
               "LOKASIRECRUITER" Then
            SchemaGuessCanonicalHeader = _
                "Lokasi HR"

        ElseIf keyText Like _
               "FILTERKATEGORISASIFPTK*" Or _
               keyText Like _
               "FILTERKATEGORIFPTK*" Or _
               keyText Like _
               "KATEGORISASIFPTK*" Then

            SchemaGuessCanonicalHeader = _
                "Filter Kategorisasi FPTK"

        ElseIf keyText Like "STATUSKARYAWAN*" Or _
               keyText Like "EMPLOYMENTSTATUS*" Or _
               keyText Like "EMPLOYEESTATUS*" Then

            SchemaGuessCanonicalHeader = _
                "Status Karyawan"

        ElseIf keyText = "KODEBU" Or _
               keyText = "KODEBU2" Then

            SchemaGuessCanonicalHeader = _
                "Kode BU"

        ElseIf keyText = _
               "FPTKAVAILABILITY" Or _
               keyText = _
               "AVAILABILITYFPTK" Then
            SchemaGuessCanonicalHeader = _
                "FPTK Availability"

        End If

    End If


    ' DB Sourcing aliases.
    If StrComp( _
        databaseKey, _
        "DB Sourcing", _
        vbTextCompare) = 0 Then

        If keyText = "MODEL" Or _
           keyText Like "MODELREKRUTMEN*" Or _
           keyText Like "RECRUITMENTMODEL*" Or _
           keyText Like "MODEL*LIHAT*DROPDOWN*" Or _
           keyText Like "MODEL*DROPDOWN*" Then

            SchemaGuessCanonicalHeader = _
                "Model Rekrutmen " & _
                "(Lihat di Sheet Flow Map Model)"

            Exit Function

        ElseIf keyText = "KODEUNIK" Or _
               keyText Like "KODEUNIKCOPYVALUE*" Then

            SchemaGuessCanonicalHeader = _
                "Kode Unik " & _
                "(copy value dari FPTK)"

            Exit Function

        ElseIf keyText = "POSITION" Or _
               keyText Like "POSISI*" Then

            SchemaGuessCanonicalHeader = _
                "Posisi"

            Exit Function

        ElseIf keyText = "RECRUITER" Or _
               keyText = "PICRECRUITER" Or _
               keyText = "PICREKRUTER" Then

            SchemaGuessCanonicalHeader = _
                "Rekruter"

            Exit Function

        ElseIf keyText = "SOURCINGSOURCE" Or _
               keyText = "SOURCE" Then

            SchemaGuessCanonicalHeader = _
                "Sumber Sourcing"

            Exit Function

        ElseIf keyText = "FMCG" Or _
               keyText Like "PERNAHDIFMCG*" Or _
               keyText Like _
                    "BERPENGALAMANDIINDUSTRIFMCG*" Or _
               keyText Like "FMCGEXPERIENCE*" Or _
               keyText Like "PENGALAMANFMCG*" Or _
               keyText Like "EXPERIENCEINFMCG*" Then

            SchemaGuessCanonicalHeader = _
                "Pernah di FMCG?"

            Exit Function

        End If

    End If

End Function

Private Function SchemaSemanticHeaderKey( _
    ByVal rawHeader As String) As String

    Dim sourceText As String
    Dim cutPosition As Long
    Dim markerPosition As Long

    Dim markers As Variant
    Dim markerItem As Variant

    sourceText = _
        UCase$( _
            Trim$( _
                rawHeader))

    sourceText = Replace$( _
        sourceText, _
        Chr$(160), _
        " ")

    If Len(sourceText) = 0 Then
        Exit Function
    End If

    ' Potong catatan dalam kurung / bracket.
    cutPosition = InStr( _
        1, _
        sourceText, _
        "(", _
        vbBinaryCompare)

    If cutPosition = 0 Then

        cutPosition = InStr( _
            1, _
            sourceText, _
            "[", _
            vbBinaryCompare)

    End If

    If cutPosition > 1 Then

        sourceText = _
            Left$( _
                sourceText, _
                cutPosition - 1)

    End If

    ' Potong catatan instruksi yang berada tanpa kurung.
    markers = Array( _
        " LIHAT DROPDOWN", _
        " LIHAT DI SHEET", _
        " LIHAT SHEET", _
        " DROPDOWN", _
        " COPY VALUE DARI", _
        " COPY VALUE", _
        " ISI SESUAI")

    For Each markerItem In markers

        markerPosition = InStr( _
            1, _
            sourceText, _
            CStr(markerItem), _
            vbTextCompare)

        If markerPosition > 1 Then

            sourceText = _
                Left$( _
                    sourceText, _
                    markerPosition - 1)

        End If

    Next markerItem

    SchemaSemanticHeaderKey = _
        SchemaNormalize( _
            sourceText)

End Function

Private Function SchemaRowsSorted( _
    ByVal wsConfig As Worksheet, _
    ByVal databaseKey As String, _
    ByVal includeInactive As Boolean) As Collection

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim rows() As Long
    Dim positions() As Long
    Dim itemCount As Long
    Dim itemIndex As Long
    Dim i As Long
    Dim j As Long
    Dim tempRow As Long
    Dim tempPosition As Long
    Dim result As New Collection

    lastRow = SchemaLastUsedRow(wsConfig)

    For rowIndex = SCHEMA_FIRST_ROW To lastRow

        If StrComp( _
            CStr(wsConfig.Cells( _
                rowIndex, C_DATABASE).Value2), _
            databaseKey, _
            vbTextCompare) = 0 Then

            If includeInactive Or _
               SchemaYes(wsConfig.Cells( _
                    rowIndex, C_ACTIVE).Value2) Then

                itemCount = itemCount + 1

            End If

        End If

    Next rowIndex

    If itemCount = 0 Then
        Set SchemaRowsSorted = result
        Exit Function
    End If

    ReDim rows(1 To itemCount)
    ReDim positions(1 To itemCount)

    For rowIndex = SCHEMA_FIRST_ROW To lastRow

        If StrComp( _
            CStr(wsConfig.Cells( _
                rowIndex, C_DATABASE).Value2), _
            databaseKey, _
            vbTextCompare) = 0 Then

            If includeInactive Or _
               SchemaYes(wsConfig.Cells( _
                    rowIndex, C_ACTIVE).Value2) Then

                itemIndex = itemIndex + 1
                rows(itemIndex) = rowIndex

                positions(itemIndex) = CLng(val( _
                    wsConfig.Cells( _
                        rowIndex, C_POSITION).Value2))

                If positions(itemIndex) <= 0 Then
                    positions(itemIndex) = 100000 + itemIndex
                End If

            End If

        End If

    Next rowIndex

    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount

            If positions(j) < positions(i) Or _
               (positions(j) = positions(i) And _
                rows(j) < rows(i)) Then

                tempPosition = positions(i)
                positions(i) = positions(j)
                positions(j) = tempPosition

                tempRow = rows(i)
                rows(i) = rows(j)
                rows(j) = tempRow

            End If

        Next j
    Next i

    For i = 1 To itemCount
        result.Add rows(i)
    Next i

    Set SchemaRowsSorted = result

End Function

Private Sub SchemaShiftPositions( _
    ByVal wsConfig As Worksheet, _
    ByVal databaseKey As String, _
    ByVal startPosition As Long)

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim currentPosition As Long

    lastRow = SchemaLastUsedRow(wsConfig)

    For rowIndex = SCHEMA_FIRST_ROW To lastRow
        If StrComp(CStr(wsConfig.Cells(rowIndex, C_DATABASE).Value2), _
            databaseKey, vbTextCompare) = 0 Then
            currentPosition = CLng(val( _
                wsConfig.Cells(rowIndex, C_POSITION).Value2))
            If currentPosition >= startPosition Then
                wsConfig.Cells(rowIndex, C_POSITION).Value = _
                    currentPosition + 1
            End If
        End If
    Next rowIndex

End Sub

Private Function SchemaMaxPosition( _
    ByVal wsConfig As Worksheet, _
    ByVal databaseKey As String) As Long

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim valueItem As Long

    lastRow = SchemaLastUsedRow(wsConfig)

    For rowIndex = SCHEMA_FIRST_ROW To lastRow
        If StrComp(CStr(wsConfig.Cells(rowIndex, C_DATABASE).Value2), _
            databaseKey, vbTextCompare) = 0 Then
            valueItem = CLng(val( _
                wsConfig.Cells(rowIndex, C_POSITION).Value2))
            If valueItem > SchemaMaxPosition Then
                SchemaMaxPosition = valueItem
            End If
        End If
    Next rowIndex

End Function

Private Function SchemaUniqueFieldID( _
    ByVal wsConfig As Worksheet, _
    ByVal databaseKey As String, _
    ByVal baseID As String) As String

    Dim candidate As String
    Dim sequenceNumber As Long

    candidate = baseID

    Do While SchemaFindConfigRow( _
        wsConfig, databaseKey, candidate) > 0
        sequenceNumber = sequenceNumber + 1
        candidate = baseID & "_" & CStr(sequenceNumber)
    Loop

    SchemaUniqueFieldID = candidate

End Function


Private Function SchemaDatabaseFieldCount( _
    ByVal wsConfig As Worksheet, _
    ByVal databaseKey As String) As Long

    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = SchemaLastUsedRow(wsConfig)

    For rowIndex = SCHEMA_FIRST_ROW To lastRow

        If StrComp( _
            CStr(wsConfig.Cells( _
                rowIndex, C_DATABASE).Value2), _
            databaseKey, _
            vbTextCompare) = 0 Then

            SchemaDatabaseFieldCount = _
                SchemaDatabaseFieldCount + 1

        End If

    Next rowIndex

End Function

Private Function SchemaClampAddPosition( _
    ByVal wsConfig As Worksheet, _
    ByVal databaseKey As String, _
    ByVal requestedPosition As Long) As Long

    Dim maximumPosition As Long

    maximumPosition = _
        SchemaDatabaseFieldCount( _
            wsConfig, databaseKey) + 1

    If requestedPosition <= 0 Then
        requestedPosition = maximumPosition
    End If

    If requestedPosition < 1 Then requestedPosition = 1

    If requestedPosition > maximumPosition Then
        requestedPosition = maximumPosition
    End If

    SchemaClampAddPosition = requestedPosition

End Function

Private Function SchemaClampMovePosition( _
    ByVal wsConfig As Worksheet, _
    ByVal databaseKey As String, _
    ByVal requestedPosition As Long) As Long

    Dim maximumPosition As Long

    maximumPosition = SchemaDatabaseFieldCount( _
        wsConfig, databaseKey)

    If maximumPosition <= 0 Then
        Err.Raise vbObjectError + 8125, _
            "SchemaClampMovePosition", _
            "Database belum memiliki field."
    End If

    If requestedPosition <= 0 Then
        Err.Raise vbObjectError + 8110, _
            "SchemaMoveField", _
            "Posisi baru harus lebih besar dari 0."
    End If

    If requestedPosition > maximumPosition Then
        requestedPosition = maximumPosition
    End If

    SchemaClampMovePosition = requestedPosition

End Function

Private Function SchemaClampPhysicalInsertPosition( _
    ByVal ws As Worksheet, _
    ByVal requestedPosition As Long) As Long

    Dim maximumPosition As Long
    Dim headerRow As Long

    headerRow = SchemaHeaderRow(ws.Name)
    maximumPosition = SchemaLastHeaderColumn(ws) + 1

    If SchemaLastHeaderColumn(ws) = 1 And _
       Len(Trim$(CStr(ws.Cells(headerRow, 1).Value2))) = 0 Then

        maximumPosition = 1

    End If

    If requestedPosition <= 0 Then
        requestedPosition = maximumPosition
    End If

    If requestedPosition < 1 Then requestedPosition = 1

    If requestedPosition > maximumPosition Then
        requestedPosition = maximumPosition
    End If

    SchemaClampPhysicalInsertPosition = requestedPosition

End Function

Private Sub SchemaInsertBlankColumnSafe( _
    ByVal ws As Worksheet, _
    ByVal columnNumber As Long)

    If ws Is Nothing Then Exit Sub

    If columnNumber < 1 Or _
       columnNumber > ws.Columns.Count Then

        Err.Raise vbObjectError + 8126, _
            "SchemaInsertBlankColumnSafe", _
            "Posisi kolom di luar batas Excel."

    End If

    ws.Columns(columnNumber).Insert _
        Shift:=xlToRight

End Sub

Private Sub SchemaMoveColumnSafe( _
    ByVal ws As Worksheet, _
    ByVal sourceColumn As Long, _
    ByVal destinationColumn As Long)

    Dim insertColumn As Long

    If ws Is Nothing Then Exit Sub
    If sourceColumn = destinationColumn Then Exit Sub

    If sourceColumn < 1 Or _
       sourceColumn > ws.Columns.Count Or _
       destinationColumn < 1 Or _
       destinationColumn > ws.Columns.Count Then

        Err.Raise vbObjectError + 8127, _
            "SchemaMoveColumnSafe", _
            "Posisi perpindahan kolom di luar batas Excel."

    End If

    insertColumn = destinationColumn

    If sourceColumn < destinationColumn Then
        insertColumn = destinationColumn + 1
    End If

    ws.Columns(sourceColumn).Cut
    ws.Columns(insertColumn).Insert _
        Shift:=xlToRight

    Application.CutCopyMode = False

End Sub

Private Sub SchemaNormalizePositionsFromSheet( _
    ByVal wsConfig As Worksheet, _
    ByVal databaseKey As String, _
    ByVal wsData As Worksheet)

    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim headerName As String
    Dim configRow As Long
    Dim headerRow As Long

    headerRow = SchemaHeaderRow(databaseKey)
    lastColumn = SchemaLastHeaderColumn(wsData)

    For columnIndex = 1 To lastColumn

        headerName = Trim$(CStr( _
            wsData.Cells(headerRow, columnIndex).Value2))

        If Len(headerName) > 0 Then

            configRow = SchemaFindConfigRow( _
                wsConfig, databaseKey, headerName)

            If configRow > 0 Then

                wsConfig.Cells( _
                    configRow, C_POSITION).Value = _
                    columnIndex

                wsConfig.Cells( _
                    configRow, C_UPDATED_AT).Value = Now

            End If

        End If

    Next columnIndex

End Sub

' ============================================================
' DEFAULTS / NORMALIZATION
' ============================================================
Private Function SchemaMakeFieldID( _
    ByVal headerName As String) As String

    Dim resultText As String
    Dim sourceText As String
    Dim i As Long
    Dim ch As String

    sourceText = UCase$(Trim$(headerName))

    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If ch Like "[A-Z0-9]" Then
            resultText = resultText & ch
        ElseIf Len(resultText) > 0 Then
            If Right$(resultText, 1) <> "_" Then
                resultText = resultText & "_"
            End If
        End If
    Next i

    Do While Right$(resultText, 1) = "_"
        resultText = Left$(resultText, Len(resultText) - 1)
    Loop

    If Len(resultText) = 0 Then resultText = "FIELD"
    SchemaMakeFieldID = resultText

End Function

Private Function SchemaNormalize( _
    ByVal rawText As String) As String

    Dim resultText As String
    Dim sourceText As String
    Dim i As Long
    Dim ch As String

    sourceText = UCase$(rawText)

    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If ch Like "[A-Z0-9]" Then resultText = resultText & ch
    Next i

    SchemaNormalize = resultText

End Function

Private Function SchemaNormalizeWithSpaces( _
    ByVal rawText As String) As String

    Dim resultText As String
    Dim sourceText As String
    Dim i As Long
    Dim ch As String
    Dim lastWasSpace As Boolean

    sourceText = UCase$(rawText)

    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If ch Like "[A-Z0-9]" Then
            resultText = resultText & ch
            lastWasSpace = False
        ElseIf Len(resultText) > 0 And Not lastWasSpace Then
            resultText = resultText & " "
            lastWasSpace = True
        End If
    Next i

    SchemaNormalizeWithSpaces = Trim$(resultText)

End Function

Private Function SchemaAppendAlias( _
    ByVal aliasesText As String, _
    ByVal newAlias As String) As String

    Dim item As Variant
    Dim resultText As String
    Dim requestedKey As String
    Dim existsAlready As Boolean

    requestedKey = SchemaNormalize(newAlias)

    For Each item In Split(aliasesText, "|")

        If Len(Trim$(CStr(item))) > 0 Then

            If Len(resultText) > 0 Then
                resultText = resultText & "|"
            End If

            resultText = resultText & Trim$(CStr(item))

            If SchemaNormalize(CStr(item)) = requestedKey Then
                existsAlready = True
            End If

        End If

    Next item

    If Len(Trim$(newAlias)) > 0 And Not existsAlready Then

        If Len(resultText) > 0 Then
            resultText = resultText & "|"
        End If

        resultText = resultText & Trim$(newAlias)

    End If

    SchemaAppendAlias = resultText

End Function

Private Function SchemaMergeAliases( _
    ByVal aliasesText As String, _
    ByVal additionalAliases As String) As String

    Dim item As Variant
    Dim resultText As String

    resultText = aliasesText

    For Each item In Split(additionalAliases, "|")
        resultText = SchemaAppendAlias(resultText, CStr(item))
    Next item

    SchemaMergeAliases = resultText

End Function

Private Function SchemaInferDataType( _
    ByVal headerName As String) As String

    Dim keyText As String
    keyText = SchemaNormalize(headerName)

    If InStr(1, keyText, "DATE", vbTextCompare) > 0 Or _
       InStr(1, keyText, "TANGGAL", vbTextCompare) > 0 Or _
       InStr(1, keyText, "ESTIMASIJOIN", vbTextCompare) > 0 Then
        SchemaInferDataType = "Date"
    ElseIf keyText = "VACANCY" Or keyText = "IPK" Or _
           InStr(1, keyText, "NUMBER", vbTextCompare) > 0 Or _
           InStr(1, keyText, "JUMLAH", vbTextCompare) > 0 Then
        SchemaInferDataType = "Number"
    Else
        SchemaInferDataType = "Text"
    End If

End Function

Private Function SchemaInferControlType( _
    ByVal headerName As String) As String

    If Len(SchemaDefaultDropdown(vbNullString, headerName)) > 0 Then
        SchemaInferControlType = "ComboBox"
    Else
        SchemaInferControlType = "TextBox"
    End If

End Function
Private Function SchemaDefaultDropdown( _
    ByVal databaseKey As String, _
    ByVal headerName As String) As String

    Select Case SchemaNormalize(headerName)
        Case "STATUS"
            SchemaDefaultDropdown = "OP,Closed,Cancel"
        Case "STATUSKARYAWAN"
            SchemaDefaultDropdown = "Permanent,Contract,Internship"
        Case "KEBUTUHANLAPTOP"
            SchemaDefaultDropdown = "Ya,Tidak"

        Case "FPTKAVAILABILITY"
            SchemaDefaultDropdown = "Y,N"

    End Select

End Function
Private Function SchemaDefaultAliases( _
    ByVal databaseKey As String, _
    ByVal headerName As String) As String

    Dim keyText As String

    keyText = _
        SchemaNormalize( _
            headerName)

    If StrComp( _
        databaseKey, _
        "DB Kode Posisi", _
        vbTextCompare) = 0 Then

        Select Case keyText

            Case "KODE"
                SchemaDefaultAliases = _
                    "Kode Angka|Position Code|Kode Posisi"

            Case "POSITION"
                SchemaDefaultAliases = _
                    "Posisi|Nama Posisi"

            Case "LOCATION"
                SchemaDefaultAliases = _
                    "Lokasi|Lokasi Kerja"

            Case "BUSINESSUNIT"
                SchemaDefaultAliases = _
                    "BU|PT Business Unit|PT / Business Unit"

            Case "DIVISIONCHRIS"
                SchemaDefaultAliases = _
                    "Divisi|Division"

            Case "DEPARTMENTCHRIS"
                SchemaDefaultAliases = _
                    "Department|Departemen"

            Case "DIRECTORATE"
                SchemaDefaultAliases = _
                    "Direktorat"

        End Select

    ElseIf StrComp( _
        databaseKey, _
        "FPTK", _
        vbTextCompare) = 0 Then

        Select Case keyText

            Case "ALASANPERMINTAANFPTK", "CATEGORY"
                SchemaDefaultAliases = _
                    "Alasan FPTK|Alasan Permintaan FPTK|" & _
                    "Status FPTK"

            Case "STATUS"
                SchemaDefaultAliases = _
                    "Status Rekrutmen|Status Vacancy"

            Case "KODEANGKA"
                SchemaDefaultAliases = _
                    "KODE|Position Code|Kode Posisi"

            Case "POSISI"
                SchemaDefaultAliases = _
                    "Position|Posisi - Kebutuhan TA|" & _
                    "Nama Posisi"

            Case "BUSINESSUNIT"
                SchemaDefaultAliases = _
                    "PT Business Unit|PT / Business Unit"

            Case "DIVISI"
                SchemaDefaultAliases = _
                    "Divisi (Sesuai SO)|Division|" & _
                    "Division Chris"

            Case "LOKASIKERJA"
                SchemaDefaultAliases = _
                    "Location|Lokasi|Penempatan"

            Case "FPTKDATEKODE"
                SchemaDefaultAliases = _
                    "FPTK Date|FPTK Date (Pake Ini)|" & _
                    "FPTK Date (Code)"

            Case "WEEKFPTKDATEKODE"
                SchemaDefaultAliases = _
                    "Week FPTK Date|" & _
                    "week fptk date (Kode)"

            Case "FILTERKATEGORISASIFPTK"
                SchemaDefaultAliases = _
                    "Filter Kategorisasi FPTK2|" & _
                    "Filter Kategorisasi FPTK 2|" & _
                    "Filter Kategori FPTK|" & _
                    "Kategorisasi FPTK"

            Case "JUMLAHSLA"
                SchemaDefaultAliases = _
                    "SLA Days|SLA"

            Case "DEADLINEPEMENUHANSLA"
                SchemaDefaultAliases = _
                    "Deadline SLA|SLA Deadline"

            Case "DETAILSLA"
                SchemaDefaultAliases = _
                    "SLA Status|Status SLA"

            Case "LEVELFPTK"
                SchemaDefaultAliases = _
                    "Level FPTK"

            Case "FILTERKATEGORISASIFPTK"
                SchemaDefaultAliases = _
                    "Filter Kategorisasi FPTK2|" & _
                    "Filter Kategorisasi FPTK 2|" & _
                    "Filter Kategori FPTK|" & _
                    "Kategorisasi FPTK"

            Case "STATUSKARYAWAN"
                SchemaDefaultAliases = _
                    "Employment Status|Employee Status|" & _
                    "Status Employee"

            Case "KODEBU"
                SchemaDefaultAliases = _
                    "Kode BU2|Kode BU 2|BU Code"

            Case "KETERANGANLULUSSLA"
                SchemaDefaultAliases = _
                    "Keterangan ""1"""

            Case "KETERANGANTIDAKLULUSSLA"
                SchemaDefaultAliases = _
                    "Keterangan ""0"""

            Case "KETERANGANCANCEL"
                SchemaDefaultAliases = _
                    "Keterangan Cancel [Kosong]"

            Case "KEBUTUHANLAPTOP"
                SchemaDefaultAliases = _
                    "Kebutuhan Laptop (V)"

            Case "LOKASIHR"
                SchemaDefaultAliases = _
                    "HR Location|Lokasi Recruiter"

            Case "FPTKAVAILABILITY"
                SchemaDefaultAliases = _
                    "Availability FPTK|" & _
                    "FPTK Available"

        End Select

    ElseIf StrComp( _
        databaseKey, _
        "DB Sourcing", _
        vbTextCompare) = 0 Then

        Select Case keyText

            Case "MODELREKRUTMENLIHATDISHEETFLOWMAPMODEL"
                SchemaDefaultAliases = _
                    "Model|Model Rekrutmen|" & _
                    "Model (Lihat Dropdown)|" & _
                    "Model Rekrutmen (Lihat Dropdown)|" & _
                    "Recruitment Model"

            Case "KODEUNIKCOPYVALUEDARIFPTK"
                SchemaDefaultAliases = _
                    "Kode Unik|" & _
                    "Kode Unik Copy Value Dari FPTK"

            Case "POSISI"
                SchemaDefaultAliases = _
                    "Position|" & _
                    "Posisi - Kebutuhan TA"

            Case "REKRUTER"
                SchemaDefaultAliases = _
                    "Recruiter|PIC Recruiter|" & _
                    "PIC Rekruter"

            Case "SUMBERSOURCING"
                SchemaDefaultAliases = _
                    "Sourcing Source|Source"

            Case "PERNAHDIFMCG"
                SchemaDefaultAliases = _
                    "FMCG|" & _
                    "FMCG Experience|" & _
                    "Pengalaman FMCG|" & _
                    "Experience in FMCG|" & _
                    "Berpengalaman di industri FMCG|" & _
                    "Berpengalaman di Industri FMCG " & _
                    "(Berdasarkan seluruh riwayat kerja)"

        End Select

    End If

End Function

Private Function SchemaDefaultRequired( _
    ByVal databaseKey As String, _
    ByVal headerName As String) As Boolean

    Dim keyText As String
    keyText = SchemaNormalize(headerName)

    If StrComp(databaseKey, "DB Kode Posisi", vbTextCompare) = 0 Then
        SchemaDefaultRequired = _
            (keyText = "POSITION" Or keyText = "LOCATION" Or _
             keyText = "BUSINESSUNIT")
    ElseIf StrComp(databaseKey, "FPTK", vbTextCompare) = 0 Then
        SchemaDefaultRequired = _
            (keyText = "POSISI" Or keyText = "BUSINESSUNIT")
    ElseIf StrComp(databaseKey, "DB Sourcing", vbTextCompare) = 0 Then
        SchemaDefaultRequired = _
            (keyText = "NAMA" Or keyText = "POSISI")
    End If

End Function

Private Function SchemaDefaultReadOnly( _
    ByVal databaseKey As String, _
    ByVal headerName As String) As Boolean

    Dim keyText As String
    keyText = SchemaNormalize(headerName)

    SchemaDefaultReadOnly = _
        (keyText = "KODE" Or keyText = "KODEANGKA" Or _
         keyText = "KODEUNIK" Or keyText = "NO")

End Function

Private Function SchemaDefaultKeyField( _
    ByVal databaseKey As String, _
    ByVal headerName As String) As Boolean

    Dim keyText As String
    keyText = SchemaNormalize(headerName)

    If StrComp(databaseKey, "DB Kode Posisi", vbTextCompare) = 0 Then
        SchemaDefaultKeyField = _
            (keyText = "KODE" Or keyText = "POSITION" Or _
             keyText = "LOCATION" Or keyText = "BUSINESSUNIT")
    ElseIf StrComp(databaseKey, "FPTK", vbTextCompare) = 0 Then
        SchemaDefaultKeyField = _
            (keyText = "KODEUNIK" Or keyText = "POSISI")
    ElseIf StrComp(databaseKey, "DB Sourcing", vbTextCompare) = 0 Then
        SchemaDefaultKeyField = _
            (keyText = "NO" Or keyText = "KODEUNIK" Or _
             keyText = "NAMA")
    End If

End Function

Public Sub SchemaCleanupGhostFPTKFields()

    Dim wsConfig As Worksheet
    Dim wsData As Worksheet
    Dim physicalMap As Object

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim columnNumber As Long
    Dim removedCount As Long

    Dim fieldID As String
    Dim currentHeader As String
    Dim canonicalHeader As String
    Dim aliases As String

    Dim answer As VbMsgBoxResult

    answer = MsgBox( _
        "Bersihkan registry FPTK custom yang tidak memiliki " & _
        "kolom fisik pada header row " & _
        CStr(FPTK_SCHEMA_HEADER_ROW) & "?" & _
        vbCrLf & vbCrLf & _
        "Core field tidak akan dihapus.", _
        vbYesNo + vbQuestion, _
        "Cleanup Ghost Schema")

    If answer <> vbYes Then Exit Sub

    EnsureSchemaConfig

    Set wsConfig = _
        ThisWorkbook.Worksheets(SCHEMA_SHEET)

    Set wsData = _
        SchemaGetRequiredDataSheet("FPTK")

    Set physicalMap = _
        SchemaBuildPhysicalHeaderMap( _
            wsData, _
            "FPTK")

    lastRow = SchemaLastUsedRow(wsConfig)

    For rowIndex = lastRow _
        To SCHEMA_FIRST_ROW Step -1

        If StrComp( _
            Trim$(CStr(wsConfig.Cells( _
                rowIndex, _
                C_DATABASE).Value2)), _
            "FPTK", _
            vbTextCompare) = 0 Then

            If SchemaYes( _
                wsConfig.Cells( _
                    rowIndex, _
                    C_CUSTOM).Value2) Then

                fieldID = CStr( _
                    wsConfig.Cells( _
                        rowIndex, _
                        C_FIELD_ID).Value2)

                currentHeader = CStr( _
                    wsConfig.Cells( _
                        rowIndex, _
                        C_HEADER_NAME).Value2)

                canonicalHeader = CStr( _
                    wsConfig.Cells( _
                        rowIndex, _
                        C_CANONICAL_HEADER).Value2)

                aliases = CStr( _
                    wsConfig.Cells( _
                        rowIndex, _
                        C_ALIASES).Value2)

                columnNumber = _
                    SchemaColumnFromPhysicalMap( _
                        physicalMap, _
                        fieldID, _
                        currentHeader, _
                        canonicalHeader, _
                        aliases)

                If columnNumber = 0 Then

                    wsConfig.rows(rowIndex).Delete
                    removedCount = removedCount + 1

                End If

            End If

        End If

    Next rowIndex

    Set mSchemaCanonicalCache = Nothing

    SchemaFormatConfig wsConfig

    MsgBox _
        CStr(removedCount) & _
        " ghost field FPTK dibersihkan.", _
        vbInformation, _
        "Cleanup Ghost Schema"

    Set physicalMap = Nothing
    Set wsData = Nothing
    Set wsConfig = Nothing

End Sub

' ============================================================
' GENERAL HELPERS
' ============================================================
Private Function SchemaYes( _
    ByVal rawValue As Variant) As Boolean

    Dim textValue As String
    textValue = UCase$(Trim$(CStr(rawValue)))

    SchemaYes = _
        (textValue = "YES" Or textValue = "TRUE" Or _
         textValue = "1" Or textValue = "Y" Or _
         textValue = "YA")

End Function

Private Function SchemaYesNo( _
    ByVal boolValue As Boolean) As String

    If boolValue Then
        SchemaYesNo = "YES"
    Else
        SchemaYesNo = "NO"
    End If

End Function

Public Function SchemaSafeControlNamePublic( _
    ByVal fieldID As String) As String

    Dim resultText As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(fieldID)
        ch = Mid$(fieldID, i, 1)
        If ch Like "[A-Za-z0-9_]" Then resultText = resultText & ch
    Next i

    If Len(resultText) = 0 Then resultText = "FIELD"
    SchemaSafeControlNamePublic = resultText

End Function

Private Function SchemaGetOrCreateSheet( _
    ByVal sheetName As String) As Worksheet

    Dim normalizedName As String

    normalizedName = Trim$(sheetName)

    On Error Resume Next
    Set SchemaGetOrCreateSheet = _
        ThisWorkbook.Worksheets(normalizedName)
    On Error GoTo 0

    If Not SchemaGetOrCreateSheet Is Nothing Then
        Exit Function
    End If

    If Not SchemaIsSystemSheetName(normalizedName) Then

        Err.Raise vbObjectError + 8120, _
            "SchemaGetOrCreateSheet", _
            "Sheet database '" & normalizedName & _
            "' tidak ditemukan. Schema tidak membuat sheet database baru."

    End If

    Set SchemaGetOrCreateSheet = _
        ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
                ThisWorkbook.Worksheets.Count))

    SchemaGetOrCreateSheet.Name = normalizedName

End Function

Private Function SchemaGetRequiredDataSheet( _
    ByVal databaseKey As String) As Worksheet

    Dim canonicalKey As String
    Dim ws As Worksheet

    canonicalKey = SchemaNormalizeDatabaseKey(databaseKey)

    If Len(canonicalKey) = 0 Then

        Err.Raise vbObjectError + 8123, _
            "SchemaGetRequiredDataSheet", _
            "Database tidak dikenali: " & databaseKey

    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(canonicalKey)
    On Error GoTo 0

    If ws Is Nothing Then

        For Each ws In ThisWorkbook.Worksheets

            If StrComp( _
                Trim$(ws.Name), _
                canonicalKey, _
                vbTextCompare) = 0 Then

                Set SchemaGetRequiredDataSheet = ws
                Exit Function

            End If

        Next ws

        Err.Raise vbObjectError + 8124, _
            "SchemaGetRequiredDataSheet", _
            "Sheet '" & canonicalKey & _
            "' tidak ditemukan. Tidak ada sheet baru yang dibuat."

    End If

    Set SchemaGetRequiredDataSheet = ws

End Function

Private Function SchemaNormalizeDatabaseKey( _
    ByVal databaseKey As String) As String

    Dim keyText As String

    keyText = SchemaNormalize(databaseKey)

    Select Case keyText

        Case "FPTK"
            SchemaNormalizeDatabaseKey = "FPTK"

        Case "DBKODEPOSISI", "DBKODE", "KODEPOSISI"
            SchemaNormalizeDatabaseKey = "DB Kode Posisi"

        Case "DBSOURCING", "SOURCING"
            SchemaNormalizeDatabaseKey = "DB Sourcing"

    End Select

End Function

Private Function SchemaIsSystemSheetName( _
    ByVal sheetName As String) As Boolean

    Select Case UCase$(Trim$(sheetName))

        Case UCase$(SCHEMA_SHEET), _
             UCase$(SCHEMA_LOG_SHEET), _
             "SETTINGS"

            SchemaIsSystemSheetName = True

    End Select

End Function

Private Function SchemaSheetExists( _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SchemaSheetExists = Not ws Is Nothing

End Function

Private Function SchemaLastUsedRow( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range

    On Error Resume Next
    Set foundCell = ws.Cells.Find( _
        What:="*", After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False)
    On Error GoTo 0

    If foundCell Is Nothing Then
        SchemaLastUsedRow = 1
    Else
        SchemaLastUsedRow = foundCell.Row
    End If

End Function

Private Function SchemaLastHeaderColumn( _
    ByVal ws As Worksheet) As Long

    Dim foundCell As Range
    Dim headerRow As Long

    If ws Is Nothing Then
        SchemaLastHeaderColumn = 1
        Exit Function
    End If

    headerRow = SchemaHeaderRow(ws.Name)

    On Error Resume Next
    Set foundCell = ws.rows(headerRow).Find( _
        What:="*", _
        After:=ws.Cells(headerRow, 1), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)
    On Error GoTo 0

    If foundCell Is Nothing Then
        SchemaLastHeaderColumn = 1
    Else
        SchemaLastHeaderColumn = foundCell.Column
    End If

End Function

Private Sub SchemaStyleHeader( _
    ByVal ws As Worksheet)

    Dim lastColumn As Long
    Dim headerRow As Long

    If ws Is Nothing Then Exit Sub

    headerRow = SchemaHeaderRow(ws.Name)
    lastColumn = SchemaLastHeaderColumn(ws)

    With ws.Range( _
        ws.Cells(headerRow, 1), _
        ws.Cells(headerRow, lastColumn))

        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True

    End With

    If headerRow = 2 Then
        ws.rows(1).RowHeight = 42
        ws.rows(2).RowHeight = 36
    End If

End Sub

Private Sub SchemaFormatConfig( _
    ByVal ws As Worksheet)

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, C_CANONICAL_HEADER))
        .Font.Bold = True
        .Interior.Color = RGB(44, 62, 80)
        .Font.Color = vbWhite
    End With

    ws.Columns("A:R").AutoFit

End Sub

Private Sub SchemaUnlistAllTables( _
    ByVal ws As Worksheet)

    Dim tableIndex As Long

    For tableIndex = ws.ListObjects.Count To 1 Step -1
        On Error Resume Next
        ws.ListObjects(tableIndex).Unlist
        On Error GoTo 0
    Next tableIndex

End Sub

Private Sub SchemaSetColumnHiddenSafe( _
    ByVal ws As Worksheet, _
    ByVal columnNumber As Long, _
    ByVal hiddenState As Boolean)

    If ws Is Nothing Then Exit Sub
    If columnNumber <= 0 Then Exit Sub

    On Error Resume Next
    ws.Columns(columnNumber).EntireColumn.Hidden = hiddenState
    On Error GoTo 0

End Sub

Private Sub SchemaBackupSheet( _
    ByVal ws As Worksheet)

End Sub

Public Sub SchemaDeleteGeneratedBackupSheets()

    Dim ws As Worksheet
    Dim sheetIndex As Long
    Dim deletedCount As Long
    Dim answer As VbMsgBoxResult
    Dim oldAlerts As Boolean

    answer = MsgBox( _
        "Hapus seluruh sheet backup lama dengan prefix BK_?" & _
        vbCrLf & vbCrLf & _
        "Sheet database utama tidak akan dihapus.", _
        vbYesNo + vbQuestion, _
        "Cleanup Backup Sheet")

    If answer <> vbYes Then Exit Sub

    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error GoTo CleanupError

    For sheetIndex = ThisWorkbook.Worksheets.Count To 1 Step -1

        Set ws = ThisWorkbook.Worksheets(sheetIndex)

        If UCase$(Left$(ws.Name, 3)) = "BK_" Then

            ws.Visible = xlSheetVisible
            ws.Delete
            deletedCount = deletedCount + 1

        End If

    Next sheetIndex

    Application.DisplayAlerts = oldAlerts

    MsgBox _
        CStr(deletedCount) & _
        " sheet backup lama berhasil dihapus.", _
        vbInformation, _
        "Cleanup Backup Sheet"

    Exit Sub

CleanupError:

    Application.DisplayAlerts = oldAlerts

    MsgBox _
        "Cleanup backup sheet gagal." & vbCrLf & _
        "Error " & CStr(Err.Number) & _
        ": " & Err.Description, _
        vbExclamation, _
        "Cleanup Backup Sheet"

End Sub

Private Sub SchemaWriteChangeLog( _
    ByVal databaseKey As String, _
    ByVal actionName As String, _
    ByVal oldValue As String, _
    ByVal newValue As String, _
    ByVal fieldReference As String)

    Dim ws As Worksheet
    Dim NextRow As Long

    Set ws = SchemaGetOrCreateSheet(SCHEMA_LOG_SHEET)

    If Len(Trim$(CStr(ws.Cells(1, 1).Value2))) = 0 Then
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "User"
        ws.Cells(1, 3).Value = "Database"
        ws.Cells(1, 4).Value = "Action"
        ws.Cells(1, 5).Value = "Field"
        ws.Cells(1, 6).Value = "Old Value"
        ws.Cells(1, 7).Value = "New Value"
        ws.rows(1).Font.Bold = True
    End If

    NextRow = SchemaLastUsedRow(ws) + 1
    If NextRow < 2 Then NextRow = 2

    ws.Cells(NextRow, 1).Value = Now
    ws.Cells(NextRow, 2).Value = Environ$("Username")
    ws.Cells(NextRow, 3).Value = databaseKey
    ws.Cells(NextRow, 4).Value = actionName
    ws.Cells(NextRow, 5).Value = fieldReference
    ws.Cells(NextRow, 6).Value = oldValue
    ws.Cells(NextRow, 7).Value = newValue

End Sub









