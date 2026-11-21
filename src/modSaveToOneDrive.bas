Attribute VB_Name = "modSaveToOneDrive"
Option Explicit

#If VBA7 Then

    Private Declare PtrSafe Function CreateFileW _
        Lib "kernel32" ( _
        ByVal lpFileName As LongPtr, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        ByVal lpSecurityAttributes As LongPtr, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As LongPtr) As LongPtr

    Private Declare PtrSafe Function WriteFile _
        Lib "kernel32" ( _
        ByVal hFile As LongPtr, _
        ByRef lpBuffer As Any, _
        ByVal nNumberOfBytesToWrite As Long, _
        ByRef lpNumberOfBytesWritten As Long, _
        ByVal lpOverlapped As LongPtr) As Long

    Private Declare PtrSafe Function SetFilePointer _
        Lib "kernel32" ( _
        ByVal hFile As LongPtr, _
        ByVal lDistanceToMove As Long, _
        ByVal lpDistanceToMoveHigh As LongPtr, _
        ByVal dwMoveMethod As Long) As Long

    Private Declare PtrSafe Function SetEndOfFile _
        Lib "kernel32" ( _
        ByVal hFile As LongPtr) As Long

    Private Declare PtrSafe Function FlushFileBuffers _
        Lib "kernel32" ( _
        ByVal hFile As LongPtr) As Long

    Private Declare PtrSafe Function CloseHandle _
        Lib "kernel32" ( _
        ByVal hObject As LongPtr) As Long

#Else

    Private Declare Function CreateFileW _
        Lib "kernel32" ( _
        ByVal lpFileName As Long, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        ByVal lpSecurityAttributes As Long, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long

    Private Declare Function WriteFile _
        Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByRef lpBuffer As Any, _
        ByVal nNumberOfBytesToWrite As Long, _
        ByRef lpNumberOfBytesWritten As Long, _
        ByVal lpOverlapped As Long) As Long

    Private Declare Function SetFilePointer _
        Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVal lDistanceToMove As Long, _
        ByVal lpDistanceToMoveHigh As Long, _
        ByVal dwMoveMethod As Long) As Long

    Private Declare Function SetEndOfFile _
        Lib "kernel32" ( _
        ByVal hFile As Long) As Long

    Private Declare Function FlushFileBuffers _
        Lib "kernel32" ( _
        ByVal hFile As Long) As Long

    Private Declare Function CloseHandle _
        Lib "kernel32" ( _
        ByVal hObject As Long) As Long

#End If

' ============================================================
' WINDOWS CONSTANTS
' ============================================================

Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const FILE_SHARE_DELETE As Long = &H4

Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_BEGIN As Long = 0

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const INVALID_SET_FILE_POINTER As Long = -1

Private Const FILE_BUFFER_SIZE As Long = 1048576


Public Sub SaveToOneDrive()

    Dim originalPath As String
    Dim outputSetting As String
    Dim outputPath As String
    Dim outputFolder As String

    Dim sourceExtension As String
    Dim tempSourcePath As String
    Dim tempXlsxPath As String

    Dim exportWorkbook As Workbook
    Dim fso As Object

    Dim outputAlreadyExists As Boolean

    Dim oldDisplayAlerts As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldAutomationSecurity As Long
    Dim oldCalculation As XlCalculation
    Dim oldStatusBar As Variant

    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ErrorHandler

    ' ========================================================
    ' 1. SIMPAN KONDISI APPLICATION
    ' ========================================================
    oldDisplayAlerts = Application.DisplayAlerts
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldAutomationSecurity = Application.AutomationSecurity
    oldCalculation = Application.Calculation
    oldStatusBar = Application.StatusBar

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Mempersiapkan workbook..."

    ' Tidak mematikan calculation di sini agar formula
    ' terakhir tetap tersimpan sesuai hasil compile.
    ' Application.Calculation tetap disimpan untuk restore.

    ' ========================================================
    ' 2. PERSIAPKAN SCHEMA
    ' ========================================================
    SchemaB02PrepareAll

    ' ========================================================
    ' 3. PASTIKAN WORKBOOK UTAMA SUDAH PERNAH DISIMPAN
    ' ========================================================
    If Len(Trim$(ThisWorkbook.Path)) = 0 Then

        RestoreApplicationState _
            oldDisplayAlerts, _
            oldScreenUpdating, _
            oldEnableEvents, _
            oldAutomationSecurity, _
            oldCalculation, _
            oldStatusBar

        MsgBox _
            "Workbook belum memiliki lokasi penyimpanan." & _
            vbCrLf & _
            "Simpan workbook terlebih dahulu menggunakan Save As.", _
            vbExclamation, _
            "Save Workbook"

        Exit Sub

    End If

    originalPath = ThisWorkbook.FullName

    ' ========================================================
    ' 4. SIMPAN WORKBOOK UTAMA
    ' ========================================================
    Application.StatusBar = _
        "Menyimpan workbook utama..."

    ThisWorkbook.Save

    DoEvents

    ' ========================================================
    ' 5. BACA OUTPUT FILE PATH DARI SETTINGS
    ' ========================================================
    outputSetting = Trim$( _
        ReadSettingValue("OutputFilePath"))

    If Len(outputSetting) = 0 Then

        RestoreApplicationState _
            oldDisplayAlerts, _
            oldScreenUpdating, _
            oldEnableEvents, _
            oldAutomationSecurity, _
            oldCalculation, _
            oldStatusBar

        MsgBox _
            "Workbook utama berhasil disimpan." & _
            vbCrLf & vbCrLf & _
            "Namun OutputFilePath belum diisi " & _
            "pada sheet Settings.", _
            vbExclamation, _
            "OutputFilePath Belum Diisi"

        Exit Sub

    End If

    ' ========================================================
    ' 6. OUTPUT HARUS FOLDER LOKAL / ONEDRIVE SYNC LOKAL
    ' ========================================================
    If IsWebPath(outputSetting) Then

        Err.Raise _
            vbObjectError + 1001, _
            "SaveToOneDrive", _
            "OutputFilePath tidak boleh berupa URL." & _
            vbCrLf & _
            "Gunakan path folder OneDrive lokal di komputer."

    End If

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    ' ========================================================
    ' 7. BENTUK OUTPUT PATH XLSX
    '
    ' Setting berupa full file:
    ' D:\Output\Master - Database FIX.xlsx
    '
    ' Setting berupa folder:
    ' D:\Output
    '
    ' Jika berupa folder, nama workbook utama akan dipakai,
    ' tetapi extension otomatis menjadi XLSX.
    ' ========================================================
    outputPath = BuildOutputXlsxPath( _
        outputSetting)

    outputFolder = fso.GetParentFolderName( _
        outputPath)

    If Len(Trim$(outputFolder)) = 0 Then

        Err.Raise _
            vbObjectError + 1002, _
            "SaveToOneDrive", _
            "Folder output tidak valid:" & _
            vbCrLf & outputSetting

    End If

    ' ========================================================
    ' 8. BUAT FOLDER OUTPUT JIKA BELUM ADA
    ' ========================================================
    EnsureFolderExists outputFolder

    ' ========================================================
    ' 9. OUTPUT TIDAK BOLEH SAMA DENGAN WORKBOOK UTAMA
    ' ========================================================
    If SameFilePath(originalPath, outputPath) Then

        Err.Raise _
            vbObjectError + 1003, _
            "SaveToOneDrive", _
            "Lokasi file output sama dengan workbook utama."

    End If

    ' ========================================================
    ' 10. CEK OUTPUT SEDANG TERBUKA DI EXCEL INI
    ' ========================================================
    If IsWorkbookOpenByPath(outputPath) Then

        Err.Raise _
            vbObjectError + 1004, _
            "SaveToOneDrive", _
            "File output sedang terbuka." & _
            vbCrLf & _
            "Tutup file berikut terlebih dahulu:" & _
            vbCrLf & outputPath

    End If

    outputAlreadyExists = fso.FileExists( _
        outputPath)

    ' ========================================================
    ' 11. CEK FILE EXISTING TIDAK READ-ONLY
    ' ========================================================
    If outputAlreadyExists Then

        If IsFileReadOnly(outputPath) Then

            Err.Raise _
                vbObjectError + 1005, _
                "SaveToOneDrive", _
                "File output berstatus Read-Only." & _
                vbCrLf & _
                "File tidak dapat diperbarui:" & _
                vbCrLf & outputPath

        End If

    End If

    ' ========================================================
    ' 12. AMBIL EXTENSION WORKBOOK SUMBER
    '
    ' Temporary source harus memiliki extension yang sama
    ' dengan workbook utama agar SaveCopyAs valid.
    ' ========================================================
    sourceExtension = GetFileExtension( _
        ThisWorkbook.Name)

    If Len(sourceExtension) = 0 Then

        Err.Raise _
            vbObjectError + 1006, _
            "SaveToOneDrive", _
            "Extension workbook utama tidak ditemukan."

    End If

    ' ========================================================
    ' 13. BUAT DUA FILE TEMPORARY
    '
    ' tempSourcePath:
    ' Salinan workbook dengan format asli.
    '
    ' tempXlsxPath:
    ' Hasil konversi XLSX tanpa macro.
    ' ========================================================
    tempSourcePath = BuildTemporaryPath( _
        "Source", _
        sourceExtension)

    tempXlsxPath = BuildTemporaryPath( _
        "Output", _
        ".xlsx")

    SafeDeleteTemporaryFile tempSourcePath
    SafeDeleteTemporaryFile tempXlsxPath

    ' ========================================================
    ' 14. BUAT SALINAN TEMPORARY DENGAN FORMAT ASLI
    ' ========================================================
    Application.StatusBar = _
        "Membuat salinan sementara workbook..."

    ThisWorkbook.SaveCopyAs tempSourcePath

    DoEvents

    If Not fso.FileExists(tempSourcePath) Then

        Err.Raise _
            vbObjectError + 1007, _
            "SaveToOneDrive", _
            "Salinan sementara workbook gagal dibuat."

    End If

    If fso.GetFile(tempSourcePath).Size <= 0 Then

        Err.Raise _
            vbObjectError + 1008, _
            "SaveToOneDrive", _
            "Salinan sementara workbook berukuran 0 KB."

    End If

    ' ========================================================
    ' 15. MATIKAN SEMUA MACRO SAAT TEMPORARY DIBUKA
    '
    ' Nilai 3:
    ' msoAutomationSecurityForceDisable
    ' ========================================================
    Application.AutomationSecurity = 3
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Application.StatusBar = _
        "Membuka salinan dengan macro dinonaktifkan..."

    Set exportWorkbook = Workbooks.Open( _
        fileName:=tempSourcePath, _
        UpdateLinks:=0, _
        ReadOnly:=False, _
        AddToMru:=False, _
        Notify:=False)

    ' ========================================================
    ' 16. KONVERSI KE XLSX DI FOLDER TEMPORARY
    '
    ' BUKAN menyimpan langsung ke output existing.
    ' Dengan demikian file SharePoint existing belum disentuh.
    ' ========================================================
    Application.StatusBar = _
        "Mengonversi salinan menjadi XLSX tanpa macro..."

    exportWorkbook.CheckCompatibility = False

    exportWorkbook.SaveAs _
        fileName:=tempXlsxPath, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False, _
        Local:=True

    DoEvents

    ' ========================================================
    ' 17. TUTUP WORKBOOK TEMPORARY
    ' ========================================================
    exportWorkbook.Close SaveChanges:=False
    Set exportWorkbook = Nothing

    If Not fso.FileExists(tempXlsxPath) Then

        Err.Raise _
            vbObjectError + 1009, _
            "SaveToOneDrive", _
            "File XLSX temporary gagal dibuat."

    End If

    If fso.GetFile(tempXlsxPath).Size <= 0 Then

        Err.Raise _
            vbObjectError + 1010, _
            "SaveToOneDrive", _
            "File XLSX temporary berukuran 0 KB."

    End If

    If outputAlreadyExists Then

        Application.StatusBar = _
            "Memperbarui isi file existing..."

        ReplaceExistingFileContent _
            tempXlsxPath, _
            outputPath

    Else

        Application.StatusBar = _
            "Membuat file output untuk pertama kali..."

        FileCopy tempXlsxPath, outputPath

    End If

    DoEvents

    ' ========================================================
    ' 19. VALIDASI FILE OUTPUT
    ' ========================================================
    If Not fso.FileExists(outputPath) Then

        Err.Raise _
            vbObjectError + 1011, _
            "SaveToOneDrive", _
            "File XLSX output tidak ditemukan setelah proses."

    End If

    If fso.GetFile(outputPath).Size <= 0 Then

        Err.Raise _
            vbObjectError + 1012, _
            "SaveToOneDrive", _
            "File XLSX output berukuran 0 KB."

    End If

    ' ========================================================
    ' 20. HAPUS HANYA FILE TEMPORARY
    ' ========================================================
    SafeDeleteTemporaryFile tempSourcePath
    SafeDeleteTemporaryFile tempXlsxPath

    ' ========================================================
    ' 21. KEMBALIKAN KONDISI EXCEL
    ' ========================================================
    RestoreApplicationState _
        oldDisplayAlerts, _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldAutomationSecurity, _
        oldCalculation, _
        oldStatusBar

    ' ========================================================
    ' 22. TAMPILKAN HASIL
    ' ========================================================
    If outputAlreadyExists Then

        MsgBox _
            "Workbook berhasil disimpan." & _
            vbCrLf & vbCrLf & _
            "Workbook utama:" & _
            vbCrLf & originalPath & _
            vbCrLf & vbCrLf & _
            "File XLSX existing berhasil diperbarui:" & _
            vbCrLf & outputPath & _
            vbCrLf & vbCrLf & _
            "Tunggu sinkronisasi OneDrive selesai.", _
            vbInformation, _
            "Save Berhasil"

    Else

        MsgBox _
            "Workbook berhasil disimpan." & _
            vbCrLf & vbCrLf & _
            "Workbook utama:" & _
            vbCrLf & originalPath & _
            vbCrLf & vbCrLf & _
            "File XLSX baru berhasil dibuat:" & _
            vbCrLf & outputPath & _
            vbCrLf & vbCrLf & _
            "Tunggu sinkronisasi OneDrive selesai.", _
            vbInformation, _
            "Save Berhasil"

    End If

    Exit Sub

ErrorHandler:

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    ' Tutup workbook temporary apabila masih terbuka.
    If Not exportWorkbook Is Nothing Then
        exportWorkbook.Close SaveChanges:=False
    End If

    Set exportWorkbook = Nothing

    ' Hanya file temporary yang dibersihkan.
    SafeDeleteTemporaryFile tempSourcePath
    SafeDeleteTemporaryFile tempXlsxPath

    RestoreApplicationState _
        oldDisplayAlerts, _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldAutomationSecurity, _
        oldCalculation, _
        oldStatusBar

    On Error GoTo 0

    MsgBox _
        "Gagal membuat atau memperbarui file XLSX." & _
        vbCrLf & vbCrLf & _
        "Error " & errorNumber & ":" & _
        vbCrLf & errorDescription & _
        vbCrLf & vbCrLf & _
        "File output existing tidak dihapus." & _
        vbCrLf & _
        "Pastikan file output tidak sedang terbuka " & _
        "dan sinkronisasi OneDrive tidak terkunci.", _
        vbCritical, _
        "Save Gagal"

End Sub

' ============================================================
' UPDATE ISI FILE EXISTING SECARA IN-PLACE
'
' PENTING:
' Fungsi ini tidak menghapus file destination.
' Fungsi ini membuka file destination dengan OPEN_EXISTING.
'
' Proses:
' 1. Buka temporary XLSX untuk dibaca.
' 2. Buka destination existing menggunakan Windows API.
' 3. Pindahkan posisi tulis ke byte pertama.
' 4. Tulis seluruh isi temporary ke destination.
' 5. SetEndOfFile memotong sisa isi lama.
' 6. FlushFileBuffers memastikan data ditulis ke disk.
' ============================================================
Private Sub ReplaceExistingFileContent( _
    ByVal sourceFile As String, _
    ByVal destinationFile As String)

    Dim sourceFileNumber As Integer
    Dim sourceOpened As Boolean

    Dim sourceSize As Long
    Dim remainingBytes As Long
    Dim bytesToRead As Long
    Dim bytesWritten As Long

    Dim fileBuffer() As Byte
    Dim apiResult As Long
    Dim apiError As Long

    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

#If VBA7 Then

    Dim destinationHandle As LongPtr

#Else

    Dim destinationHandle As Long

#End If

    Dim destinationOpened As Boolean

    On Error GoTo ErrorHandler

    If Len(Dir$(sourceFile)) = 0 Then

        Err.Raise _
            vbObjectError + 1200, _
            "ReplaceExistingFileContent", _
            "File sumber temporary tidak ditemukan:" & _
            vbCrLf & sourceFile

    End If

    If Len(Dir$(destinationFile)) = 0 Then

        Err.Raise _
            vbObjectError + 1201, _
            "ReplaceExistingFileContent", _
            "File destination existing tidak ditemukan:" & _
            vbCrLf & destinationFile

    End If

    sourceSize = FileLen(sourceFile)

    If sourceSize <= 0 Then

        Err.Raise _
            vbObjectError + 1202, _
            "ReplaceExistingFileContent", _
            "File sumber temporary berukuran 0 KB."

    End If

    ' ========================================================
    ' BUKA FILE SUMBER UNTUK DIBACA
    ' ========================================================
    sourceFileNumber = FreeFile

    Open sourceFile _
        For Binary Access Read _
        As #sourceFileNumber

    sourceOpened = True

    ' ========================================================
    ' BUKA FILE DESTINATION EXISTING
    '
    ' OPEN_EXISTING memastikan file yang sama dibuka.
    ' Tidak membuat file baru dan tidak menghapus file lama.
    ' ========================================================
    destinationHandle = CreateFileW( _
        StrPtr(destinationFile), _
        GENERIC_WRITE, _
        FILE_SHARE_READ Or _
            FILE_SHARE_WRITE Or _
            FILE_SHARE_DELETE, _
        0, _
        OPEN_EXISTING, _
        FILE_ATTRIBUTE_NORMAL, _
        0)

    If destinationHandle = INVALID_HANDLE_VALUE Then

        apiError = Err.LastDllError

        Err.Raise _
            vbObjectError + 1203, _
            "ReplaceExistingFileContent", _
            "File destination tidak dapat dibuka untuk update." & _
            vbCrLf & _
            "Kemungkinan file sedang terbuka atau dikunci." & _
            vbCrLf & vbCrLf & _
            destinationFile & _
            vbCrLf & vbCrLf & _
            "Windows error: " & CStr(apiError)

    End If

    destinationOpened = True

    ' ========================================================
    ' PINDAHKAN POSISI TULIS KE AWAL FILE
    ' ========================================================
    apiResult = SetFilePointer( _
        destinationHandle, _
        0, _
        0, _
        FILE_BEGIN)

    If apiResult = INVALID_SET_FILE_POINTER Then

        apiError = Err.LastDllError

        Err.Raise _
            vbObjectError + 1204, _
            "ReplaceExistingFileContent", _
            "Gagal mengatur posisi tulis file destination." & _
            vbCrLf & _
            "Windows error: " & CStr(apiError)

    End If

    ' ========================================================
    ' TULIS FILE DALAM BUFFER 1 MB
    ' ========================================================
    remainingBytes = sourceSize

    Do While remainingBytes > 0

        If remainingBytes > FILE_BUFFER_SIZE Then

            bytesToRead = FILE_BUFFER_SIZE

        Else

            bytesToRead = remainingBytes

        End If

        ReDim fileBuffer(0 To bytesToRead - 1)

        Get #sourceFileNumber, , fileBuffer

        bytesWritten = 0

        apiResult = WriteFile( _
            destinationHandle, _
            fileBuffer(0), _
            bytesToRead, _
            bytesWritten, _
            0)

        If apiResult = 0 Then

            apiError = Err.LastDllError

            Err.Raise _
                vbObjectError + 1205, _
                "ReplaceExistingFileContent", _
                "Gagal menulis isi baru ke file existing." & _
                vbCrLf & _
                "Windows error: " & CStr(apiError)

        End If

        If bytesWritten <> bytesToRead Then

            Err.Raise _
                vbObjectError + 1206, _
                "ReplaceExistingFileContent", _
                "Jumlah byte yang ditulis tidak lengkap." & _
                vbCrLf & _
                "Seharusnya: " & CStr(bytesToRead) & _
                vbCrLf & _
                "Tertulis: " & CStr(bytesWritten)

        End If

        remainingBytes = _
            remainingBytes - bytesToRead

        DoEvents

    Loop

    ' ========================================================
    ' POTONG SISA ISI FILE LAMA
    '
    ' Diperlukan jika ukuran file baru lebih kecil
    ' daripada ukuran file existing sebelumnya.
    ' ========================================================
    apiResult = SetEndOfFile( _
        destinationHandle)

    If apiResult = 0 Then

        apiError = Err.LastDllError

        Err.Raise _
            vbObjectError + 1207, _
            "ReplaceExistingFileContent", _
            "Gagal memotong sisa data file lama." & _
            vbCrLf & _
            "Windows error: " & CStr(apiError)

    End If

    ' ========================================================
    ' PAKSA DATA DITULIS KE DISK
    ' ========================================================
    apiResult = FlushFileBuffers( _
        destinationHandle)

    If apiResult = 0 Then

        apiError = Err.LastDllError

        Err.Raise _
            vbObjectError + 1208, _
            "ReplaceExistingFileContent", _
            "Gagal menyelesaikan penulisan file existing." & _
            vbCrLf & _
            "Windows error: " & CStr(apiError)

    End If

CleanExit:

    On Error Resume Next

    If sourceOpened Then
        Close #sourceFileNumber
    End If

    If destinationOpened Then
        CloseHandle destinationHandle
    End If

    On Error GoTo 0

    Exit Sub

ErrorHandler:

    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description

    On Error Resume Next

    If sourceOpened Then
        Close #sourceFileNumber
    End If

    If destinationOpened Then
        CloseHandle destinationHandle
    End If

    On Error GoTo 0

    Err.Raise _
        savedErrorNumber, _
        "ReplaceExistingFileContent", _
        savedErrorDescription

End Sub

' ============================================================
' BENTUK OUTPUT PATH MENJADI XLSX
' ============================================================
Private Function BuildOutputXlsxPath( _
    ByVal outputSetting As String) As String

    Dim fso As Object
    Dim outputFolder As String
    Dim outputFileName As String

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    outputSetting = Trim$(outputSetting)

    If IsExcelFilePath(outputSetting) Then

        outputFolder = fso.GetParentFolderName( _
            outputSetting)

        outputFileName = fso.GetBaseName( _
            outputSetting) & ".xlsx"

        BuildOutputXlsxPath = fso.BuildPath( _
            outputFolder, _
            outputFileName)

    Else

        outputFolder = RemoveTrailingSlash( _
            outputSetting)

        outputFileName = fso.GetBaseName( _
            ThisWorkbook.Name) & ".xlsx"

        BuildOutputXlsxPath = fso.BuildPath( _
            outputFolder, _
            outputFileName)

    End If

End Function

' ============================================================
' BUAT PATH FILE TEMPORARY
' ============================================================
Private Function BuildTemporaryPath( _
    ByVal filePrefix As String, _
    ByVal fileExtension As String) As String

    Dim tempFolder As String
    Dim tempFileName As String
    Dim uniqueNumber As Long

    tempFolder = Environ$("TEMP")

    If Len(Trim$(tempFolder)) = 0 Then
        tempFolder = ThisWorkbook.Path
    End If

    If Left$(fileExtension, 1) <> "." Then
        fileExtension = "." & fileExtension
    End If

    uniqueNumber = CLng( _
        Timer * 1000)

    tempFileName = _
        "~Database_" & _
        filePrefix & "_" & _
        Format$(Now, "yyyymmdd_hhnnss") & _
        "_" & _
        Format$(uniqueNumber, "00000000") & _
        fileExtension

    BuildTemporaryPath = _
        tempFolder & _
        Application.PathSeparator & _
        tempFileName

End Function

' ============================================================
' HAPUS HANYA FILE TEMPORARY
' ============================================================
Private Sub SafeDeleteTemporaryFile( _
    ByVal temporaryPath As String)

    On Error Resume Next

    If Len(Trim$(temporaryPath)) > 0 Then

        If Len(Dir$(temporaryPath)) > 0 Then
            Kill temporaryPath
        End If

    End If

    On Error GoTo 0

End Sub

' ============================================================
' CEK FILE BERSTATUS READ-ONLY
' ============================================================
Private Function IsFileReadOnly( _
    ByVal filePath As String) As Boolean

    Dim fileAttributes As VbFileAttribute

    On Error GoTo ErrorHandler

    fileAttributes = GetAttr(filePath)

    IsFileReadOnly = _
        ((fileAttributes And vbReadOnly) = vbReadOnly)

    Exit Function

ErrorHandler:

    IsFileReadOnly = False

End Function

' ============================================================
' AMBIL EXTENSION FILE BESERTA TITIK
'
' Contoh:
' Master Database.xlsm
'
' Hasil:
' .xlsm
' ============================================================
Private Function GetFileExtension( _
    ByVal fileName As String) As String

    Dim dotPosition As Long

    dotPosition = InStrRev( _
        fileName, ".")

    If dotPosition > 0 Then

        GetFileExtension = LCase$( _
            Mid$(fileName, dotPosition))

    End If

End Function

' ============================================================
' CEK APAKAH SETTING BERUPA FULL FILE PATH EXCEL
' ============================================================
Private Function IsExcelFilePath( _
    ByVal pathValue As String) As Boolean

    Dim fso As Object
    Dim extensionName As String

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    pathValue = Trim$(pathValue)

    extensionName = LCase$( _
        fso.GetExtensionName(pathValue))

    Select Case extensionName

        Case "xls", "xlsx", "xlsm", "xlsb"
            IsExcelFilePath = True

        Case Else
            IsExcelFilePath = False

    End Select

End Function

' ============================================================
' CEK APAKAH PATH BERUPA URL WEB
' ============================================================
Private Function IsWebPath( _
    ByVal pathValue As String) As Boolean

    Dim normalizedPath As String

    normalizedPath = LCase$( _
        Trim$(pathValue))

    If Left$(normalizedPath, 7) = "http://" Then

        IsWebPath = True
        Exit Function

    End If

    If Left$(normalizedPath, 8) = "https://" Then
        IsWebPath = True
    End If

End Function

' ============================================================
' CEK APAKAH WORKBOOK OUTPUT SEDANG TERBUKA
' DALAM INSTANCE EXCEL YANG SAMA
' ============================================================
Private Function IsWorkbookOpenByPath( _
    ByVal filePath As String) As Boolean

    Dim currentWorkbook As Workbook

    For Each currentWorkbook In Application.Workbooks

        If SameFilePath( _
            currentWorkbook.FullName, _
            filePath) Then

            IsWorkbookOpenByPath = True
            Exit Function

        End If

    Next currentWorkbook

End Function

' ============================================================
' CEK APAKAH PATH BERADA DI ONEDRIVE
' ============================================================
Private Function IsOneDrivePath( _
    ByVal filePath As String) As Boolean

    Dim normalizedPath As String
    Dim oneDrivePath As String
    Dim oneDriveCommercialPath As String
    Dim oneDriveConsumerPath As String

    normalizedPath = LCase$( _
        Trim$(filePath))

    If IsWebPath(normalizedPath) Then

        If InStr( _
            1, _
            normalizedPath, _
            ".sharepoint.com", _
            vbTextCompare) > 0 Then

            IsOneDrivePath = True
            Exit Function

        End If

        If InStr( _
            1, _
            normalizedPath, _
            "onedrive.live.com", _
            vbTextCompare) > 0 Then

            IsOneDrivePath = True
            Exit Function

        End If

    End If

    oneDrivePath = Environ$("OneDrive")

    oneDriveCommercialPath = _
        Environ$("OneDriveCommercial")

    oneDriveConsumerPath = _
        Environ$("OneDriveConsumer")

    If Len(oneDrivePath) > 0 Then

        If InStr( _
            1, _
            filePath, _
            oneDrivePath, _
            vbTextCompare) = 1 Then

            IsOneDrivePath = True
            Exit Function

        End If

    End If

    If Len(oneDriveCommercialPath) > 0 Then

        If InStr( _
            1, _
            filePath, _
            oneDriveCommercialPath, _
            vbTextCompare) = 1 Then

            IsOneDrivePath = True
            Exit Function

        End If

    End If

    If Len(oneDriveConsumerPath) > 0 Then

        If InStr( _
            1, _
            filePath, _
            oneDriveConsumerPath, _
            vbTextCompare) = 1 Then

            IsOneDrivePath = True
            Exit Function

        End If

    End If

    If InStr( _
        1, _
        filePath, _
        "OneDrive", _
        vbTextCompare) > 0 Then

        IsOneDrivePath = True

    End If

End Function

' ============================================================
' BUAT FOLDER TERMASUK PARENT FOLDER
' ============================================================
Private Sub EnsureFolderExists( _
    ByVal folderPath As String)

    Dim fso As Object
    Dim parentFolder As String

    folderPath = RemoveTrailingSlash( _
        folderPath)

    If Len(folderPath) = 0 Then

        Err.Raise _
            vbObjectError + 1300, _
            "EnsureFolderExists", _
            "Folder path kosong."

    End If

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        Exit Sub
    End If

    parentFolder = fso.GetParentFolderName( _
        folderPath)

    If Len(parentFolder) > 0 Then

        If Not fso.FolderExists(parentFolder) Then
            EnsureFolderExists parentFolder
        End If

    End If

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

End Sub

' ============================================================
' HAPUS SLASH TERAKHIR DARI FOLDER
'
' Root drive seperti C:\ tidak diubah menjadi C:
' ============================================================
Private Function RemoveTrailingSlash( _
    ByVal folderPath As String) As String

    folderPath = Trim$(folderPath)

    Do While Len(folderPath) > 0

        If Len(folderPath) = 3 Then

            If Mid$(folderPath, 2, 1) = ":" Then
                Exit Do
            End If

        End If

        If Right$(folderPath, 1) <> "\" And _
           Right$(folderPath, 1) <> "/" Then

            Exit Do

        End If

        folderPath = Left$( _
            folderPath, _
            Len(folderPath) - 1)

    Loop

    RemoveTrailingSlash = folderPath

End Function

' ============================================================
' CEK APAKAH DUA PATH SAMA
' ============================================================
Private Function SameFilePath( _
    ByVal firstPath As String, _
    ByVal secondPath As String) As Boolean

    Dim fso As Object
    Dim normalizedFirst As String
    Dim normalizedSecond As String

    firstPath = Trim$(firstPath)
    secondPath = Trim$(secondPath)

    Set fso = CreateObject( _
        "Scripting.FileSystemObject")

    On Error Resume Next

    normalizedFirst = _
        fso.GetAbsolutePathName(firstPath)

    normalizedSecond = _
        fso.GetAbsolutePathName(secondPath)

    On Error GoTo 0

    If Len(normalizedFirst) = 0 Then

        normalizedFirst = Replace( _
            firstPath, _
            "/", _
            "\")

    End If

    If Len(normalizedSecond) = 0 Then

        normalizedSecond = Replace( _
            secondPath, _
            "/", _
            "\")

    End If

    SameFilePath = _
        (StrComp( _
            normalizedFirst, _
            normalizedSecond, _
            vbTextCompare) = 0)

End Function

' ============================================================
' KEMBALIKAN KONDISI APPLICATION
' ============================================================
Private Sub RestoreApplicationState( _
    ByVal oldDisplayAlerts As Boolean, _
    ByVal oldScreenUpdating As Boolean, _
    ByVal oldEnableEvents As Boolean, _
    ByVal oldAutomationSecurity As Long, _
    ByVal oldCalculation As XlCalculation, _
    ByVal oldStatusBar As Variant)

    On Error Resume Next

    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.AutomationSecurity = oldAutomationSecurity
    Application.Calculation = oldCalculation
    Application.StatusBar = oldStatusBar

    On Error GoTo 0

End Sub

' ============================================================
' READ SETTING VALUE
'
' Struktur sheet Settings:
'
' Kolom A = Nama setting
' Kolom B = Nilai setting
'
' Contoh full path:
'
' A2 = OutputFilePath
' B2 = D:\Output Recruitment\Master - Database FIX.xlsx
'
' Contoh folder:
'
' A2 = OutputFilePath
' B2 = D:\Output Recruitment
' ============================================================
Private Function ReadSettingValue( _
    ByVal settingName As String) As String

    Dim wsSettings As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim currentSettingName As String

    On Error Resume Next

    Set wsSettings = _
        ThisWorkbook.Worksheets("Settings")

    On Error GoTo 0

    If wsSettings Is Nothing Then
        Exit Function
    End If

    lastRow = wsSettings.Cells( _
        wsSettings.rows.Count, _
        "A").End(xlUp).Row

    For rowIndex = 1 To lastRow

        currentSettingName = Trim$(CStr( _
            wsSettings.Cells( _
                rowIndex, _
                "A").Value2))

        If StrComp( _
            currentSettingName, _
            settingName, _
            vbTextCompare) = 0 Then

            ReadSettingValue = Trim$(CStr( _
                wsSettings.Cells( _
                    rowIndex, _
                    "B").Value2))

            Exit Function

        End If

    Next rowIndex

End Function

