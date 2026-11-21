Attribute VB_Name = "m15_ProgressManagement"
Option Explicit

Private mProgressStartedAt As Date
Private mLastVisualRefresh As Double
Private mLastVisualStage As String

Private Const PROGRESS_VISUAL_INTERVAL_SECONDS As Double = 1#

Public Sub InitProgress( _
    ByVal totalSteps As Long, _
    ByVal titleText As String, _
    ByVal showProgress As Boolean)

    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean

    gTotalSteps = totalSteps
    gCurrentStep = 0
    gShowProgress = showProgress

    mProgressStartedAt = Now
    mLastVisualRefresh = 0
    mLastVisualStage = vbNullString

    If Not gShowProgress Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Progress")
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
                ThisWorkbook.Worksheets.Count))

        ws.Name = "Progress"

    End If

    oldScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = True

    ThisWorkbook.Activate
    ws.Activate

    ws.Cells.Clear

    ws.Range("A1").Value = titleText
    ws.Range("A3").Value = "Status"
    ws.Range("A4").Value = "Progress"
    ws.Range("A5").Value = "Current File"
    ws.Range("A6").Value = "Current Stage"
    ws.Range("A7").Value = "Detail"
    ws.Range("A8").Value = "Started At"
    ws.Range("A9").Value = "Last Updated"
    ws.Range("A10").Value = "Lead Time"

    ws.Range("B3").Value = "Preparing..."
    ws.Range("B4").Value = "0%"
    ws.Range("B5").Value = "-"
    ws.Range("B6").Value = "Initialization"
    ws.Range("B7").Value = "-"
    ws.Range("B8").Value = mProgressStartedAt
    ws.Range("B9").Value = Now
    ws.Range("B10").Value = "00:00:00"

    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A3:A10").Font.Bold = True

    ws.Range("B8:B9").numberFormat = _
        "dd/mm/yyyy hh:mm:ss"

    ' Jangan AutoFit pada setiap heartbeat.
    ws.Columns("A").columnWidth = 20
    ws.Columns("B").columnWidth = 85

    ws.Range("B3:B10").WrapText = True
    ws.rows("3:10").AutoFit

    Application.DisplayStatusBar = True
    Application.StatusBar = _
        titleText & " - Preparing..."

    DoEvents

    Application.ScreenUpdating = oldScreenUpdating

End Sub

Public Sub UpdateProgress( _
    ByVal statusText As String, _
    ByVal currentFile As String)

    Dim pct As Double
    Dim progressText As String

    If Not gShowProgress Then Exit Sub

    gCurrentStep = gCurrentStep + 1

    If gTotalSteps <= 0 Then
        gTotalSteps = 1
    End If

    pct = gCurrentStep / gTotalSteps

    If pct > 1 Then
        pct = 1
    End If

    progressText = _
        Format$(pct, "0%") & _
        " (" & CStr(gCurrentStep) & _
        "/" & CStr(gTotalSteps) & ")"

    WriteProgressHeartbeat _
        statusText, _
        currentFile, _
        statusText, _
        "Step advanced", _
        progressText, _
        True

End Sub

Public Sub HeartbeatProgress( _
    ByVal statusText As String, _
    ByVal currentFile As String, _
    ByVal stageText As String, _
    ByVal detailText As String)

    Dim progressText As String
    Dim forceVisual As Boolean

    If Not gShowProgress Then Exit Sub

    If gTotalSteps <= 0 Then
        gTotalSteps = 1
    End If

    progressText = _
        Format$(gCurrentStep / gTotalSteps, "0%") & _
        " (" & CStr(gCurrentStep) & _
        "/" & CStr(gTotalSteps) & ")"

    forceVisual = _
        ProgressShouldRefreshVisual(stageText)

    WriteProgressHeartbeat _
        statusText, _
        currentFile, _
        stageText, _
        detailText, _
        progressText, _
        forceVisual

End Sub

Private Function ProgressShouldRefreshVisual( _
    ByVal stageText As String) As Boolean

    Dim currentTick As Double
    Dim elapsedTick As Double

    currentTick = Timer

    If StrComp( _
        stageText, _
        mLastVisualStage, _
        vbTextCompare) <> 0 Then

        ProgressShouldRefreshVisual = True
        Exit Function

    End If

    elapsedTick = currentTick - mLastVisualRefresh

    If elapsedTick < 0 Then
        elapsedTick = elapsedTick + 86400
    End If

    ProgressShouldRefreshVisual = _
        (elapsedTick >= _
            PROGRESS_VISUAL_INTERVAL_SECONDS)

End Function

Private Sub WriteProgressHeartbeat( _
    ByVal statusText As String, _
    ByVal currentFile As String, _
    ByVal stageText As String, _
    ByVal detailText As String, _
    ByVal progressText As String, _
    ByVal refreshWorkbookCells As Boolean)

    Dim ws As Worksheet
    Dim elapsedSeconds As Long
    Dim elapsedText As String

    If Not gShowProgress Then Exit Sub

    If mProgressStartedAt = 0 Then
        mProgressStartedAt = Now
    End If

    elapsedSeconds = DateDiff( _
        "s", _
        mProgressStartedAt, _
        Now)

    elapsedText = Format$( _
        TimeSerial( _
            elapsedSeconds \ 3600, _
            (elapsedSeconds Mod 3600) \ 60, _
            elapsedSeconds Mod 60), _
        "hh:nn:ss")

    ' Status bar selalu diperbarui karena murah.
    Application.StatusBar = _
        stageText & _
        " | " & detailText & _
        " | " & progressText & _
        " | " & currentFile

    If refreshWorkbookCells Then

        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("Progress")
        On Error GoTo 0

        If Not ws Is Nothing Then

            ws.Range("B3").Value = statusText
            ws.Range("B4").Value = progressText
            ws.Range("B5").Value = currentFile
            ws.Range("B6").Value = stageText
            ws.Range("B7").Value = detailText
            ws.Range("B9").Value = Now
            ws.Range("B10").Value = elapsedText

            ws.Range("B9").numberFormat = _
                "dd/mm/yyyy hh:mm:ss"

            ProgressForceWorkbookRepaint ws

            mLastVisualRefresh = Timer
            mLastVisualStage = stageText

        End If

    Else

        DoEvents

    End If

End Sub

Private Sub ProgressForceWorkbookRepaint( _
    ByVal ws As Worksheet)

    Dim oldScreenUpdating As Boolean

    If ws Is Nothing Then Exit Sub

    oldScreenUpdating = Application.ScreenUpdating

    On Error Resume Next

    Application.ScreenUpdating = True

    ThisWorkbook.Activate
    ws.Activate

    DoEvents

    Application.ScreenUpdating = oldScreenUpdating

    On Error GoTo 0

End Sub

Public Sub FinishProgress( _
    ByVal finalStatus As String)

    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean

    If gShowProgress Then

        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("Progress")
        On Error GoTo 0

        If Not ws Is Nothing Then

            ws.Range("B3").Value = finalStatus
            ws.Range("B4").Value = _
                "100% (" & _
                CStr(gTotalSteps) & "/" & _
                CStr(gTotalSteps) & ")"

            ws.Range("B6").Value = "Finished"
            ws.Range("B7").Value = finalStatus
            ws.Range("B9").Value = Now
            ws.Range("B9").numberFormat = _
                "dd/mm/yyyy hh:mm:ss"

            oldScreenUpdating = _
                Application.ScreenUpdating

            Application.ScreenUpdating = True

            ThisWorkbook.Activate
            ws.Activate

            DoEvents

            Application.ScreenUpdating = _
                oldScreenUpdating

        End If

        Application.StatusBar = finalStatus
        DoEvents

    End If

    Application.StatusBar = False

End Sub


