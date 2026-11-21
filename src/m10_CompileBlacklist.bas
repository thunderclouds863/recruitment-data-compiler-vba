Attribute VB_Name = "m10_CompileBlacklist"
Option Explicit

Public Sub Compile_Blacklist_Candidate(folderPath As String)
    Dim wsTarget As Worksheet
    Set wsTarget = GetOrCreateWorksheet("Blacklist Candidate")
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    LoadExistingKeysBlacklist wsTarget, dict
    
    Dim extension As Variant
    extension = Array("*.xlsx", "*.xlsm")
    Dim ext As Variant
    
    For Each ext In extension
        Dim fileName As String
        fileName = Dir(folderPath & ext)
        Do While fileName <> ""
            If Left(fileName, 2) <> "~$" Then
                UpdateProgress "Compiling Blacklist Candidate...", fileName
                
                Dim wbSource As Object
                Dim openError As String
                Set wbSource = OpenSourceWorkbook(folderPath & fileName, openError)
                
                If Not wbSource Is Nothing Then
                    If SheetExists(wbSource, "Blacklist Candidate") Then
                        Dim wsSource As Object
                        Set wsSource = wbSource.Worksheets("Blacklist Candidate")
                        
                        Dim lastRow As Long
                        lastRow = GetLastUsedRowByValue(wsSource)
                        
                        If lastRow >= 2 Then
                            Dim lastCol As Long
                            lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
                            If lastCol < 9 Then lastCol = 9
                            If lastCol > 9 Then lastCol = 9
                            
                            Dim r As Long, c As Long
                            For r = 2 To lastRow
                                If RowHasAnyValue(wsSource, r, 1, lastCol) Then
                                    Dim keyStr As String
                                    keyStr = NormalizeKey(wsSource.Cells(r, 1).Value)
                                    
                                    If keyStr <> "" And Not dict.Exists(keyStr) Then
                                        Dim targetRow As Long
                                        targetRow = GetNextAppendRow(wsTarget, 2)
                                        
                                        For c = 1 To lastCol
                                            wsTarget.Cells(targetRow, c).Value = wsSource.Cells(r, c).Value
                                        Next c
                                        
                                        dict.Add keyStr, targetRow
                                    End If
                                End If
                            Next r
                        End If
                    End If
                    CloseSourceWorkbook wbSource
                End If
            End If
            fileName = Dir()
        Loop
    Next ext
    
    Set dict = Nothing
End Sub

Private Sub LoadExistingKeysBlacklist(ByVal wsTarget As Worksheet, ByRef dict As Object)
    Dim lastRow As Long
    lastRow = GetLastUsedRowByValue(wsTarget)
    If lastRow < 2 Then Exit Sub
    
    Dim existingData As Variant
    existingData = wsTarget.Range("A2:A" & lastRow).Value
    
    If IsArray(existingData) Then
        Dim r As Long
        For r = 1 To UBound(existingData, 1)
            Dim keyStr As String
            keyStr = NormalizeKey(existingData(r, 1))
            If keyStr <> "" And Not dict.Exists(keyStr) Then
                dict.Add keyStr, r
            End If
        Next r
    End If
End Sub

