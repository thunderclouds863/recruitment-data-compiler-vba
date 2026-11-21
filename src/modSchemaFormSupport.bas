Attribute VB_Name = "modSchemaFormSupport"
Option Explicit

' ============================================================
' DYNAMIC FORM SUPPORT FOR SCHEMA-DRIVEN DATABASES
' ============================================================

Public Sub SchemaTagControl( _
    ByVal inputControl As Object, _
    ByVal labelControl As Object, _
    ByVal databaseKey As String, _
    ByVal fieldReference As String)

    Dim fieldID As String

    fieldID = SchemaResolveFieldID( _
        databaseKey, fieldReference)

    If Len(fieldID) = 0 Then Exit Sub

    On Error Resume Next

    inputControl.Tag = _
        "SCHEMA|" & databaseKey & "|" & _
        fieldID & "|INPUT"

    labelControl.Tag = _
        "SCHEMA|" & databaseKey & "|" & _
        fieldID & "|LABEL"

    On Error GoTo 0

End Sub

Public Sub SchemaApplyTaggedControls( _
    ByVal formObject As Object, _
    ByVal databaseKey As String)

    Dim controlItem As Object
    Dim tagParts() As String
    Dim fieldID As String
    Dim roleName As String
    Dim activeState As Boolean
    Dim headerName As String

    For Each controlItem In formObject.Controls

        If Left$(CStr(controlItem.Tag), 7) = _
           "SCHEMA|" Then

            tagParts = Split( _
                CStr(controlItem.Tag), "|")

            If UBound(tagParts) >= 3 Then

                If StrComp( _
                    tagParts(1), _
                    databaseKey, _
                    vbTextCompare) = 0 Then

                    fieldID = tagParts(2)
                    roleName = UCase$(tagParts(3))

                    activeState = SchemaFieldIsActive( _
                        databaseKey, fieldID)

                    headerName = SchemaCurrentHeader( _
                        databaseKey, fieldID)

                    controlItem.Visible = _
                        activeState

                    If roleName = "LABEL" Then

                        On Error Resume Next
                        controlItem.Caption = headerName
                        On Error GoTo 0

                    End If

                End If

            End If

        End If

    Next controlItem

End Sub

Public Function SchemaAppendCustomFields( _
    ByVal formObject As Object, _
    ByVal databaseKey As String, _
    ByVal startTop As Single, _
    ByVal leftPosition As Single, _
    ByVal totalWidth As Single, _
    Optional ByVal readOnlyMode As Boolean = False) As Single

    Dim fieldList As Collection
    Dim item As Variant
    Dim parts() As String

    Dim fieldID As String
    Dim headerName As String
    Dim activeText As String
    Dim customText As String
    Dim showText As String
    Dim readOnlyText As String
    Dim controlType As String
    Dim dropdownValues As String

    Dim labelObject As Object
    Dim inputObject As Object
    Dim currentTop As Single
    Dim safeName As String
    Dim values As Variant
    Dim valueItem As Variant

    Set fieldList = SchemaGetFieldList( _
        databaseKey, False)

    currentTop = startTop

    For Each item In fieldList

        parts = Split(CStr(item), "|")

        If UBound(parts) >= 8 Then

            fieldID = parts(0)
            headerName = parts(1)
            activeText = UCase$(Trim$(parts(2)))
            customText = UCase$(Trim$(parts(4)))
            showText = UCase$(Trim$(parts(5)))
            readOnlyText = UCase$(Trim$(parts(6)))
            controlType = UCase$(Trim$(parts(7)))
            dropdownValues = parts(8)

            If activeText = "YES" And _
               customText = "YES" And _
               showText = "YES" Then

                safeName = Left$( _
                    SchemaSafeControlNamePublic( _
                        fieldID), _
                    24)

                Set labelObject = formObject.Controls.Add( _
                    "Forms.Label.1", _
                    "schLbl_" & safeName, _
                    True)

                With labelObject

                    .Left = leftPosition
                    .Top = currentTop
                    .Width = 165
                    .Height = 16
                    .Caption = headerName
                    .Font.Size = 8
                    .Font.Bold = True
                    .BackStyle = 0
                    .Tag = _
                        "SCHEMA|" & databaseKey & "|" & _
                        fieldID & "|LABEL"

                End With

                If controlType = "COMBOBOX" Or _
                   Len(Trim$(dropdownValues)) > 0 Then

                    Set inputObject = _
                        formObject.Controls.Add( _
                            "Forms.ComboBox.1", _
                            "schInp_" & safeName, _
                            True)

                    inputObject.Style = _
                        fmStyleDropDownCombo

                    If Len(Trim$(dropdownValues)) > 0 Then

                        values = Split( _
                            dropdownValues, "^")

                        If UBound(values) = 0 Then

                            values = Split( _
                                dropdownValues, ";")

                        End If

                        If UBound(values) = 0 Then

                            values = Split( _
                                dropdownValues, ",")

                        End If

                        For Each valueItem In values

                            inputObject.AddItem _
                                Trim$(CStr(valueItem))

                        Next valueItem

                    End If

                Else

                    Set inputObject = _
                        formObject.Controls.Add( _
                            "Forms.TextBox.1", _
                            "schInp_" & safeName, _
                            True)

                End If

                With inputObject

                    .Left = leftPosition + 170
                    .Top = currentTop - 1
                    .Width = totalWidth - 175
                    .Height = 18
                    .Font.Size = 8
                    .Tag = _
                        "SCHEMA|" & databaseKey & "|" & _
                        fieldID & "|DYNAMIC"

                End With

                On Error Resume Next

                inputObject.Locked = _
                    readOnlyMode Or _
                    (readOnlyText = "YES")

                If inputObject.Locked Then

                    inputObject.backColor = _
                        RGB(236, 240, 241)

                End If

                On Error GoTo 0

                currentTop = currentTop + 23

            End If

        End If

    Next item

    SchemaAppendCustomFields = currentTop

End Function

Public Sub SchemaLoadTaggedControls( _
    ByVal formObject As Object, _
    ByVal databaseKey As String, _
    ByVal rowNumber As Long)

    Dim controlItem As Object
    Dim tagParts() As String
    Dim fieldID As String
    Dim valueItem As Variant

    For Each controlItem In formObject.Controls

        If Left$(CStr(controlItem.Tag), 7) = _
           "SCHEMA|" Then

            tagParts = Split( _
                CStr(controlItem.Tag), "|")

            If UBound(tagParts) >= 3 Then

                If StrComp( _
                    tagParts(1), _
                    databaseKey, _
                    vbTextCompare) = 0 And _
                   UCase$(tagParts(3)) = "DYNAMIC" Then

                    fieldID = tagParts(2)

                    valueItem = SchemaGetValue( _
                        databaseKey, _
                        fieldID, _
                        rowNumber)

                    On Error Resume Next
                    controlItem.Value = valueItem
                    On Error GoTo 0

                End If

            End If

        End If

    Next controlItem

End Sub

Public Sub SchemaSaveTaggedControls( _
    ByVal formObject As Object, _
    ByVal databaseKey As String, _
    ByVal rowNumber As Long)

    Dim controlItem As Object
    Dim tagParts() As String
    Dim fieldID As String

    For Each controlItem In formObject.Controls

        If Left$(CStr(controlItem.Tag), 7) = _
           "SCHEMA|" Then

            tagParts = Split( _
                CStr(controlItem.Tag), "|")

            If UBound(tagParts) >= 3 Then

                If StrComp( _
                    tagParts(1), _
                    databaseKey, _
                    vbTextCompare) = 0 And _
                   UCase$(tagParts(3)) = "DYNAMIC" Then

                    fieldID = tagParts(2)

                    SchemaSetValue _
                        databaseKey, _
                        fieldID, _
                        rowNumber, _
                        controlItem.Value

                End If

            End If

        End If

    Next controlItem

End Sub

Public Sub SchemaClearTaggedControls( _
    ByVal formObject As Object, _
    ByVal databaseKey As String)

    Dim controlItem As Object
    Dim tagParts() As String

    For Each controlItem In formObject.Controls

        If Left$(CStr(controlItem.Tag), 7) = _
           "SCHEMA|" Then

            tagParts = Split( _
                CStr(controlItem.Tag), "|")

            If UBound(tagParts) >= 3 Then

                If StrComp( _
                    tagParts(1), _
                    databaseKey, _
                    vbTextCompare) = 0 And _
                   UCase$(tagParts(3)) = "DYNAMIC" Then

                    On Error Resume Next
                    controlItem.Value = ""
                    On Error GoTo 0

                End If

            End If

        End If

    Next controlItem

End Sub

Private Function SchemaSafeControlNamePublic( _
    ByVal fieldID As String) As String

    Dim resultText As String
    Dim characterIndex As Long
    Dim currentCharacter As String

    For characterIndex = 1 To Len(fieldID)

        currentCharacter = Mid$( _
            fieldID, characterIndex, 1)

        If currentCharacter Like _
           "[A-Za-z0-9_]" Then

            resultText = resultText & _
                currentCharacter

        End If

    Next characterIndex

    If Len(resultText) = 0 Then
        resultText = "FIELD"
    End If

    SchemaSafeControlNamePublic = resultText

End Function
