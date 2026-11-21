Attribute VB_Name = "modKeyboardState"
Option Explicit

' ============================================================
' WINDOWS KEYBOARD STATE
'
' Letakkan code ini di STANDARD MODULE, bukan di UserForm.
' Contoh nama module: modKeyboardState
' ============================================================

#If VBA7 Then

    Private Declare PtrSafe Function GetKeyState Lib "user32" ( _
        ByVal nVirtKey As Long) As Integer

#Else

    Private Declare Function GetKeyState Lib "user32" ( _
        ByVal nVirtKey As Long) As Integer

#End If

Public Function IsShiftKeyDown() As Boolean

    IsShiftKeyDown = _
        ((GetKeyState(vbKeyShift) And &H8000) <> 0)

End Function




