Attribute VB_Name = "SysFunction"
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare PtrSafe Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Function CapsLock() As Boolean
Dim Res As Long
Dim KBState(0 To 255) As Byte
Res = GetKeyboardState(KBState(0))
CapsLock = KBState(&H14) And 1
End Function
