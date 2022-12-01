Attribute VB_Name = "setcomboWidth"


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const CB_SETDROPPEDWIDTH = &H160
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub SetComboWidth(oComboBox As ComboBox, lWidth As Long)
' lWidth 是宽度，单位是 pixels
SendMessage oComboBox.hwnd, CB_SETDROPPEDWIDTH, lWidth, 0
End Sub
Public Sub SetCombo(ByRef oo As Form)
Dim i As Long
    For i = 0 To oo.Controls.Count - 1
        If TypeOf oo.Controls(i) Is ComboBox Then
            'Call SetComboWidth(oo.Controls(i), 270)
            Call SetComboHeight(oo.Controls(i), 270)
        End If
    Next i
End Sub

Public Sub SetComboHeight(oComboBox As ComboBox, lNewHeight As Long)
Dim oldscalemode As Integer
If TypeOf oComboBox.Parent Is Frame Then Exit Sub
oldscalemode = oComboBox.Parent.ScaleMode
oComboBox.Parent.ScaleMode = vbPixels
MoveWindow oComboBox.hwnd, oComboBox.Left, _
oComboBox.Top, oComboBox.Width, lNewHeight, 1
oComboBox.Parent.ScaleMode = oldscalemode
End Sub

