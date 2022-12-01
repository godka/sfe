VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSkinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function AttachSkin(ByVal hwnd As Long) As Long
Dim i As Long
If GetINILong("run", "debug") <> 99 Then
    AttachSkin = Attach(hwnd)

End If
End Function

Public Function DetachSkin(ByVal hwnd As Long) As Long
    DetachSkin = Detach(hwnd)
End Function


