Attribute VB_Name = "sfe_add"
'不知道为什么，vc编的dll只能命名成这样才可以链接
Public Declare Function GetPNGInfo Lib "png.dll" Alias "_GetPNGInfo@12" (ByVal filename As String, w As Long, h As Long) As Long
Public Declare Function GetPNGData Lib "png.dll" Alias "_GetPNGData@8" (ByVal filename As String, data As Any) As Long

Public Declare Function BYTErol Lib "sfeAdd.dll" (ByVal a As Byte, ByVal n As Byte) As Byte
Public Declare Function BYTEror Lib "sfeAdd.dll" (ByVal a As Byte, ByVal n As Byte) As Byte
Public Declare Function INTrol Lib "sfeAdd.dll" (ByVal a As Integer, ByVal n As Integer) As Integer
Public Declare Function INTror Lib "sfeAdd.dll" (ByVal a As Integer, ByVal n As Integer) As Integer
Public Declare Sub getbyteItem Lib "sfeAdd.dll" (ByRef a As Byte, ByVal Length As Long, ByVal tkey As Integer)
'Public Declare Function RoRForWord Lib "math2.dll" (ByVal a As Integer, ByVal n As Integer) As Integer
'Public Declare Function RoLForWord Lib "math2.dll" (ByVal a As Integer, ByVal n As Integer) As Integer
'Public Declare Function RoRForByte Lib "math2.dll" (ByVal a As Byte, ByVal n As Byte) As Byte
'Public Declare Function RoLForByte Lib "math2.dll" (ByVal a As Byte, ByVal n As Byte) As Byte
'Public Declare Function XoRWord Lib "math2.dll" (ByVal a As Integer, c1 As Byte, c2 As Byte) As Integer
Public Declare Sub convertCOLOR Lib "sfeAdd.dll" (ByRef mcolor_RGB As Long, ByRef data As Long, ByVal WW As Long, ByVal HH As Long, ByVal MaskColor1 As Long)
Public Declare Sub convertCOLOR2 Lib "sfeAdd.dll" (ByRef mcolor_RGB As Long, ByRef data As Long, ByVal WW As Long, ByVal HH As Long, ByVal MaskColor1 As Long)
Public Declare Function get256 Lib "sfeAdd.dll" (ByRef mcolor_RGB As Long, ByVal d As Long) As Byte

