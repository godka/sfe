Attribute VB_Name = "inifile"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public iniFileName As String '配置文件的名称，一般在窗体load事件中初始化

'获取Ini的值，注意DefString表示如果不存在对应的KeyWord就设置此项为DefString，为空时不处理
Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, Optional ByVal DefString As String) As String
    Dim ResultString As String * 144, Temp%
    Dim s$, i%
    Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, iniFileName)
    '检索关键词的值
    If Temp% > 0 Then '关键词的值不为空
        For i = 1 To 144
            If Asc(Mid$(ResultString, i, 1)) <> 0 Then
                s = s & Mid$(ResultString, i, 1)
            End If
        Next
    Else
        Temp% = WritePrivateProfileString(SectionName, KeyWord, DefString, iniFileName) '将缺省值写入INI文件
        s = DefString
    End If
    GetIniS = s
End Function
'写入字符串值，返回值如果是0表示操作失败
Public Function SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) As Boolean
    SetIniS = WritePrivateProfileString(SectionName, KeyWord, ValStr, iniFileName)
End Function
'清除 Section"段"
Public Function DelIniSec(ByVal SectionName As String) As Boolean
    DelIniSec = WritePrivateProfileString(SectionName, 0&, "", iniFileName)
End Function
''清除KeyWord"键"
Public Function DelIniKey(ByVal SectionName As String, ByVal KeyWord As String) As Boolean
    DelIniKey = WritePrivateProfileString(SectionName, KeyWord, 0&, iniFileName)
End Function

