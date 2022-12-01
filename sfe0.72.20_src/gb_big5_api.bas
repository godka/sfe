Attribute VB_Name = "gb_big5_api"
'简繁转换声明
Private Declare Function LCMapStringW Lib "kernel32.dll" _
    (ByVal Locale As Long, ByVal dwMapFlags As Long, ByVal lpSrcStr As Long, _
    ByVal cchSrc As Long, ByVal lpDestStr As Long, ByVal cchDest As Long) As Long
 
Private Const LCMAP_BYTEREV As Long = &H800
Private Const LCMAP_FULLWIDTH As Long = &H800000
Private Const LCMAP_HALFWIDTH As Long = &H400000
Private Const LCMAP_HIRAGANA As Long = &H100000
Private Const LCMAP_KATAKANA As Long = &H200000
Private Const LCMAP_LINGUISTIC_CASING As Long = &H1000000
Private Const LCMAP_LOWERCASE As Long = &H100
Private Const LCMAP_SIMPLIFIED_CHINESE As Long = &H2000000
Private Const LCMAP_SORTKEY As Long = &H400
Private Const LCMAP_TRADITIONAL_CHINESE As Long = &H4000000
Private Const LCMAP_UPPERCASE As Long = &H200
 
'主语言ID
Private Const LANG_CHINESE As Long = &H4
 
'次语言ID
Private Const SUBLANG_CHINESE_TRADITIONAL As Long = &H1
Private Const SUBLANG_CHINESE_SIMPLIFIED As Long = &H2
Private Const SUBLANG_CHINESE_HONGKONG As Long = &H3
Private Const SUBLANG_CHINESE_SINGAPORE As Long = &H4
Private Const SUBLANG_CHINESE_MACAU As Long = &H5
 
'排序方式
Private Const SORT_CHINESE_PRCP As Long = &H0
Private Const SORT_CHINESE_BIG5 As Long = &H0
Private Const SORT_CHINESE_UNICODE As Long = &H1
Private Const SORT_CHINESE_PRC As Long = &H2
Private Const SORT_CHINESE_BOPOMOFO As Long = &H3
 
'生成LCID
Private Const LCID_CHINESE_SIMPLIFIED As Long = (LANG_CHINESE Or SUBLANG_CHINESE_SIMPLIFIED * &H400) _
    And &HFFFF& Or SORT_CHINESE_PRCP * &H10000
Private Const LCID_CHINESE_TRADITIONAL As Long = (LANG_CHINESE Or SUBLANG_CHINESE_TRADITIONAL * &H400) _
    And &HFFFF& Or SORT_CHINESE_BIG5 * &H10000
Public Function GBBIG5(sStr As String, iConver As Integer) As String
' sStr 需要转换的文本
' iConver 要转化的类型，1 BIG5-->GB， 2 GB-->BIG5
    On Error Resume Next
    Dim Str
    If iConver = 1 Then           'BIG5-->GB
        Str = StrConv(sStr, vbFromUnicode, &H804)
        Str = StrConv(Str, vbUnicode, &H404)
        'Call LCMapStringW(LCID_CHINESE_SIMPLIFIED, LCMAP_SIMPLIFIED_CHINESE, _
            ByVal StrPtr(STR), Len(STR), ByVal StrPtr(GBBIG5), Len(GBBIG5))
        GBBIG5 = String$(Len(Str), 0)
        Call LCMapStringW(LCID_CHINESE_SIMPLIFIED, LCMAP_SIMPLIFIED_CHINESE, _
            ByVal StrPtr(Str), Len(Str), ByVal StrPtr(GBBIG5), Len(GBBIG5))
    ElseIf iConver = 2 Then       'GB-->BIG5
        Str = String$(Len(sStr), 0)
        Call LCMapStringW(LCID_CHINESE_TRADITIONAL, LCMAP_TRADITIONAL_CHINESE, _
            ByVal StrPtr(sStr), Len(sStr), ByVal StrPtr(Str), Len(Str))
        Str = StrConv(Str, vbFromUnicode, &H404)
        GBBIG5 = StrConv(Str, vbUnicode, &H804)
    End If
End Function

Public Sub UnicodetoBIG5(data As String, lengthdata As Long, big5data() As Byte)
Dim lengthb As Long
Dim testgb() As Byte
Dim GBstr As String
Dim i As Long


If Charset = "GBK" Then    ' 操作系统字符集为gbk

    'GBstr = GBBIG5(data, 2)
    GBstr = data 'StrConv(GBstr, vbFromUnicode)   ' unicode 转换为gbk
    
    lengthdata = LenB(GBstr)
    If lengthdata <= 0 Then Exit Sub
    ReDim big5data(lengthdata - 1)
    For i = 0 To lengthdata - 1
        big5data(i) = AscB(MidB(GBstr, i + 1, 1))  ' 复制到字节数组
    Next i
Else      ' 操作系统字符集为big5
    GBstr = data 'StrConv(data, vbFromUnicode)   ' 直接转换成big5字符串
    lengthdata = LenB(GBstr)
    ReDim big5data(lengthdata - 1)
    For i = 0 To lengthdata - 1
        big5data(i) = AscB(MidB(GBstr, i + 1, 1))  ' 复制到字节数组
    Next i
End If
    

End Sub
Public Function Big5toUnicode(data() As Byte, lengthdata As Long) As String
Dim tmpdata() As Byte
Dim lengthb
Dim tmpstr As String
Dim i As Long
    lengthb = getlengthb(data, lengthdata)     ' 计算实际长度（不包括后面的若干个0的）
    If lengthb <= 0 Then Exit Function
    ReDim tmpdata(lengthb - 1)
    For i = 0 To lengthb - 1
        tmpdata(i) = data(i)               ' 复制data中实际数据，去掉后面的0
    Next i
    tmpstr = tmpdata
    
    'StrConv(tmpdata, vbUnicode)
    
If Charset = "GBK" Then        ' 操作系统字符集为gbk
    Big5toUnicode = tmpdata 'StrConv(tmpdata, vbUnicode)   ' gbk转换为unicode
    'Big5toUnicode = GBBIG5(Big5toUnicode, 1)
Else            ' 操作系统字符集为big5
    Big5toUnicode = tmpdata 'StrConv(tmpdata, vbUnicode)  ' big5转换为unicode
End If
End Function
Public Function JtoF(data As String) As String
Dim Str As String
    'sStr = data
    'Str = String$(Len(data), 0)
    'Call LCMapStringW(LCID_CHINESE_TRADITIONAL, LCMAP_TRADITIONAL_CHINESE, ByVal StrPtr(sStr), Len(sStr), ByVal StrPtr(Str), Len(Str))
    JtoF = data                   ' 把数组直接赋值给字符串
End Function

' gbk到unicode的转换
' 输入
'       data           保存big5的数组
'       lengthdata     data数组长度
' 返回值
'       转换后的unicode字符串

Public Function GBKtoUnicode(ss As String) As String
Dim i As Long
Dim data() As Byte
Dim lengthb
    'lengthb = getlengthb(data, lengthdata)     ' 计算实际长度（不包括后面的若干个0的）
    'If lengthb = 0 Then Exit Function
    'ReDim tmpdata(lengthb - 1)
    
    lengthb = LenB(ss)
    If lengthb = 0 Then
        GBKtoUnicode = ""
        Exit Function
    End If
    ReDim data(lengthb - 1)
    
    For i = 0 To lengthb - 1
        data(i) = AscB(MidB(ss, i + 1, 1))
    Next i
    
    
If Charset = "BIG5" Then        ' 操作系统字符集为big5
    GBKtoUnicode = StrConv(data, vbUnicode)   ' gbk转换为unicode
    'GBKtoUnicode = GBBIG5(GBKtoUnicode, 2)
Else                            ' 操作系统字符集为gbk
    GBKtoUnicode = StrConv(data, vbUnicode)  ' big5转换为unicode
End If
End Function
' 程序中gbk字符能正确显示，在big5下，用于控件、菜单等5
Public Function StrUnicode(ss As String) As String
If Charset = "GBK" Then
    StrUnicode = ss
Else
    StrUnicode = ss 'GBKtoUnicode(StrConv(ss, vbFromUnicode))
    ' StrUnicode = JtoF(ss)
End If
End Function


' 程序中gbk字符能正确显示，在big5下, 用于程序字符串
Public Function StrUnicode2(ss As String) As String
If Charset = "GBK" Then
    StrUnicode2 = ss
Else
    'StrUnicode = GBKtoUnicode(StrConv(ss, vbFromUnicode))
     StrUnicode2 = ss 'JtoF(ss)
End If
End Function
Public Function LoadResStr(id As Long) As String
Dim buffer() As Byte
Dim tmpstr As String
    
    buffer = LoadResData(id, 6)         ' 按照字节数组形式读资源文件
    If Charset = "GBK" Then
        LoadResStr = buffer          ' 当前操作系统字符集采用GBK则直接赋值（直接把字节数据赋值给字符串并不进行unicode转换）
    Else
        tmpstr = buffer
        LoadResStr = JtoF(tmpstr)    ' 把简体unicode转换为繁体unicode（big5繁体系统中不能显示简体字）
    End If

End Function


' 计算byte数组实际长度，到0为止。只用于判断以0结尾的字节字串实际长度
' 输入
'      data()              输入的数组
'      lengthdata          数组长度
' 返回值
'      数组实际长度（后面为0）
Public Function getlengthb(data() As Byte, lengthdata As Long) As Long
Dim i As Long
    getlengthb = lengthdata
     'For i = 0 To lengthdata - 1
   '    If data(i) = 0 Then
   '        getlengthb = i
   '
   '        Exit Function
   '    End If
   'Next i
   'getlengthb = lengthdata
End Function
