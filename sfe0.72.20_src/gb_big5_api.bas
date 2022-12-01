Attribute VB_Name = "gb_big5_api"
'��ת������
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
 
'������ID
Private Const LANG_CHINESE As Long = &H4
 
'������ID
Private Const SUBLANG_CHINESE_TRADITIONAL As Long = &H1
Private Const SUBLANG_CHINESE_SIMPLIFIED As Long = &H2
Private Const SUBLANG_CHINESE_HONGKONG As Long = &H3
Private Const SUBLANG_CHINESE_SINGAPORE As Long = &H4
Private Const SUBLANG_CHINESE_MACAU As Long = &H5
 
'����ʽ
Private Const SORT_CHINESE_PRCP As Long = &H0
Private Const SORT_CHINESE_BIG5 As Long = &H0
Private Const SORT_CHINESE_UNICODE As Long = &H1
Private Const SORT_CHINESE_PRC As Long = &H2
Private Const SORT_CHINESE_BOPOMOFO As Long = &H3
 
'����LCID
Private Const LCID_CHINESE_SIMPLIFIED As Long = (LANG_CHINESE Or SUBLANG_CHINESE_SIMPLIFIED * &H400) _
    And &HFFFF& Or SORT_CHINESE_PRCP * &H10000
Private Const LCID_CHINESE_TRADITIONAL As Long = (LANG_CHINESE Or SUBLANG_CHINESE_TRADITIONAL * &H400) _
    And &HFFFF& Or SORT_CHINESE_BIG5 * &H10000
Public Function GBBIG5(sStr As String, iConver As Integer) As String
' sStr ��Ҫת�����ı�
' iConver Ҫת�������ͣ�1 BIG5-->GB�� 2 GB-->BIG5
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


If Charset = "GBK" Then    ' ����ϵͳ�ַ���Ϊgbk

    'GBstr = GBBIG5(data, 2)
    GBstr = data 'StrConv(GBstr, vbFromUnicode)   ' unicode ת��Ϊgbk
    
    lengthdata = LenB(GBstr)
    If lengthdata <= 0 Then Exit Sub
    ReDim big5data(lengthdata - 1)
    For i = 0 To lengthdata - 1
        big5data(i) = AscB(MidB(GBstr, i + 1, 1))  ' ���Ƶ��ֽ�����
    Next i
Else      ' ����ϵͳ�ַ���Ϊbig5
    GBstr = data 'StrConv(data, vbFromUnicode)   ' ֱ��ת����big5�ַ���
    lengthdata = LenB(GBstr)
    ReDim big5data(lengthdata - 1)
    For i = 0 To lengthdata - 1
        big5data(i) = AscB(MidB(GBstr, i + 1, 1))  ' ���Ƶ��ֽ�����
    Next i
End If
    

End Sub
Public Function Big5toUnicode(data() As Byte, lengthdata As Long) As String
Dim tmpdata() As Byte
Dim lengthb
Dim tmpstr As String
Dim i As Long
    lengthb = getlengthb(data, lengthdata)     ' ����ʵ�ʳ��ȣ���������������ɸ�0�ģ�
    If lengthb <= 0 Then Exit Function
    ReDim tmpdata(lengthb - 1)
    For i = 0 To lengthb - 1
        tmpdata(i) = data(i)               ' ����data��ʵ�����ݣ�ȥ�������0
    Next i
    tmpstr = tmpdata
    
    'StrConv(tmpdata, vbUnicode)
    
If Charset = "GBK" Then        ' ����ϵͳ�ַ���Ϊgbk
    Big5toUnicode = tmpdata 'StrConv(tmpdata, vbUnicode)   ' gbkת��Ϊunicode
    'Big5toUnicode = GBBIG5(Big5toUnicode, 1)
Else            ' ����ϵͳ�ַ���Ϊbig5
    Big5toUnicode = tmpdata 'StrConv(tmpdata, vbUnicode)  ' big5ת��Ϊunicode
End If
End Function
Public Function JtoF(data As String) As String
Dim Str As String
    'sStr = data
    'Str = String$(Len(data), 0)
    'Call LCMapStringW(LCID_CHINESE_TRADITIONAL, LCMAP_TRADITIONAL_CHINESE, ByVal StrPtr(sStr), Len(sStr), ByVal StrPtr(Str), Len(Str))
    JtoF = data                   ' ������ֱ�Ӹ�ֵ���ַ���
End Function

' gbk��unicode��ת��
' ����
'       data           ����big5������
'       lengthdata     data���鳤��
' ����ֵ
'       ת�����unicode�ַ���

Public Function GBKtoUnicode(ss As String) As String
Dim i As Long
Dim data() As Byte
Dim lengthb
    'lengthb = getlengthb(data, lengthdata)     ' ����ʵ�ʳ��ȣ���������������ɸ�0�ģ�
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
    
    
If Charset = "BIG5" Then        ' ����ϵͳ�ַ���Ϊbig5
    GBKtoUnicode = StrConv(data, vbUnicode)   ' gbkת��Ϊunicode
    'GBKtoUnicode = GBBIG5(GBKtoUnicode, 2)
Else                            ' ����ϵͳ�ַ���Ϊgbk
    GBKtoUnicode = StrConv(data, vbUnicode)  ' big5ת��Ϊunicode
End If
End Function
' ������gbk�ַ�����ȷ��ʾ����big5�£����ڿؼ����˵���5
Public Function StrUnicode(ss As String) As String
If Charset = "GBK" Then
    StrUnicode = ss
Else
    StrUnicode = ss 'GBKtoUnicode(StrConv(ss, vbFromUnicode))
    ' StrUnicode = JtoF(ss)
End If
End Function


' ������gbk�ַ�����ȷ��ʾ����big5��, ���ڳ����ַ���
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
    
    buffer = LoadResData(id, 6)         ' �����ֽ�������ʽ����Դ�ļ�
    If Charset = "GBK" Then
        LoadResStr = buffer          ' ��ǰ����ϵͳ�ַ�������GBK��ֱ�Ӹ�ֵ��ֱ�Ӱ��ֽ����ݸ�ֵ���ַ�����������unicodeת����
    Else
        tmpstr = buffer
        LoadResStr = JtoF(tmpstr)    ' �Ѽ���unicodeת��Ϊ����unicode��big5����ϵͳ�в�����ʾ�����֣�
    End If

End Function


' ����byte����ʵ�ʳ��ȣ���0Ϊֹ��ֻ�����ж���0��β���ֽ��ִ�ʵ�ʳ���
' ����
'      data()              ���������
'      lengthdata          ���鳤��
' ����ֵ
'      ����ʵ�ʳ��ȣ�����Ϊ0��
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
