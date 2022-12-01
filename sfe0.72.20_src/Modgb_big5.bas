Attribute VB_Name = "modgb_big5"
Option Explicit


' ���ֲ�ͬ����ת��ģ��
' �����б�
'     LoadResStr   ���ַ�����Դ��ת��Ϊ���ʵ��ַ���
'     JtoF            ����unicodeת��Ϊ����unicode
'     UnicodetoBIG5   unicode��big5��ת��
'     Big5toUnicode   big5��unicode��ת��
'     LoadMB          ������ļ�
'
'
'  ��Ҫ�趨ȫ�ֱ���     Charset  = "GBK"  or "BIG5"
'

Public gbk_big5(128 To 255, 255, 1) As Byte
Public big5_gbk(128 To 255, 255, 1) As Byte
Public unicodeF_J(255, 255, 1) As Byte


' ���ַ�����Դ��ת��Ϊ���ʵ��ַ���
' ����
'      id:       ��Դid
'
' ����ֵ��
'      ת������ַ���
'
' ˵��: ��Դ�ļ����ü���д�ģ������ڴ�����߱���ʱ��ת����unicode�ˣ������ڷ���ϵͳ�в�����ʾ���еļ��֣�
'       ���Ʒ���ϵͳ��û����Щ�ַ�����˱������unicode���� to unicode�����ת������ʵֱ������������Դ�ļ�
'        Ҳ����
      
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
Private Function getlengthb(Data() As Byte, lengthdata As Long) As Long
Dim i As Long
   For i = 0 To lengthdata - 1
       If Data(i) = 0 Then
           getlengthb = i
           
           Exit Function
       End If
   Next i
   getlengthb = lengthdata
End Function

' ����unicodeת��Ϊ����unicode
' ����
'      data       ��ת�����ַ���
' ����ֵ
'      ת������ַ���


Public Function JtoF(Data As String) As String
Dim i As Long
Dim lengthb As Long
Dim tmpdata() As Byte
Dim tmpresult() As Byte
Dim b0 As Byte, b1 As Byte

    lengthb = LenB(Data)
    If lengthb = 0 Then
        JtoF = ""
        Exit Function
    End If
    ReDim tmpdata(lengthb - 1)
    ReDim tmpresult(lengthb - 1)
    
    For i = 0 To lengthb - 1                 ' ���ַ������Ƶ��ֽ�����
        tmpdata(i) = AscB(MidB(Data, i + 1, 1))
    Next i

    For i = 0 To lengthb - 1 Step 2          ' unicodeΪ�����ֽ�һ���ַ�
        b0 = tmpdata(i)
        b1 = tmpdata(i + 1)
        If unicodeF_J(b0, b1, 0) = 0 And unicodeF_J(b0, b1, 1) = 0 Then   ' Ϊ0��ʾ��unicode�ַ�û�ж�Ӧ�ķ��壬����ת��
            tmpresult(i) = b0
            tmpresult(i + 1) = b1
        Else
            tmpresult(i) = unicodeF_J(b0, b1, 0)            ' ���ת��
            tmpresult(i + 1) = unicodeF_J(b0, b1, 1)
        End If
    Next i
    JtoF = tmpresult                         ' ������ֱ�Ӹ�ֵ���ַ���
End Function


' unicode��big5��ת��
' ����
'        data    ��ת�����ַ���
' ���
'        lengthdata   ���big5���鳤��
'        big5data()   ����ת�����big5����

Public Sub UnicodetoBIG5(Data As String, lengthdata As Long, big5data() As Byte)
Dim lengthb As Long
Dim lengthUnicode As Long
Dim result As Long
Dim testgb() As Byte
Dim GBstr As String
Dim i As Long
Dim p As Long
Dim b0 As Byte, b1 As Byte
Dim bf0 As Byte, bf1 As Byte


If Charset = "GBK" Then    ' ����ϵͳ�ַ���Ϊgbk
    
    GBstr = StrConv(Data, vbFromUnicode)   ' unicode ת��Ϊgbk
    
    lengthb = LenB(GBstr)
    lengthdata = lengthb
    ReDim testgb(lengthb - 1)
    For i = 0 To lengthb - 1
        testgb(i) = AscB(MidB(GBstr, i + 1, 1))  ' gbk��ֵ���ֽ�����
    Next i
    
    
    p = 0
    ReDim big5data(lengthb - 1)
    
    While p < lengthb
        b0 = testgb(p)
        If b0 < 128 Then                ' ���ֽڲ��Ǻ��֣���ת����ֱ�Ӹ�ֵ
            big5data(p) = b0
            p = p + 1
        Else
            b1 = testgb(p + 1)
            bf0 = gbk_big5(b0, b1, 0)
            bf1 = gbk_big5(b0, b1, 1)
            If bf0 > 0 And bf1 > 0 Then  '  ' ת����������ת��ΪBIG5
                big5data(p) = bf0
                big5data(p + 1) = bf1
            Else
                big5data(p) = AscB("?")    ' û�����ã�����
                big5data(p + 1) = AscB("?")
            End If
            p = p + 2
        End If
    Wend
Else      ' ����ϵͳ�ַ���Ϊbig5
    GBstr = StrConv(Data, vbFromUnicode)   ' ֱ��ת����big5�ַ���
    lengthdata = LenB(GBstr)
    ReDim big5data(lengthdata - 1)
    For i = 0 To lengthdata - 1
        big5data(i) = AscB(MidB(GBstr, i + 1, 1))  ' ���Ƶ��ֽ�����
    Next i
End If
    

End Sub


' big5��unicode��ת��
' ����
'       data           ����big5������
'       lengthdata     data���鳤��
' ����ֵ
'       ת�����unicode�ַ���

Public Function Big5toUnicode(Data() As Byte, lengthdata As Long) As String
Dim tmpdata() As Byte
Dim p As Long
Dim b0 As Byte, b1 As Byte
Dim bf0 As Byte, bf1 As Byte
Dim lengthb
    p = 0
    lengthb = getlengthb(Data, lengthdata)     ' ����ʵ�ʳ��ȣ���������������ɸ�0�ģ�
    If lengthb = 0 Then Exit Function
    ReDim tmpdata(lengthb - 1)
    
    
    
If Charset = "GBK" Then        ' ����ϵͳ�ַ���Ϊgbk
    While p < lengthb
        b0 = Data(p)
        If b0 < 128 Then        ' ���ֽڲ��Ǻ��֣���ת����ֱ�Ӹ�ֵ
            tmpdata(p) = b0
            p = p + 1
        Else
            b1 = Data(p + 1)
            bf0 = big5_gbk(b0, b1, 0)
            bf1 = big5_gbk(b0, b1, 1)
            If bf0 > 0 Or bf1 > 0 Then    ' ת����������ת��Ϊgbk
                tmpdata(p) = bf0
                tmpdata(p + 1) = bf1
             Else
                 tmpdata(p) = AscB("?")    ' û�����ã�����
                 tmpdata(p + 1) = AscB("?")
             End If
            p = p + 2
        End If
    Wend
    Big5toUnicode = StrConv(tmpdata, vbUnicode)   ' gbkת��Ϊunicode
Else            ' ����ϵͳ�ַ���Ϊbig5
    Dim i As Long
    For i = 0 To lengthb - 1
        tmpdata(i) = Data(i)               ' ����data��ʵ�����ݣ�ȥ�������0
    Next i
    Big5toUnicode = StrConv(tmpdata, vbUnicode)  ' big5ת��Ϊunicode
End If
End Function

' ������gbk�ַ�����ȷ��ʾ����big5�£����ڿؼ����˵���5
Public Function StrUnicode(ss As String) As String
If Charset = "GBK" Then
    StrUnicode = ss
Else
    StrUnicode = GBKtoUnicode(StrConv(ss, vbFromUnicode))
    ' StrUnicode = JtoF(ss)
End If
End Function


' ������gbk�ַ�����ȷ��ʾ����big5��, ���ڳ����ַ���
Public Function StrUnicode2(ss As String) As String
If Charset = "GBK" Then
    StrUnicode2 = ss
Else
    'StrUnicode = GBKtoUnicode(StrConv(ss, vbFromUnicode))
     StrUnicode2 = JtoF(ss)
End If
End Function


' gbk��unicode��ת��
' ����
'       data           ����big5������
'       lengthdata     data���鳤��
' ����ֵ
'       ת�����unicode�ַ���

Public Function GBKtoUnicode(ss As String) As String
Dim i As Long
Dim Data() As Byte
Dim tmpdata() As Byte
Dim p As Long
Dim b0 As Byte, b1 As Byte
Dim bf0 As Byte, bf1 As Byte
Dim lengthb
    p = 0
    'lengthb = getlengthb(data, lengthdata)     ' ����ʵ�ʳ��ȣ���������������ɸ�0�ģ�
    'If lengthb = 0 Then Exit Function
    'ReDim tmpdata(lengthb - 1)
    
    lengthb = LenB(ss)
    If lengthb = 0 Then
        GBKtoUnicode = ""
        Exit Function
    End If
    ReDim Data(lengthb - 1)
    ReDim tmpdata(lengthb - 1)
    
    For i = 0 To lengthb - 1
        Data(i) = AscB(MidB(ss, i + 1, 1))
    Next i
    
    
If Charset = "BIG5" Then        ' ����ϵͳ�ַ���Ϊbig5
    While p < lengthb
        b0 = Data(p)
        If b0 < 128 Then        ' ���ֽڲ��Ǻ��֣���ת����ֱ�Ӹ�ֵ
            tmpdata(p) = b0
            p = p + 1
        Else
            b1 = Data(p + 1)
            bf0 = gbk_big5(b0, b1, 0)
            bf1 = gbk_big5(b0, b1, 1)
            If bf0 > 0 Or bf1 > 0 Then    ' ת����������ת��Ϊgbk
                tmpdata(p) = bf0
                tmpdata(p + 1) = bf1
            Else
                 tmpdata(p) = AscB("?")   ' û�����ã�����
                 tmpdata(p + 1) = AscB("?")
            End If
            p = p + 2
        End If
    Wend
    GBKtoUnicode = StrConv(tmpdata, vbUnicode)   ' gbkת��Ϊunicode
    GBKtoUnicode = JtoF(GBKtoUnicode)
Else                            ' ����ϵͳ�ַ���Ϊgbk
    For i = 0 To lengthb - 1
        tmpdata(i) = Data(i)               ' ����data��ʵ�����ݣ�ȥ�������0
    Next i
    GBKtoUnicode = StrConv(tmpdata, vbUnicode)  ' big5ת��Ϊunicode
End If
End Function


' ��ȡ����ļ�
' 1 �ɹ�
' 0 ʧ��
Public Function LoadMB() As Long
Dim i As Long, j As Long
Dim filenum As Long
Dim filename As String
Dim b0 As Byte, b1 As Byte
Dim c0 As Byte, c1 As Byte
    
    
    filename = App.Path & "\mb.dat"        ' ����ļ���
    If Dir(filename) = "" Then
        MsgBox "file " & filename & " not exist"
        LoadMB = 0
        Exit Function
    End If
    
    filenum = FreeFile()
    
    Open App.Path & "\mb.dat" For Binary Access Read As #filenum
    For i = &HA0 To &HFE
        For j = &H40 To &HFE
            If j <= &H7E Or j >= &HA1 Then
                Get #filenum, , b0
                Get #filenum, , b1
                big5_gbk(i, j, 0) = b0
                big5_gbk(i, j, 1) = b1
            End If
        Next j
    Next i
            
            
    For i = &H81 To &HFE
        For j = &H40 To &HFE
            If j <> &H7F Then
                Get #filenum, , b0
                Get #filenum, , b1
                gbk_big5(i, j, 0) = b0
                gbk_big5(i, j, 1) = b1
            End If
        Next j
    Next i

    For i = &H81 To &HFE
        For j = &H40 To &HFE
            If j <> &H7F Then
                Get #filenum, , b0
                Get #filenum, , b1
                Get #filenum, , c0
                Get #filenum, , c1
                unicodeF_J(b0, b1, 0) = c0
                unicodeF_J(b0, b1, 1) = c1
            End If
        Next j
    Next i
    Close filenum
    LoadMB = 1

End Function


