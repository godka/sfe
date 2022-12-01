Attribute VB_Name = "basBIG2GB"

'*****************************************************
'//
'//  basBIG2GB.bas  1999/08/19
'//
'//  ����:�¹�ǿ  alone@telekbird.com.cn
'//  ԭ�����ݹ�����  http://www.quanqiu.com/vb
'//
'//  ���������ɵĿ�����ʹ�ñ�����
'//  ��������ѳ����е�BUG������
'//
'*****************************************************
'
'
'
'����˵��

'Sub InitDATA
'��ʼ����������
'�״ε���GB2BIG��BIG2GB����֮ǰ�����Call InitDATA
'����BIG5Order()�д������BIG5�뺺�ֶ�Ӧ��GB2312��ĵ� ANSI �ַ����롣
'����GBOrder()�д������GB2312�뺺�ֶ�Ӧ��BIG5��ĵ� ANSI �ַ����롣
'ʹ��Chr(ANSI �ַ�����)���ɵõ���Ӧ�ĺ���
'
'Function GB2BIG(strGB As String) As String
'GB2312�� -> BIG5��
'
'Function BIG2GB(strBIG As String) As String
'BIG5�� -> GB2312��

'Function CheckBIG(strSource As String) As Boolean
'�ж�һ���������Ƿ���BIG5�뺺�� , ������������Զ�ʶ��
'����True��ʾ����BIG5��
'����False��ʾ����BIG5�� , �������һ�����Ϊ��GB��
'
'
'��Դ�ļ������ɷ������ResourceĿ¼�µ�BuildDATA.vbp��Ŀ


Option Explicit

Private GBOrder(8177) As Integer
Private BIG5Order(14757) As Integer
Private InitOK As Boolean
Private ByteDataGB() As Byte
Private ByteDataBIG() As Byte



Public Sub InitDATA()
On Error GoTo ERROR_HANDLE
Dim h As Long
Dim i, j As Integer
InitOK = True

ByteDataGB = LoadResData(101, "INS")
ByteDataBIG = LoadResData(102, "INS")

For i = LBound(ByteDataGB) To UBound(ByteDataGB) / 2
    GBOrder(i) = Val("&H" & Hex(ByteDataGB(2 * i + 1)) & Hex(ByteDataGB(2 * i)))
Next i
For i = LBound(ByteDataBIG) To UBound(ByteDataBIG) / 2
    BIG5Order(i) = Val("&H" & Hex(ByteDataBIG(2 * i + 1)) & Hex(ByteDataBIG(2 * i)))
Next i
Exit Sub
ERROR_HANDLE:
    InitOK = False
End Sub

Public Function GB2BIG(strGB As String) As String
On Error Resume Next
Dim ByteGB() As Byte
Dim ByteTemp(1) As Byte
Dim leng As Long, idx As Long
Dim strOut As String
Dim Offset As Long

If Not InitOK Then Call InitDATA
If Not InitOK Then
    GB2BIG = strGB
    Exit Function
End If

ByteGB = StrConv(strGB, vbFromUnicode)
leng = UBound(ByteGB)
idx = 0

Do While idx <= leng
    ByteTemp(0) = ByteGB(idx)
    ByteTemp(1) = ByteGB(idx + 1)
    Offset = GBOffset(ByteTemp)
    If isGB(ByteTemp) And (Offset >= 0) And (Offset <= 8177) Then
        strOut = strOut & Chr(GBOrder(Offset))
        idx = idx + 2
    Else
        strOut = strOut & Chr(ByteTemp(0))
        idx = idx + 1
    End If
    Loop

GB2BIG = strOut
End Function

Public Function BIG2GB(ByteBIG() As Byte) As String
On Error Resume Next
'Dim ByteBIG() As Byte
Dim ByteTemp(1) As Byte
Dim leng As Long, idx As Long
Dim strOut As String
Dim Offset As Long

If Not InitOK Then Call InitDATA
If Not InitOK Then
    BIG2GB = ByteBIG
    Exit Function
End If

'ByteBIG = StrConv(strBIG, vbFromUnicode)
leng = UBound(ByteBIG)
idx = 0
Do While idx <= leng
    ByteTemp(0) = ByteBIG(idx)
    ByteTemp(1) = ByteBIG(idx + 1)
    Offset = BIG5Offset(ByteTemp)
    If isBIG(ByteTemp) And (Offset >= 0) And (Offset <= 14757) Then
        strOut = strOut & Chr(BIG5Order(Offset))
        idx = idx + 1
    Else
        strOut = strOut & Chr(ByteTemp(0))
    End If
    idx = idx + 1
Loop
BIG2GB = strOut
End Function

Public Function CheckBIG(strSource As String) As Boolean
Dim idx As Long
Dim ByteTemp() As Byte
CheckBIG = False
For idx = 1 To Len(strSource)
    ByteTemp = StrConv(Mid(strSource, idx, 1), vbFromUnicode)
    If UBound(ByteTemp) > 0 Then
        If (ByteTemp(1) >= 64) And (ByteTemp(1) <= 126) Then
            CheckBIG = True
            Exit For
        End If
    End If
Next idx
End Function

Private Function GBOffset(ChrString() As Byte) As Long
'On Error GoTo ERROR_HANDLE
Dim Dl, Dh
    Dl = ChrString(0)
    Dh = ChrString(1)
    GBOffset = (Dl - 161) * 94 + (Dh - 161)
'    Exit Function
'ERROR_HANDLE:
'    GBOffset = -1
End Function

Private Function BIG5Offset(ChrString() As Byte) As Long
'On Error GoTo ERROR_HANDLE
Dim Dl, Dh
    Dl = ChrString(0)
    Dh = ChrString(1)
    If (Dh >= 64) And (Dh <= 126) Then _
        BIG5Offset = (Dl - 161) * 157 + (Dh - 64)
    If (Dh >= 161) And (Dh <= 254) Then _
        BIG5Offset = (Dl - 161) * 157 + 63 + (Dh - 161)
'    Exit Function
'ERROR_HANDLE:
'    BIG5Offset = -1
End Function

Private Function isGB(ChrString() As Byte) As Boolean
'On Error GoTo ERRORHANDLE
If UBound(ChrString) >= 1 Then
    If (ChrString(0) <= 161) And (ChrString(0) >= 247) Then
        isGB = False
    Else
        If (ChrString(1) <= 161) And (ChrString(1) >= 254) Then
            isGB = False
        Else
            isGB = True
        End If
    End If
Else
    isGB = False
End If
'Exit Function
'ERRORHANDLE:
'    isGB = False
End Function

Private Function isBIG(ChrString() As Byte) As Boolean
'On Error GoTo ERRORHANDLE
If UBound(ChrString) >= 1 Then
    If ChrString(0) < 161 Then
        isBIG = False
    Else
        If ((ChrString(1) >= 64) And (ChrString(1) <= 126)) Or ((ChrString(1) >= 161) And (ChrString(1) <= 254)) Then
            isBIG = True
        Else
            isBIG = False
        End If
    End If
Else
    isBIG = False
End If
'Exit Function
'ERRORHANDLE:
'    isBIG = False
End Function










