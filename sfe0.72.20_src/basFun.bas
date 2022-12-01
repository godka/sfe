Attribute VB_Name = "basFun"
Option Explicit


Public Function RndLong(i As Long) As Long
    RndLong = Int(Rnd * (i))
End Function


Public Function fmt(ByVal str As String, ByVal Length As Long) As String
    If Length > Len(str) Then
        fmt = String(Length - Len(str), " ")
    Else
        fmt = ""
    End If
    fmt = str & fmt
End Function
Public Function GetINISection(strSection As String)
    Dim MyKeys As String * 10000
    Dim EachElement
    Dim EachKey
    GetPrivateProfileSection strSection, MyKeys, 10000, G_Var.iniFileName
    ReDim FiftyItem(0)
    EachKey = Split(MyKeys, Chr(0))
    For Each EachElement In EachKey
        If EachElement = "" Then
            Exit For
        End If
        FiftyItem(UBound(FiftyItem)) = EachElement
        ReDim Preserve FiftyItem(UBound(FiftyItem) + 1)
    Next
End Function


Public Function GetINILong(strSection As String, StrKey As String) As Long
    GetINILong = GetPrivateProfileInt(strSection, StrKey, -65536, G_Var.iniFileName)
'    If GetINILong = -65536 Then
'        Err.Raise vbObjectError + 1, , "INI Section " & strSection & " Key " & StrKey & " Not found"
'    End If
End Function

Public Function GetINIStr(strSection As String, StrKey As String) As String
Dim tmpStr() As Byte
Dim returnval As Long
Dim ttt As String
    ReDim tmpStr(255)
    returnval = GetPrivateProfileString(strSection, StrKey, "", tmpStr(0), 256, G_Var.iniFileName)
    If returnval > 0 Then
        ReDim Preserve tmpStr(returnval - 1)
        If Charset = "GBK" Then
            GetINIStr = StrConv(tmpStr, vbUnicode)
        Else
            ttt = tmpStr
            GetINIStr = GBKtoUnicode(ttt)
        End If
    Else
        GetINIStr = ""
    End If
   ' GetINIStr = tmpstr
End Function

Public Sub PutINIStr(strSection As String, StrKey As String, strVal As String)
Dim returnval As Long
    returnval = WritePrivateProfileString(strSection, StrKey, strVal, G_Var.iniFileName)

End Sub



' 打开二进制文件
' status = "R"   读
' status = "W"   写，备份文件
' Status = "WN"  写新文件，可以比原来文件小，备份文件
Public Function OpenBin(filename As String, status As String) As Long
   OpenBin = FreeFile()
   Select Case UCase(status)
   Case "R"
       If Dir(filename) = "" Then
           Err.Raise vbObjectError + 1, , "File " & filename & " not exist"
       End If
       Open filename For Binary Access Read As OpenBin
   Case "W"
       FileCopy filename, filename & ".bak"
       Open filename For Binary Access Write As OpenBin
   Case "WN"
       If Dir(filename & ".bak") <> "" Then
           Kill filename & ".bak"
       End If
       If Dir(filename) <> "" Then
           Name filename As filename & ".bak"
       End If
       Open filename For Binary Access Write As OpenBin
   End Select
   
End Function

' 转换form中所有空间的caption字符集, 用于新的窗口处理字符串

Public Sub ConvertForm(frm As Form)
Dim i As Long
    frm.Caption = StrUnicode(frm.Caption)
 
    For i = 0 To frm.Controls.Count - 1
         Call SetCaption(frm.Controls(i))
    Next i

   
End Sub

' 设置窗体的字符串信息和窗体标题，用于50指令解析窗口
Public Sub Set50Form(frm As Form, id As Long)
Dim s1 As String

    s1 = GetINIStr("Kdef50", "sub" & id)
    If s1 = "" Then
        s1 = GetINIStr("Kdef50", "Other")
    End If

    frm.Caption = StrUnicode2("50指令") & id & " - " & s1

    On Error Resume Next
    frm.txtNote.Text = StrUnicode(frm.txtNote.Text)

End Sub



' 设置控件的caption属性
Public Sub SetCaption(oo As Object)
    If TypeOf oo Is Menu Then
        oo.Caption = StrUnicode(oo.Caption)
    End If
    If TypeOf oo Is CommandButton Then
        oo.Caption = StrUnicode(oo.Caption)
        oo.ToolTipText = StrUnicode(oo.ToolTipText)
    End If
    If TypeOf oo Is label Then
        oo.Caption = StrUnicode(oo.Caption)
        oo.ToolTipText = StrUnicode(oo.ToolTipText)
    End If
    If TypeOf oo Is CheckBox Then
        oo.Caption = StrUnicode(oo.Caption)
    End If
    If TypeOf oo Is OptionButton Then
        oo.Caption = StrUnicode(oo.Caption)
    End If
    If TypeOf oo Is Frame Then
        oo.Caption = StrUnicode(oo.Caption)
    End If
    If TypeOf oo Is PictureBox Then
        oo.ToolTipText = StrUnicode(oo.ToolTipText)
    End If
    If TypeOf oo Is ListBox Then
        oo.ToolTipText = StrUnicode(oo.ToolTipText)
    End If
    
    
End Sub


Public Function Max_V(X1 As Variant, X2 As Variant) As Variant
    Max_V = IIf(X1 > X2, X1, X2)
End Function

Public Function Min_V(X1 As Variant, X2 As Variant) As Variant
    Min_V = IIf(X1 < X2, X1, X2)
End Function

' 控制x在（xmin，xmax）范围
Public Function RangeValue(X As Variant, ByVal Xmin As Long, ByVal Xmax As Long) As Variant
    If X > Xmax Then
        X = Xmax
    End If
    If X < Xmin Then
        X = Xmin
    End If
    RangeValue = X
End Function


'  去掉字符串前后多余的空格，中间多个连续的空格变成一个空格
Public Function SubSpace(s As String) As String
Dim tmps As String
Dim pos As String
    tmps = Trim(s)
    Do
        pos = InStr(tmps, "  ")
        If pos = 0 Then Exit Do
        tmps = Mid(tmps, 1, pos - 1) & Mid(tmps, pos + 1)
    Loop
    SubSpace = tmps
End Function


Public Function Byte2String(s() As Byte) As String
Dim i As Long
Dim Length As Long
Dim tmpbyte() As Byte
Dim MaxArray As Long
    Length = 0
    MaxArray = UBound(s, 1)
    For i = 0 To MaxArray
        If s(i) = 0 Then
            Length = i
            Exit For
        End If
    Next i
    If Length = 0 Then Length = MaxArray + 1
    ReDim tmpbyte(Length - 1)
    For i = 0 To Length - 1
        tmpbyte(i) = s(i)
    Next i
    Byte2String = StrConv(tmpbyte, vbUnicode)
    
End Function

Public Function HexInt(i As Integer) As String
Dim Length As Long
    HexInt = Hex(i)
    Length = Len(HexInt)
    If Length < 4 Then
        HexInt = String(4 - Length, "0") & HexInt
    End If
End Function


