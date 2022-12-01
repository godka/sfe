VERSION 5.00
Begin VB.Form frmWarEditNew 
   Caption         =   "战斗编辑"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13125
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   875
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   375
      Left            =   11640
      TabIndex        =   12
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   375
      Left            =   10560
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      Height          =   5175
      Left            =   9480
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   9
      Top             =   1560
      Width           =   5295
   End
   Begin VB.CommandButton CmdGetExcel 
      Caption         =   "导出excel"
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton CmdPutExcel 
      Caption         =   "导入excel"
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ListBox ListItem 
      Columns         =   3
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      ToolTipText     =   "双击修改"
      Top             =   1560
      Width           =   9375
   End
   Begin VB.CommandButton cmdLoadRecord 
      Caption         =   "读取战斗"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveRecord 
      Caption         =   "保存战斗"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "增加战斗"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除战斗"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox ComboType 
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox ComboNumber 
      Height          =   300
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label10"
      Height          =   255
      Left            =   11040
      TabIndex        =   13
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   9600
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmWarEditNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mapdata() As Integer
Dim MapIdx() As Long
Dim mapBig As Long
Const LineColor = vbBlack
Dim Tx As Long, Ty As Long
Dim PersonK As Long, EnemyK As Long

Dim MouseLock As Boolean
Private TypeNumber As Long

'Private TypeName() As String
Private TypeDataItem() As Long
Private Wlong() As Long
Dim Warperson() As WarType, Warenemy() As WarType, WarPersonNum As Long, WarEnemyNum As Long
Private Type DataItem_type
    ByteNum As Long
    isStr As Long
    isName As Long
    isMap As Long
    TagType As Long
    TypeNum As Long
    Ref As Long
    Name As String
    note As String
    offset As Long
End Type

Private Type WarType
    X As Long
    Y As Long
    ItemPos As Long
    XPos As Long
    YPos As Long
    PersonID As Long
End Type

Private Type Data_type
    Int As Integer
    str As String
End Type


Private Type TypeData_Type
    Name As String
    ItemNumber As Long
    DataItem() As DataItem_type
    DataNumber As Long
    DataV() As Data_type
    DataName() As String
    
End Type

Private TypeData() As TypeData_Type

Private Sub CmdGetExcel_Click()
Dim i As Integer
Dim kuang As OPENFILENAME
Dim filename As String
    kuang.lStructSize = Len(kuang)
    kuang.hwndOwner = Me.hWnd
    kuang.hInstance = App.hInstance
    kuang.lpstrFile = Space(254)
    kuang.nMaxFile = 255
    kuang.lpstrFileTitle = Space(254)
    kuang.nMaxFileTitle = 255
    kuang.lpstrInitialDir = App.Path
    kuang.flags = 6148
    '过虑对话框文件类型
    kuang.lpstrFilter = "xls文件(*.xls)" + Chr$(0) + "*.xls" + Chr$(0)
    '对话框标题栏文字
    kuang.lpstrTitle = "保存文件的路径及文件名..."
    i = GetSaveFileName(kuang) '显示保存文件对话框
    If i >= 1 Then '取得对话中用户选择输入的文件名及路径
        filename = kuang.lpstrFile
        filename = Left(filename, InStr(filename, Chr(0)) - 1)
    End If
    If Len(filename) = 0 Then Exit Sub
    '保存代码
    getexcel filename
    MsgBox "Saved in " & filename & ".xls"
End Sub

Private Sub cmdLoadRecord_Click()
Load_W
End Sub

Private Sub CmdPutExcel_Click()
Dim ofn As OPENFILENAME
Dim Rtn As String
Dim tmpStr As String
Dim filenum As Long
Dim i As Long, j As Long, k As Long
    tmpStr = "xls文件|*.xls"
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = Replace$(tmpStr, "|", Chr$(0)) + vbNullChar + vbNullChar
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = App.Path
    ofn.lpstrTitle = "Open File"
    ofn.flags = 6148

    Rtn = GetOpenFileName(ofn)

    If Rtn < 1 Then Exit Sub
    
    putexcel ofn.lpstrFile
    ComboType.Clear
    For i = 0 To TypeNumber - 1
        ComboType.AddItem TypeData(i).Name
    Next i
    ComboType.ListIndex = 0
    ComboType_click
End Sub

Private Sub cmdSaveRecord_Click()
Dim i As Long, j As Long, k As Long
Dim filenum As Long
Dim tmpbyte() As Byte
Dim Length As Long

    'If ComboRecord.ListIndex = -1 Then Exit Sub
    filenum = OpenBin(G_Var.JYPath & G_Var.WarDefine, "WN")
    For i = 0 To TypeNumber - 1
        For j = 0 To TypeData(i).DataNumber - 1
            For k = 0 To TypeData(i).ItemNumber - 1
                If TypeData(i).DataItem(k).isStr = 0 Then
                    Put #filenum, , TypeData(i).DataV(k, j).Int
                Else
                    Call UnicodetoBIG5(TypeData(i).DataV(k, j).str, Length, tmpbyte)
                    ReDim Preserve tmpbyte(TypeData(i).DataItem(k).ByteNum - 1)
                    Put #filenum, , tmpbyte
                End If
            Next k
        Next j
    Next i
    Close #filenum
End Sub

Private Sub Command1_Click()
Dim i As Long, typeN As Long, Index As Long
    mapBig = mapBig + 1
    pic1.Height = 64 * mapBig
    pic1.Width = 64 * mapBig
    pic1.ScaleHeight = 64 * mapBig
    pic1.ScaleWidth = 64 * mapBig
    Index = ComboNumber.ListIndex
    If Index = -1 Then Exit Sub
    typeN = ComboType.ListIndex
    If typeN = -1 Then Exit Sub
    'ListItem.Clear
    For i = 0 To TypeData(typeN).ItemNumber - 1
        If TypeData(typeN).DataItem(i).isMap = 1 Then
            drawScene TypeData(typeN).DataV(i, Index).Int
            Exit For
        End If
    Next i
    Label1.Caption = mapBig
End Sub
Private Sub Command2_Click()
Dim i As Long, typeN As Long, Index As Long
    mapBig = mapBig - 1
    pic1.Height = 64 * mapBig
    pic1.Width = 64 * mapBig
    pic1.ScaleHeight = 64 * mapBig
    pic1.ScaleWidth = 64 * mapBig
    Index = ComboNumber.ListIndex
    If Index = -1 Then Exit Sub
    typeN = ComboType.ListIndex
    If typeN = -1 Then Exit Sub
    'ListItem.Clear
    For i = 0 To TypeData(typeN).ItemNumber - 1
        If TypeData(typeN).DataItem(i).isMap = 1 Then
            drawScene TypeData(typeN).DataV(i, Index).Int
            Exit For
        End If
    Next i
    Label1.Caption = mapBig
End Sub
Private Sub Form_Load()
    Call ConvertForm(Me)
    mapBig = 7
Dim filenum As Long
Dim idxlong As Long
Dim tmp As Long, i As Long

    filenum = OpenBin(G_Var.JYPath & G_Var.WarMapDefIDX, "R")
        idxlong = LOF(filenum) / 4
        ReDim MapIdx(idxlong - 1)
        MapIdx(0) = 0
        For i = 1 To idxlong - 1
            Get filenum, , MapIdx(i)
        Next i
    Close (filenum)
    
    Tx = 0: Ty = 0
    pic1.Height = 64 * mapBig
    pic1.Width = 64 * mapBig
    pic1.ScaleHeight = 64 * mapBig
    pic1.ScaleWidth = 64 * mapBig
    Label10.Caption = "(0,0)"
    PersonK = -1: EnemyK = 1
    Label1.Caption = mapBig
Load_W_Type
Load_W
MouseLock = True
c_Skinner.AttachSkin Me.hWnd
End Sub
Private Sub Load_W_Type()
Dim i As Long
Dim j As Long
Dim k As Long
Dim ll As Long
Dim num As Long
Dim tmpStr() As String
Dim tmpstr2() As String
Dim NumArray As Long
Dim NumType As Long

    TypeNumber = GetINILong("W_Modify", "TypeNumber")
    ReDim TypeData(TypeNumber - 1)
    
    ReDim TypeDataItem(TypeNumber - 1)
    For i = 0 To TypeNumber - 1
        TypeData(i).Name = GetINIStr("W_Modify", "TypeName" & i)
        TypeDataItem(i) = GetINILong("W_Modify", "TypeDataItem" & i)
    Next i
    
    For i = 0 To TypeNumber - 1
        num = 0
        For j = 0 To TypeDataItem(i) - 1
            tmpStr = Split(SubSpace(GetINIStr("W_Modify", "Data(" & i & "," & j & ")")), " ")
            num = num + CLng(tmpStr(0)) * CLng(tmpStr(1))
        Next j
        TypeData(i).ItemNumber = num
        ReDim TypeData(i).DataItem(num - 1)
        num = 0
        j = 0
        Do While j < TypeDataItem(i)
            tmpStr = Split(SubSpace(GetINIStr("W_Modify", "Data(" & i & "," & j & ")")), " ")
            NumArray = CLng(tmpStr(0))
            NumType = CLng(tmpStr(1))
            For k = 1 To NumArray
                TypeData(i).DataItem(num).ByteNum = CLng(tmpStr(2))
                TypeData(i).DataItem(num).isStr = CLng(tmpStr(3))
                TypeData(i).DataItem(num).isName = CLng(tmpStr(4))
                TypeData(i).DataItem(num).Ref = CLng(tmpStr(5))
                TypeData(i).DataItem(num).Name = tmpStr(6) & IIf(NumArray > 1, k, "")
                TypeData(i).DataItem(num).note = (tmpStr(7))
                TypeData(i).DataItem(num).isMap = (tmpStr(8))
                TypeData(i).DataItem(num).TagType = (tmpStr(9))
                TypeData(i).DataItem(num).TypeNum = (tmpStr(10))
                num = num + 1
                For ll = 2 To NumType
                    tmpstr2 = Split(SubSpace(GetINIStr("W_Modify", "Data(" & i & "," & j + ll - 1 & ")")), " ")
                    TypeData(i).DataItem(num).ByteNum = CLng(tmpstr2(2))
                    TypeData(i).DataItem(num).isStr = CLng(tmpstr2(3))
                    TypeData(i).DataItem(num).isName = CLng(tmpstr2(4))
                    TypeData(i).DataItem(num).Ref = CLng(tmpstr2(5))
                    TypeData(i).DataItem(num).Name = tmpstr2(6) & IIf(NumArray > 1, k, "")
                    TypeData(i).DataItem(num).note = tmpstr2(7)
                    TypeData(i).DataItem(num).isMap = tmpstr2(8)
                    TypeData(i).DataItem(num).TagType = (tmpStr(9))
                    TypeData(i).DataItem(num).TypeNum = (tmpStr(10))
                    num = num + 1
                Next ll
            Next k
            j = j + NumType
        Loop
    Next i
    
End Sub

Private Sub Load_W()
Dim i As Long
Dim j As Long
Dim k As Long
Dim filenum As Long
Dim DataLong As Long
Dim offset As Long
Dim tmpbyte() As Byte

    filenum = OpenBin(G_Var.JYPath & G_Var.WarDefine, "R")
    For i = 0 To TypeNumber - 1
        DataLong = 0
        For j = 0 To TypeData(i).ItemNumber - 1
            DataLong = DataLong + TypeData(i).DataItem(j).ByteNum
        Next j
        
        TypeData(i).DataNumber = LOF(filenum) / DataLong
        ReDim TypeData(i).DataV(TypeData(i).ItemNumber - 1, TypeData(i).DataNumber - 1)
        ReDim TypeData(i).DataName(TypeData(i).DataNumber - 1)
        Seek #filenum, 1
        For j = 0 To TypeData(i).DataNumber - 1
            For k = 0 To TypeData(i).ItemNumber - 1
                If TypeData(i).DataItem(k).isStr = 0 Then
                    Get #filenum, , TypeData(i).DataV(k, j).Int
                Else
                    ReDim tmpbyte(TypeData(i).DataItem(k).ByteNum - 1)
                    Get #filenum, , tmpbyte
                    TypeData(i).DataV(k, j).str = Big5toUnicode(tmpbyte, TypeData(i).DataItem(k).ByteNum)
                End If
                If TypeData(i).DataItem(k).isName = 1 Then
                    TypeData(i).DataName(j) = TypeData(i).DataV(k, j).str
                End If
            Next k
        Next j
    Next i
    
    Close #filenum
    
    ComboType.Clear
    For i = 0 To TypeNumber - 1
        ComboType.AddItem TypeData(i).Name
    Next i
    ComboType.ListIndex = 0

End Sub
Private Sub ComboType_click()
Dim Index As Long
Dim i As Long
    Index = ComboType.ListIndex
    If Index = -1 Then Exit Sub
    
    ComboNumber.Clear
    For i = 0 To TypeData(Index).DataNumber - 1
        ComboNumber.AddItem i & " " & TypeData(Index).DataName(i)
    Next i
    ComboNumber.ListIndex = 0

End Sub
Private Sub ComboNumber_click()
Dim typeN As Long
Dim Index As Long
Dim i As Long, a As Long
Dim tmpStr As String
'Dim warnum As Long
    Index = ComboNumber.ListIndex
    If Index = -1 Then Exit Sub
    typeN = ComboType.ListIndex
    If typeN = -1 Then Exit Sub
    ListItem.Clear
    For i = 0 To TypeData(typeN).ItemNumber - 1
        If TypeData(typeN).DataItem(i).isMap = 1 Then
            drawScene TypeData(typeN).DataV(i, Index).Int
        End If
        ListItem.AddItem GenListstr(typeN, Index, i)
    Next i
    ListItem.ListIndex = 0
    
End Sub

' 生成list字符串
' i type    j   datanumber     k itemnumber
Private Function GenListstr(i As Long, j As Long, k As Long) As String
Dim tmpStr As String
Dim Tstring As String
Dim ll As Long
        tmpStr = TypeData(i).DataItem(k).Name
        ll = LenB(StrConv(TypeData(i).DataItem(k).Name, vbFromUnicode))
         If ll < 18 Then
            tmpStr = tmpStr & Space(15 - ll)
         End If
        
        If TypeData(i).DataItem(k).isStr = 1 Then
            tmpStr = tmpStr & TypeData(i).DataV(k, j).str
        Else
            tmpStr = tmpStr & fmt(TypeData(i).DataV(k, j).Int, 6)
        End If
        
        
        If TypeData(i).DataItem(k).Ref >= 0 And TypeData(i).DataV(k, j).Int >= 0 Then
            'tmpstr = tmpstr & " " & TypeData(TypeData(i).DataItem(k).Ref).DataName(TypeData(i).DataV(k, j).Int)
            Select Case TypeData(i).DataItem(k).Ref
            Case 0
            Case 1
                Tstring = Person(TypeData(i).DataV(k, j).Int).Name1
            Case 2
                Tstring = Things(TypeData(i).DataV(k, j).Int).Name1
            Case 3
                Tstring = Big5toUnicode(Scene(TypeData(i).DataV(k, j).Int).Name1, 10)
            Case 4
                Tstring = WuGong(TypeData(i).DataV(k, j).Int).Name1
            End Select
            tmpStr = tmpStr & " " & Tstring
        End If
    GenListstr = tmpStr
End Function

Private Sub ListItem_Click()
Dim i As Long, j As Long, k As Long
    i = ComboType.ListIndex
    j = ComboNumber.ListIndex
    k = ListItem.ListIndex
    If i < 0 Or j < 0 Or k < 0 Then Exit Sub
    DrawMassPerson (k)
    MDIMain.StatusBar1.Panels(1).Text = TypeData(i).DataItem(k).note
    
End Sub


Private Sub ListItem_DblClick()
Dim i As Long, j As Long, k As Long
Dim num As Long
Dim tmp As String
    i = ComboType.ListIndex
    j = ComboNumber.ListIndex
    k = ListItem.ListIndex
    If i < 0 Or j < 0 Or k < 0 Then Exit Sub
    Load frmChangeWValue
    
    frmChangeWValue.Label1.Caption = TypeData(i).DataItem(k).Name
    If TypeData(i).DataItem(k).isStr = 0 Then
        frmChangeWValue.Text1.Text = TypeData(i).DataV(k, j).Int
    Else
        frmChangeWValue.Text1.Text = TypeData(i).DataV(k, j).str
    End If
    If TypeData(i).DataItem(k).Ref > 0 Then
        frmChangeWValue.Combo1.Clear
        frmChangeWValue.Combo1.AddItem LoadResStr(10602)
        'For num = 0 To TypeData(TypeData(i).DataItem(k).Ref).DataNumber - 1
        '    frmChangeWValue.Combo1.AddItem num & TypeData(TypeData(i).DataItem(k).Ref).DataName(num)
        'Next num
        Select Case TypeData(i).DataItem(k).Ref
        Case 0
        Case 1
            For num = 0 To PersonNum - 1
                frmChangeWValue.Combo1.AddItem num & Person(num).Name1
            Next num
        Case 2
            For num = 0 To Thingsnum - 1
                frmChangeWValue.Combo1.AddItem num & Things(num).Name1
            Next num
        Case 3
            For num = 0 To Scenenum - 1
                frmChangeWValue.Combo1.AddItem num & Big5toUnicode(Scene(num).Name1, 10)
            Next num
        Case 4
            For num = 0 To WuGongnum - 1
                frmChangeWValue.Combo1.AddItem num & WuGong(num).Name1
            Next num
        End Select
        frmChangeWValue.Combo1.ListIndex = TypeData(i).DataV(k, j).Int + 1
        frmChangeWValue.Text1.Enabled = False
    Else
        frmChangeWValue.Combo1.Visible = False
    End If
    
    frmChangeWValue.Show vbModal
    
    If frmChangeWValue.OK = 1 Then
        If TypeData(i).DataItem(k).isStr = 0 Then
            TypeData(i).DataV(k, j).Int = frmChangeWValue.Text1.Text
        Else
            TypeData(i).DataV(k, j).str = frmChangeWValue.Text1.Text
        End If
        
        ListItem.List(k) = GenListstr(i, j, k)
'        ListItem.ForeColor = vbRed
    End If
    drawPix
    Unload frmChangeWValue
End Sub

Public Sub getexcel(filename As String)
Dim i As Long, j As Long, k As Long, m As Long
Dim tmp() As String
Dim tmp2() As Variant
Dim Excel As Object
Dim Book As Object
Dim Sheet As Object

'Start a new workbook in Excel
Set Excel = CreateObject("Excel.Application")
'表的sheet总数
Excel.SheetsInNewWorkbook = TypeNumber
Set Book = Excel.Workbooks.Add

For m = 0 To TypeNumber - 1
    '判断横竖
    Set Sheet = Excel.Worksheets(m + 1)
    Sheet.Name = TypeData(m).Name
    If TypeData(m).ItemNumber > 256 Then
        
        'read name in memory ,heritos
        ReDim tmp(TypeData(m).ItemNumber, 1)
        For i = 0 To TypeData(m).ItemNumber - 1
            tmp(i, 0) = TypeData(m).DataItem(i).Name
        Next i

        'read int in memory,vertious
        ReDim tmp2(TypeData(m).ItemNumber, TypeData(m).DataNumber)
        For k = 0 To TypeData(m).ItemNumber - 1
            For j = 0 To TypeData(m).DataNumber - 1
                If TypeData(m).DataItem(j).isStr = 0 Then
                    tmp2(k, j) = TypeData(m).DataV(k, j).Int
                Else
                    tmp2(k, j) = TypeData(m).DataV(k, j).str
                End If
            Next j
        Next k
        Sheet.Range("A1").Resize(TypeData(m).ItemNumber, 1).Value = tmp
        Sheet.Range("B1").Resize(TypeData(m).ItemNumber, TypeData(m).DataNumber).Value = tmp2
    Else
    
        'read name in memory ,heritos
        ReDim tmp(1, TypeData(m).ItemNumber)
        For i = 0 To TypeData(m).ItemNumber - 1
            tmp(0, i) = TypeData(m).DataItem(i).Name
        Next i

        'read int in memory,vertious
        ReDim tmp2(TypeData(m).DataNumber, TypeData(m).ItemNumber)
        For j = 0 To TypeData(m).ItemNumber - 1
            For k = 0 To TypeData(m).DataNumber - 1
                If TypeData(m).DataItem(j).isStr = 0 Then
                    tmp2(k, j) = TypeData(m).DataV(j, k).Int
                Else
                    tmp2(k, j) = TypeData(m).DataV(j, k).str
                End If
            Next k
        Next j
        Sheet.Range("A1").Resize(1, TypeData(m).ItemNumber).Value = tmp
        Sheet.Range("A2").Resize(TypeData(m).DataNumber, TypeData(m).ItemNumber).Value = tmp2
    End If
Next m
    

Excel.ActiveWorkbook.SaveAs filename
Excel.Quit
Set Excel = Nothing
End Sub
Public Sub putexcel(filename As String)
Dim i As Long, j As Long, k As Long, m As Long
'Dim tmp() As String
Dim tmp2() As Variant
Dim Excel As Object
Dim Book As Object
Dim Sheet As Object

'Start a new workbook in Excel
Set Excel = CreateObject("Excel.Application")
Set Book = Excel.Workbooks.Open(filename)
'逆操作
For m = 0 To TypeNumber - 1
    Set Sheet = Excel.Worksheets(m + 1)
    If TypeData(m).ItemNumber > 256 Then
        TypeData(m).ItemNumber = Sheet.UsedRange.Rows.Count
        TypeData(m).DataNumber = Sheet.UsedRange.Columns.Count - 1
        ReDim TypeData(m).DataV(TypeData(m).ItemNumber - 1, TypeData(m).DataNumber - 1)
        tmp2 = Sheet.Range("B1").Resize(Sheet.UsedRange.Rows.Count, Sheet.UsedRange.Columns.Count - 1).Value
        For k = 0 To TypeData(m).ItemNumber - 1
            For j = 0 To TypeData(m).DataNumber - 1
                If TypeData(m).DataItem(j).isStr = 0 Then
                    TypeData(m).DataV(k, j).Int = tmp2(k + 1, j + 1)
                Else
                    TypeData(m).DataV(k, j).str = tmp2(k + 1, j + 1)
                End If
                If TypeData(m).DataItem(k).isName = 1 Then
                    TypeData(m).DataName(j) = TypeData(m).DataV(k, j).str
                End If
            Next j
        Next k
    Else
        tmp2 = Sheet.Range("A2").Resize(Sheet.UsedRange.Rows.Count - 1, Sheet.UsedRange.Columns.Count).Value
        TypeData(m).ItemNumber = Sheet.UsedRange.Columns.Count
        TypeData(m).DataNumber = Sheet.UsedRange.Rows.Count - 1
        ReDim TypeData(m).DataName(TypeData(m).DataNumber - 1)
        ReDim TypeData(m).DataV(TypeData(m).ItemNumber - 1, TypeData(m).DataNumber - 1)
        For j = 0 To TypeData(m).ItemNumber - 1
            For k = 0 To TypeData(m).DataNumber - 1
                If TypeData(m).DataItem(j).isStr = 0 Then
                    TypeData(m).DataV(j, k).Int = tmp2(k + 1, j + 1)
                Else
                    TypeData(m).DataV(j, k).str = tmp2(k + 1, j + 1)
                End If
                If TypeData(m).DataItem(j).isName = 1 Then
                    TypeData(m).DataName(k) = TypeData(m).DataV(j, k).str
                End If
            Next k
        Next j
    End If
Next m
Excel.Quit
Set Excel = Nothing
End Sub

Public Sub drawP(ByVal X As Long, ByVal Y As Long, ByVal colorS As Long, ByVal flag As Boolean)
Dim i As Long
    
    For i = 0 To mapBig - 1
            pic1.Line (X + i, Y)-(X + i, Y + mapBig), colorS
    Next i
    If flag = True Then
        pic1.Line (X, Y + mapBig)-(X + mapBig, Y + mapBig), LineColor
        pic1.Line (X, Y)-(X, Y + mapBig), LineColor
        pic1.Line (X, Y)-(X + mapBig, Y), LineColor
        pic1.Line (X + mapBig, Y)-(X + mapBig, Y + mapBig), LineColor
    End If
End Sub
Public Sub loadWmap(warnum1 As Long)
Dim filenum As Long
Dim idxlong As Long
Dim tmp As Long, i As Long
    
    filenum = OpenBin(G_Var.JYPath & G_Var.WarMapDefGRP, "R")
        ReDim Mapdata(0 To 63, 0 To 63, 0 To 1)
        Seek filenum, MapIdx(warnum1) + 1
        Get filenum, , Mapdata
    Close (filenum)
    
End Sub
Public Function getcolor(ByVal data As Long) As Long
    If ((data > 468 And data < 472) Or (data > 351 And data < 357) Or (data > 363 And data < 368) Or (data > 392 And data < 399)) Then
        getcolor = RGB(222, 222, 222)
    ElseIf ((data > 661 And data < 675)) Then
        getcolor = RGB(64, 28, 4)
    ElseIf ((data > 1 And data < 35 And data <> 6) Or (data > 104 And data < 151) Or (data > 194 And data < 224) Or (data > 530 And data < 544) Or (data > 674 And data < 679) Or (data > 356 And data < 392) Or (data > 41 And data < 70) Or (data > 36 And data < 41)) Then
        getcolor = RGB(156, 116, 60)
    ElseIf ((data > 154 And data < 191) Or data = 511) Then
        getcolor = RGB(52, 52, 252)
    ElseIf ((data > 305 And data < 331) Or (data > 497 And data < 518) Or (data > 543 And data < 627) Or (data > 678 And data < 699)) Then
        getcolor = RGB(108, 108, 108)
    ElseIf (data <> 0) Then
        getcolor = RGB(28, 104, 16)
    End If
    
End Function
'draw war scene in right pic,ugly script for terrible type
Public Sub drawScene(ByVal Scenenum As Long)
Dim i As Long, j As Long
Dim typeN As Long, Index As Long

Dim warpersonNumY As Long, warenemyNumY As Long
Dim WarPersonID As Long, WarEnemyID As Long
    Index = ComboNumber.ListIndex
    If Index = -1 Then Exit Sub
    typeN = ComboType.ListIndex
    If typeN = -1 Then Exit Sub
    
    loadWmap (Scenenum)
    For i = 0 To 64 - 1
        For j = 0 To 64 - 1
            If Mapdata(i, j, 1) / 2 <> 0 And Mapdata(i, j, 1) / 2 <> -1 Then
                drawP i * mapBig, j * mapBig, getcolor(Mapdata(i, j, 1) / 2), True
            Else
                drawP i * mapBig, j * mapBig, getcolor(Mapdata(i, j, 0) / 2), True
            End If
        Next j
    Next i
    WarPersonNum = 0: WarEnemyNum = 0
    ReDim Warperson(0)
    ReDim Warenemy(0)
    For i = 0 To TypeData(typeN).ItemNumber - 1
        Select Case TypeData(typeN).DataItem(i).TypeNum
        Case 1
            WarPersonNum = WarPersonNum + 1
            'ReDim Preserve Warperson(WarPersonNum - 1)
            Warperson(WarPersonNum - 1).X = TypeData(typeN).DataV(i, Index).Int
            Warperson(WarPersonNum - 1).XPos = i
        Case 2
            warpersonNumY = warpersonNumY + 1
            Warperson(warpersonNumY - 1).Y = TypeData(typeN).DataV(i, Index).Int
            Warperson(warpersonNumY - 1).YPos = i
        Case 3
            WarEnemyNum = WarEnemyNum + 1
            'ReDim Preserve Warenemy(WarEnemyNum - 1)
            Warenemy(WarEnemyNum - 1).X = TypeData(typeN).DataV(i, Index).Int
            Warenemy(WarEnemyNum - 1).XPos = i
        Case 4
            warenemyNumY = warenemyNumY + 1
            Warenemy(warenemyNumY - 1).Y = TypeData(typeN).DataV(i, Index).Int
            Warenemy(warenemyNumY - 1).YPos = i
        Case 5
            WarPersonID = WarPersonID + 1
            ReDim Preserve Warperson(WarPersonID - 1)
            Warperson(WarPersonID - 1).PersonID = TypeData(typeN).DataV(i, Index).Int
            Warperson(WarPersonID - 1).ItemPos = i
        Case 6
            WarEnemyID = WarEnemyID + 1
            ReDim Preserve Warenemy(WarEnemyID - 1)
            Warenemy(WarEnemyID - 1).PersonID = TypeData(typeN).DataV(i, Index).Int
            Warenemy(WarEnemyID - 1).ItemPos = i
        End Select
    Next i
    
    drawPix
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     'drawP X * mapBig, Y * mapBig, vbYellow, True
Dim i As Long, j As Long, k As Long
    i = ComboType.ListIndex
    j = ComboNumber.ListIndex
    'K = ListItem.ListIndex
    
    If MouseLock = False Then
        If Mapdata(Tx, Ty, 1) / 2 <> 0 And Mapdata(Tx, Ty, 1) / 2 <> -1 Then
            drawP Tx * mapBig, Ty * mapBig, getcolor(Mapdata(Tx, Ty, 1) / 2), True
        Else
            drawP Tx * mapBig, Ty * mapBig, getcolor(Mapdata(Tx, Ty, 0) / 2), True
        End If
        Tx = Int(X / mapBig): Ty = Int(Y / mapBig)
        
        If PersonK >= 0 And EnemyK < 0 Then
            TypeData(i).DataV(Warperson(PersonK).XPos, j).Int = Tx
            TypeData(i).DataV(Warperson(PersonK).YPos, j).Int = Ty
            Warperson(PersonK).X = Tx: Warperson(PersonK).Y = Ty
            ListItem.List(Warperson(PersonK).XPos) = GenListstr(i, j, Warperson(PersonK).XPos)
            ListItem.List(Warperson(PersonK).YPos) = GenListstr(i, j, Warperson(PersonK).YPos)
        Else
            TypeData(i).DataV(Warenemy(EnemyK).XPos, j).Int = Tx
            TypeData(i).DataV(Warenemy(EnemyK).YPos, j).Int = Ty
            Warenemy(EnemyK).X = Tx: Warenemy(EnemyK).Y = Ty
            ListItem.List(Warenemy(EnemyK).XPos) = GenListstr(i, j, Warenemy(EnemyK).XPos)
            ListItem.List(Warenemy(EnemyK).YPos) = GenListstr(i, j, Warenemy(EnemyK).YPos)
        End If
        
        drawPix
        drawP Int(X / mapBig) * mapBig, Int(Y / mapBig) * mapBig, vbYellow, True
    End If
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Or Button = 2 Or Button = 3 Then Exit Sub
    Label10.Caption = "(" & Int(X / mapBig) & "," & Int(Y / mapBig) & ")"
End Sub

Private Sub DrawMassPerson(k As Long)
Dim i As Long

    MouseLock = True
    drawPix
    'judge item in pic1
    For i = 0 To WarPersonNum - 1
        If Warperson(i).ItemPos = k Or Warperson(i).XPos = k Or Warperson(i).YPos = k Then
            If Warperson(i).X > 0 And Warperson(i).Y > 0 And Warperson(i).PersonID >= 0 Then
                drawP Warperson(i).X * mapBig, Warperson(i).Y * mapBig, vbYellow, True
                Tx = Warperson(i).X: Ty = Warperson(i).Y: PersonK = i: EnemyK = -1
                MouseLock = False
                Exit For
            End If
        End If
    Next i
    For i = 0 To WarEnemyNum - 1
        If Warenemy(i).ItemPos = k Or Warenemy(i).XPos = k Or Warenemy(i).YPos = k Then
            If Warenemy(i).X > 0 And Warenemy(i).Y > 0 And Warenemy(i).PersonID >= 0 Then
                drawP Warenemy(i).X * mapBig, Warenemy(i).Y * mapBig, vbYellow, True
                Tx = Warenemy(i).X: Ty = Warenemy(i).Y: EnemyK = i: PersonK = -1
                MouseLock = False
                Exit For
            End If
        End If
    Next i
End Sub
Public Sub drawPix()
Dim i As Long
    For i = 0 To WarPersonNum - 1
        If Warperson(i).X > 0 And Warperson(i).Y > 0 And Warperson(i).PersonID >= 0 Then
            drawP Warperson(i).X * mapBig, Warperson(i).Y * mapBig, vbBlue, True
        End If
    Next i
    
    For i = 0 To WarEnemyNum - 1
        If Warenemy(i).X > 0 And Warenemy(i).Y > 0 And Warenemy(i).PersonID >= 0 Then
            drawP Warenemy(i).X * mapBig, Warenemy(i).Y * mapBig, vbRed, True
        End If
    Next i
End Sub
Private Sub CmdAdd_Click()
Dim Index As Long
Dim CurrentID As Long
Dim i As Long
Dim tmpStr As String
    Index = ComboType.ListIndex
    If Index < 0 Then Exit Sub
    CurrentID = ComboNumber.ListIndex
    If CurrentID < 0 Then Exit Sub
    tmpStr = LoadResStr(10509) & Trim(TypeData(Index).Name) & " (" & LoadResStr(10511) & ")?"
    
    If MsgBox(tmpStr, vbYesNo, Me.Caption) = vbYes Then
        TypeData(Index).DataNumber = TypeData(Index).DataNumber + 1
        ReDim Preserve TypeData(Index).DataV(TypeData(Index).ItemNumber - 1, TypeData(Index).DataNumber - 1)
        ReDim Preserve TypeData(Index).DataName(TypeData(Index).DataNumber - 1)
        For i = 0 To TypeData(Index).ItemNumber - 1
            TypeData(Index).DataV(i, TypeData(Index).DataNumber - 1) = TypeData(Index).DataV(i, CurrentID)
        Next i
        TypeData(Index).DataName(TypeData(Index).DataNumber - 1) = TypeData(Index).DataName(CurrentID)
            
        ComboType_click
        ComboNumber.ListIndex = TypeData(Index).DataNumber - 1

    End If
    
    
End Sub

Private Sub cmdDelete_Click()
Dim Index As Long
Dim CurrentID As Long
Dim i As Long
    Index = ComboType.ListIndex
    If Index < 0 Then Exit Sub
    CurrentID = ComboNumber.ListIndex
    If CurrentID < 0 Then Exit Sub
    
    If MsgBox(LoadResStr(10510) & TypeData(Index).Name, vbYesNo, Me.Caption) = vbYes Then
        TypeData(Index).DataNumber = TypeData(Index).DataNumber - 1
        ReDim Preserve TypeData(Index).DataV(TypeData(Index).ItemNumber - 1, TypeData(Index).DataNumber - 1)
        ReDim Preserve TypeData(Index).DataName(TypeData(Index).DataNumber - 1)
    End If
    
    ComboType_click
    ComboNumber.ListIndex = 0
    
    
End Sub
