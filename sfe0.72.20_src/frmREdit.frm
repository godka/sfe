VERSION 5.00
Begin VB.Form frmREdit 
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmREdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   705
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdPutExcel 
      Caption         =   "导入excel"
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton CmdGetExcel 
      Caption         =   "导出excel"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveRecord 
      Caption         =   "save"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoadRecord 
      Caption         =   "read"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox ComboNumber 
      Height          =   345
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.ComboBox ComboRecord 
      Height          =   345
      Left            =   8760
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox ComboType 
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2175
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
      Height          =   6720
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "双击修改"
      Top             =   1560
      Width           =   10455
   End
End
Attribute VB_Name = "frmREdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormOldWidth As Long
    '保存窗体的原始宽度
Private FormOldHeight As Long
    '保存窗体的原始高度
'Private AllRIDX() As String
'Private ALLRGRP() As String
'Private AllRNote() As String

Private TypeNumber As Long

'Private TypeName() As String
Private TypeDataItem() As Long

Private Type DataItem_type
    ByteNum As Long
    isStr As Long
    isName As Long
    Ref As Long
    Name As String
    note As String
    offset As Long
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




Private Sub Load_R_Type()
Dim i As Long
Dim j As Long
Dim k As Long
Dim ll As Long
Dim num As Long
Dim tmpStr() As String
Dim tmpstr2() As String
Dim NumArray As Long
Dim NumType As Long

    TypeNumber = GetINILong("R_Modify", "TypeNumber")
    ReDim TypeData(TypeNumber - 1)
    
    ReDim TypeDataItem(TypeNumber - 1)
    For i = 0 To TypeNumber - 1
        TypeData(i).Name = GetINIStr("R_Modify", "TypeName" & i)
        TypeDataItem(i) = GetINILong("R_Modify", "TypeDataItem" & i)
    Next i
    
    For i = 0 To TypeNumber - 1
        num = 0
        For j = 0 To TypeDataItem(i) - 1
            tmpStr = Split(SubSpace(GetINIStr("R_Modify", "Data(" & i & "," & j & ")")), " ")
            num = num + CLng(tmpStr(0)) * CLng(tmpStr(1))
        Next j
        TypeData(i).ItemNumber = num
        ReDim TypeData(i).DataItem(num - 1)
        num = 0
        j = 0
        Do While j < TypeDataItem(i)
            tmpStr = Split(SubSpace(GetINIStr("R_Modify", "Data(" & i & "," & j & ")")), " ")
            NumArray = CLng(tmpStr(0))
            NumType = CLng(tmpStr(1))
            For k = 1 To NumArray
                TypeData(i).DataItem(num).ByteNum = CLng(tmpStr(2))
                TypeData(i).DataItem(num).isStr = CLng(tmpStr(3))
                TypeData(i).DataItem(num).isName = CLng(tmpStr(4))
                TypeData(i).DataItem(num).Ref = CLng(tmpStr(5))
                TypeData(i).DataItem(num).Name = tmpStr(6) & IIf(NumArray > 1, k, "")
                TypeData(i).DataItem(num).note = (tmpStr(7))
                num = num + 1
                For ll = 2 To NumType
                    tmpstr2 = Split(SubSpace(GetINIStr("R_Modify", "Data(" & i & "," & j + ll - 1 & ")")), " ")
                    TypeData(i).DataItem(num).ByteNum = CLng(tmpstr2(2))
                    TypeData(i).DataItem(num).isStr = CLng(tmpstr2(3))
                    TypeData(i).DataItem(num).isName = CLng(tmpstr2(4))
                    TypeData(i).DataItem(num).Ref = CLng(tmpstr2(5))
                    TypeData(i).DataItem(num).Name = tmpstr2(6) & IIf(NumArray > 1, k, "")
                    TypeData(i).DataItem(num).note = tmpstr2(7)
                    num = num + 1
                Next ll
            Next k
            j = j + NumType
        Loop
    Next i
    
End Sub



Private Sub Load_R()
Dim i As Long
Dim j As Long
Dim k As Long
Dim filenum As Long
Dim Idx() As Long
Dim DataLong As Long
Dim offset As Long
Dim tmpbyte() As Byte
    ReDim Idx(TypeNumber)
    Idx(0) = 0
    filenum = OpenBin(G_Var.JYPath & G_Var.RIDX(ComboRecord.ListIndex), "R")
    For i = 1 To TypeNumber
        Get #filenum, , Idx(i)
'        MsgBox Idx(i)
    Next i
    Close #filenum
    
    filenum = OpenBin(G_Var.JYPath & G_Var.RGRP(ComboRecord.ListIndex), "R")

    For i = 0 To TypeNumber - 1
        DataLong = 0
        For j = 0 To TypeData(i).ItemNumber - 1
            DataLong = DataLong + TypeData(i).DataItem(j).ByteNum
        Next j
        If (Idx(i + 1) - Idx(i)) Mod DataLong <> 0 Then
            MsgBox "R* data format error" & i
            Exit Sub
        Else
            TypeData(i).DataNumber = (Idx(i + 1) - Idx(i)) / DataLong
        End If
        ReDim TypeData(i).DataV(TypeData(i).ItemNumber - 1, TypeData(i).DataNumber - 1)
        ReDim TypeData(i).DataName(TypeData(i).DataNumber - 1)
        offset = Idx(i)
        Seek #filenum, offset + 1
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

Private Sub CmdGetExcel_Click()

Dim i As Integer
Dim kuang As OPENFILENAME
Dim filename As String
    kuang.lStructSize = Len(kuang)
    kuang.hwndOwner = Me.hwnd
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
    'filenum = OpenBin(ofn.lpstrFile, "R")
End Sub

Private Sub cmdLoadRecord_Click()
On Error GoTo Label6:
    Load_R
Exit Sub
Label6:
MsgBox "error,cannot find the file:" & G_Var.RGRP(ComboRecord.ListIndex) & Chr(13) & "error,cannot find the file:" & G_Var.RIDX(ComboRecord.ListIndex), 64, "error"
ComboRecord.ListIndex = 0
End Sub

Private Sub CmdPutExcel_Click()
Dim ofn As OPENFILENAME
Dim Rtn As String
Dim tmpStr As String
Dim filenum As Long
Dim i As Long, j As Long, k As Long
    tmpStr = "xls文件|*.xls"
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hwnd
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
    
    'ComboNumber.ListIndex = 0
End Sub

Private Sub cmdSaveRecord_Click()
Dim Idx() As Long
Dim i As Long, j As Long, k As Long
Dim filenum As Long
Dim tmpbyte() As Byte
Dim Length As Long

    If ComboRecord.ListIndex = -1 Then Exit Sub
    ReDim Idx(TypeNumber)
    Idx(0) = 0
    filenum = OpenBin(G_Var.JYPath & G_Var.RGRP(ComboRecord.ListIndex), "WN")
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
        Idx(i + 1) = Loc(filenum)
    Next i
    Close #filenum
    
    filenum = OpenBin(G_Var.JYPath & G_Var.RIDX(ComboRecord.ListIndex), "W")
    For i = 1 To TypeNumber
        Put #filenum, , Idx(i)
    Next i
    Close #filenum
    
    Call ReadRR(0)
        
   
    
    
End Sub

Private Sub ComboNumber_click()
Dim typeN As Long
Dim Index As Long
Dim i As Long
Dim tmpStr As String
    Index = ComboNumber.ListIndex
    If Index = -1 Then Exit Sub
    typeN = ComboType.ListIndex
    If typeN = -1 Then Exit Sub
    ListItem.Clear
    For i = 0 To TypeData(typeN).ItemNumber - 1
        ListItem.AddItem GenListstr(typeN, Index, i)
    Next i
    ListItem.ListIndex = 0
End Sub

' 生成list字符串
' i type    j   datanumber     k itemnumber
Private Function GenListstr(i As Long, j As Long, k As Long) As String
Dim tmpStr As String
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
            tmpStr = tmpStr & " " & TypeData(TypeData(i).DataItem(k).Ref).DataName(TypeData(i).DataV(k, j).Int)
        End If
    GenListstr = tmpStr
End Function




Private Sub ComboType_click()
Dim Index As Long
Dim i As Long
    Index = ComboType.ListIndex
    If Index = -1 Then Exit Sub
    
    ComboNumber.Clear
    For i = 0 To TypeData(Index).DataNumber - 1
        ComboNumber.AddItem i & TypeData(Index).DataName(i)
    Next i
    ComboNumber.ListIndex = 0

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim i As Long
    Call ConvertForm(Me)
    Me.Caption = LoadResStr(224)
    cmdLoadRecord.Caption = LoadResStr(10501)
    cmdSaveRecord.Caption = LoadResStr(10502)
    
    CmdAdd.Caption = LoadResStr(10507)
    cmdDelete.Caption = LoadResStr(10508)
    
    ComboRecord.Clear
    Select Case GetINIStr("run", "style")
        Case "kys"
            For i = 0 To 6
                ComboRecord.AddItem G_Var.RecordNote(i)
            Next i
        Case "DOS"
            For i = 0 To 3
                ComboRecord.AddItem G_Var.RecordNote(i)
            Next i
    End Select
    ComboRecord.ListIndex = 1
    Call ResizeInit(Me)  '在程序装入时必须加入
    c_Skinner.AttachSkin Me.hwnd
    'Call SetCombo(Me)
    Load_R_Type
    Load_R
End Sub



Private Sub ListItem_Click()
Dim i As Long, j As Long, k As Long
    i = ComboType.ListIndex
    j = ComboNumber.ListIndex
    k = ListItem.ListIndex
    If i < 0 Or j < 0 Or k < 0 Then Exit Sub
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
    Load frmChangeRValue
    
    frmChangeRValue.Label1.Caption = TypeData(i).DataItem(k).Name
    If TypeData(i).DataItem(k).isStr = 0 Then
        frmChangeRValue.Text1.Text = TypeData(i).DataV(k, j).Int
    Else
        frmChangeRValue.Text1.Text = TypeData(i).DataV(k, j).str
    End If
    If TypeData(i).DataItem(k).Ref >= 0 Then
        frmChangeRValue.Combo1.Clear
        frmChangeRValue.Combo1.AddItem LoadResStr(10602)
        For num = 0 To TypeData(TypeData(i).DataItem(k).Ref).DataNumber - 1
            frmChangeRValue.Combo1.AddItem num & TypeData(TypeData(i).DataItem(k).Ref).DataName(num)
        Next num
        frmChangeRValue.Combo1.ListIndex = TypeData(i).DataV(k, j).Int + 1
        frmChangeRValue.Text1.Enabled = False
    Else
        frmChangeRValue.Combo1.Visible = False
    End If
    
    frmChangeRValue.Show vbModal
    
    If frmChangeRValue.OK = 1 Then
        If TypeData(i).DataItem(k).isStr = 0 Then
             TypeData(i).DataV(k, j).Int = frmChangeRValue.Text1.Text
        Else
            TypeData(i).DataV(k, j).str = frmChangeRValue.Text1.Text
        End If
        
        ListItem.List(k) = GenListstr(i, j, k)
'        ListItem.ForeColor = vbRed
    End If
    
    Unload frmChangeRValue
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



'在调用ResizeForm前先调用本函数
Public Sub ResizeInit(FormName As Form)
   Dim Obj As Control
   FormOldWidth = FormName.ScaleWidth
   FormOldHeight = FormName.ScaleHeight
   On Error Resume Next
   For Each Obj In FormName
     Obj.Tag = Obj.Left & " " & Obj.Top & " " _
           & Obj.Width & " " & Obj.Height & " "
   Next Obj
   On Error GoTo 0
End Sub

'按比例改变表单内各元件的大小，在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(FormName As Form)
   Dim pos(4) As Double
   Dim i As Long, TempPos As Long, startpos As Long
   Dim Obj As Control
   Dim ScaleX As Double, ScaleY As Double

   ScaleX = FormName.ScaleWidth / FormOldWidth
   '保存窗体宽度缩放比例
   ScaleY = FormName.ScaleHeight / FormOldHeight
   '保存窗体高度缩放比例
   On Error Resume Next
   For Each Obj In FormName
     startpos = 1
     For i = 0 To 4
      '读取控件的原始位置与大小

       TempPos = InStr(startpos, Obj.Tag, " ", vbTextCompare)
       If TempPos > 0 Then
         pos(i) = Mid(Obj.Tag, startpos, TempPos - startpos)
         startpos = TempPos + 1
       Else
         pos(i) = 0
       End If
       '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
       Obj.Move pos(0) * ScaleX, pos(1) * ScaleY, _
                pos(2) * ScaleX, pos(3) * ScaleY
     Next i
   Next Obj
   On Error GoTo 0
End Sub

