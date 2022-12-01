VERSION 5.00
Begin VB.Form frmInitEdit 
   Caption         =   "初始属性修改"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   9345
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameInit 
      Caption         =   "Frame2"
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   6975
      Begin VB.ComboBox ComboInit 
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtbase 
         Height          =   330
         Left            =   2280
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtRND 
         Height          =   330
         Left            =   3480
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtQQ 
         Height          =   330
         Left            =   5400
         TabIndex        =   13
         Text            =   "Text3"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "基础值"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "随机范围"
         Height          =   495
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "秘笈值"
         Height          =   255
         Left            =   5400
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FrameShow 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   6975
      Begin VB.ComboBox ComboShownum 
         Height          =   345
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox ComboShowVar 
         Height          =   345
         Left            =   1440
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtFormat 
         Height          =   330
         Left            =   3600
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtVarMax 
         Height          =   330
         Left            =   5280
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "最大值"
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "显示字串"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "显示变量"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "显示顺序"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "Modify"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "确定"
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmInitEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private zfilenum As Long
Private zfilename As String

Private Initnum As Long

Private Type InitType
    BaseAddr As Long
    Basedata As Long
    RndAddr As Long
    Rnddata As Byte
    MaxAddr As Long
    MaxData As Byte
    QQAddr As Long
    QQdata As Integer
    Name As String
    VarAddr As Long
End Type

Private InitData() As InitType


Private Shownum As Long

Private Type Showtype
    AddrVar1 As Long
    AddrVar2 As Long
    Var As Long
    AddrStr As Long
    StrByte(8) As Byte
    Str As String
    AddrMax As Long
    MaxData As Byte
End Type

Private ShowData() As Showtype

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub CmdModify_Click()
    Modify_it
End Sub

Private Sub cmdok_Click()

Dim I As Long
Dim tmpstr As String
Dim tmpstr2 As String
Dim tmpbyte() As Byte
    If MsgBox(LoadResStr(10131), vbYesNo) = vbNo Then Exit Sub
    zfilenum = OpenBin(G_Var.JYPath & G_Var.EXE, "W")
    
    For I = 0 To Initnum - 1
        If InitData(I).BaseAddr >= 0 Then
            Put #zfilenum, InitData(I).BaseAddr + 1, InitData(I).Basedata
        End If
        If InitData(I).RndAddr >= 0 Then
            Put #zfilenum, InitData(I).RndAddr + 1, InitData(I).Rnddata
        End If
        If InitData(I).QQAddr >= 0 Then
            Put #zfilenum, InitData(I).QQAddr + 1, InitData(I).QQdata
        End If
    Next I
        
    For I = 0 To Shownum - 1
        Put #zfilenum, ShowData(I).AddrVar1 + 1, ShowData(I).Var
        Put #zfilenum, ShowData(I).AddrVar2 + 1, ShowData(I).Var
        Call UnicodetoBIG5(ShowData(I).Str, 9, tmpbyte)
        Put #zfilenum, ShowData(I).AddrStr + 1, tmpbyte
        Put #zfilenum, , CByte(0)
        Put #zfilenum, ShowData(I).AddrMax + 1, ShowData(I).MaxData
    Next I
        
    Close (zfilenum)

End Sub


Private Sub ComboInit_click()
Dim Index As Long
    Index = ComboInit.ListIndex
    If Index < 0 Then Exit Sub
    txtbase.Text = InitData(Index).Basedata
    txtRND.Text = InitData(Index).Rnddata
    'txtMax.Text = InitData(index).MaxData
    txtQQ.Text = InitData(Index).QQdata
End Sub

Private Sub ComboShownum_click()
Dim Index As Long
Dim I As Long
    Index = ComboShownum.ListIndex
    If Index < 0 Then Exit Sub
    For I = 0 To Initnum - 1
        If ShowData(Index).Var = InitData(I).VarAddr Then
            ComboShowVar.ListIndex = I
            Exit For
        End If
    Next I
    txtFormat = ShowData(Index).Str
    txtVarMax.Text = ShowData(Index).MaxData
End Sub


Private Sub Form_Load()

    Me.Caption = LoadResStr(218)
    
    cmdok.Caption = LoadResStr(102)
    cmdcancel.Caption = LoadResStr(103)
    
    Label1.Caption = LoadResStr(10301)
    Label2.Caption = LoadResStr(10302)
    Label3.Caption = LoadResStr(10303)
    Label4.Caption = LoadResStr(10304)
    Label5.Caption = LoadResStr(10307)
    Label6.Caption = LoadResStr(10306)
    Label7.Caption = LoadResStr(10305)
    
    CmdModify.Caption = LoadResStr(10308)
    
    FrameInit.Caption = LoadResStr(10309)
    FrameShow.Caption = LoadResStr(10310)
    

    Load_init
    c_Skinner.AttachSkin Me.hWnd
End Sub

Private Sub Load_init()
Dim I As Long
Dim tmpstr() As String
Dim tmplong As Long
Dim fixupOffset As Long
    Initnum = GetINILong("InitProperty", "initnum")
    ReDim InitData(Initnum - 1)
    For I = 0 To Initnum - 1
        tmpstr = Split(GetINIStr("InitProperty", "Addr" & I), ",")
        InitData(I).BaseAddr = tmpstr(0)
        InitData(I).RndAddr = tmpstr(1)
        InitData(I).MaxAddr = tmpstr(2)
        InitData(I).QQAddr = tmpstr(3)
        InitData(I).VarAddr = tmpstr(4)
        InitData(I).Name = tmpstr(5)
    Next I

    'zfilename = G_Var.JYPath & "\z.dat"
    
    zfilenum = OpenBin(G_Var.JYPath & G_Var.EXE, "R")
    For I = 0 To Initnum - 1
        If InitData(I).BaseAddr >= 0 Then
            Get #zfilenum, InitData(I).BaseAddr + 1, InitData(I).Basedata
        End If
        If InitData(I).RndAddr >= 0 Then
            Get #zfilenum, InitData(I).RndAddr + 1, InitData(I).Rnddata
        End If
        If InitData(I).MaxAddr >= 0 Then
            Get #zfilenum, InitData(I).MaxAddr + 1, InitData(I).MaxData
        End If
        If InitData(I).QQAddr >= 0 Then
            Get #zfilenum, InitData(I).QQAddr + 1, InitData(I).QQdata
        End If
        
    Next I
    ComboInit.Clear
    For I = 0 To Initnum - 1
        ComboInit.AddItem InitData(I).Name
    Next I

    ComboInit.ListIndex = 0
    
    fixupOffset = 8 * GetINILong("newZ", "PageAdd")

    Shownum = GetINILong("InitProperty", "initShownum")
    ReDim ShowData(Shownum - 1)
    For I = 0 To Shownum - 1
        tmpstr = Split(GetINIStr("InitProperty", "InitShow" & I), ",")
        ShowData(I).AddrVar1 = tmpstr(0) + fixupOffset
        ShowData(I).AddrVar2 = tmpstr(1) + fixupOffset
        ShowData(I).AddrStr = tmpstr(2)
        ShowData(I).AddrMax = tmpstr(3)
    Next I
    
    For I = 0 To Shownum - 1
        Get #zfilenum, ShowData(I).AddrVar1 + 1, ShowData(I).Var
        Get #zfilenum, ShowData(I).AddrVar2 + 1, tmplong
        If tmplong <> ShowData(I).Var Then
            MsgBox "Ini show data error:var1<>var2 num=" & I
        End If
        Get #zfilenum, ShowData(I).AddrStr + 1, ShowData(I).StrByte
        ShowData(I).Str = Big5toUnicode(ShowData(I).StrByte, 9)
        Get #zfilenum, ShowData(I).AddrMax + 1, ShowData(I).MaxData
    Next I
    
    Close #zfilenum

    ComboShownum.Clear
    For I = 0 To Shownum - 1
       ComboShownum.AddItem I
    Next I
    
    ComboShowVar.Clear
    For I = 0 To Initnum - 1
        ComboShowVar.AddItem InitData(I).Name
    Next I
    
    ComboShownum.ListIndex = 0

End Sub



Private Sub Modify_it()
Dim Index As Long
    Index = ComboInit.ListIndex
    If Index > -1 Then
        InitData(Index).Basedata = txtbase.Text
        InitData(Index).Rnddata = txtRND.Text
        InitData(Index).QQdata = txtQQ.Text
    End If
    Index = ComboShownum.ListIndex
    If Index >= -1 Then
        ShowData(Index).Var = InitData(ComboShowVar.ListIndex).VarAddr
        ShowData(Index).Str = txtFormat.Text
        ShowData(Index).MaxData = txtVarMax.Text
    End If
End Sub

