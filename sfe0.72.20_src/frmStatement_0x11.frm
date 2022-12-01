VERSION 5.00
Begin VB.Form frmStatement_0x11 
   Caption         =   "修改场景贴图"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form2"
   ScaleHeight     =   6315
   ScaleWidth      =   9420
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox ComboLevel 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmStatement_0x11.frx":0000
      Left            =   1440
      List            =   "frmStatement_0x11.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtPicnum 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   18
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox ComboAddress 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtpic 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      Height          =   4575
      Left            =   4440
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   8
      Top             =   1440
      Width           =   3135
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5655
      LargeChange     =   20
      Left            =   4080
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.ComboBox Combomanual 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtid 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      Height          =   2535
      Left            =   1320
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   2
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtx 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txty 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "贴图编号"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "图层"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "指令id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "手工确认当前场景"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "坐标X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "坐标Y"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "场景"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmStatement_0x11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Long
Dim kk As Statement
Dim picdata() As RLEPic
Dim picnum As Long

' 17



Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
   
    If ComboAddress.ListIndex = 0 Then
        kk.data(0) = &HFFFE
    Else
        kk.data(0) = ComboAddress.ListIndex - 1
    End If
    kk.data(1) = ComboLevel.Text
    kk.data(2) = txtX.Text
    kk.data(3) = txtY.Text
    kk.data(4) = txtPicnum.Text
    
    Unload Me
End Sub



Private Sub txtpicnum_change()
    If Combomanual.ListIndex < 0 Then Exit Sub
        pic1.Cls
        Call ShowPicDIB(picdata(txtPicnum.Text / 2), pic1.hDC, pic1.ScaleWidth / 2, pic1.ScaleHeight / 2)
        pic1.Refresh
End Sub

Private Sub Combomanual_Click()
Dim I As Long
    I = Combomanual.ListIndex
    If I = -1 Then
        Exit Sub
    End If
    Call LoadSMap(I, picdata, picnum)
    VScroll1.Min = 0
    VScroll1.Max = (picnum - 1)
    txtpicnum_change
End Sub

Private Sub Form_Load()
Dim I As Long

     Me.Caption = LoadResStr(528)
    Label1.Caption = LoadResStr(1102)
    Label2.Caption = LoadResStr(1313)
    Label3.Caption = LoadResStr(1314)
    Label4.Caption = LoadResStr(1302)
    Label5.Caption = LoadResStr(510)
    Label6.Caption = LoadResStr(529)
    Label7.Caption = LoadResStr(2101)
   
    cmdok.Caption = LoadResStr(102)
    cmdcancel.Caption = LoadResStr(103)




    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    
    Combomanual.Clear
    ComboAddress.Clear
    ComboAddress.AddItem "---" & LoadResStr(509) & "---"
    For I = 1 To Scenenum
        Combomanual.AddItem I - 1 & Big5toUnicode(Scene(I - 1).Name1, 10)
        ComboAddress.AddItem I - 1 & Big5toUnicode(Scene(I - 1).Name1, 10)
    Next I
    
    
    txtid.Text = kk.id & "(" & Hex(kk.id) & ")"
    
    If kk.data(0) = &HFFFE Then
        ComboAddress.ListIndex = 0
    Else
        ComboAddress.ListIndex = kk.data(0) + 1
    End If
    ComboLevel.Text = kk.data(1)
   
    txtX.Text = kk.data(2)
    txtY.Text = kk.data(3)
    txtPicnum.Text = kk.data(4)
    c_Skinner.AttachSkin Me.hWnd
End Sub

Private Sub txtpic_Change()
    If Combomanual.ListIndex = -1 Then Exit Sub
    pic2.Cls
    Call ShowPicDIB(picdata(txtpic.Text / 2), pic2.hDC, pic2.ScaleWidth / 2, pic2.ScaleHeight / 2)
    pic2.Refresh
End Sub

Private Sub VScroll1_Change()
    txtpic.Text = VScroll1.Value * 2
End Sub







