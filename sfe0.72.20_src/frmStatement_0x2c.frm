VERSION 5.00
Begin VB.Form frmStatement_0x2c 
   Caption         =   "���Ŷ���"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   LinkTopic       =   "Form2"
   ScaleHeight     =   6765
   ScaleWidth      =   10230
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox pic22 
      AutoRedraw      =   -1  'True
      Height          =   2535
      Left            =   3480
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   21
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtstart2 
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
      TabIndex        =   20
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtend2 
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
      TabIndex        =   19
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txt2 
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
      TabIndex        =   18
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdshow2 
      Caption         =   "��ʾ����"
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
      Left            =   3480
      TabIndex        =   17
      Top             =   2880
      Width           =   975
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
      Left            =   7440
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      Height          =   4695
      Left            =   6480
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   10
      Top             =   1920
      Width           =   3255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6015
      LargeChange     =   20
      Left            =   6120
      TabIndex        =   9
      Top             =   600
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
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
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
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
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
      Left            =   8760
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
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
      Left            =   8760
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      Height          =   2535
      Left            =   240
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   4
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtstart 
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
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtend 
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
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txt1 
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
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "��ʾ����"
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
      Left            =   240
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "����2"
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
      Left            =   3480
      TabIndex        =   26
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "����1"
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
      Left            =   360
      TabIndex        =   25
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "��ʼ��ͼ���"
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
      Left            =   3360
      TabIndex        =   24
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "������ͼ���"
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
      Left            =   3360
      TabIndex        =   23
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "��ʾλ�ã������¼����"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   22
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ָ��id"
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
      TabIndex        =   16
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "�ֹ�ȷ�ϵ�ǰ����"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "��ʼ��ͼ���"
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
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "������ͼ���"
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
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "��ʾλ�ã������¼����"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "frmStatement_0x2c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Long
Dim kk As Statement
Dim picdata() As RLEPic
Dim picnum As Long




Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
   
   
    kk.data(0) = txt1.Text
    kk.data(1) = txtstart.Text
    kk.data(2) = txtend.Text
    kk.data(3) = txt2.Text
    kk.data(4) = txtstart2.Text
    kk.data(5) = txtend2.Text
    
    Unload Me
End Sub

Private Sub cmdshow_Click()
Dim I As Long
Dim time1 As Single
    If Combomanual.ListIndex < 0 Then Exit Sub
    For I = txtstart.Text To txtend.Text Step 2
        pic1.Cls
        Call ShowPicDIB(picdata(I / 2), pic1.hDC, pic1.ScaleWidth / 2, pic1.ScaleHeight / 2)
        pic1.Refresh
        time1 = Timer
        While (Timer - time1) < 0.15
            DoEvents
        Wend
    Next I
End Sub

Private Sub cmdshow2_Click()
Dim I As Long
Dim time1 As Single
    If Combomanual.ListIndex < 0 Then Exit Sub
    For I = txtstart2.Text To txtend2.Text Step 2
        pic22.Cls
        Call ShowPicDIB(picdata(I / 2), pic22.hDC, pic22.ScaleWidth / 2, pic22.ScaleHeight / 2)
        pic22.Refresh
        time1 = Timer
        While (Timer - time1) < 0.15
            DoEvents
        Wend
    Next I

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
End Sub

Private Sub Form_Load()
Dim I As Long
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    
    Combomanual.Clear
    For I = 1 To Scenenum
        Combomanual.AddItem I - 1 & Big5toUnicode(Scene(I - 1).Name1, 10)
    Next I
    
    
    txtid.Text = kk.id & "(" & Hex(kk.id) & ")"
    
    txt1.Text = kk.data(0)
   
    txtstart.Text = kk.data(1)
    txtend.Text = kk.data(2)
    
    txt2.Text = kk.data(3)
   
    txtstart2.Text = kk.data(4)
    txtend2.Text = kk.data(5)
    
    Me.Caption = LoadResStr(4401)
    Label1.Caption = LoadResStr(1102)
    Label4.Caption = LoadResStr(1302)
    Label9.Caption = LoadResStr(4402)
    Label5.Caption = LoadResStr(2703)
    Label2.Caption = LoadResStr(2701)
    Label3.Caption = LoadResStr(2702)
    cmdShow.Caption = LoadResStr(4401)
    Label10.Caption = LoadResStr(4403)
    Label6.Caption = LoadResStr(2703)
    Label8.Caption = LoadResStr(2701)
    Label7.Caption = LoadResStr(2702)
    cmdshow2.Caption = LoadResStr(4401)

    cmdok.Caption = LoadResStr(102)
    cmdcancel.Caption = LoadResStr(103)
  
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







