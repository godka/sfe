VERSION 5.00
Begin VB.Form frmstatement_0x03 
   Caption         =   "�޸ĳ����¼�"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   LinkTopic       =   "Form2"
   ScaleHeight     =   6660
   ScaleWidth      =   10230
   StartUpPosition =   2  '��Ļ����
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
      Left            =   6000
      TabIndex        =   37
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      Height          =   3255
      Left            =   6120
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   36
      Top             =   1080
      Width           =   2775
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5295
      LargeChange     =   20
      Left            =   5760
      TabIndex        =   35
      Top             =   1080
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4800
      Top             =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "�����¼���Ϣ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   5535
      Begin VB.PictureBox pic1 
         AutoRedraw      =   -1  'True
         Height          =   1515
         Left            =   3480
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   97
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   117
         TabIndex        =   39
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton txtX2 
         Caption         =   "*2"
         Height          =   375
         Left            =   3360
         TabIndex        =   38
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtY 
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
         Left            =   2280
         TabIndex        =   30
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox txtX 
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
         Left            =   2280
         TabIndex        =   28
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox txtspeed 
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
         Left            =   2280
         TabIndex        =   26
         Top             =   3240
         Width           =   855
      End
      Begin VB.ComboBox ComboPass 
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtnum 
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
         Left            =   2280
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtkdef1 
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
         Left            =   2280
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtkdef2 
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
         Left            =   2280
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtkdef3 
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
         Left            =   2280
         TabIndex        =   12
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtpic1 
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
         Left            =   2280
         TabIndex        =   11
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtpic2 
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
         Left            =   2280
         TabIndex        =   10
         Text            =   " "
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtpic3 
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
         Left            =   2280
         TabIndex        =   9
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "�¼����Ϊ[-1]��ʾȡ���¼���"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   5040
         Width           =   4335
      End
      Begin VB.Label Label16 
         Caption         =   "����������Ϊ[-2]���ʾ���ֵ�ǰֵ����"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   4800
         Width           =   5175
      End
      Begin VB.Label Label15 
         Caption         =   "������Y"
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
         Left            =   240
         TabIndex        =   29
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "������X"
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
         Left            =   240
         TabIndex        =   27
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "��ͼ�����ٶ�"
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
         Left            =   240
         TabIndex        =   25
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "�Ƿ����ͨ��"
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
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "���"
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
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "ʹ�ÿո񴥷��¼����"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "ʹ����Ʒ�����¼����"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "���Ǿ��������¼����"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "��ͼ1(���Ϊ2�ı���)"
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
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "��ͼ2"
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
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "��ͼ3"
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
         Left            =   240
         TabIndex        =   17
         Top             =   2880
         Width           =   615
      End
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.ComboBox ComboSkey 
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
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   2535
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
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1815
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
      Left            =   840
      TabIndex        =   2
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
      Left            =   8400
      TabIndex        =   1
      Top             =   4800
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
      Left            =   8400
      TabIndex        =   0
      Top             =   5640
      Width           =   1335
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
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "����"
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
      Left            =   1800
      TabIndex        =   31
      Top             =   120
      Width           =   735
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
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "�����¼�����"
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
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmstatement_0x03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Long
Dim kk As Statement
Dim picdata() As RLEPic
Dim picnum As Long
Dim showpicnum As Long



Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    If ComboAddress.ListIndex = 0 Then
        kk.data(0) = &HFFFE
    Else
        kk.data(0) = ComboAddress.ListIndex - 1
    End If
    If ComboSkey.ListIndex = 0 Then
        kk.data(1) = &HFFFE
    Else
        kk.data(1) = ComboSkey.ListIndex - 1
    End If
    
     If ComboPass.ListIndex = 0 Then
        kk.data(2) = &HFFFE
    Else
        kk.data(2) = ComboPass.ListIndex - 1
    End If
   
   
   kk.data(3) = txtnum.Text
   
     kk.data(4) = txtkdef1.Text
     kk.data(5) = txtkdef2.Text
     kk.data(6) = txtkdef3.Text
    
     kk.data(7) = txtpic1.Text
     kk.data(8) = txtpic2.Text
    kk.data(9) = txtpic3.Text
    
    kk.data(10) = txtspeed.Text
    kk.data(11) = txtX.Text
    kk.data(12) = txtY.Text
    
    Unload Me
End Sub

Private Sub Combomanual_Click()
Dim i As Long
    i = Combomanual.ListIndex
    If i = -1 Then
        Timer1.Enabled = False
        Exit Sub
    End If
    Timer1.Enabled = True
    Call LoadSMap(i, picdata, picnum)
    VScroll1.Min = 0
    VScroll1.Max = (picnum - 1)
End Sub

Private Sub Form_Load()
Dim i As Long

    
    Me.Caption = LoadResStr(1301)
    Label1.Caption = LoadResStr(1102)
    Label2.Caption = LoadResStr(510)
    Label3.Caption = LoadResStr(512)
    Label4.Caption = LoadResStr(1302)
    Frame1.Caption = LoadResStr(1303)
    Label5.Caption = LoadResStr(1304)
    Label6.Caption = LoadResStr(1305)
    Label7.Caption = LoadResStr(1306)
    Label8.Caption = LoadResStr(1307)
    Label9.Caption = LoadResStr(1308)
    Label10.Caption = LoadResStr(1309)
    Label11.Caption = LoadResStr(1310)
    Label12.Caption = LoadResStr(1311)
    Label13.Caption = LoadResStr(1312)
    Label14.Caption = LoadResStr(1313)
    Label15.Caption = LoadResStr(1314)
    Label16.Caption = LoadResStr(1315)
    Label17.Caption = LoadResStr(1316)
    cmdok.Caption = LoadResStr(102)
    cmdcancel.Caption = LoadResStr(103)
    
    
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    Timer1.Enabled = False
    showpicnum = 0
    ComboAddress.Clear
    ComboAddress.AddItem "---" & LoadResStr(509) & "---"
    Combomanual.Clear
    For i = 1 To Scenenum
        ComboAddress.AddItem i - 1 & Big5toUnicode(Scene(i - 1).Name1, 10)
        Combomanual.AddItem i - 1 & Big5toUnicode(Scene(i - 1).Name1, 10)
    Next i
    
    ComboSkey.Clear
    ComboSkey.AddItem "---" & LoadResStr(511) & "---"
    For i = 1 To 200
        ComboSkey.AddItem i - 1
    Next i
    
    ComboPass.Clear
    ComboPass.AddItem LoadResStr(1317)
    ComboPass.AddItem LoadResStr(1318)
    ComboPass.AddItem LoadResStr(1319)
    
    
    
    txtid.Text = kk.id & "(" & Hex(kk.id) & ")"
    
    If kk.data(0) = &HFFFE Then
        ComboAddress.ListIndex = 0
    Else
        ComboAddress.ListIndex = kk.data(0) + 1
    End If
    
    If kk.data(1) = &HFFFE Then
        ComboSkey.ListIndex = 0
    Else
        ComboSkey.ListIndex = kk.data(1) + 1
    End If
    
    If kk.data(2) = &HFFFE Then
        ComboPass.ListIndex = 0
    Else
        ComboPass.ListIndex = kk.data(2) + 1
    End If
    
 
    txtnum.Text = kk.data(3)
   
    txtkdef1.Text = kk.data(4)
    txtkdef2.Text = kk.data(5)
    txtkdef3.Text = kk.data(6)
    
    txtpic1.Text = kk.data(7)
    txtpic2.Text = kk.data(8)
    txtpic3.Text = kk.data(9)
    
    txtspeed.Text = kk.data(10)
    txtX.Text = kk.data(11)
    txtY.Text = kk.data(12)
    
    c_Skinner.AttachSkin Me.hWnd
End Sub

Private Sub txtpic_Change()
    If Combomanual.ListIndex = -1 Then Exit Sub
    pic2.Cls
    Call ShowPicDIB(picdata(txtpic.Text / 2), pic2.hDC, pic2.ScaleWidth / 2, pic2.ScaleHeight - 10)
    pic2.Refresh
End Sub

Private Sub txtX2_Click()
txtpic1 = txtpic1 * 2
txtpic2 = txtpic2 * 2
txtpic3 = txtpic3 * 2
End Sub

Private Sub VScroll1_Change()
    txtpic.Text = VScroll1.Value * 2
End Sub

Private Sub Picture3_Click()

End Sub

Private Sub Timer1_Timer()
    If Combomanual.ListIndex = -1 Then Exit Sub
    Select Case showpicnum
    Case 0
       pic1.Cls
       If txtpic1.Text > 0 Then
           Call ShowPicDIB(picdata(txtpic1.Text / 2), pic1.hDC, pic1.ScaleWidth / 2, pic1.ScaleHeight - 10)
       End If
       pic1.Refresh
       showpicnum = showpicnum + 1
       Exit Sub
    Case 1
       pic1.Cls
       If txtpic2.Text > 0 Then
           Call ShowPicDIB(picdata(txtpic2.Text / 2), pic1.hDC, pic1.ScaleWidth / 2, pic1.ScaleHeight - 10)
       End If
       pic1.Refresh
       showpicnum = showpicnum + 1
       Exit Sub
    Case 2
       pic1.Cls
       If txtpic3.Text > 0 Then
           Call ShowPicDIB(picdata(txtpic3.Text / 2), pic1.hDC, pic1.ScaleWidth / 2, pic1.ScaleHeight - 10)
       End If
       pic1.Refresh
       showpicnum = 0
       Exit Sub
    
    End Select
    
End Sub

