VERSION 5.00
Begin VB.Form frmswitchcolor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Switch_Color"
   ClientHeight    =   7350
   ClientLeft      =   1590
   ClientTop       =   1305
   ClientWidth     =   11535
   DrawWidth       =   20
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   48
      Top             =   5280
      Width           =   3855
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2760
         TabIndex        =   53
         Text            =   "0"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   390
         Left            =   1800
         TabIndex        =   52
         Text            =   "0"
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "转换指定图片"
         Height          =   375
         Left            =   240
         TabIndex        =   51
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "转换当前图片"
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "转换所有图片"
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         Height          =   255
         Left            =   2520
         TabIndex        =   54
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.CommandButton ori 
      Caption         =   "复原"
      Height          =   375
      Left            =   2640
      TabIndex        =   47
      Top             =   4920
      Width           =   1095
   End
   Begin VB.PictureBox picbak 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   7080
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   46
      Top             =   2040
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "测试"
      Height          =   375
      Left            =   1200
      TabIndex        =   45
      Top             =   4920
      Width           =   1095
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   3960
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   9
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   43
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   8
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   42
      Top             =   3960
      Width           =   735
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   7
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   41
      Top             =   3480
      Width           =   735
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   6
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   40
      Top             =   3000
      Width           =   735
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   5
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   39
      Top             =   2520
      Width           =   735
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   4
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   38
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox color1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   30
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox color1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   29
      Top             =   3960
      Width           =   735
   End
   Begin VB.PictureBox color1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   28
      Top             =   3480
      Width           =   735
   End
   Begin VB.PictureBox color1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   27
      Top             =   3000
      Width           =   735
   End
   Begin VB.PictureBox color1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   26
      Top             =   2520
      Width           =   735
   End
   Begin VB.PictureBox color1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   21
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox color1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   19
      Top             =   600
      Width           =   735
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   18
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   17
      Top             =   1080
      Width           =   735
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   16
      Top             =   600
      Width           =   735
   End
   Begin VB.PictureBox color2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox color1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   14
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox color1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin VB.PictureBox color1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   705
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton OK 
      Caption         =   "确定"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox piclarge 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   3960
      ScaleHeight     =   485
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label Label1 
      Caption         =   "替换颜色"
      Height          =   375
      Index           =   9
      Left            =   1440
      TabIndex        =   37
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "替换颜色"
      Height          =   375
      Index           =   8
      Left            =   1440
      TabIndex        =   36
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "选择颜色"
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   35
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "选择颜色"
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   34
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "选择颜色"
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   33
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "选择颜色"
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   32
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "选择颜色"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   31
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "替换颜色"
      Height          =   375
      Index           =   7
      Left            =   1440
      TabIndex        =   25
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "替换颜色"
      Height          =   375
      Index           =   6
      Left            =   1440
      TabIndex        =   24
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "替换颜色"
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   23
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "替换颜色"
      Height          =   375
      Index           =   4
      Left            =   1440
      TabIndex        =   22
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "选择颜色"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "选择颜色"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "替换颜色"
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "选择颜色"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "替换颜色"
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "选择颜色"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "替换颜色"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6840
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "替换颜色"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "选择颜色"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmswitchcolor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public YES As Long
Public Pflag As Long
Public PBegin As Long
Public PEnd As Long
Private picScale As Long
Private data() As Long
Private WW As Long
Private HH As Long

'Public color1, color2 As Long

Private Sub Showpic()

   
Dim temp As Long
Dim dib As New clsDIB
    
    WW = g_PP.Width
    HH = g_PP.Height
    dib.CreateDIB pic.Width, pic.Height
    pic.BackColor = MaskColor
    WW = pic.Width
    HH = pic.Height
    temp = BitBlt(dib.CompDC, 0, 0, WW, HH, pic.hDC, 0, 0, &HCC0020)
    
    'Picnum = num
    'If num >= 0 Then
        Call genPicData(g_PP, dib.addr, WW, HH, 0, 0)
    'End If
        ' 复制到dib上
    temp = BitBlt(pic.hDC, 0, 0, WW, HH, dib.CompDC, 0, 0, &HCC0020)
    'piclarge.Width = WW
    'piclarge.height = HH
    picbak.PaintPicture pic.Image, 0, 0, WW, HH, 0, 0, WW, HH
    piclarge.PaintPicture pic.Image, 0, 0, WW * 4, HH * 4, 0, 0, WW, HH
End Sub

Private Sub cancel_Click()
YES = 0
Unload Me
End Sub

Private Sub Command1_Click()
        Changepic
End Sub

Private Sub color1_Click(Index As Integer)
    frmgetpix.Top = Me.Top + color1(Index).Top + color1(Index).Height
    frmgetpix.Left = Me.Left + color1(Index).Left + color1(Index).Width + 100

    frmgetpix.Show vbModal
    If frmgetpix.YES = 1 Then
        color1(Index).BackColor = frmgetpix.color3
    End If
End Sub

Private Sub color2_Click(Index As Integer)
    frmcolor.Top = Me.Top + color2(Index).Top + color2(Index).Height
    frmcolor.Left = Me.Left + color2(Index).Left + color2(Index).Width + 100
    frmcolor.Show vbModal
    If frmcolor.YES = 1 Then
        color2(Index).BackColor = frmcolor.Color
    End If
    'Changepic
End Sub

Private Sub Form_Load()
Dim I As Long, j As Long
Dim rr As Long, gg As Long, bb As Long
    

    Me.Caption = StrUnicode(Me.Caption)
    For I = 0 To Me.Controls.Count - 1
        Call SetCaption(Me.Controls(I))
    Next I
    YES = 0
    
    picbak.BackColor = MaskColor
    Showpic

    If MsgBox(StrUnicode2("是否使用颜色缓存？"), vbQuestion + vbOKCancel) = vbOK Then
        For I = 0 To 9
            color1(I).BackColor = colorA(I)
            color2(I).BackColor = colorB(I)
        Next I
    End If
    c_Skinner.AttachSkin Me.hWnd

End Sub

Private Sub OK_Click()
Dim I As Long
    Label2.Caption = "switching..."
    DoEvents
'color1 = ShapeColor.FillColor
'color2 = Shapeswitch.FillColor
    For I = 0 To 9
        colorA(I) = color1(I).BackColor
        colorB(I) = color2(I).BackColor
    Next I

    If Option1.Value = True Then
        Pflag = 1
    ElseIf Option2.Value = True Then
        Pflag = 2
    ElseIf Option3.Value = True Then
        Pflag = 3
        PBegin = Text1
        PEnd = Text2
    End If

YES = 1
Unload Me

End Sub



Public Sub Changepic()
Dim I, j, k As Long
    
    For I = 0 To 9
        colorA(I) = color1(I).BackColor
        colorB(I) = color2(I).BackColor
    Next I
    WW = g_PP.Width
    HH = g_PP.Height
    'pic.Width = WW
    'pic.height = HH
    For I = 0 To WW - 1
        For j = 0 To HH - 1
            For k = 0 To 9
                If pic.Point(I, j) = color1(k).BackColor Then
                    'MsgBox 11
                    pic.PSet (I, j), color2(k).BackColor
                    Exit For
                End If
                DoEvents
            Next k
        Next j
    Next I
    'WW = pic.Width
    'HH = pic.height
    piclarge.PaintPicture pic.Image, 0, 0, WW * 4, HH * 4, 0, 0, WW, HH
    'MsgBox "11"
End Sub

Private Sub Option1_Click()
Text1.Enabled = False
Text2.Enabled = False
End Sub
Private Sub Option2_Click()
Text1.Enabled = False
Text2.Enabled = False
End Sub
Private Sub Option3_Click()
Text1.Enabled = True
Text2.Enabled = True
End Sub

Private Sub ori_Click()
pic.PaintPicture picbak.Image, 0, 0, WW, HH, 0, 0, WW, HH
piclarge.PaintPicture picbak.Image, 0, 0, WW * 4, HH * 4, 0, 0, WW, HH
End Sub
