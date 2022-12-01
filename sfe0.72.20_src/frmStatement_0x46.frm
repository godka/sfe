VERSION 5.00
Begin VB.Form frmStatement_0x46 
   Caption         =   "显示字幕"
   ClientHeight    =   6765
   ClientLeft      =   4800
   ClientTop       =   1590
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   9900
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmStatement_0x46.frx":0000
      Top             =   5520
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   375
      Left            =   8400
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   375
      Left            =   8400
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox combotxt 
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   600
      Width           =   7815
   End
   Begin VB.Frame Frame4 
      Caption         =   "显示颜色Color"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7815
      Begin VB.TextBox userColor 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.PictureBox PicPalette 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   3720
         ScaleHeight     =   255
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   255
         TabIndex        =   6
         ToolTipText     =   "单击选择颜色"
         Top             =   240
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "选择背景色"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "选择前景色"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "设置颜色"
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1800
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt2 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2640
         Top             =   1440
         Width           =   855
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2640
         Top             =   2400
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "对话"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmStatement_0x46"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Index, I, j As Long
Dim kk As Statement
Dim picdata() As RLEPic
Dim rr As Long, gg As Long, bb As Long
Dim Color As Long
Dim picnum As Long

Private Sub Command1_Click()
        userColor.Text = Long2int(txt1.Text * 256 + txt2.Text)

End Sub

Private Sub Command2_Click()
kk.data(0) = combotxt.ListIndex
kk.data(1) = userColor
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    For j = 0 To 15
        For I = 0 To 15
            rr = (mcolor_RGB(I + j * 16) \ 65536) And &HFF&
            gg = (mcolor_RGB(I + j * 16) \ 256) And &HFF
            bb = mcolor_RGB(I + j * 16) And &HFF
            
            PicPalette.Line (I * 16, j * 16)-((I + 1) * 16, (j + 1) * 16), RGB(rr, gg, bb), BF
        Next I
    Next j
    For I = 0 To numtalk - 1
'        combotxt.AddItem i
        combotxt.AddItem I & ":" & Talk(I)
    Next I
    combotxt.ListIndex = kk.data(0)
    userColor.Text = kk.data(1)
        Color = Int2Long(userColor.Text) \ 256
        rr = (mcolor_RGB(Color) \ 65536) And &HFF&
        gg = (mcolor_RGB(Color) \ 256) And &HFF
        bb = mcolor_RGB(Color) And &HFF
        
        Shape1.FillColor = RGB(rr, gg, bb)
        txt1.Text = Color
        
        Color = Int2Long(userColor.Text) And &HFF
        rr = (mcolor_RGB(Color) \ 65536) And &HFF&
        gg = (mcolor_RGB(Color) \ 256) And &HFF
        bb = mcolor_RGB(Color) And &HFF
        
        Shape2.FillColor = RGB(rr, gg, bb)
        txt2.Text = Color
            c_Skinner.AttachSkin Me.hWnd

End Sub
Private Sub PicPalette_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rr As Long, gg As Long, bb As Long
Dim Color As Long
Dim colorRGB As Long
    Color = (x \ 16) + (y \ 16) * 16
    rr = (mcolor_RGB(Color) \ 65536) And &HFF&
    gg = (mcolor_RGB(Color) \ 256) And &HFF
    bb = mcolor_RGB(Color) And &HFF

    colorRGB = RGB(rr, gg, bb)

    If Option1.Value = True Then
        Shape1.FillColor = colorRGB
        txt1.Text = Color
    Else
        Shape2.FillColor = colorRGB
        txt2.Text = Color
    End If
        c_Skinner.AttachSkin Me.hWnd

End Sub
