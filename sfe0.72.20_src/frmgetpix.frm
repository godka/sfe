VERSION 5.00
Begin VB.Form frmgetpix 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.VScrollBar HH1 
      Height          =   4695
      LargeChange     =   5
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar WW1 
      Height          =   255
      LargeChange     =   5
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   4575
   End
   Begin VB.CommandButton cancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton OK 
      Caption         =   "确定"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox color2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5160
      ScaleHeight     =   345
      ScaleWidth      =   1065
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4635
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmgetpix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public YES As Long
Public cok As Boolean
Public color3 As Long
Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
YES = 0
cok = False
WW1.Max = 256
HH1.Max = 256
pic1.BackColor = MaskColor
pic1.Cls
pic1.PaintPicture frmswitchcolor.piclarge.Image, 0, 0 ', frmswitchcolor.piclarge.Width, frmswitchcolor.piclarge.Height
c_Skinner.AttachSkin Me.hWnd
End Sub

Private Sub HH1_Change()
pic1.Cls
pic1.PaintPicture frmswitchcolor.piclarge.Image, -WW1.Value * 15, -HH1.Value * 15 ', frmswitchcolor.piclarge.Width, frmswitchcolor.piclarge.Height

End Sub

Private Sub OK_Click()
    color3 = color2.BackColor
    YES = 1
    Unload Me
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    color2.BackColor = pic1.Point(x, y)
    cok = True

End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If cok = False Then
    color2.BackColor = pic1.Point(x, y)
End If
End Sub

Private Sub WW1_Change()
pic1.Cls
pic1.PaintPicture frmswitchcolor.piclarge.Image, -WW1.Value * 15, -HH1.Value * 15 ', frmswitchcolor.piclarge.Width, frmswitchcolor.piclarge.Height

End Sub
