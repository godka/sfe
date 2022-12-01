VERSION 5.00
Begin VB.Form frmcolor 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton OK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox PicPalette 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   0
      ToolTipText     =   "单击选择颜色"
      Top             =   0
      Width           =   3855
   End
   Begin VB.Shape color2 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3960
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "frmcolor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cok As Boolean
Public Color As Long
Public YES As Long
Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim I As Long, j As Long
Dim rr As Long, gg As Long, bb As Long
    cok = False
    Me.Height = PicPalette.Height
    Me.Caption = StrUnicode(Me.Caption)
    For I = 0 To Me.Controls.Count - 1
        Call SetCaption(Me.Controls(I))
    Next I
    YES = 0
    
    For j = 0 To 15
        For I = 0 To 15
            rr = (mcolor_RGB(I + j * 16) \ 65536) And &HFF&
            gg = (mcolor_RGB(I + j * 16) \ 256) And &HFF
            bb = mcolor_RGB(I + j * 16) And &HFF
            
            PicPalette.Line (I * 16, j * 16)-((I + 1) * 16, (j + 1) * 16), RGB(rr, gg, bb), BF
        Next I
    Next j
    c_Skinner.AttachSkin Me.hWnd
End Sub

Private Sub PicPalette_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If cok = True Then Exit Sub
Dim rr As Long, gg As Long, bb As Long
Dim Color As Long
    Color = (x \ 16) + (y \ 16) * 16
            rr = (mcolor_RGB(Color) \ 65536) And &HFF&
            gg = (mcolor_RGB(Color) \ 256) And &HFF
            bb = mcolor_RGB(Color) And &HFF

    color2.FillColor = RGB(rr, gg, bb)
    cok = False
End Sub

Private Sub OK_Click()
YES = 1
Color = color2.FillColor
Unload Me
End Sub

Private Sub PicPalette_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rr As Long, gg As Long, bb As Long
Dim Color As Long
    Color = (x \ 16) + (y \ 16) * 16
            rr = (mcolor_RGB(Color) \ 65536) And &HFF&
            gg = (mcolor_RGB(Color) \ 256) And &HFF
            bb = mcolor_RGB(Color) And &HFF

    color2.FillColor = RGB(rr, gg, bb)
    cok = True
End Sub
