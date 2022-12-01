VERSION 5.00
Begin VB.Form title 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "bspojian"
   ClientHeight    =   4950
   ClientLeft      =   8085
   ClientTop       =   3525
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "ËÎÌå"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "picn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   599
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Me.Width = 300 * 15
Me.Height = 225 * 15
Me.ForeColor = vbWhite
Me.CurrentX = 100
Me.CurrentY = 200
Me.FontSize = 10
End Sub

Private Sub Form_Paint()

'Randomize
Me.PaintPicture LoadResPicture(101, 0), 0, 0
Me.Print "sfe0.72.20"
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Main
Unload Me
End Sub
