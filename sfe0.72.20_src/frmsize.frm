VERSION 5.00
Begin VB.Form frmsize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置贴图编号"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton OK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Text            =   "0"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Text            =   "2"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "结束贴图"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "起始贴图"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "放大倍数"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "frmsize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pic1, pic2, pscale As Long
Public cok As Long
Private Sub cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim I As Long
    Me.Caption = StrUnicode(Me.Caption)
    For I = 0 To Me.Controls.Count - 1
        Call SetCaption(Me.Controls(I))
    Next I
cok = 0
c_Skinner.AttachSkin Me.hWnd
End Sub

Private Sub OK_Click()
pic1 = Val(Text3)
pic2 = Val(Text2)
pscale = Val(Text1)
cok = 1
Unload Me
End Sub
