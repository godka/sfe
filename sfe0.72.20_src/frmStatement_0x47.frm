VERSION 5.00
Begin VB.Form frmStatement_0x47 
   Caption         =   "跳转场景"
   ClientHeight    =   2970
   ClientLeft      =   5790
   ClientTop       =   1755
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   7980
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
      TabIndex        =   8
      Text            =   "frmStatement_0x47.frx":0000
      Top             =   1920
      Width           =   5175
   End
   Begin VB.CommandButton CmdCnl 
      Caption         =   "取消"
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "确定"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox TextY 
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Text            =   "0"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TextX 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox ComboScene 
      Height          =   300
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "坐标Y"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "坐标X"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "场景"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmStatement_0x47"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kk As Statement
Private Sub CmdCnl_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
kk.data(0) = ComboScene.ListIndex
kk.data(1) = TextX
kk.data(2) = TextY
Unload Me
End Sub


Private Sub Form_Load()
    Index = frmmain.listkdef.ListIndex
Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
   Call ConvertForm(Me)
For I = 0 To Scenenum - 1
    ComboScene.AddItem (I & Big5toUnicode(Scene(I).Name1, 10))
Next I
ComboScene.ListIndex = kk.data(0)
TextX = kk.data(1)
TextY = kk.data(2)
    c_Skinner.AttachSkin Me.hWnd

End Sub
