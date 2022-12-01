VERSION 5.00
Begin VB.Form frm50_0x09 
   Caption         =   "50指令9"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   9690
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frm50_0x09.frx":0000
      Top             =   2160
      Width           =   7455
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "确定"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin sfe72.userVar userS 
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
   End
   Begin sfe72.userVar userFormat 
      Height          =   1095
      Left            =   2280
      TabIndex        =   5
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
      Value           =   1
   End
   Begin sfe72.userVar userX 
      Height          =   1095
      Left            =   4320
      TabIndex        =   6
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
      Value           =   1
   End
   Begin VB.Label Label6 
      Caption         =   "X"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "字符串变量S"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "字符串变量format"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frm50_0x09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Long
Dim kk As Statement


Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
 
    kk.data(1) = userX.Value
    kk.data(2) = userS.Text
    kk.data(3) = userFormat.Text
    kk.data(4) = userX.Text
    kk.data(5) = 0
    kk.data(6) = 0
    
    Unload Me
    
End Sub

 

Private Sub Form_Load()
Dim num50 As Long
Dim I As Long
Dim s1 As String
    Call ConvertForm(Me)
    
    
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)

    userS.Text = kk.data(2)
    userFormat.Text = kk.data(3)
    userX.Text = kk.data(4)
 
    
    userX.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    
    Call Set50Form(Me, kk.data(0))

    userS.Showtype = False
    userS.SetCombo
    userFormat.Showtype = False
    userFormat.SetCombo
    userX.SetCombo
c_Skinner.AttachSkin Me.hWnd
End Sub

 
