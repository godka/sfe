VERSION 5.00
Begin VB.Form frm50_0x11 
   Caption         =   "50ָ��11"
   ClientHeight    =   2835
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
   ScaleHeight     =   2835
   ScaleWidth      =   9690
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frm50_0x11.frx":0000
      Top             =   1920
      Width           =   4215
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin sfe72.userVar userX 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
   End
   Begin sfe72.userVar userA 
      Height          =   1095
      Left            =   2640
      TabIndex        =   5
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
      Value           =   1
   End
   Begin sfe72.userVar userB 
      Height          =   1095
      Left            =   5040
      TabIndex        =   7
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
      Value           =   1
   End
   Begin VB.Label Label8 
      Caption         =   "+"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "�ַ�������B"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "="
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "�ַ�������X"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "�ַ�������A"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frm50_0x11"
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
 
    kk.data(1) = userX.Text
    kk.data(2) = userA.Text
    kk.data(3) = userB.Text
    kk.data(4) = 0
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

    userX.Text = kk.data(1)
    userA.Text = kk.data(2)
    userB.Text = kk.data(3)
 
    Call Set50Form(Me, kk.data(0))

    userX.Showtype = False
    userX.SetCombo
    userA.Showtype = False
    userA.SetCombo
    userB.Showtype = False
    userB.SetCombo
c_Skinner.AttachSkin Me.hWnd
End Sub

 
Private Sub userVar1_GotFocus()

End Sub

