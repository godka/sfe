VERSION 5.00
Begin VB.Form frm50_0x08 
   Caption         =   "50ָ��8"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
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
   ScaleHeight     =   3075
   ScaleWidth      =   9255
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frm50_0x08.frx":0000
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin sfe72.userVar userID 
      Height          =   1095
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
   End
   Begin sfe72.userVar userX 
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
      Value           =   1
   End
   Begin VB.Label Label1 
      Caption         =   "="
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "�Ի�ID"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "�ַ�������X"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frm50_0x08"
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
 
    kk.data(1) = userID.Value
    kk.data(2) = userID.Text
    kk.data(3) = userX.Text
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

    userID.Text = kk.data(2)
    userX.Text = kk.data(3)
 
    
    userID.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    
    Call Set50Form(Me, kk.data(0))

    userID.SetCombo
    userX.Showtype = False
    userX.SetCombo
c_Skinner.AttachSkin Me.hWnd
End Sub

 
 
