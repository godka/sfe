VERSION 5.00
Begin VB.Form frm50_0x19 
   Caption         =   "50ָ��18"
   ClientHeight    =   3525
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
   ScaleHeight     =   3525
   ScaleWidth      =   9690
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frm50_0x19.frx":0000
      Top             =   2040
      Width           =   4335
   End
   Begin sfe72.userVar userID 
      Height          =   1215
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2778
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
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2143
   End
   Begin VB.Label Label7 
      Caption         =   "="
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "����ID"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "����X"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frm50_0x19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Long
Dim kk As Statement
Dim OffsetName As Collection


 
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
    
    userID.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    
    userID.Text = kk.data(2)
    userID.SetCombo
    
    userX.Text = kk.data(3)
    userX.Showtype = False
    userX.SetCombo
    

    Call Set50Form(Me, kk.data(0))
 c_Skinner.AttachSkin Me.hWnd
End Sub

 
