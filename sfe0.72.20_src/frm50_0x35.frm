VERSION 5.00
Begin VB.Form frm50_0x35 
   Caption         =   "50ָ��35"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
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
   ScaleHeight     =   3435
   ScaleWidth      =   9510
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
      TabIndex        =   4
      Text            =   "frm50_0x35.frx":0000
      Top             =   2280
      Width           =   5295
   End
   Begin sfe72.userVar userX 
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2143
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin sfe72.userVar userVarX 
      Height          =   1215
      Left            =   2880
      TabIndex        =   5
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2143
   End
   Begin sfe72.userVar userVarY 
      Height          =   1215
      Left            =   4920
      TabIndex        =   6
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2143
   End
   Begin VB.Label Label2 
      Caption         =   "���X"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "���Y"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "����X"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frm50_0x35"
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
 
    kk.data(1) = Val(userX.Text)
    kk.data(2) = Val(userVarX.Text)
    kk.data(3) = Val(userVarY.Text)
    kk.data(4) = 0
    kk.data(5) = 0
    kk.data(6) = 0

    
    Unload Me
    
End Sub

 




Private Sub Form_Load()
Dim num50 As Long
Dim i As Long, j As Long
Dim rr As Long, gg As Long, bb As Long
Dim Color As Long
Dim s1 As String
    Call ConvertForm(Me)
    
    
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    
    
    userX.Text = kk.data(1)
    userX.Showtype = False
    userX.SetCombo

    userVarX.Text = kk.data(2)
    userVarX.Showtype = False
    userVarX.SetCombo

    userVarY.Text = kk.data(3)
    userVarY.Showtype = False
    userVarY.SetCombo

       
    Call Set50Form(Me, kk.data(0))
     c_Skinner.AttachSkin Me.hWnd
End Sub

 
