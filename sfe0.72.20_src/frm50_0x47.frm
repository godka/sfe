VERSION 5.00
Begin VB.Form frm50_0x47 
   Caption         =   "50ָ��47"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   5100
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frm50_0x47.frx":0000
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin sfe72.userVar userX 
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2143
   End
   Begin VB.Label Label1 
      Caption         =   "ս�����"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frm50_0x47"
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
 
    kk.data(1) = userX.Value
    kk.data(2) = userX.Text
    kk.data(3) = 0
    kk.data(4) = 0
    kk.data(5) = 0
    kk.data(6) = 0

    
    Unload Me
    
End Sub

 




Private Sub Form_Load()
Dim num50 As Long
Dim I As Long, j As Long
Dim rr As Long, gg As Long, bb As Long
Dim Color As Long
Dim s1 As String
    Call ConvertForm(Me)
    
    
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    
    
    userX.Text = kk.data(1)
    
    userX.SetCombo

       
    Call Set50Form(Me, kk.data(0))
        c_Skinner.AttachSkin Me.hWnd
End Sub

 

