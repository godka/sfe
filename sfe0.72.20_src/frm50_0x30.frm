VERSION 5.00
Begin VB.Form frm50_0x30 
   Caption         =   "50ָ��30"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   10395
   StartUpPosition =   2  '��Ļ����
   Begin sfe72.userVar userI 
      Height          =   1215
      Left            =   2640
      TabIndex        =   9
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frm50_0x30.frx":0000
      Top             =   1920
      Width           =   9735
   End
   Begin sfe72.userVar userID 
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2355
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin sfe72.userVar userX 
      Height          =   1335
      Left            =   5400
      TabIndex        =   0
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "ս���������"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "��������ƫ��I"
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "���Ա���X"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm50_0x30"
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
 
    kk.data(1) = userI.Value * 2 + userID.Value * 1
    kk.data(2) = userID.Text
    kk.data(3) = userI.Text
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


    
    
    
    
    
    
    
    userID.Text = kk.data(2)
    
    userID.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    userI.Value = IIf((kk.data(1) And &H2) > 0, 1, 0)
    
    
    
    userID.SetCombo

    userI.Text = kk.data(3)
    userI.SetCombo
     
    userX.Text = kk.data(4)
    userX.Showtype = False
    userX.SetCombo

    Call Set50Form(Me, kk.data(0))
 c_Skinner.AttachSkin Me.hWnd
End Sub

 
Public Function GetOffsetname(id As Long) As String
    On Error GoTo Label1:
    GetOffsetname = OffsetName.Item("ID" & id)
     Exit Function
Label1:
    GetOffsetname = ""
    
End Function


