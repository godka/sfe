VERSION 5.00
Begin VB.Form frm50_0x31 
   Caption         =   "50ָ��31"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   10065
   StartUpPosition =   2  '��Ļ����
   Begin sfe72.userVar userI 
      Height          =   1455
      Left            =   5400
      TabIndex        =   12
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2566
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frm50_0x31.frx":0000
      Top             =   2040
      Width           =   9735
   End
   Begin VB.ComboBox ComboType 
      Height          =   300
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin sfe72.userVar userID 
      Height          =   1335
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2355
   End
   Begin sfe72.UserVar2 UserII 
      Height          =   1695
      Left            =   5640
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2990
   End
   Begin sfe72.userVar userX 
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2355
   End
   Begin sfe72.UserVar2 UserK 
      Height          =   1095
      Left            =   4320
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin VB.Label Label5 
      Caption         =   "��������ƫ��I"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ս���������"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "���Ա���X"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1095
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
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "frm50_0x31"
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
 
    kk.data(1) = userI.Value * 2 + userID.Value * 1 + userX.Value * 4
    kk.data(2) = userID.Text
    kk.data(3) = userI.Text
    kk.data(4) = userX.Text
    kk.data(5) = 0
    kk.data(6) = 0

    
    Unload Me
    
End Sub


Private Sub Form_Load()
Dim num50 As Long
Dim i As Long
Dim s1 As String
    Call ConvertForm(Me)
    
    
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)


    
    
    
    
    
    userID.Text = kk.data(2)
    
    userID.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    userI.Value = IIf((kk.data(1) And &H2) > 0, 1, 0)
    userX.Value = IIf((kk.data(1) And &H4) > 0, 1, 0)
    
    
    
    
    userID.SetCombo

    userI.Text = kk.data(3)
    userI.SetCombo
     
    userX.Text = kk.data(4)
    
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



