VERSION 5.00
Begin VB.Form frm50_0x29 
   Caption         =   "50ָ��29"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   10215
   StartUpPosition =   2  '��Ļ����
   Begin sfe72.userVar userT 
      Height          =   1095
      Left            =   2040
      TabIndex        =   12
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin sfe72.userVar userN 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
   End
   Begin sfe72.userVar userX 
      Height          =   1095
      Left            =   4680
      TabIndex        =   10
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frm50_0x29.frx":0000
      Top             =   1560
      Width           =   9855
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "����"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "�书"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "�Ƿ���ʾ����"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "����ֵ"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "����"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "ս�����"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frm50_0x29"
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
    kk.data(1) = userN.Value + userT.Value * 2
    kk.data(2) = userN.Text
    kk.data(3) = userT.Text
    kk.data(4) = userX.Text
    If Option1 = True Then
       kk.data(5) = 1
     Else
       kk.data(5) = 0
    End If
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

    
    userN.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    userT.Value = IIf((kk.data(1) And &H2) > 0, 1, 0)

    
    
    userN.Text = kk.data(2)
    userN.SetCombo
    userT.Text = kk.data(3)
    userT.SetCombo
    
    userX.Showtype = False
    userX.Text = kk.data(4)
    userX.SetCombo
    
    
    
    

    Call Set50Form(Me, kk.data(0))
Option1 = True
c_Skinner.AttachSkin Me.hWnd
End Sub

