VERSION 5.00
Begin VB.Form frm50_0x45 
   Caption         =   "50ָ��45"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   8880
   StartUpPosition =   2  '��Ļ����
   Begin VB.OptionButton Option2 
      Caption         =   "��"
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "��"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frm50_0x45.frx":0000
      Top             =   2280
      Width           =   6615
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin sfe72.userVar userID 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
   End
   Begin sfe72.userVar userE 
      Height          =   1455
      Left            =   4680
      TabIndex        =   1
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2566
   End
   Begin VB.Label Label3 
      Caption         =   "��˸��ɫ  "
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "�Ƿ���˸ "
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "������ɫ "
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frm50_0x45"
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
    kk.data(1) = userID.Value + userE.Value * 4
    kk.data(2) = userID.Text
    If Option1 = True Then
        kk.data(3) = 0
      Else
         kk.data(3) = 1
    End If
    
    kk.data(4) = userE.Text
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
    
    userE.Value = IIf((kk.data(1) And &H4) > 0, 1, 0)


    
    
    userID.Text = kk.data(2)
    userID.SetCombo
    If kk.data(3) = 0 Then
      Option1 = True
     Else
      Option2 = True
    End If
    
    
    
    userE.Text = kk.data(4)
    userE.SetCombo
    
    
    
    

    Call Set50Form(Me, kk.data(0))
c_Skinner.AttachSkin Me.hWnd
End Sub


