VERSION 5.00
Begin VB.Form frm50_0x00 
   Caption         =   "50ָ��0"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
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
   ScaleHeight     =   2970
   ScaleWidth      =   7695
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
      TabIndex        =   7
      Text            =   "frm50_0x00.frx":0000
      Top             =   2040
      Width           =   3375
   End
   Begin sfe72.userVar userVar 
      Height          =   975
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1720
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   " = "
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "����ֵ"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "�������"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frm50_0x00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim kk As Statement


Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    kk.data(1) = userVar.Text
    kk.data(2) = txtValue.Text
    kk.data(3) = 0
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
    

    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(frmmain.listkdef.ListIndex + 1)
    
    Call Set50Form(Me, kk.data(0))
    
    userVar.Text = kk.data(1)
    txtValue.Text = kk.data(2)
    
    userVar.Showtype = False
    userVar.SetCombo
    c_Skinner.AttachSkin Me.hWnd
    
End Sub

Private Sub Label1_Click()

End Sub

