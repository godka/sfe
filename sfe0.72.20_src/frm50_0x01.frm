VERSION 5.00
Begin VB.Form frm50_0x01 
   Caption         =   "50ָ��1"
   ClientHeight    =   4020
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
   ScaleHeight     =   4020
   ScaleWidth      =   9255
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frm50_0x01.frx":0000
      Top             =   2880
      Width           =   4215
   End
   Begin sfe72.userVar userX 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
   End
   Begin VB.ComboBox comboType 
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin sfe72.userVar userI 
      Height          =   1215
      Left            =   2280
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
   End
   Begin sfe72.userVar userY 
      Height          =   1215
      Left            =   4800
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
   End
   Begin VB.Label Label7 
      Caption         =   ") = "
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "("
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "ֵY"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "�±�I"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "����/�ַ�������X"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frm50_0x01"
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
    kk.data(1) = userI.Value + userY.Value * 2
    kk.data(2) = ComboType.ListIndex
    kk.data(3) = userX.Text
    kk.data(4) = userI.Text
    kk.data(5) = userY.Text
    kk.data(6) = 0

    
    Unload Me
    
End Sub

 

Private Sub Form_Load()
Dim num50 As Long
Dim I As Long
Dim s1 As String
    Call ConvertForm(Me)
    
    ComboType.Clear
    ComboType.AddItem StrUnicode2("��ȡ16λ��")
    ComboType.AddItem StrUnicode2("��ȡ8λ�ֽ�")
    
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    
    Call Set50Form(Me, kk.data(0))
    
    ComboType.ListIndex = kk.data(2)
    userX.Text = kk.data(3)
    userI.Text = kk.data(4)
    userY.Text = kk.data(5)
    
    userI.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    userY.Value = IIf((kk.data(1) And &H2) > 0, 1, 0)
      
      
    userX.Showtype = False
    userX.SetCombo
    userY.SetCombo
    userI.SetCombo

   c_Skinner.AttachSkin Me.hWnd

End Sub

 
