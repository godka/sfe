VERSION 5.00
Begin VB.Form frm50_0x06 
   Caption         =   "50ָ��6"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
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
   ScaleHeight     =   2220
   ScaleWidth      =   6540
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frm50_0x06.frx":0000
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox txtname 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtVar 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "������(���10��Ӣ���ַ�)"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "�������"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frm50_0x06"
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
Dim x(4) As Integer
Dim I As Long
    kk.data(1) = txtVar.Text
    
    Call SetXchar(Trim(txtname.Text), x(0), x(1), x(2), x(3), x(4))
    For I = 0 To 4
        kk.data(2 + I) = x(I)
    Next I
    
    Unload Me
    
End Sub

Private Sub Form_Load()
Dim num50 As Long
Dim I As Long
Dim s1 As String
    Call ConvertForm(Me)
    
    
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    txtVar.Text = kk.data(1)

    txtname.Text = GetXchar(kk.data(2), kk.data(3), kk.data(4), kk.data(5), kk.data(6))
    
    Call Set50Form(Me, kk.data(0))
    c_Skinner.AttachSkin Me.hWnd
End Sub

 
 
 
