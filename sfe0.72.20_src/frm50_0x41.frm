VERSION 5.00
Begin VB.Form frm50_0x41 
   Caption         =   "50ָ��41"
   ClientHeight    =   3990
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
   ScaleHeight     =   3990
   ScaleWidth      =   9690
   StartUpPosition =   2  '��Ļ����
   Begin sfe72.userVar userN 
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1935
      _extentx        =   3413
      _extenty        =   2143
   End
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
      Text            =   "frm50_0x41.frx":0000
      Top             =   3000
      Width           =   5175
   End
   Begin VB.ComboBox ComboType 
      Height          =   345
      ItemData        =   "frm50_0x41.frx":0047
      Left            =   1560
      List            =   "frm50_0x41.frx":0049
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2775
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
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
      _extentx        =   3413
      _extenty        =   2143
   End
   Begin sfe72.userVar userY 
      Height          =   1215
      Left            =   4800
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
      _extentx        =   3413
      _extenty        =   2778
   End
   Begin VB.Label Label2 
      Caption         =   "������X"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "������Y"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "��ͼ���N"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "��ͼ���"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm50_0x41"
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
 
    kk.data(1) = userX.Value + userY.Value * 2 + userN.Value * 4
    kk.data(2) = ComboType.ListIndex
    kk.data(3) = userX.Text
    kk.data(4) = userY.Text
    kk.data(5) = userN.Text
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

    ComboType.Clear
    ComboType.AddItem StrUnicode2("������ͼ")
    ComboType.AddItem StrUnicode2("ͷ����ͼ")
    ComboType.AddItem StrUnicode2("��Ʒ��ͼ")
    
    ComboType.ListIndex = kk.data(2)
    
    userX.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    userY.Value = IIf((kk.data(1) And &H2) > 0, 1, 0)
    userN.Value = IIf((kk.data(1) And &H4) > 0, 1, 0)
    
    userX.Text = kk.data(3)
    userX.SetCombo
    userY.Text = kk.data(4)
    userY.SetCombo
    
    userN.Text = kk.data(5)
    userN.SetCombo

    Call Set50Form(Me, kk.data(0))
 c_Skinner.AttachSkin Me.hWnd
End Sub

 
