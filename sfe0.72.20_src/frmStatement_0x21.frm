VERSION 5.00
Begin VB.Form frmStatement_0x21 
   Caption         =   "ѧ���书"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   5865
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox Check1 
      Caption         =   "��Ļ����ʾѧ���书"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ComboBox ComboWugong 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox ComboPerson 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtid 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   855
   End
   Begin VB.Label label3 
      Caption         =   "�书"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "ָ��id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frmStatement_0x21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 33
Dim Index As Long
Dim kk As Statement

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    kk.data(0) = Comboperson.ListIndex

    kk.data(1) = ComboWugong.ListIndex
    kk.data(2) = Check1.Value
    Unload Me
End Sub

Private Sub Form_Load()
Dim I As Long
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    
    Comboperson.Clear
    For I = 0 To PersonNum - 1
        Comboperson.AddItem I & "(" & Hex(I) & ")" & Person(I).Name1
    Next I
    ComboWugong.Clear
    For I = 0 To WuGongnum - 1
        ComboWugong.AddItem I & "(" & Hex(I) & ")" & WuGong(I).Name1
    Next I
    
    txtid.Text = kk.id & "(" & Hex(kk.id) & ")"
    Comboperson.ListIndex = kk.data(0)
    ComboWugong.ListIndex = kk.data(1)
    Check1.Value = kk.data(2)
    
    
    Me.Caption = LoadResStr(547)
    Label1.Caption = LoadResStr(1102)
    Label2.Caption = LoadResStr(2001)
    Label3.Caption = LoadResStr(3301)
    cmdok.Caption = LoadResStr(102)
    cmdcancel.Caption = LoadResStr(103)
    Check1.Caption = LoadResStr(3302)
    
    
        c_Skinner.AttachSkin Me.hWnd

End Sub





