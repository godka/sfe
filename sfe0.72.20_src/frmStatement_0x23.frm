VERSION 5.00
Begin VB.Form frmStatement_0x23 
   Caption         =   "���������书"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   6930
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtPower 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1200
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtWugongnum 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
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
      Left            =   4320
      TabIndex        =   2
      Top             =   1200
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
      Left            =   4320
      TabIndex        =   1
      Top             =   240
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
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "�书������"
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
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label label4 
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
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "�书���"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1200
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
      Top             =   720
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmStatement_0x23"
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
    kk.data(0) = Comboperson.ListIndex
    kk.data(1) = txtWugongnum.Text
    kk.data(2) = ComboWugong.ListIndex
    kk.data(3) = txtPower.Text
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
    txtWugongnum.Text = kk.data(1)
    ComboWugong.ListIndex = kk.data(2)
    txtPower.Text = kk.data(3)
    
    
    Me.Caption = LoadResStr(3501)
    Label1.Caption = LoadResStr(1102)
    Label2.Caption = LoadResStr(2001)
    Label3.Caption = LoadResStr(3502)
     Label4.Caption = LoadResStr(550)
      Label5.Caption = LoadResStr(551)
    cmdok.Caption = LoadResStr(102)
    cmdcancel.Caption = LoadResStr(103)

        c_Skinner.AttachSkin Me.hWnd

    
    
End Sub




