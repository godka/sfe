VERSION 5.00
Begin VB.Form frmStatement_0x0a 
   Caption         =   "��������"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form2"
   ScaleHeight     =   2295
   ScaleWidth      =   5910
   StartUpPosition =   2  '��Ļ����
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
      Left            =   1080
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   240
      Width           =   855
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
      TabIndex        =   2
      Top             =   360
      Width           =   1335
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
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2655
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
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   855
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
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmStatement_0x0a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 10

Dim Index As Long
Dim kk As Statement

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    kk.data(0) = Comboperson.ListIndex
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

    
    txtid.Text = kk.id & "(" & Hex(kk.id) & ")"
    Comboperson.ListIndex = kk.data(0)
    
    
    Me.Caption = LoadResStr(521)
    Label1.Caption = LoadResStr(1102)
    Label2.Caption = LoadResStr(2001)
   
    cmdok.Caption = LoadResStr(102)
    cmdcancel.Caption = LoadResStr(103)

    c_Skinner.AttachSkin Me.hWnd
    
    
End Sub


