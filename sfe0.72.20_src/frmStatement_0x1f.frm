VERSION 5.00
Begin VB.Form frmStatement_0x1f 
   Caption         =   "钱是否够"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   2550
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
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
      Left            =   960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1080
      Width           =   855
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
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "确定"
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
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "取消"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "钱是否够"
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
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "指令id"
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
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmStatement_0x1f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 31



Dim Index As Long
Dim kk As Statement

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    kk.data(0) = Text1.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    txtid.Text = kk.id & "(" & Hex(kk.id) & ")"
    Text1.Text = kk.data(0)
    
    
    
    Me.Caption = LoadResStr(3101)
    Label1.Caption = LoadResStr(1102)
    Label2.Caption = LoadResStr(3101)
    cmdok.Caption = LoadResStr(102)
    cmdcancel.Caption = LoadResStr(103)
        c_Skinner.AttachSkin Me.hWnd

End Sub


