VERSION 5.00
Begin VB.Form define 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自定义调用事件"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7620
   Icon            =   "define.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   4080
      TabIndex        =   6
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "放弃"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "现有自定义"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "需要自定义为"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "事件编号"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "define"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
MsgBox "blank"
Unload Me
End Sub
