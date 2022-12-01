VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   12825
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox ListItem 
      Columns         =   3
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6720
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "双击修改"
      Top             =   1560
      Width           =   10455
   End
   Begin VB.ComboBox ComboType 
      Height          =   300
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox ComboNumber 
      Height          =   300
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdLoadRecord 
      Caption         =   "read"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveRecord 
      Caption         =   "save"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton CmdGetExcel 
      Caption         =   "导出excel"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton CmdPutExcel 
      Caption         =   "导入excel"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
