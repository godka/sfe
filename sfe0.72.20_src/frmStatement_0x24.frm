VERSION 5.00
Begin VB.Form frmStatement_0x24 
   Caption         =   "设置性别"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5115
   LinkTopic       =   "Form2"
   ScaleHeight     =   2250
   ScaleWidth      =   5115
   StartUpPosition =   2  '屏幕中心
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
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
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
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "判断jmp"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      Caption         =   "判断性别"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frmStatement_0x24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Index As Long
Dim kk As Statement
Private Sub Form_Load()
    Call ConvertForm(Me)
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    If kk.data(0) < 256 Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    c_Skinner.AttachSkin Me.hWnd
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    If Option1.Value = True Then
        kk.data(0) = 0
    Else
        kk.data(0) = 256
    End If
    Unload Me
End Sub
