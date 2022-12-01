VERSION 5.00
Begin VB.Form frmheadnum 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "调整头像"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3870
   LinkTopic       =   "None"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   2160
      ScaleHeight     =   795
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "说话人头像"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1695
      Begin VB.PictureBox pic1 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         ScaleHeight     =   85
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   77
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox Comboperson 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmheadnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OK As Long, NHeadnum As Long


Private Sub Comboperson_click()
Dim i, offset As Long
If Comboperson.ListIndex = 0 Then Exit Sub
       pic1.Cls
 '      pic1.Refresh
If G_Var.EditMode = "classic" Then
        Call ShowPicDIB(HeadPic(Comboperson.ListIndex - 1), pic1.hDC, 0, pic1.ScaleHeight)
Else
    
         offset = KGoffset(Comboperson.ListIndex - 1)
        Call DrawPng(G_Var.JYPath & G_Var.NewHeadGRP, offset, frmheadnum.pic1, frmheadnum.Picture1, 0, 0)
 End If
End Sub

Private Sub Command2_Click()
OK = 1
NHeadnum = Comboperson.Text
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Long
OK = 0
    Me.Caption = StrUnicode(Me.Caption)
    For i = 0 To Me.Controls.Count - 1
        Call SetCaption(Me.Controls(i))
    Next i
    Comboperson.AddItem ("-2")
If G_Var.EditMode = "classic" Then
    For i = 0 To Headnum - 1
        Comboperson.AddItem (i)
    Next i
Else
    For i = 0 To NewHeadNum - 1
        Comboperson.AddItem (i)
    Next i
End If
Comboperson.ListIndex = 0
    c_Skinner.AttachSkin Me.hWnd

End Sub
