VERSION 5.00
Begin VB.Form frmTeam 
   Caption         =   "Kys_Edit"
   ClientHeight    =   3840
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   9270
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar ExpLevel 
      Height          =   255
      Left            =   120
      Max             =   100
      TabIndex        =   23
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox TextExp 
      Height          =   270
      Left            =   2760
      TabIndex        =   21
      Text            =   "TextExp"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存所有数据"
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox TextExtra 
      Height          =   270
      Left            =   6720
      TabIndex        =   15
      Text            =   "TextExtra"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox ComboMagic 
      Height          =   300
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ComboBox ComboThing 
      Height          =   300
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.HScrollBar MatchLevel 
      Height          =   255
      Left            =   120
      Max             =   100
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox TextEffect 
      Height          =   270
      Left            =   2760
      TabIndex        =   3
      Text            =   "TextEffect"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.HScrollBar EffectLevel 
      Height          =   255
      Left            =   120
      Max             =   100
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox ComboTeam 
      Height          =   300
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.HScrollBar TeamLevel 
      Height          =   255
      Left            =   120
      Max             =   100
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "升级经验"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "经验值"
      Height          =   255
      Left            =   2760
      TabIndex        =   22
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "帧数"
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "人物"
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "加成"
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "武功"
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "武器"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "武器配合"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "武功动画效果帧数"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "离队队友列表"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim teamR() As Integer
Dim effectR() As Integer
Dim matchR() As Integer
Dim expR() As Integer
Dim match1(), match2(), match0() As Integer
Dim idnum(0 To 5) As Integer


Private Sub ComboMagic_Click()
On Error Resume Next
    matchR((MatchLevel.Value) * 3 + 1) = ComboMagic.ListIndex
End Sub

Private Sub ComboTeam_Click()
On Error Resume Next
    teamR(TeamLevel.Value) = ComboTeam.ListIndex
End Sub

Private Sub ComboThing_click()
On Error Resume Next
    matchR((MatchLevel.Value) * 3) = ComboThing.ListIndex
End Sub

Private Sub Command1_Click()
On Error Resume Next
'1 by 1
    filenum = OpenBin(G_Var.JYPath & G_Var.Leave, "W")
        For i = 0 To idnum(0)
            Put filenum, , teamR(i)
        Next i
    Close (filenum)

    filenum = OpenBin(G_Var.JYPath & G_Var.Effect, "W")
        For i = 0 To idnum(1)
            Put filenum, , effectR(i)
        Next i
    Close (filenum)
    
    filenum = OpenBin(G_Var.JYPath & G_Var.Match, "W")
        For i = 0 To idnum(2)
            Put filenum, , matchR(i)
        Next i
    Close (filenum)
    
    filenum = OpenBin(G_Var.JYPath & G_Var.Exp, "W")
        For i = 0 To idnum(3)
            Put filenum, , expR(i)
        Next i
    Close (filenum)
MsgBox "done~", vbOKOnly, "sfe0.72"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub EffectLevel_Change()
On Error Resume Next
    TextEffect = effectR(EffectLevel.Value)
    Label2.Caption = EffectLevel.Value
End Sub

Private Sub ExpLevel_Change()
On Error Resume Next
    TextExp = expR(ExpLevel.Value)
    Label14.Caption = ExpLevel.Value
End Sub

Private Sub Form_Load()
Dim i, j As Long
Dim match1(), match2(), match0() As Integer
    For i = 0 To Me.Controls.Count - 1
         Call SetCaption(Me.Controls(i))
    Next i
    Call SetCombo(Me)
    On Error Resume Next
    'list of leave
    filenum = OpenBin(G_Var.JYPath & G_Var.Leave, "R")
   '     idnum(0) = LOF(filenum) / 4 - 1
        idnum(0) = 99
        TeamLevel.Max = idnum(0)
        TeamLevel.Value = 0
        ReDim teamR(idnum(0)) As Integer
        For i = 0 To idnum(0)
            Get filenum, , teamR(i)
        Next i
    Close (filenum)
    For i = 0 To PersonNum - 1
        ComboTeam.AddItem ((i) & Person(i).Name1)
    Next i
    For i = 0 To Thingsnum - 1
        ComboThing.AddItem ((i) & Things(i).Name1)
    Next i
     For i = 0 To WuGongnum - 1
        ComboMagic.AddItem ((i) & WuGong(i).Name1)
     Next i
'list of effect
    filenum = OpenBin(G_Var.JYPath & G_Var.Effect, "R")
 '       idnum(1) = LOF(filenum) / 2 - 1
        idnum(1) = 199
        EffectLevel.Max = idnum(1)
        EffectLevel.Value = 0
        ReDim effectR(idnum(1)) As Integer
        For i = 0 To idnum(1)
            Get filenum, , effectR(i)
        Next i
    Close (filenum)
    ComboTeam.ListIndex = teamR(0)
    TextEffect = effectR(0)
    Label2.Caption = EffectLevel.Value
    Label1.Caption = TeamLevel.Value


'list of match
    filenum = OpenBin(G_Var.JYPath & G_Var.Match, "R")
 '       idnum(2) = LOF(filenum) / 2 - 1
        idnum(2) = 100 * 3
        MatchLevel.Max = idnum(2) / 3 - 1
        MatchLevel.Value = 0
        ReDim matchR(idnum(2)) As Integer
        For i = 0 To idnum(2)
            Get filenum, , matchR(i)
        Next i
    Close (filenum)
    Label6.Caption = MatchLevel.Value
    ComboThing.ListIndex = matchR((MatchLevel.Value) * 3)
    ComboMagic.ListIndex = matchR((MatchLevel.Value) * 3 + 1)
    TextExtra = matchR((MatchLevel.Value) * 3 + 2)

'list of exp
    filenum = OpenBin(G_Var.JYPath & G_Var.Exp, "R")
'        idnum(3) = LOF(filenum) / 2 - 1
        idnum(3) = 99
        'MsgBox idnum(3)
        ExpLevel.Max = idnum(3)
        ExpLevel.Value = 0
        ReDim expR(idnum(3)) As Integer
        For i = 0 To idnum(3)
            Get filenum, , expR(i)
        Next i
    Close (filenum)
    Label14.Caption = ExpLevel.Value
    TextExp = expR(ExpLevel.Value)
    c_Skinner.AttachSkin Me.hwnd
End Sub

Private Sub MatchLevel_Change()
On Error Resume Next
    Label6.Caption = MatchLevel.Value
    ComboThing.ListIndex = matchR((MatchLevel.Value) * 3)
    ComboMagic.ListIndex = matchR((MatchLevel.Value) * 3 + 1)
    TextExtra = matchR((MatchLevel.Value) * 3 + 2)
End Sub

Private Sub TeamLevel_Change()
On Error Resume Next
    Label1.Caption = TeamLevel.Value
   ComboTeam.ListIndex = teamR(TeamLevel.Value)
End Sub


Private Sub TextEffect_Change()
On Error Resume Next
        effectR(EffectLevel.Value) = TextEffect
End Sub

Private Sub TextEffect_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
On Error Resume Next
        effectR(EffectLevel.Value) = TextEffect
End If
End Sub


Private Sub TextExp_Change()
On Error Resume Next
        expR(ExpLevel.Value) = TextExp
End Sub

Private Sub TextExtra_Change()
On Error Resume Next
     matchR((MatchLevel.Value) * 3 + 2) = TextExtra
End Sub

Private Sub TextExtra_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
     matchR((MatchLevel.Value) * 3 + 2) = TextExtra
End If
End Sub
