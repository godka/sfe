VERSION 5.00
Begin VB.Form frminsertstatement 
   Caption         =   "插入指令"
   ClientHeight    =   4725
   ClientLeft      =   5040
   ClientTop       =   3540
   ClientWidth     =   9345
   LinkTopic       =   "Form2"
   ScaleHeight     =   4725
   ScaleWidth      =   9345
   Begin VB.ListBox combostatment 
      Height          =   2400
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   7095
   End
   Begin VB.Frame Frame2 
      Caption         =   "跳转方向"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   3840
      Width           =   7095
      Begin VB.OptionButton Option4 
         Caption         =   "向上跳转"
         Height          =   375
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "向下跳转"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "条件指令转移选择"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   7095
      Begin VB.OptionButton Option2 
         Caption         =   "否（条件不满足）跳转"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "是（条件满足）跳转"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "确定"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "请选择要插入的指令："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frminsertstatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OK As Long

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
Dim tmpstat As Statement
Dim labelstat As Statement
Dim Index As Long
    OK = 1
    Index = combostatment.ListIndex
    If Index < 0 Then Exit Sub
    Set tmpstat = New Statement
    If Index = KdefNum + 1 Then
        tmpstat.id = &HFFFF
        tmpstat.DataNum = 0
        KdefInfo(frmmain.Combokdef.ListIndex).kdef.Add tmpstat, , frmmain.listkdef.ListIndex + 1
        Call re_Analysis(frmmain.Combokdef.ListIndex)
        Unload Me
        Exit Sub
    End If
    tmpstat.id = Index
    tmpstat.DataNum = StatAttrib(Index).Length - 1
    If StatAttrib(tmpstat.id).isGoto = 0 Then
        KdefInfo(frmmain.Combokdef.ListIndex).kdef.Add tmpstat, , frmmain.listkdef.ListIndex + 1
    Else
        If Option3.Value = True Then
            tmpstat.yesOffset = StatAttrib(tmpstat.id).yesOffset
            tmpstat.noOffset = StatAttrib(tmpstat.id).noOffset
            If Option1.Value = True Then
                tmpstat.isGoto = 1
                tmpstat.data(tmpstat.yesOffset - 1) = 1
            Else
                tmpstat.isGoto = 2
                tmpstat.data(tmpstat.noOffset - 1) = 1
            End If
            tmpstat.gotoLabel = "Label" & KdefInfo(frmmain.Combokdef.ListIndex).numlabel
            KdefInfo(frmmain.Combokdef.ListIndex).numlabel = KdefInfo(frmmain.Combokdef.ListIndex).numlabel + 1
            KdefInfo(frmmain.Combokdef.ListIndex).kdef.Add tmpstat, , frmmain.listkdef.ListIndex + 1
            Set labelstat = New Statement
            labelstat.islabel = True
            labelstat.note = tmpstat.gotoLabel
            Set tmpstat = New Statement
            tmpstat.id = 0
            tmpstat.DataNum = 0
            KdefInfo(frmmain.Combokdef.ListIndex).kdef.Add tmpstat, , frmmain.listkdef.ListIndex + 2
            KdefInfo(frmmain.Combokdef.ListIndex).kdef.Add labelstat, , frmmain.listkdef.ListIndex + 3
        Else
            tmpstat.yesOffset = StatAttrib(tmpstat.id).yesOffset
            tmpstat.noOffset = StatAttrib(tmpstat.id).noOffset
            If Option1.Value = True Then
                tmpstat.isGoto = 1
                tmpstat.data(tmpstat.yesOffset - 1) = -tmpstat.DataNum - 2
            Else
                tmpstat.isGoto = 2
                tmpstat.data(tmpstat.noOffset - 1) = -tmpstat.DataNum - 2
            End If
            tmpstat.gotoLabel = "Label" & KdefInfo(frmmain.Combokdef.ListIndex).numlabel
            KdefInfo(frmmain.Combokdef.ListIndex).numlabel = KdefInfo(frmmain.Combokdef.ListIndex).numlabel + 1
            
            Set labelstat = New Statement
            labelstat.islabel = True
            labelstat.note = tmpstat.gotoLabel
            KdefInfo(frmmain.Combokdef.ListIndex).kdef.Add labelstat, , frmmain.listkdef.ListIndex + 1
            Set labelstat = New Statement
            labelstat.id = 0
            labelstat.DataNum = 0
            KdefInfo(frmmain.Combokdef.ListIndex).kdef.Add labelstat, , frmmain.listkdef.ListIndex + 2
            
            KdefInfo(frmmain.Combokdef.ListIndex).kdef.Add tmpstat, , frmmain.listkdef.ListIndex + 3

        End If
    End If
    Call re_Analysis(frmmain.Combokdef.ListIndex)
    Unload Me
End Sub

Private Sub ComboStatment_click()
Dim Index As Long
    Index = combostatment.ListIndex
    If Index < 0 Then Exit Sub
    If Index = KdefNum + 1 Then Exit Sub
    If StatAttrib(Index).isGoto = 1 Then
        Frame1.Visible = True
    End If
End Sub

Private Sub Form_Load()
Dim i As Long
    OK = 0
    Me.Caption = LoadResStr(1001)
    Label1.Caption = LoadResStr(1002)
    
    Frame1.Caption = LoadResStr(1002)
    Option1.Caption = LoadResStr(1004)
    Option2.Caption = LoadResStr(1005)
    cmdok.Caption = LoadResStr(102)
     cmdcancel.Caption = LoadResStr(103)
    
    For i = 0 To KdefNum
        combostatment.AddItem i & "(" & Hex(i) & ")：" & StatAttrib(i).notes
    Next i
    combostatment.AddItem -1 & "：" & LoadResStr(1006)
    Frame1.Visible = False
    combostatment.ListIndex = 0
    c_Skinner.AttachSkin Me.hWnd
End Sub

