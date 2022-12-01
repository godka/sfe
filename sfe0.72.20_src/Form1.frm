VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "金庸群侠传事件&对话修改器"
   ClientHeight    =   9780
   ClientLeft      =   2745
   ClientTop       =   1545
   ClientWidth     =   16005
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   652
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1067
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   9615
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   8175
      Begin VB.ListBox listkdef 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8985
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   34
         Top             =   360
         Width           =   7935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "事件编辑"
      Enabled         =   0   'False
      Height          =   7575
      Left            =   8280
      TabIndex        =   13
      Top             =   2160
      Width           =   4455
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   7080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command4 
         Caption         =   "X 2"
         Height          =   375
         Left            =   3120
         TabIndex        =   35
         Top             =   3360
         Width           =   1215
      End
      Begin VB.ComboBox Combokdef 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton cmdmodifykdef 
         Caption         =   "修改事件"
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
         Left            =   600
         TabIndex        =   27
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmddeletekdef 
         Caption         =   "删除事件"
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
         Left            =   600
         TabIndex        =   26
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdinsertkdef 
         Caption         =   "增加事件"
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
         Left            =   600
         TabIndex        =   25
         Top             =   1080
         Width           =   975
      End
      Begin VB.ListBox Liststatement 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5160
         Left            =   1800
         TabIndex        =   24
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdinsertstatement 
         Caption         =   "插入指令"
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
         Left            =   600
         TabIndex        =   23
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdmodifystatement 
         Caption         =   "修改指令"
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
         Left            =   600
         TabIndex        =   22
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmddeletestatement 
         Caption         =   "删除指令"
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
         Left            =   600
         TabIndex        =   21
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdmodifyword 
         Caption         =   "修改指令字"
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
         Left            =   3120
         TabIndex        =   20
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdwizard 
         Caption         =   "指令编辑指导"
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
         Left            =   2760
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdcopy 
         Caption         =   "复制到剪贴板"
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
         Left            =   360
         TabIndex        =   18
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton cmdpaste 
         Caption         =   "从剪贴板复制"
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
         Left            =   360
         TabIndex        =   17
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton cmdup 
         Caption         =   "上"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmddown 
         Caption         =   "下"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3120
         TabIndex        =   14
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "指令编辑"
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
         Left            =   1800
         TabIndex        =   32
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "编辑指令字"
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
         Left            =   3120
         TabIndex        =   31
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label txtword 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   30
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "事件指令"
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
         TabIndex        =   29
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
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
      Left            =   12960
      TabIndex        =   10
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdExportKDef 
      Caption         =   "导出全部事件"
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
      Left            =   12960
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdExportTalk 
      Caption         =   "导出全部对话"
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
      Left            =   12960
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdsaveFile 
      Caption         =   "save文件"
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
      Left            =   12960
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "对话编辑"
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
      Height          =   1935
      Left            =   8280
      TabIndex        =   0
      Tag             =   "301"
      Top             =   120
      Width           =   6375
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "12"
         Top             =   1395
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "重排星号"
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdinserttalk 
         Caption         =   "增加对话"
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
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmddeletetalk 
         Caption         =   "删除对话"
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
         Left            =   5040
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdmodifytalk 
         Caption         =   "修改对话"
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
         Left            =   5040
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txttalk 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   2
         Top             =   600
         Width           =   4695
      End
      Begin VB.ComboBox Combotalk 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lbltalk 
         Caption         =   "选择星号（默认为12）"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public listlabel As Long

Private Sub cmd_about_Click()
'    frmAbout.Show vbModal
End Sub


Private Sub cmdaddWord_Click()

End Sub

Private Sub cmdCopy_Click()
Dim Index As Long
Dim j As Long
Dim stat As Statement
Dim tmpstr As String
Dim datastr As String
     
    Index = Combokdef.ListIndex
    If Index < 0 Then Exit Sub
    
    datastr = ";Kdefnum=" & Index & vbCrLf
    For Each stat In KdefInfo(Index).kdef
        If stat.islabel = True Then
            tmpstr = ";" & stat.note & vbCrLf
        Else
        
            tmpstr = "  " & stat.id
            For j = 0 To stat.DataNum - 1
                tmpstr = tmpstr & " " & stat.data(j)
            Next j
            tmpstr = fmt(tmpstr, 30)
            tmpstr = tmpstr & "   ;" & stat.note & vbCrLf
        End If
        datastr = datastr & tmpstr
     Next stat
    Clipboard.Clear
    Clipboard.SetText datastr
    
End Sub

Private Sub cmddeletekdef_Click()
'    If MsgBox(LoadResStr(129), vbOKCancel, LoadResStr(121)) = vbOK Then
        If Combokdef.ListIndex = numkdef - 1 Then
            numkdef = numkdef - 1
            ReDim Preserve KdefInfo(numkdef - 1)
            ReDim Preserve KDEFIDX(numkdef)
            Combokdef.RemoveItem numkdef
            Combokdef.ListIndex = numkdef - 1
        Else
            MsgBox LoadResStr(130)
        End If
  '   End If
End Sub

Private Sub cmddeletestatement_Click()
Dim Index As Long
Dim kdef As Collection
Dim stat As Statement
Dim note As String
Dim i As Long
    If listlabel = 1 Then Exit Sub
    Index = listkdef.ListIndex
    If listkdef.ListCount = 1 Then    ' 剩下最后一条指令不能删除
        Exit Sub
    End If
    Set kdef = KdefInfo(Combokdef.ListIndex).kdef
    If Index < 0 Then Exit Sub
    If MsgBox(LoadResStr(131) & kdef(Index + 1).note, vbOKCancel) = vbCancel Then Exit Sub
    If kdef(Index + 1).islabel = True Then
        MsgBox LoadResStr(132)
        Exit Sub
    End If
    If kdef(Index + 1).isGoto <> 0 Then
        note = kdef(Index + 1).gotoLabel
        kdef.Remove Index + 1
        For i = 1 To kdef.Count
            If kdef(i).islabel = True Then
                If kdef(i).note = note Then
                    kdef.Remove i
                    Exit For
                End If
            End If
        Next i
'        If kdef(Index + 1).gotoLabel = kdef(Index + 2).note Then
'            If kdef(Index + 2).islabel = True Then
'                kdef.Remove Index + 2
'                kdef.Remove Index + 1
'            Else
'                MsgBox LoadResStr(133)
'            End If
'        ElseIf kdef(Index + 1).gotoLabel = kdef(Index).note Then
'            If Index > 0 Then
'                If kdef(Index).islabel = True Then
'                    kdef.Remove Index + 1
'                    kdef.Remove Index
'                Else
'                    MsgBox LoadResStr(133)
'                End If
'            End If
'        End If
    Else
        kdef.Remove Index + 1
    End If
    cmdmodifykdef_Click
    listkdef.Clear
    For Each stat In kdef
        listkdef.AddItem stat.note
    Next
    If listkdef.ListCount > Index + 1 Then
        listkdef.ListIndex = Index
    Else
        listkdef.ListIndex = listkdef.ListCount - 1
    End If
    listkdef.SetFocus
End Sub

Private Sub cmddeletetalk_Click()
'    If MsgBox(LoadResStr(120), vbOKCancel, LoadResStr(121)) = vbOK Then
        If Combotalk.ListIndex = numtalk - 1 Then
 '           numtalk = 2
            If numtalk <= 1 Then numtalk = 2
            numtalk = numtalk - 1
            ReDim Preserve Talk(numtalk - 1)
            ReDim Preserve TalkIdx(numtalk)
            refresh_combotalk
        Else
            MsgBox LoadResStr(122), , LoadResStr(121)
        End If
 '    End If
        
End Sub

Private Sub cmddown_Click()
Dim stat As Statement
Dim Index As Long
    Index = listkdef.ListIndex
    If Index = -1 Then Exit Sub
    If Index = KdefInfo(Combokdef.ListIndex).kdef.Count - 1 Then Exit Sub
     
    Set stat = KdefInfo(Combokdef.ListIndex).kdef(Index + 1)
    
    If Index >= 1 Then
        If KdefInfo(Combokdef.ListIndex).kdef(Index + 2).islabel = True And KdefInfo(Combokdef.ListIndex).kdef(Index).gotoLabel = KdefInfo(Combokdef.ListIndex).kdef(Index + 2).note Then
            Exit Sub
        End If
        If KdefInfo(Combokdef.ListIndex).kdef(Index).islabel = True And KdefInfo(Combokdef.ListIndex).kdef(Index + 2).gotoLabel = KdefInfo(Combokdef.ListIndex).kdef(Index).note Then
            Exit Sub
        End If
    End If
    If Index < KdefInfo(Combokdef.ListIndex).kdef.Count - 2 Then
        If stat.islabel = True And KdefInfo(Combokdef.ListIndex).kdef(Index + 3).gotoLabel = stat.note Then
            Exit Sub
        End If
        If KdefInfo(Combokdef.ListIndex).kdef(Index + 3).islabel = True And KdefInfo(Combokdef.ListIndex).kdef(Index + 3).note = stat.gotoLabel Then
            Exit Sub
        End If
    End If
    
    
    
    KdefInfo(Combokdef.ListIndex).kdef.Remove Index + 1
    KdefInfo(Combokdef.ListIndex).kdef.Add stat, , , Index + 1
     
    Call re_Analysis(Combokdef.ListIndex)
    cmdmodifykdef_Click
    
    
    listkdef.Clear
    For Each stat In KdefInfo(Combokdef.ListIndex).kdef
        listkdef.AddItem stat.note
    Next
    listkdef.ListIndex = Index + 1
     listkdef.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExportKDef_Click()
    mnu_Exportkdef_Click
End Sub

Private Sub cmdExportTalk_Click()
    mnu_exporttalkfile_Click
End Sub

Private Sub cmdinsertkdef_Click()
Dim tmpstat As Statement
    numkdef = numkdef + 1
    ReDim Preserve KdefInfo(numkdef - 1)
    ReDim Preserve KDEFIDX(numkdef)
    
    Set KdefInfo(numkdef - 1).kdef = New Collection
'    Set tmpstat = New Statement
'    tmpstat.Id = &HFFFF
'    tmpstat.DataNum = 0
'    KdefInfo(numkdef - 1).kdef.Add tmpstat
 '   Call re_Analysis(numkdef - 1)
 '   cmdmodifykdef_Click
    KdefInfo(numkdef - 1).DataLong = 2
    ReDim KdefInfo(numkdef - 1).data(1)
    KdefInfo(numkdef - 1).data(0) = 0
    
    KdefInfo(numkdef - 1).data(1) = &HFFFF
    
 
    Combokdef.AddItem numkdef - 1

    MsgBox LoadResStr(308) & numkdef - 1 & LoadResStr(134)
   
End Sub



Private Sub cmdinsertstatement_Click()
Dim stat As Statement
Dim Index As Long
    Index = listkdef.ListIndex
    If Index = -1 Then Exit Sub
    frminsertstatement.Show vbModal
    If frminsertstatement.OK = 1 Then
        cmdmodifykdef_Click
    
    
    
        listkdef.Clear
        For Each stat In KdefInfo(Combokdef.ListIndex).kdef
            listkdef.AddItem stat.note
        Next
        listkdef.ListIndex = Index
        ModifyStatementWizard False
        cmdmodifystatement_Click
        listkdef.SetFocus
    End If
End Sub

Private Sub cmdinserttalk_Click()
    numtalk = numtalk + 1
    ReDim Preserve Talk(numtalk - 1)
    ReDim Preserve TalkIdx(numtalk)
    Talk(numtalk - 1) = LoadResStr(123)
    refresh_combotalk
    Combotalk.ListIndex = numtalk - 1
    MsgBox LoadResStr(124) & numtalk - 1 & LoadResStr(125)
End Sub

Private Sub cmdmodifykdef_Click()
Dim ii As Long

    ii = Combokdef.ListIndex
    If ii < 0 Then Exit Sub
    
    Call modifykdef(ii)
    
    
End Sub

Private Sub cmdmodifystatement_Click()
Dim Index As Long
Dim tmpstate As Statement
Dim stat As Statement
Dim i As Long
    If listlabel = 1 Then Exit Sub
    Index = listkdef.ListIndex
    If Index = -1 Then Exit Sub
    Set tmpstate = KdefInfo(Combokdef.ListIndex).kdef(Index + 1)
    tmpstate.id = Liststatement.List(0)
    For i = 0 To tmpstate.DataNum - 1
        tmpstate.data(i) = Liststatement.List(i + 1)
    Next i
    
    Call re_Analysis(Combokdef.ListIndex)
    
    
    listkdef.Clear
    
    
    For Each stat In KdefInfo(Combokdef.ListIndex).kdef
        listkdef.AddItem stat.note
    Next
    On Error Resume Next
    If (listkdef.ListCount - 1 - Index) < 17 Then
        listkdef.ListIndex = Index + listkdef.ListCount - 1 - Index
    Else
        listkdef.ListIndex = Index + 17
    End If
    listkdef.Selected(Index) = True
    listkdef.SetFocus
    
End Sub

Private Sub CmdModifyTalk_Click()
    Talk(Combotalk.ListIndex) = txttalk.Text
    refresh_combotalk
End Sub

Private Sub cmdmodifyword_Click()
If Text1.Text = "" Then
MsgBox "数据不能为空", vbOKOnly, "警告"
Exit Sub
End If
Dim Index As Long
    Index = Liststatement.ListIndex
    If Index >= 0 Then
        Liststatement.List(Liststatement.ListIndex) = Text1.Text
    End If
End Sub

Private Sub cmdPaste_Click()
Dim datastr As String
Dim tmparray() As String
Dim tmpdata() As String
Dim Index As Long
Dim i As Long, j As Long
Dim num As Long
Dim pos As Long

    Index = Combokdef.ListIndex
    If Index < 0 Then Exit Sub
    datastr = Clipboard.GetText
    tmparray = Split(datastr, vbCrLf)

    
    num = 0
    For i = 0 To UBound(tmparray)
        pos = InStr(1, tmparray(i), ";")
        If pos = 0 Then
        
        ElseIf pos = 1 Then
            tmparray(i) = ""
        Else
            tmparray(i) = Mid(tmparray(i), 1, pos - 1)
        End If
        tmparray(i) = SubSpace(tmparray(i))
        If tmparray(i) <> "" Then
            tmpdata = Split(tmparray(i))
            num = num + UBound(tmpdata) + 1
        End If
    Next i
    
    KdefInfo(Index).numlabel = 0
    KdefInfo(Index).DataLong = num
    ReDim KdefInfo(Index).data(KdefInfo(Index).DataLong - 1)

    num = 0
    For i = 0 To UBound(tmparray)
        If tmparray(i) <> "" Then
            tmpdata = Split(tmparray(i))
            For j = 0 To UBound(tmpdata)
                KdefInfo(Index).data(num) = CLng(tmpdata(j))
                num = num + 1
            Next j
        End If
    Next i
    refresh_list

End Sub

Private Sub cmdSavefile_Click()
Dim time As Long
    If MsgBox(LoadResStr(128), vbOKCancel, Me.Caption) = vbOK Then
        'time = GetTickCount
        Call savetalk
'        MsgBox G_Var.KDEFGRP
        Call savekdef(G_Var.KDEFGRP)
        If GetINIStr("run", "Mode") = "pic" Then
            saveName
        End If
        'MsgBox GetTickCount - time
    End If
End Sub

Private Sub cmdup_Click()
Dim stat As Statement
Dim Index As Long
    Index = listkdef.ListIndex
    If Index = -1 Then Exit Sub
    If Index = 0 Then Exit Sub
     
     
    Set stat = KdefInfo(Combokdef.ListIndex).kdef(Index + 1)
    If Index >= 2 Then
        If stat.islabel = True And KdefInfo(Combokdef.ListIndex).kdef(Index - 1).gotoLabel = stat.note Then
            Exit Sub
        End If
        If KdefInfo(Combokdef.ListIndex).kdef(Index - 1).islabel = True And KdefInfo(Combokdef.ListIndex).kdef(Index - 1).note = stat.gotoLabel Then
            Exit Sub
        End If
    End If
    If Index < KdefInfo(Combokdef.ListIndex).kdef.Count - 1 Then
        If KdefInfo(Combokdef.ListIndex).kdef(Index + 2).islabel = True And KdefInfo(Combokdef.ListIndex).kdef(Index).gotoLabel = KdefInfo(Combokdef.ListIndex).kdef(Index + 2).note Then
            Exit Sub
        End If
        If KdefInfo(Combokdef.ListIndex).kdef(Index).islabel = True And KdefInfo(Combokdef.ListIndex).kdef(Index + 2).gotoLabel = KdefInfo(Combokdef.ListIndex).kdef(Index).note Then
            Exit Sub
        End If
    End If
    
    
    
    KdefInfo(Combokdef.ListIndex).kdef.Remove Index + 1
    KdefInfo(Combokdef.ListIndex).kdef.Add stat, , Index
     
    Call re_Analysis(Combokdef.ListIndex)
    cmdmodifykdef_Click
    
    
    listkdef.Clear
    For Each stat In KdefInfo(Combokdef.ListIndex).kdef
        listkdef.AddItem stat.note
    Next
    listkdef.ListIndex = Index - 1
    listkdef.SetFocus

End Sub

Private Sub cmdwizard_Click()
    ModifyStatementWizard True
End Sub

Private Sub Combokdef_click()
    Call refresh_list
End Sub

Private Sub Combotalk_Click()
    txttalk.Text = Talk(Combotalk.ListIndex)
End Sub



Private Sub Command1_Click()
txt2kdef (App.Path & "\1.txt")
End Sub

Private Sub Command2_Click()
Dim a As Integer
Dim tmpstr As String
Dim NewStr As String
Dim NewSubStr As String
Dim tmp() As String
Dim i As Long
Dim Width As Long
    a = Text2.Text * 2 - 1
    tmpstr = txttalk.Text
    tmp = Split(tmpstr, "*")
    tmpstr = ""
    For i = 0 To UBound(tmp)
        tmpstr = tmpstr & tmp(i)
    Next i
    NewStr = ""
    While Len(tmpstr) > 0
        NewSubStr = ""
        Width = 0
        Do
            If Len(tmpstr) = 0 Then
                Exit Do
            End If
            NewSubStr = NewSubStr & left(tmpstr, 1)
            If Abs(Asc(left(tmpstr, 1))) < 128 Then
                Width = Width + 1
            Else
                Width = Width + 2
            End If
            tmpstr = Mid(tmpstr, 2)
            
            If Width >= a Then
                Exit Do
            End If
        Loop
        
        If Len(tmpstr) > 0 Then
            If Width = a Then
                If Abs(Asc((left(tmpstr, 1)))) < 128 Then
                    NewSubStr = NewSubStr & left(tmpstr, 1)
                    tmpstr = Mid(tmpstr, 2)
                End If
            End If
        End If
        If Len(tmpstr) > 0 Then
            NewStr = NewStr & NewSubStr & "*"
        Else
            NewStr = NewStr & NewSubStr
        End If
    Wend
    txttalk.Text = NewStr
End Sub




Private Sub Command4_Click()
On Error Resume Next
Text1 = Text1 * 2
txtword.Caption = Text1
Text1.SetFocus
End Sub

'Private Sub Command5_Click()
'faq.Show vbNormal
'End Sub

Private Sub define_Click()
 define.Show vbNormal

End Sub

Private Sub Command6_Click()
define.Show vbNormal
End Sub

Private Sub Command7_Click()
Judgelabel
End Sub

Private Sub Form_Load()
'iniFileName = App.Path & "\txt.ini"

'If Charset = "BIG5" Then
'Text3.Visible = True
'Else

'txtNote.Visible = True
'End If
'Command3.Caption = StrUnicode("导入")
Dim i As Long
Dim ctl As Control
    'mnu_War.Enabled = False
    'mnu_z_modify.Enabled = False
    For i = 0 To Me.Controls.Count - 1
        Call SetCaption(Me.Controls(i))
    Next i
    'On Error Resume Next
    If readtalkidx(G_Var.JYPath & G_Var.nameidx, G_Var.JYPath & G_Var.Namegrp) = False Then
        readtalkidx G_Var.JYPath & G_Var.TalkIdx, G_Var.JYPath & G_Var.TalkGRP
    Else
        readtalkidx G_Var.JYPath & G_Var.nameidx, G_Var.JYPath & G_Var.Namegrp
    End If
    listlabel = 0
    LoadFormRes
    App.title = Me.Caption
    Call ReadStatementAttrib
    'MsgBox 1
    readKdeffile
    'MsgBox 2
    c_Skinner.AttachSkin Me.hWnd
End Sub

Private Sub refresh_list()
Dim stat As Statement
Dim Index As Long
    Index = listkdef.ListIndex
    listkdef.Clear
    Frame3.Caption = "正在编辑事件-No." & Val(Combokdef.ListIndex)
    Call DatatoKdef(Combokdef.ListIndex)
    For Each stat In KdefInfo(Combokdef.ListIndex).kdef
        listkdef.AddItem stat.note
    Next
    
End Sub


Private Sub refresh_combotalk()
Dim i As Long
Dim Index As Long
    Index = Combotalk.ListIndex

    Combotalk.Clear
    For i = 0 To numtalk - 1
        Combotalk.AddItem i & ":" & Talk(i)
    Next i
    If Index > Combotalk.ListCount - 1 Then
        Combotalk.ListIndex = Combotalk.ListCount - 1
    Else
        Combotalk.ListIndex = Index
    End If
End Sub



Private Sub listkdef_Click()
Dim tmp As String
Dim i As Long
Dim Index As Long
Dim Length As Long
Dim kk As Statement
Dim flag As Boolean
'Flag = False
    Liststatement.Clear
    Index = listkdef.ListIndex
    Set kk = KdefInfo(Combokdef.ListIndex).kdef.Item(Index + 1)
    
    If kk.islabel = True Then
        'MsgBox LoadResStr(132)
        tmp = Me.listkdef.Text
        listlabel = 1
        For i = 0 To Me.listkdef.ListCount - 1
            Set kk = KdefInfo(Combokdef.ListIndex).kdef.Item(i + 1)
            If StrComp(kk.gotoLabel, tmp) = 0 Then
                Me.listkdef.Selected(i) = True
                Exit For
            End If
        Next i
        listkdef.ListIndex = Index
        'Flag = True
    Else
        'If Flag = False Then
            If kk.gotoLabel <> "" Then
                Judgelabel
            End If
            listlabel = 0
            Liststatement.Clear
            Liststatement.AddItem kk.id

            Length = kk.DataNum
            'MsgBox Flag
            For i = 0 To Length - 1
                Liststatement.AddItem kk.data(i)
            Next i
        'End If
        'MsgBox 1
    End If
End Sub

Private Sub listkdef_DblClick()
        ModifyStatementWizard True
End Sub

Private Sub listkdef_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 2 Then   '   检查是否单击了鼠标右键。
      PopupMenu MDIMain.mnu_KdefPopup    '   把文件菜单显示为一个弹出式菜单。
   End If

End Sub

Private Sub Liststatement_Click()
    txtword.Caption = Liststatement.Text
 
End Sub

Private Sub mnu_MMapEdit_Click()
    frmMMapEdit.Show vbModal
End Sub
Public Sub mnu_kdef_txt2grp_Click()
Dim ofn As OPENFILENAME
Dim Rtn As String
Dim filenum As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "txt文件(*.txt)" + Chr$(0) + "*.txt" + Chr$(0)
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = App.Path
    ofn.lpstrTitle = "Open File"
    ofn.flags = 6148

    Rtn = GetOpenFileName(ofn)

    If Rtn < 1 Then Exit Sub
    
    txt2kdef (ofn.lpstrFile)
End Sub
Public Sub mnu_replacetalk_Click()
    If numkdef = 0 Then Exit Sub
    frmReplacetalk1.Show vbModal
    Combokdef_click
End Sub

Public Sub mnu_copy_Click()
Dim i As Long

    If listkdef.SelCount = 0 Then Exit Sub
    
    Set ClipboardStatement = Nothing
    Set ClipboardStatement = New Collection
    
    For i = 0 To listkdef.ListCount - 1
        If listkdef.Selected(i) = True Then
            ClipboardStatement.Add KdefInfo(Combokdef.ListIndex).kdef.Item(i + 1)
        End If
    Next i
    MDIMain.mnu_kdef_Paste.Enabled = True
End Sub

Public Sub mnu_copyall_Click()
Dim i As Long
If (MsgBox(StrUnicode2("你确定要复制所有事件?"), vbOKCancel) = vbOK) Then
    
    Set ClipboardKdef = Nothing
    Set ClipboardKdef = New Collection
    
    For i = 0 To listkdef.ListCount - 1
        ClipboardKdef.Add KdefInfo(Combokdef.ListIndex).kdef.Item(i + 1)
    Next i
    MDIMain.mnu_kdef_PasteALL.Enabled = True
End If
End Sub



Private Sub mnu_Exit_Click()
    Unload Me
End Sub

Private Sub mnu_Exportkdef_Click()
Dim filenum As Long
Dim exportfilename As String
Dim i As Long
Dim j As Long
Dim stat As Statement
Dim isShow As Boolean
    exportfilename = App.Path & "\kdefout.txt"
    If Dir(exportfilename) <> "" Then
        Kill exportfilename
    End If
    filenum = FreeFile()
    Open exportfilename For Output As #filenum
    For i = 0 To numkdef - 1
        KdefInfo(i).numlabel = 0
        Call DatatoKdef(i)
        Print #filenum, LoadResStr(135) & i & "(" & Hex(i) & ")"
        For Each stat In KdefInfo(i).kdef
            If stat.islabel = True Then
                Print #filenum, " " & LoadResStr(136); stat.note
            Else
                If stat.id = 68 Or stat.id = -1 Then
                    'Print #filenum, stat.id;
                    'For j = 0 To stat.DataNum - 1
                    '    Print #filenum, stat.data(j);
                    'Next j
                    'Print #filenum, " " & LoadResStr(136); stat.note
                    Print #filenum, stat.note
                End If
            End If
            DoEvents
        Next stat
        Print #filenum,
    Next i
    Close (filenum)
    MsgBox LoadResStr(137) & exportfilename, , LoadResStr(138)
End Sub

Private Sub mnu_exporttalkfile_Click()
Dim exportfilename As String
Dim filenum As Long
Dim i As Long
    
    exportfilename = App.Path & "\talk.txt"
    If Dir(exportfilename) <> "" Then
        Kill exportfilename
    End If
    filenum = FreeFile()
    Open exportfilename For Output As #filenum
    Write #filenum, numtalk, "Talk Number"
    For i = 0 To numtalk - 1
        Write #filenum, i, Hex(i), Talk(i)
    Next i
    Close (filenum)
    MsgBox LoadResStr(126) & exportfilename, , LoadResStr(127)
End Sub


Public Sub mnu_paste_Click()
Dim i As Long
Dim Index As Long
Dim stat As Statement
    Index = listkdef.ListIndex
    If Index = -1 Then Exit Sub
    ' paste
    For i = 1 To ClipboardStatement.Count
       KdefInfo(Combokdef.ListIndex).kdef.Add ClipboardStatement(i), , Index + i
    Next i
    
    Call re_Analysis(Combokdef.ListIndex)
    
    
    cmdmodifykdef_Click
    
    
    listkdef.Clear
    For Each stat In KdefInfo(Combokdef.ListIndex).kdef
        listkdef.AddItem stat.note
    Next
    listkdef.ListIndex = Index
    
    
    
End Sub

Public Sub mnu_pasteall_Click()
Dim i As Long
Dim Index As Long
Dim stat As Statement
    ' paste
    If MsgBox(LoadResStr(319), vbYesNo) = vbYes Then
        While KdefInfo(Combokdef.ListIndex).kdef.Count > 0
            KdefInfo(Combokdef.ListIndex).kdef.Remove 1
        Wend
        For i = 1 To ClipboardKdef.Count
           KdefInfo(Combokdef.ListIndex).kdef.Add ClipboardKdef(i)
        Next i
        
        Call re_Analysis(Combokdef.ListIndex)
        
        
        cmdmodifykdef_Click
        
        
        listkdef.Clear
        For Each stat In KdefInfo(Combokdef.ListIndex).kdef
            listkdef.AddItem stat.note
        Next

       
   End If
    
End Sub



Private Sub readKdeffile()
Dim i As Long
    Call readtalk
    refresh_combotalk
    'Combotalk.ListIndex = 0
    'mnu_Savefile.Enabled = True
    'mnu_exporttalkfile.Enabled = True
    'mnu_Exportkdef.Enabled = True
    'mnu_replacetalk.Enabled = True
    'mnu_War.Enabled = True
    'mnu_z_modify.Enabled = True
    'mnu_Map.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True
    'MsgBox G_Var.EditMode
    If G_Var.EditMode = "classic" Then
        Call LoadPicFile(G_Var.JYPath & G_Var.HeadIDX, G_Var.JYPath & G_Var.HeadGRP, HeadPic, Headnum)
        Call SetHeadtoPerson
    Else
        Call LoadKGPicFile(G_Var.JYPath & G_Var.NewHeadGRP)
    End If
    
    Call ReadKdef
    
    
    'Call SetHeadtoPerson

    'MsgBox numkdef
    Combokdef.Clear
    ReDim kdefs(numkdef)
    
    For i = 0 To numkdef - 1
        Set KdefInfo(i).kdef = New Collection
        Combokdef.AddItem i
'        Call Analysis(i)
    Next i
    Combokdef.ListIndex = 0
    
End Sub

Private Sub mnu_savekdeffile_Click()
    Call savekdef(G_Var.KDEFGRP)
'    MsgBox G_Var.KDEFGRP
End Sub



Private Sub mnu_setpath_Click()
    Getgamepath.Show vbModal
    Me.Caption = LoadResStr(101) & "-" & G_Var.JYPath
    'mnu_readfile.Enabled = True
End Sub



Private Sub ModifyStatementWizard(flag As Boolean)
Dim Index As Long
Dim kk As Statement
Dim num50 As Long
    Index = listkdef.ListIndex
    If Index < 0 Then Exit Sub
    Set kk = KdefInfo(Combokdef.ListIndex).kdef.Item(Index + 1)
    
    Select Case kk.id
    'case 0 太简单不需要
    
    Case &H1
        If GetINIStr("run", "Mode") = "grp" Then
            frmStatement_0x01.Show vbModal
            Call refresh_combotalk
        End If
    Case &H2
        frmstatement_0x02.Show vbModal
        
    Case &H3
        frmstatement_0x03.Show vbModal
        
    Case &H4
        frmStatement_0x04.Show vbModal
    'case 5 太简单不需要
        
    Case &H6
        frmStatement_0x06.Show vbModal
    'case 7 太简单不需要
    'case 8 太简单不需要
    'case 9 太简单不需要
    Case &HA
        frmStatement_0x0a.Show vbModal
    'case B 太简单不需要
    'case C 太简单不需要
    'case D 太简单不需要
    'case E 太简单不需要
    'case F 太简单不需要

    Case &H10
        frmStatement_0x10.Show vbModal
    
    Case &H11
        frmStatement_0x11.Show vbModal
    
    Case &H12
        frmStatement_0x12.Show vbModal
    
    Case &H13
        frmStatement_0x13.Show vbModal
        
    'case 14 太简单不需要
    
    Case &H15
        frmStatement_0x15.Show vbModal
    
    'case 16 太简单不需要
    
    Case &H17
        frmStatement_0x17.Show vbModal
    
    Case &H19
        frmstatement_0x19.Show vbModal
    
    Case &H1A
        frmStatement_0x1a.Show vbModal

    Case &H1B
        frmStatement_0x1B.Show vbModal
        
    Case &H1C
        frmStatement_0x1c.Show vbModal
        
    Case &H1D
        frmStatement_0x1d.Show vbModal
        
    Case &H1E
        frmStatement_0x1e.Show vbModal
        
    Case &H1F
        frmStatement_0x1f.Show vbModal
        
        
    Case &H20
        frmstatement_0x20.Show vbModal
    
    Case &H21
        frmStatement_0x21.Show vbModal
    
    Case &H22
        frmStatement_0x22.Show vbModal
    
    Case &H23
        frmStatement_0x23.Show vbModal
        
    Case &H24 '暂时需要
        frmStatement_0x24.Show vbModal
        'Judgelabel
    
    Case &H25
        frmStatement_0x25.Show vbModal
        
    Case &H26
        frmStatement_0x26.Show vbModal
        
        
    Case &H27
        frmStatement_0x27.Show vbModal
        
        
    Case &H28
        frmStatement_0x28.Show vbModal
        
    Case &H29
        frmStatement_0x29.Show vbModal
        
    'case 2a 太简单不需要
    
    Case &H2B
        
        frmStatement_0x12.Show vbModal
        
    Case &H2C
        frmStatement_0x2c.Show vbModal
    
    Case &H2D
        frmStatement_0x2d.Show vbModal
    
    
    Case &H2E
        frmStatement_0x2e.Show vbModal
    
    
    Case &H2F
        frmStatement_0x2f.Show vbModal
        
        
    Case &H30
        frmStatement_0x30.Show vbModal
        
    Case &H31
        frmStatement_0x31.Show vbModal
    
    Case &H32
    
        num50 = GetINILong("kdef50", "num")
    
        If kk.data(0) < num50 Then
            frmStatement_0x32.Show vbModal
        Else
            MsgBox LoadResStr(139) & kk.id & "(" & Hex(kk.id) & ")" & LoadResStr(141)
        End If
        
        
        
    ' case 33 太简单不需要
    ' case 34
    ' case 35
    
    ' case 36
        
    Case &H38
        FrmStatement_0x38.Show vbModal
    
    ' case 39 太简单不需要
    ' case 3a
    ' case 3b
    
    Case &H3C
        frmStatement_0x3c.Show vbModal
    
    ' case 3d
    Case &H3F
        frmstatement_0x3f.Show vbModal
    Case &H44
         frmStatement_0x44.Show vbModal
    Case &H45
         frmStatement_0x45.Show vbModal
    Case &H46
         frmStatement_0x46.Show vbModal
     Case &H47
         frmStatement_0x47.Show vbModal
    'case 40
    ' case 41
    ' case 42
    Case &H0, &H5, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &H14, &H16, &H39, &HFFFF
        If flag = True Then
            MsgBox LoadResStr(139) & kk.id & "(" & Hex(kk.id) & ")" & LoadResStr(140)
        End If
    Case Else
        If flag = True Then
            MsgBox LoadResStr(139) & kk.id & "(" & Hex(kk.id) & ")" & LoadResStr(141)
        End If
    End Select
    
    listkdef_Click
    
End Sub




Private Sub LoadFormRes()
    Me.Caption = LoadResStr(221)
    'mnu_file.Caption = LoadResStr(201)
    'mnu_help.Caption = LoadResStr(202)
    'mnu_setpath.Caption = LoadResStr(203)
    'mnu_readfile.Caption = LoadResStr(204)
    cmdsaveFile.Caption = LoadResStr(205)
    cmdExportTalk.Caption = LoadResStr(206)
    cmdExportKDef.Caption = LoadResStr(207)
    cmdExit.Caption = LoadResStr(208)
    'mnu_about.Caption = LoadResStr(209)
    'mnu_War.Caption = LoadResStr(210)
    Frame1.Caption = LoadResStr(301)
    cmdinserttalk.Caption = LoadResStr(302)
    cmdmodifytalk.Caption = LoadResStr(303)
    cmddeletetalk.Caption = LoadResStr(304)
    lbltalk.Caption = LoadResStr(305)
    Frame2.Caption = LoadResStr(306)
    Label3.Caption = LoadResStr(307)
    cmdinsertkdef.Caption = LoadResStr(308)
    cmddeletekdef.Caption = LoadResStr(309)
    cmdmodifykdef.Caption = LoadResStr(310)
    cmdinsertstatement.Caption = LoadResStr(311)
    cmddeletestatement.Caption = LoadResStr(312)
    cmdmodifystatement.Caption = LoadResStr(313)
    Label1.Caption = LoadResStr(314)
    Label2.Caption = LoadResStr(315)
    cmdwizard.Caption = LoadResStr(316)
    cmdmodifyword.Caption = LoadResStr(317)
    
    'mnu_replacetalk.Caption = LoadResStr(10031)
    
    'mnu_z_modify.Caption = LoadResStr(211)
    'mnu_z_modify.Caption = LoadResStr(211)
    'mnu_z_InitEdit.Caption = LoadResStr(218)
    'mnu_z_AddIDX.Caption = LoadResStr(219)

    
    'mnu_copy.Caption = LoadResStr(212)
    'mnu_copyAll.Caption = LoadResStr(213)
    'mnu_paste.Caption = LoadResStr(214)
    'mnu_pasteAll.Caption = LoadResStr(215)
    
    'mnu_Map.Caption = LoadResStr(216)
    'mnu_MMapEdit.Caption = LoadResStr(217)
    
    
End Sub

Private Sub mnu_ShowPic_Click()
    frmSelectMap.Show vbModal
End Sub

Private Sub mnu_War_Click()
    FrmWarEdit.Show vbModal
End Sub

Private Sub mnu_z_Edit_Click()
     '   frmzModify.Show vbModal
End Sub

Private Sub mnu_z_InitEdit_Click()
    frmInitEdit.Show vbModal
End Sub





' 读 talk 文件
Private Sub readtalk()
'Dim talkfilename As String
'Dim talkfileid As String
Dim filenum As Long

Dim testb() As Byte
Dim i As Long, j As Long

Dim Length As Long
Dim offset As Long
    
 '   talkfilename = G_Var.JYPath & G_Var.TalkGRP
 '   talkfileid = G_Var.JYPath & G_Var.TalkIDX
    
    filenum = OpenBin(G_Var.JYPath & G_Var.TalkIdx, "R")
        numtalk = LOF(filenum) / 4
        ReDim TalkIdx(numtalk)
        TalkIdx(0) = 0
        For i = 1 To numtalk
            Get filenum, , TalkIdx(i)
        Next i
    Close (filenum)
    ReDim Talk(numtalk)
    
    filenum = OpenBin(G_Var.JYPath & G_Var.TalkGRP, "R")
    
    For i = 0 To numtalk - 1
        Length = TalkIdx(i + 1) - TalkIdx(i)
        offset = TalkIdx(i) + 1
        ReDim testb(Length - 1)
        For j = 0 To Length - 1
            Get filenum, offset + j, testb(j)
            testb(j) = testb(j) ' Xor &HFF
        Next j
        Talk(i) = Big5toUnicode(testb, Length - 1)
    Next i
    Close (filenum)
    
    
End Sub

' 存 talk 文件
Private Sub savetalk()
'Dim talkfilename As String
'Dim talkfileid As String
Dim outputfile As String
Dim filenum As Long
Dim filenum2 As Long

Dim testb() As Byte
Dim i As Long, j As Long

Dim Length As Long
Dim offset As Long
Dim tempb As Byte
    
 '   talkfilename = G_Var.JYPath & "\" & GetINIStr("File", "TalkGrpFilename")
 '   talkfileid = G_Var.JYPath & "\" & GetINIStr("File", "TalkIdxFIlename")
 '   Kill talkfilename
 '   Kill talkfileid
    filenum = OpenBin(G_Var.JYPath & G_Var.TalkIdx, "WN")

    'Open talkfileid For Binary Access Write As #filenum
    
    filenum2 = OpenBin(G_Var.JYPath & G_Var.TalkGRP, "WN")
    TalkIdx(0) = 0
    tempb = 0
    For i = 0 To numtalk - 1
        Call UnicodetoBIG5(Talk(i), Length, testb)
        TalkIdx(i + 1) = TalkIdx(i) + Length + 1
        For j = 0 To Length - 1
            Put #filenum2, , CByte(testb(j) Xor &HFF)
        Next j
        Put #filenum2, , tempb
        Put #filenum, , TalkIdx(i + 1)
    Next i
    Close (filenum)
    Close (filenum2)


End Sub




'  设置根据头像查询人物

Private Sub SetHeadtoPerson()
Dim i As Long
On Error GoTo Label1:
    ReDim HeadtoPerson(Headnum - 1)
    For i = 0 To Headnum - 1
         Set HeadtoPerson(i) = New Collection
    Next i
    For i = 0 To PersonNum - 1
        HeadtoPerson(Person(i).PhotoId).Add i
    Next i
    Exit Sub
Label1:
    
End Sub

Private Sub Text1_Change()
Dim a As Integer
a = 0
If Text1.Text <> "-" Then
  If IsNumeric(Text1) Then
    If Text1.Text > 32767 Then
        Text1.Text = 32767
        MsgBox "超出", vbOKOnly, "警告"
    End If
    If Text1.Text < -32767 Then
        Text1.Text = -32767
        MsgBox "超出", vbOKOnly, "警告"
    End If
  Else
    Text1.Text = ""
    Text1.SetFocus
 End If
End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'如果键码＝13就是回车键，就会发生这种事：
If KeyCode = 13 Then
    If Text1.Text = "" Then
        MsgBox "错误", vbOKOnly, "警告"
        Exit Sub
Else
End If
  Dim Index As Long
    Index = Liststatement.ListIndex
    If Index >= 0 Then
        Liststatement.List(Liststatement.ListIndex) = Text1.Text
    End If
    txtword.Caption = Text1.Text
    Text1.Text = ""
    Liststatement.SetFocus
End If
'ESC
If KeyCode = 27 Then
   Text1.Text = ""
   Liststatement.SetFocus
End If

End Sub
' 读 name 文件
Private Sub readname()
'Dim talkfilename As String
'Dim talkfileid As String
Dim filenum As Long

Dim testb() As Byte
Dim i As Long, j As Long

Dim Length As Long
Dim offset As Long
    
 '   talkfilename = G_Var.JYPath & G_Var.TalkGRP
 '   talkfileid = G_Var.JYPath & G_Var.TalkIDX
    
    filenum = OpenBin(G_Var.JYPath & G_Var.nameidx, "R")
        numname = LOF(filenum) / 4
        ReDim nameidx(numname)
        nameidx(0) = 0
        For i = 1 To numname
            Get filenum, , nameidx(i)
        Next i
    Close (filenum)
    ReDim Name(numname)
    
    filenum = OpenBin(G_Var.JYPath & G_Var.Namegrp, "R")
    
    For i = 0 To numname - 1
        Length = nameidx(i + 1) - nameidx(i)
        offset = nameidx(i) + 1
        ReDim testb(Length - 1)
        For j = 0 To Length - 1
            Get filenum, offset + j, testb(j)
            testb(j) = testb(j) Xor &HFF
        Next j
        Talk(i) = Big5toUnicode(testb, Length - 1)
    Next i
    Close (filenum)
    
    
End Sub


Private Sub Text2_Change()
If Text2.Text = "" Then
  MsgBox "数据不能为空", vbOKOnly, "警告"
  Exit Sub
End If
If IsNumeric(Text2) Then
Else
  Text2.Text = 12
  MsgBox "数据错误", vbOKOnly, "警告"
End If
End Sub
Private Sub Liststatement_KeyDown(KeyCode As Integer, Shift As Integer)
'如果键码＝46就是>键，就会发生这种事：
If KeyCode = 68 Then
 Text1.SetFocus
End If
'如果键码＝44就是<键，就会发生这种事：
If KeyCode = 65 Then
    listkdef.SetFocus
End If
'w
If KeyCode = 87 Then
   Dim Index As Long
Dim tmpstate As Statement
Dim stat As Statement
Dim i As Long
    If listlabel = 1 Then Exit Sub
    Index = listkdef.ListIndex
    If Index = -1 Then Exit Sub
    Set tmpstate = KdefInfo(Combokdef.ListIndex).kdef(Index + 1)
    tmpstate.id = Liststatement.List(0)
    For i = 0 To tmpstate.DataNum - 1
        tmpstate.data(i) = Liststatement.List(i + 1)
    Next i
    
    Call re_Analysis(Combokdef.ListIndex)
    
    
    listkdef.Clear
    
    
    For Each stat In KdefInfo(Combokdef.ListIndex).kdef
        listkdef.AddItem stat.note
    Next

    listkdef.ListIndex = Index
    listkdef.SetFocus
Dim ii As Long

    ii = Combokdef.ListIndex
    If ii < 0 Then Exit Sub
    
    Call modifykdef(ii)
End If
End Sub
Private Sub listkdef_KeyDown(KeyCode As Integer, Shift As Integer)
'如果键码＝46就是>键，就会发生这种事：
If KeyCode = 68 Then
 Liststatement.SetFocus
End If
'www
If KeyCode = 87 Then
    cmdmodifystatement_Click
    cmdmodifykdef_Click
End If
End Sub

' 读 name 文件
Private Function readtalkidx(idxname As String, grpname As String) As Boolean
'Dim talkfilename As String
'Dim talkfileid As String
Dim filenum As Long
Dim testb() As Byte
Dim i As Long, j As Long

Dim Length As Long
Dim offset As Long
'MsgBox numname
 '   talkfilename = G_Var.JYPath & G_Var.TalkGRP
 '   talkfileid = G_Var.JYPath & G_Var.TalkIDX
 'G_Var.JYPath & G_Var.nameidx,G_Var.JYPath & G_Var.Namegrp
 '   MsgBox G_Var.JYPath & G_Var.Nameidx
 On Error GoTo errorNum:
    filenum = OpenBin(idxname, "R")
        numname = LOF(filenum) / 4
        ReDim nameidx(numname)
        nameidx(0) = 0
        For i = 1 To numname
            Get filenum, , nameidx(i)
        Next i
    Close (filenum)
    ReDim nam(numname)
    
    filenum = OpenBin(grpname, "R")
    
        For i = 0 To numname - 1
        Length = nameidx(i + 1) - nameidx(i)
        offset = nameidx(i) + 1
        ReDim testb(Length - 1)
        For j = 0 To Length - 1
            Get filenum, offset + j, testb(j)
            testb(j) = testb(j) Xor &HFF
        Next j
        nam(i) = Big5toUnicode(testb, Length - 1)
    Next i
    Close (filenum)
    readtalkidx = True
    Exit Function
errorNum:
    readtalkidx = False
End Function
Private Sub Judgelabel()
Dim kk
Dim strtmp As String
Dim i As Long
Dim Index As Long
    Index = Me.listkdef.ListIndex
    kk = Split(Me.listkdef.Text, ":")
    strtmp = ":" & kk(UBound(kk))
    For i = 0 To Me.listkdef.ListCount - 1
        If StrComp(Me.listkdef.List(i), strtmp) = 0 Then
            Me.listkdef.Selected(i) = True
            Exit For
        End If
    Next i
    Me.listkdef.ListIndex = Index
End Sub
Public Sub txt2kdef(ByVal txtname As String)
Dim i As Long
Dim j As Long
Dim num As Long, Index As Long
Dim datatmp() As String
Dim DataLong As Long
Dim strtmp() As String, tmp() As String
Dim filenum As Long
Dim flag As Boolean
Dim battletmp() As String
Dim tmpstr As String
    DataLong = 0
    filenum = FreeFile()
    Open txtname For Input As filenum
        ReDim datatmp(DataLong)
        Do While Not EOF(filenum)
            Line Input #filenum, tmpstr
            If tmpstr <> "" And Mid(tmpstr, 1, 1) <> "*" And Mid(tmpstr, 1, 1) <> "＊" Then
                DataLong = DataLong + 1
                ReDim Preserve datatmp(DataLong)
                tmpstr = Replace(tmpstr, ":", "：")
                tmpstr = Replace(tmpstr, "(", "（")
                tmpstr = Replace(tmpstr, ")", "）")
                datatmp(DataLong - 1) = tmpstr
            End If
        Loop
    Close filenum
num = 0
    For i = 0 To UBound(datatmp) - 1
        strtmp = Split(datatmp(i), "：")
        If strtmp(0) = "战斗" Then
            tmp = Split(strtmp(1), "（")
            battletmp = Split(tmp(UBound(tmp) - 0), "）")
            If StrComp(battletmp(0), "胜利") = 0 Or StrComp(battletmp(0), "失败") = 0 Then
                num = num + 10 + 1
            Else
                num = num + 7 + 1
            End If
        ElseIf strtmp(0) = "黑屏" Then
            num = num + 2 + 1
        Else
            num = num + 8 + 1
        End If
    Next i
    
    Index = Combokdef.ListIndex
    If Index < 0 Then Exit Sub
    KdefInfo(Index).numlabel = 0
    KdefInfo(Index).DataLong = num + 1
    ReDim KdefInfo(Index).data(KdefInfo(Index).DataLong - 1)
      
 num = 0
    For i = 0 To UBound(datatmp) - 1
        strtmp = Split(datatmp(i), "：")
        Select Case strtmp(0)
        Case "战斗"
            
            KdefInfo(Index).data(num) = 6
            tmp = Split(strtmp(1), "（")
            On Error Resume Next
            KdefInfo(Index).data(num + 1) = getWarNum(tmp(0))
            battletmp = Split(tmp(UBound(tmp) - 0), "）")
            If StrComp(battletmp(0), "胜利") = 0 Then
                KdefInfo(Index).data(num + 2) = 4
                KdefInfo(Index).data(num + 3) = 0
                KdefInfo(Index).data(num + 4) = 0
                KdefInfo(Index).data(num + 5) = 0
                
                KdefInfo(Index).data(num + 6) = 15
                KdefInfo(Index).data(num + 7) = 0
                KdefInfo(Index).data(num + 8) = -1
                KdefInfo(Index).data(num + 9) = 0
                KdefInfo(Index).data(num + 10) = 13
                num = num + 10 + 1
            ElseIf StrComp(battletmp(0), "失败") = 0 Then
                KdefInfo(Index).data(num + 2) = 0
                KdefInfo(Index).data(num + 3) = 4
                KdefInfo(Index).data(num + 4) = 0
                KdefInfo(Index).data(num + 5) = 0
                
                KdefInfo(Index).data(num + 6) = 15
                KdefInfo(Index).data(num + 7) = 0
                KdefInfo(Index).data(num + 8) = -1
                KdefInfo(Index).data(num + 9) = 0
                KdefInfo(Index).data(num + 10) = 13
                num = num + 10 + 1
            Else
                KdefInfo(Index).data(num + 2) = 1
                KdefInfo(Index).data(num + 3) = 0
                KdefInfo(Index).data(num + 4) = 0
                KdefInfo(Index).data(num + 5) = 0
                KdefInfo(Index).data(num + 6) = 0
                KdefInfo(Index).data(num + 7) = 13
                num = num + 7 + 1
            End If

            
        Case "黑屏"
            KdefInfo(Index).data(num) = 14
            KdefInfo(Index).data(num + 1) = 0
            KdefInfo(Index).data(num + 2) = 13
            num = num + 2 + 1
        Case Else
            If G_Var.EditMode = "classic" Then
                KdefInfo(Index).data(num) = 1
                numtalk = numtalk + 1
                ReDim Preserve Talk(numtalk - 1)
                ReDim Preserve TalkIdx(numtalk)
                Talk(numtalk - 1) = strtmp(1)
                KdefInfo(Index).data(num + 1) = numtalk - 1
                flag = False
                    For j = 0 To PersonNum - 1
                        If strtmp(0) = Person(i).Name1 Then
                            KdefInfo(Index).data(num + 2) = Person(i).PhotoId
                            flag = True
                            Exit For
                        End If
                    Next j
                
                    If flag = False Then
                        On Error Resume Next
                        KdefInfo(Index).data(num + 2) = GetINILong("Person_txt", strtmp(0))
                    End If
                KdefInfo(Index).data(num + 3) = 0
                KdefInfo(Index).data(num + 4) = 0
                num = num + 4 + 1
            Else
                KdefInfo(Index).data(num + 1) = GetHeadNum(strtmp(0))
                KdefInfo(Index).data(num) = 68
                numtalk = numtalk + 1
                ReDim Preserve Talk(numtalk - 1)
                ReDim Preserve TalkIdx(numtalk)
                
                Talk(numtalk - 1) = strtmp(1)
                Combotalk.AddItem numtalk - 1 & ":" & Talk(numtalk - 1)
                flag = False
                For j = 0 To (numname - 1)
                    If nam(j) = strtmp(0) Then
                        KdefInfo(Index).data(num + 3) = j
                        flag = True
                    End If
                Next j
                If flag = False Then
                    numname = numname + 1
                    ReDim Preserve nam(numtalk - 1)
                    ReDim Preserve nameidx(numtalk)
                    nam(numname - 1) = strtmp(0)
                    KdefInfo(Index).data(num + 3) = numname - 1
                End If
                KdefInfo(Index).data(num + 2) = numtalk - 1
                KdefInfo(Index).data(num + 4) = 0
                KdefInfo(Index).data(num + 5) = 0
                KdefInfo(Index).data(num + 6) = 28515
                KdefInfo(Index).data(num + 7) = 0
                KdefInfo(Index).data(num + 8) = 0
                num = num + 8 + 1
            End If
        End Select
        KdefInfo(Index).data(num) = -1
    Next i
    refresh_list
End Sub
Public Function GetHeadNum(ByVal SurName As String)
Dim j As Long
    For j = 0 To PersonNum - 1
        If SurName = Person(j).Name1 Then
            GetHeadNum = Person(j).PhotoId
            Exit Function
            Exit For
        End If
    Next j
    
    If GetINILong("Person_txt", SurName) = -65536 Then
        frmheadnum.Label1.Caption = "错误头像:" & SurName & ",请重新选择"
            frmheadnum.Show vbModal
            If frmheadnum.OK = 1 Then
                GetHeadNum = frmheadnum.NHeadnum
                Call PutINIStr("Person_txt", SurName, frmheadnum.NHeadnum)
            End If
        Else
            GetHeadNum = GetINILong("Person_txt", SurName)
    End If
End Function
Public Function getWarNum(ByVal WarName As String)
    If GetINILong("war_txt", WarName) = -65536 Then
        frmwarNum.Label1.Caption = "错误战斗:" & WarName & ",请重新选择"
        frmwarNum.Label2.Caption = WarName
            frmwarNum.Show vbModal
            If frmwarNum.OK = 1 Then
                getWarNum = frmwarNum.KWarNum
                Call PutINIStr("war_txt", WarName, frmwarNum.KWarNum)
            End If
        Else
            getWarNum = GetINILong("war_txt", WarName)
    End If
End Function

Public Sub saveName()
'Dim talkfilename As String
'Dim talkfileid As String
Dim outputfile As String
Dim filenum As Long
Dim filenum2 As Long

Dim testb() As Byte
Dim i As Long, j As Long

'Dim nameidx() As Long
Dim Length As Long
Dim offset As Long
Dim tempb As Byte
    
 '   talkfilename = G_Var.JYPath & "\" & GetINIStr("File", "TalkGrpFilename")
 '   talkfileid = G_Var.JYPath & "\" & GetINIStr("File", "TalkIdxFIlename")
 '   Kill talkfilename
 '   Kill talkfileid
    filenum = OpenBin(G_Var.JYPath & G_Var.nameidx, "WN")

    'Open talkfileid For Binary Access Write As #filenum
    
    filenum2 = OpenBin(G_Var.JYPath & G_Var.Namegrp, "WN")
    nameidx(0) = 0
    tempb = 0
    For i = 0 To numname - 1
        'MsgBox nam(i)
        Call UnicodetoBIG5(nam(i), Length, testb)
        nameidx(i + 1) = nameidx(i) + Length + 1
        For j = 0 To Length - 1
            Put #filenum2, , CByte(testb(j) Xor &HFF)
        Next j
        Put #filenum2, , tempb
        Put #filenum, , nameidx(i + 1)
    Next i
    Close (filenum)
    Close (filenum2)


End Sub
