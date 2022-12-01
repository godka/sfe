VERSION 5.00
Begin VB.Form frmSMapEdit 
   ClientHeight    =   9210
   ClientLeft      =   7185
   ClientTop       =   2445
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSMapEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   614
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   Begin VB.ListBox ListLevel 
      Height          =   1845
      ItemData        =   "frmSMapEdit.frx":406A
      Left            =   120
      List            =   "frmSMapEdit.frx":4071
      Style           =   1  'Checkbox
      TabIndex        =   60
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Frame FrameD_Event 
      Caption         =   "修改场景事件"
      Height          =   4815
      Left            =   1800
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      Begin VB.CommandButton X2 
         Caption         =   "X2"
         Height          =   375
         Left            =   1200
         TabIndex        =   58
         Top             =   4320
         Width           =   495
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   51
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdModifyD 
         Caption         =   "确认修改"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   10
         Left            =   840
         TabIndex        =   47
         Text            =   "0"
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   9
         Left            =   840
         TabIndex        =   45
         Text            =   "0"
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   8
         Left            =   840
         TabIndex        =   43
         Text            =   "0"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   7
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   41
         Text            =   "0"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   6
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   39
         Text            =   "0"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   5
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   37
         Text            =   "0"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   35
         Text            =   "0"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   33
         Text            =   "0"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   31
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtD 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   28
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "纵坐标Y"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   49
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "横坐标X"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   48
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "动画延迟"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "动画延迟帧数"
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "开始贴图"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   44
         ToolTipText     =   "同动画开始贴图编号"
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "结束贴图"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "动画结束贴图编号*2"
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "开始贴图"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "动画开始贴图编号*2"
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "事件3"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "通过触发事件"
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "事件2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "物品触发事件"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "事件1"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "空格触发事件"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "编号"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "能否通过"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "0能通过，1不能通过"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9240
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "当前图片"
      Height          =   5415
      Left            =   0
      TabIndex        =   10
      Top             =   3720
      Width           =   1815
      Begin VB.PictureBox PicEvent 
         AutoRedraw      =   -1  'True
         Height          =   1335
         Left            =   600
         ScaleHeight     =   1275
         ScaleWidth      =   1035
         TabIndex        =   26
         Top             =   3960
         Width           =   1095
      End
      Begin VB.PictureBox PicEarth 
         AutoRedraw      =   -1  'True
         Height          =   615
         Left            =   600
         ScaleHeight     =   555
         ScaleWidth      =   1035
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.PictureBox PicBiuld 
         AutoRedraw      =   -1  'True
         Height          =   1815
         Left            =   600
         ScaleHeight     =   1755
         ScaleWidth      =   1035
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.PictureBox PicAir 
         AutoRedraw      =   -1  'True
         Height          =   855
         Left            =   600
         ScaleHeight     =   795
         ScaleWidth      =   1035
         TabIndex        =   11
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblEventValue 
         Caption         =   "-1"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblEvent 
         Caption         =   "场景事件"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "1地面"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "2建筑"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "3空中"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lbl1 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lbl2 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbl3 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "海拔"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lbl6 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lbl5 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "海拔"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   615
      End
   End
   Begin VB.ComboBox ComboScene 
      Height          =   345
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   0
      Width           =   1695
   End
   Begin VB.HScrollBar HScrollWidth 
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   7200
      Width           =   7455
   End
   Begin VB.VScrollBar VScrollHeight 
      Height          =   7335
      Left            =   9240
      Max             =   479
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.ComboBox ComboLevel 
      Height          =   345
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox PicBak 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdSelectMap 
      Caption         =   "选择贴图"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   1800
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.ListBox OutList 
         Height          =   735
         Left            =   0
         TabIndex        =   59
         Top             =   4800
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Timer RT 
         Interval        =   10
         Left            =   6960
         Top             =   0
      End
      Begin VB.Frame frame2 
         Caption         =   "批量增加海拔"
         Height          =   1455
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
         Begin VB.CommandButton Command1 
            Caption         =   "隐藏"
            Height          =   375
            Left            =   360
            TabIndex        =   57
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   330
            Left            =   840
            TabIndex        =   55
            Text            =   "1"
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "改为"
            Height          =   255
            Left            =   840
            TabIndex        =   54
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "增加"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Label8"
            Height          =   375
            Left            =   120
            TabIndex        =   56
            Top             =   600
            Width           =   615
         End
      End
   End
   Begin VB.Label Label7 
      Caption         =   "能否通过"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   30
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblMenu 
      Caption         =   "<快捷菜单>"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "操作层"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblSelectPicNum 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "frmSMapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private SelectPicNum As Long

Private Const SMapXmax = 64
Private Const SMapYmax = 64

Private Rtime As Long

Private SMapPic() As RLEPic
Private SMappicnum As Long

Private KGPic() As RLEPic
Private KGpicnum As Long

Private ss As Long           ' 当前场景代号
Private st As Long           '复制前的场景号

Private SinData0() As Integer
Private SinData1() As Integer

'retry
Private Const RetryNum = 20
Private RetrySin() As Integer

Private SceneIDX() As Long
Private SceneMapNum As Long      ' 场景地图个数，注意与r*中的数据不同

Private xx As Long
Private yy As Long

Private MouseX As Long
Private MouseY As Long

Private BlockX1 As Long, BlockY1 As Long     ' 选择块位置
Private BlockX2 As Long, BlockY2 As Long
Private SelectBlock As Long                  ' 0 未选择块，1 选择块

Private iMode As Long                      ' 0 正常   1 块操作  2 删除

Private isGrid As Long                       ' 0 不显示网格 1 显示网格
Private isShowLevel As Long                  ' 0 全部显示   1 只显示操作层
Private isScene As Long                      ' 0 不显示     1 显示场景

'Private Recordnum As Long                   ' 进度编号


Private D_Event0() As D_Event_type
Private D_IDX() As Long
Private D_mapnum As Long

Private SceneOffset() As Long
Private SceneOffsetNum As Long

Private Current_D_EventNum As Long
Private Current_D_Event_Pic As Long
Public Sub LoadScenePic(filename As String)
Dim idnum As Integer
Dim PersonNum As Long
Dim filenum As Long
Dim i As Long
'Dim idx() As Integer
    filenum = OpenBin(filename, "R")
        Get filenum, , SceneOffsetNum
        ReDim SceneOffset(SceneOffsetNum)
        SceneOffset(0) = SceneOffsetNum * 4 + 4
        For i = 1 To SceneOffsetNum
            Get filenum, , SceneOffset(i)
        Next i
    Close (filenum)

'MsgBox NewHeadNum
End Sub
Private Sub cmdModifyD_Click()
Dim numD As Long
Dim x As Long, y As Long
    
    numD = FrameD_Event.Tag
    x = txtD(9).Tag
    y = txtD(10).Tag
    
    D_Event0(numD, ss).isGo = txtD(0).Text
    D_Event0(numD, ss).id = txtD(1).Text
    D_Event0(numD, ss).EventNum1 = txtD(2).Text
    D_Event0(numD, ss).EventNum2 = txtD(3).Text
    D_Event0(numD, ss).EventNum3 = txtD(4).Text
    D_Event0(numD, ss).picnum(0) = txtD(5).Text
    D_Event0(numD, ss).picnum(1) = txtD(6).Text
    D_Event0(numD, ss).picnum(2) = txtD(7).Text
    D_Event0(numD, ss).PicDelay = txtD(8).Text
    D_Event0(numD, ss).x = txtD(9).Text
    D_Event0(numD, ss).y = txtD(10).Text
                
    SinData0(x, y, 3, ss) = -1
    SinData0(D_Event0(numD, ss).x, D_Event0(numD, ss).y, 3, ss) = numD
    cmdModifyD.Enabled = False
    showsmap
End Sub

Private Sub cmdSelectMap_Click()
    SelectPicNum = -1
    Load frmSelectMap
    frmSelectMap.txtIDX = G_Var.SMAPIDX
    frmSelectMap.txtGRP = G_Var.SMAPGRP
    frmSelectMap.cmdshow_Click
    frmSelectMap.Show
End Sub



Private Sub showsmap()
    pic1.Cls
    Draw_Smap
    Draw_smap_2
End Sub


Private Sub cmdSelectMap2_Click()
    SelectPicNum = -1
    Load frmSelectMap
    frmSelectMap.txtIDX = G_Var.SMAPIDX2
    frmSelectMap.txtGRP = G_Var.SMAPGRP2
    frmSelectMap.cmdshow_Click
    frmSelectMap.Show
End Sub

Private Sub ComboLevel_click()
    If iMode = 1 And (ComboLevel.ListIndex = 5 Or ComboLevel.ListIndex = 6) Then
        Frame2.Visible = True
    Else
        Frame2.Visible = False
    End If
    If ComboLevel.ListIndex = 4 Then
        FrameD_Event.Visible = True
    Else
        FrameD_Event.Visible = False
    End If
    If ComboLevel.ListIndex <> 7 Then
        OutList.Visible = False
    End If
    Set_Note
    showsmap
End Sub


Private Sub ComboScene_click()
    ss = ComboScene.ListIndex
    showsmap
End Sub

Private Sub Command1_Click()
Frame2.Visible = False
End Sub
Private Sub SaveKGSmap()
Dim Smaptmp(SMapXmax - 1, SMapYmax - 1, 6 - 1) As Integer

Dim filenum As Long
Dim i As Long, j As Long, k As Long
Dim ll As Integer
Dim kuang As OPENFILENAME
Dim filename As String
    kuang.lStructSize = Len(kuang)
    kuang.hwndOwner = Me.hwnd
    kuang.hInstance = App.hInstance
    kuang.lpstrFile = Space(254)
    kuang.nMaxFile = 255
    kuang.lpstrFileTitle = Space(254)
    kuang.nMaxFileTitle = 255
    kuang.lpstrInitialDir = App.Path
    kuang.flags = 6148
    '过虑对话框文件类型
    kuang.lpstrFilter = "map文件(*.map)" + Chr$(0) + "*.map" + Chr$(0)
    '对话框标题栏文字
    kuang.lpstrTitle = "保存文件的路径及文件名..."
    ll = GetSaveFileName(kuang) '显示保存文件对话框
    If ll >= 1 Then '取得对话中用户选择输入的文件名及路径
        filename = kuang.lpstrFile
        filename = Left(filename, InStr(filename, Chr(0)) - 1)
    End If
    If Len(filename) = 0 Then Exit Sub
    
    For k = 0 To 6 - 1
        For i = 0 To SMapXmax - 1
            For j = 0 To SMapYmax - 1
                Smaptmp(i, j, k) = SinData0(i, j, k, ss)
            Next j
        Next i
    Next k
        For i = 0 To SMapXmax - 1
            For j = 0 To SMapYmax - 1
                Smaptmp(i, j, 3) = -1
            Next j
        Next i
    filename = filename & ".map"
    filenum = OpenBin(filename, "WN")
        Put filenum, , Smaptmp
    Close (filenum)
    MsgBox LoadResString(10916) & filename
End Sub
Private Sub LoadKGSmap()
Dim Smaptmp(SMapXmax - 1, SMapYmax - 1, 6 - 1) As Integer
Dim ofn As OPENFILENAME
Dim Rtn As String
Dim tmpStr As String
Dim filenum As Long
Dim i As Long, j As Long, k As Long
    tmpStr = "map文件|*.map"
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = Replace$(tmpStr, "|", Chr$(0)) + vbNullChar + vbNullChar
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = App.Path
    ofn.lpstrTitle = "Open File"
    ofn.flags = 6148

    Rtn = GetOpenFileName(ofn)

    If Rtn < 1 Then Exit Sub
    
    filenum = OpenBin(ofn.lpstrFile, "R")
        Get filenum, , Smaptmp
    Close (filenum)
    For k = 0 To 6 - 1
        For i = 0 To SMapXmax - 1
            For j = 0 To SMapYmax - 1
                SinData0(i, j, k, ss) = Smaptmp(i, j, k)
            Next j
        Next i
    Next k
        For i = 0 To SMapXmax - 1
            For j = 0 To SMapYmax - 1
                SinData0(i, j, 3, ss) = -1
            Next j
        Next i
    showsmap
End Sub

Private Sub Form_Load()
Dim filenum As Long
Dim i As Long
Dim fileid As String
Dim filepic As String
    'LoadScenePic (G_Var.JYPath & G_Var.SceneMap)
    Rtime = 0
    Me.Caption = LoadResStr(226)
    If Option1.Value = True Then
       Label8.Caption = StrUnicode2("增加")
    Else
       Label8.Caption = StrUnicode2("改为")
    End If
    For i = 0 To Me.Controls.Count - 1
        Call SetCaption(Me.Controls(i))
        If TypeOf Me.Controls(i) Is ComboBox Then
            'Call SetComboWidth(oo.Controls(i), 270)
            Call SetComboHeight(Me.Controls(i), 270)
        End If
    Next i
    If GetINIStr("run", "style") = "DOS" Then
        MDIMain.mnu_SMAPMenu_LoadMap4.Enabled = False
        MDIMain.mnu_SMAPMenu_LoadMap5.Enabled = False
        MDIMain.mnu_SMAPMenu_LoadMap6.Enabled = False
    End If
    Recordnum = 1
    
    isGrid = 0
    
    Call LoadPicFile(G_Var.JYPath & G_Var.SMAPIDX, G_Var.JYPath & G_Var.SMAPGRP, SMapPic, SMappicnum)
    'Call LoadPngPicFile(G_Var.JYPath & G_Var.SceneMap, KGPic, KGpicnum)
    Load_DS
    
    
    
    ComboLevel.Clear
    ComboLevel.AddItem LoadResStr(10805)
    ComboLevel.AddItem LoadResStr(10806)
    ComboLevel.AddItem LoadResStr(10807)
    ComboLevel.AddItem LoadResStr(10808)
    ComboLevel.AddItem LoadResStr(10809)
    ComboLevel.AddItem LoadResStr(10810)
    ComboLevel.AddItem LoadResStr(10811)
    ComboLevel.AddItem LoadResStr(10917)
    ComboLevel.ListIndex = 0

    ListLevel.Clear
    ListLevel.AddItem LoadResStr(10805)
    ListLevel.AddItem LoadResStr(10806)
    ListLevel.AddItem LoadResStr(10807)
    ListLevel.AddItem LoadResStr(10808)
    ListLevel.Selected(0) = True
    'ListLevel.ListIndex = 0
    
    OutList.AddItem StrUnicode2("出口1")
    OutList.AddItem StrUnicode2("出口2")
    OutList.AddItem StrUnicode2("出口3")
    
    VScrollHeight.Max = SMapXmax - 1
    VScrollHeight.LargeChange = 5
    VScrollHeight.SmallChange = 1
    VScrollHeight.Value = SMapXmax / 2
    
    HScrollWidth.Max = SMapYmax - 1
    HScrollWidth.LargeChange = 5
    HScrollWidth.SmallChange = 1
    HScrollWidth.Value = SMapXmax / 2
    
    
    Current_D_EventNum = -1
    c_Skinner.AttachSkin Me.hwnd
    Timer1.Enabled = True
End Sub

' 读d*s*
Private Sub Load_DS()
Dim filenum As Long
Dim i As Long
Dim Sfilelength As Currency

    Call ReadRR(Recordnum)
    Select Case GetINIStr("run", "style")
        Case "DOS"
            filenum = OpenBin(G_Var.JYPath & G_Var.SIDX(Recordnum), "R")
                SceneMapNum = LOF(filenum) / 4
            Close filenum
            
            filenum = OpenBin(G_Var.JYPath & G_Var.DIDX(Recordnum), "R")
                D_mapnum = LOF(filenum) / 4
            Close filenum
        Case "kys"
            Sfilelength = 4# * 3 * SMapXmax * SMapYmax
            SceneMapNum = FileLen(G_Var.JYPath & G_Var.SGRP(Recordnum)) / (Sfilelength) '64*64*4
            D_mapnum = FileLen(G_Var.JYPath & G_Var.DGRP(Recordnum)) / (200 * 11 * 2)
    End Select
    If D_mapnum <> SceneMapNum Then
        Err.Raise vbObjectError + 1, , "frmSMapEdit error: D number not equal S number"
    End If
    
    ReDim SceneIDX(SceneMapNum)
    ReDim SinData0(SMapXmax - 1, SMapYmax - 1, 5, SceneMapNum - 1)
    
    ReDim SinData1(SMapXmax - 1, SMapYmax - 1, 5, RetryNum, SceneMapNum - 1)
    ReDim RetrySin(Scenenum)
    
    ReDim D_IDX(SceneMapNum)
    ReDim D_Event0(200 - 1, SceneMapNum - 1)
    
    If GetINIStr("run", "style") = "DOS" Then
        filenum = OpenBin(G_Var.JYPath & G_Var.SIDX(Recordnum), "R")
        For i = 0 To SceneMapNum - 1
            Get filenum, , SceneIDX(i + 1)
        Next i
        Close #filenum
    End If
    
    SceneIDX(0) = 0
    
    filenum = OpenBin(G_Var.JYPath & G_Var.SGRP(Recordnum), "R")
    Get #filenum, , SinData0
    Close (filenum)
    
    If GetINIStr("run", "style") = "DOS" Then
        filenum = OpenBin(G_Var.JYPath & G_Var.DIDX(Recordnum), "R")
        ReDim D_IDX(SceneMapNum)
        For i = 0 To SceneMapNum - 1
            Get filenum, , D_IDX(i + 1)
        Next i
        Close #filenum
    End If
    D_IDX(0) = 0
    
    filenum = OpenBin(G_Var.JYPath & G_Var.DGRP(Recordnum), "R")
    
    Get filenum, , D_Event0
    
    Close (filenum)

    ComboScene.Clear
    For i = 0 To SceneMapNum - 1
        If i < Scenenum Then
            ComboScene.AddItem i & Big5toUnicode(Scene(i).Name1, 10)
        Else
            ComboScene.AddItem i
        End If
    Next i
    ComboScene.ListIndex = 0

End Sub

' 写D*s*
Private Sub Save_DS()
Dim filenum As Long
Dim i As Long
    
    If GetINIStr("run", "style") = "DOS" Then
        filenum = OpenBin(G_Var.JYPath & G_Var.SIDX(Recordnum), "WN")
        For i = 1 To SceneMapNum
            Put filenum, , SceneIDX(i)
        Next i
        Close (filenum)
    End If
    
    filenum = OpenBin(G_Var.JYPath & G_Var.SGRP(Recordnum), "WN")
    
    Put filenum, , SinData0
    
    Close (filenum)
    
    If GetINIStr("run", "style") = "DOS" Then
        filenum = OpenBin(G_Var.JYPath & G_Var.DIDX(Recordnum), "WN")
        For i = 1 To SceneMapNum
            Put filenum, , D_IDX(i)
        Next i
        Close (filenum)
    End If
    filenum = OpenBin(G_Var.JYPath & G_Var.DGRP(Recordnum), "WN")
    
    Put filenum, , D_Event0
    
    Close (filenum)
    
    SaveRinS (Recordnum)
End Sub



' 绘场景地图

Public Sub Draw_Smap()
Dim RangeX As Long, rangeY As Long
Dim i As Long, j As Long, k As Long
Dim i1 As Long, j1 As Long
Dim X1 As Long, Y1 As Long
Dim picnum As Long
    
Dim temp As Long
Dim lineSize As Long
Dim dx1 As Long, dx2 As Long
Dim dib As New clsDIB

    On Error Resume Next
    dib.CreateDIB pic1.Width, pic1.Height
    
    RangeX = 18 + 15
    rangeY = 10 + 15
    
     For j = -rangeY To 2 * RangeX + rangeY
        For i = RangeX To 0 Step -1
           
            If j Mod 2 = 0 Then
                i1 = -RangeX + i + j \ 2
                j1 = -i + j \ 2
            Else
                i1 = -RangeX + i + j \ 2
                j1 = -i + j \ 2 + 1
            End If
            X1 = XSCALE * (i1 - j1) + pic1.Width / 2
            Y1 = YSCALE * (i1 + j1) + pic1.Height / 2
            
            If yy + j1 >= 0 And xx + i1 >= 0 And yy + j1 < SMapYmax And xx + i1 < SMapXmax Then
                dx1 = SinData0(xx + i1, yy + j1, 4, ss)
                dx2 = SinData0(xx + i1, yy + j1, 5, ss)
                '2013年3月4日0:19:56
                    picnum = SinData0(xx + i1, yy + j1, 0, ss) / 2
                
                    If picnum > 0 And picnum < SMappicnum Then
                        If ListLevel.Selected(1) Then
                            Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y)
                        End If
                    End If
                    
                    picnum = SinData0(xx + i1, yy + j1, 1, ss) / 2
                    If picnum > 0 And picnum < SMappicnum Then
                        If ListLevel.Selected(2) Then
                                Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y - dx1)
                        End If
                    End If
                    picnum = SinData0(xx + i1, yy + j1, 2, ss) / 2
                    If picnum > 0 And picnum < SMappicnum Then
                        If ListLevel.Selected(3) Then
                                Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y - dx2)
                        End If
                    End If
                
                picnum = SinData0(xx + i1, yy + j1, 3, ss)
                If picnum >= 0 Then
                    picnum = D_Event0(picnum, ss).picnum(0) / 2
                End If
                
                If picnum > 0 And picnum < SMappicnum Then
                    If ListLevel.Selected(4) Then
                        Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y - dx1)
                    End If
                End If
            End If
        Next i
    Next j
    
    
    
    picbak.Cls
     
        ' 复制到dib上
    temp = BitBlt(picbak.hDC, 0, 0, pic1.Width, pic1.Height, dib.CompDC, 0, 0, &HCC0020)
   
    picbak.ForeColor = &H808000
    
   
      For j = -rangeY To 2 * RangeX + rangeY
       For i = RangeX To 0 Step -1
           
            If j Mod 2 = 0 Then
                i1 = -RangeX + i + j \ 2
                j1 = -i + j \ 2
            Else
                i1 = -RangeX + i + j \ 2
                j1 = -i + j \ 2 + 1
            End If
            X1 = XSCALE * (i1 - j1) + pic1.Width / 2
            Y1 = YSCALE * (i1 + j1) + pic1.Height / 2
            If yy + j1 >= 0 And xx + i1 >= 0 And yy + j1 < SMapYmax And xx + i1 < SMapXmax Then
                If isGrid = 1 Then
                      picbak.Line (X1, Y1)-(X1 + XSCALE, Y1 - YSCALE)
                      picbak.Line (X1, Y1)-(X1 - XSCALE, Y1 - YSCALE)
                End If
            End If
        Next i
    Next j
    
    picbak.ForeColor = vbYellow
    picbak.FontSize = 9
   
      For j = -rangeY To 2 * RangeX + rangeY
       For i = RangeX To 0 Step -1
           
            If j Mod 2 = 0 Then
                i1 = -RangeX + i + j \ 2
                j1 = -i + j \ 2
            Else
                i1 = -RangeX + i + j \ 2
                j1 = -i + j \ 2 + 1
            End If
            X1 = XSCALE * (i1 - j1) + pic1.Width / 2
            Y1 = YSCALE * (i1 + j1) + pic1.Height / 2
           
           
           If yy + j1 >= 0 And xx + i1 >= 0 And yy + j1 < SMapYmax And xx + i1 < SMapXmax Then
               
                If SinData0(xx + i1, yy + j1, 3, ss) >= 0 Then
                    dx1 = SinData0(xx + i1, yy + j1, 4, ss)
                     picbak.CurrentX = X1 - XSCALE / 2
                     picbak.CurrentY = Y1 - YSCALE - 4 - dx1
                    
                     picbak.Print "[" & SinData0(xx + i1, yy + j1, 3, ss) & "]"
                End If
                
                If (xx + i1 = Scene(ss).InX And yy + j1 = Scene(ss).InY) Then
                    picbak.ForeColor = vbRed
                    picbak.CurrentX = X1 - XSCALE / 2
                    picbak.CurrentY = Y1 - YSCALE - 4
                    picbak.Print "(" & xx + i1 & "," & yy + j1 & ")" & StrUnicode2("入口")
                    picbak.ForeColor = vbYellow
                End If
                
                For k = 0 To 2
                    If (xx + i1 = Scene(ss).OutX(k) And yy + j1 = Scene(ss).OutY(k)) Then
                        picbak.ForeColor = vbRed
                        picbak.CurrentX = X1 - XSCALE / 2
                        picbak.CurrentY = Y1 - YSCALE - 4
                        picbak.Print "(" & xx + i1 & "," & yy + j1 & ")" & StrUnicode2("出口") & k + 1
                        picbak.ForeColor = vbYellow
                    End If
                Next k
            End If
        Next i
    Next j
End Sub


Public Sub Draw_smap_2()
Dim RangeX As Long, rangeY As Long
Dim i As Long, j As Long
Dim i1 As Long, j1 As Long
Dim X1 As Long, Y1 As Long
Dim picnum As Long
Dim ListIndex As Long

Dim temp As Long
Dim dx As Long
Dim dib As New clsDIB

    dib.CreateDIB pic1.Width, pic1.Height
    
    temp = BitBlt(dib.CompDC, 0, 0, pic1.Width, pic1.Height, picbak.hDC, 0, 0, &HCC0020)
    
    RangeX = 18 + 15
    rangeY = 10 + 15
    
    
    i1 = MouseX - xx
    j1 = MouseY - yy
    
    X1 = XSCALE * (i1 - j1) + pic1.Width / 2
    Y1 = YSCALE * (i1 + j1) + pic1.Height / 2
    picnum = SelectPicNum
    
        If picnum >= 0 And picnum < SMappicnum And iMode <> 2 Then
            If yy + j1 >= 0 And xx + i1 >= 0 And yy + j1 < SMapYmax And xx + i1 < SMapXmax Then
                Select Case ComboLevel.ListIndex
                Case 2
                    dx = SinData0(xx + i1, yy + j1, 4, ss)
                Case 3
                    dx = SinData0(xx + i1, yy + j1, 5, ss)
                Case 4
                    dx = SinData0(xx + i1, yy + j1, 4, ss)
                    dx = 0
                    picnum = 0
                Case 5
                    dx = SinData0(xx + i1, yy + j1, 4, ss)
                    dx = 0
                    picnum = 0
                Case 6
                    dx = SinData0(xx + i1, yy + j1, 5, ss)
                    dx = 0
                    picnum = 0
                End Select
               
            End If
            
            If iMode = 2 Then
                picnum = 0
            End If

            Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y - dx)
       End If
    
     If iMode = 1 And SelectBlock = 0 Then
     'If SelectBlock = 0 Then
       If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
           pic1.ForeColor = vbRed
           For j = -rangeY To 2 * RangeX + rangeY
                For i = RangeX To 0 Step -1
                 
                 If j Mod 2 = 0 Then
                     i1 = -RangeX + i + j \ 2
                     j1 = -i + j \ 2
                 Else
                     i1 = -RangeX + i + j \ 2
                     j1 = -i + j \ 2 + 1
                 End If
                 
                X1 = XSCALE * (i1 - j1) + pic1.Width / 2
                Y1 = YSCALE * (i1 + j1) + pic1.Height / 2
                 
                If i1 + xx >= MouseX - (BlockX2 - BlockX1) And i1 + xx <= MouseX And _
                   j1 + yy >= MouseY - (BlockY2 - BlockY1) And j1 + yy <= MouseY Then
                    'For ListIndex = 1 To ListLevel.ListCount - 1
                    '    If ListLevel.Selected(ListIndex) Then
                    '        picnum = SinData0(BlockX2 - MouseX + i1 + xx, BlockY2 - MouseY + j1 + yy, ListIndex - 1, st) / 2
                    '        If picnum > 0 And picnum < SMappicnum Then
                    '            Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y)
                    '        End If
                    '    End If
                    'Next ListIndex
                    Select Case ComboLevel.ListIndex
                    Case 0
                        'copy 1,2,3
                        picnum = SinData0(BlockX2 - MouseX + i1 + xx, BlockY2 - MouseY + j1 + yy, 0, st) / 2
                        If picnum > 0 And picnum < SMappicnum Then
                            Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y)
                        End If
                    
                        picnum = SinData0(BlockX2 - MouseX + i1 + xx, BlockY2 - MouseY + j1 + yy, 1, st) / 2
                        If picnum > 0 And picnum < SMappicnum Then
                            Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y)
                        End If
                        
                        picnum = SinData0(BlockX2 - MouseX + i1 + xx, BlockY2 - MouseY + j1 + yy, 2, ss) / 2
                        If picnum > 0 And picnum < SMappicnum Then
                            Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y)
                        End If
                    
                    Case 1
                        picnum = SinData0(BlockX2 - MouseX + i1 + xx, BlockY2 - MouseY + j1 + yy, 0, st) / 2
                        If picnum > 0 And picnum < SMappicnum Then
                            Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y)
                        End If
                    Case 2
                        picnum = SinData0(BlockX2 - MouseX + i1 + xx, BlockY2 - MouseY + j1 + yy, 1, st) / 2
                        If picnum > 0 And picnum < SMappicnum Then
                            Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y)
                        End If
                    Case 3
                        picnum = SinData0(BlockX2 - MouseX + i1 + xx, BlockY2 - MouseY + j1 + yy, 2, ss) / 2
                        If picnum > 0 And picnum < SMappicnum Then
                            Call genPicData(SMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - SMapPic(picnum).x, Y1 - SMapPic(picnum).y)
                        End If
                    Case 4
                    Case 5
                        
                    End Select
                End If
               Next i
         Next j
      End If
    End If
     
     
     pic1.Cls
        ' 复制到dib上
    temp = BitBlt(pic1.hDC, 0, 0, pic1.Width, pic1.Height, dib.CompDC, 0, 0, &HCC0020)
   
   
   If iMode = 1 Or (iMode = 2 And (ComboLevel.ListIndex < 4 And ComboLevel.ListIndex > 0)) Then
    'MsgBox 1
    If SelectBlock = 1 And (ComboLevel.ListIndex = 0 Or ComboLevel.ListIndex = 1 Or ComboLevel.ListIndex = 2 Or ComboLevel.ListIndex = 3 Or ComboLevel.ListIndex = 5 Or ComboLevel.ListIndex = 6) Then
       If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
'           Randomize
           pic1.ForeColor = vbRed
           For j = -rangeY To 2 * RangeX + rangeY
                For i = RangeX To 0 Step -1
                 
                 If j Mod 2 = 0 Then
                     i1 = -RangeX + i + j \ 2
                     j1 = -i + j \ 2
                 Else
                     i1 = -RangeX + i + j \ 2
                     j1 = -i + j \ 2 + 1
                 End If
                 X1 = XSCALE * (i1 - j1) + pic1.Width / 2
                 Y1 = YSCALE * (i1 + j1) + pic1.Height / 2
                 
                 
                If i1 + xx >= Min_V(BlockX1, BlockX2) And i1 + xx <= Max_V(BlockX1, BlockX2) And _
                   j1 + yy >= Min_V(BlockY1, BlockY2) And j1 + yy <= Max_V(BlockY1, BlockY2) Then
                    pic1.Line (X1, Y1)-(X1 + XSCALE, Y1 - YSCALE)
                    pic1.Line (X1, Y1)-(X1 - XSCALE, Y1 - YSCALE)
                    pic1.Line (X1, Y1 - 2 * YSCALE)-(X1 - XSCALE, Y1 - YSCALE)
                    pic1.Line (X1, Y1 - 2 * YSCALE)-(X1 + XSCALE, Y1 - YSCALE)
                End If
               Next i
         Next j
      End If
     End If
    End If
    MDIMain.StatusBar1.Panels(2).Text = " X=" & MouseX & ",Y=" & MouseY

End Sub


Public Sub Show_picture(pic As PictureBox, ByVal num As Long)
   
Dim temp As Long
Dim dib As New clsDIB
    
    dib.CreateDIB pic.Width, pic.Height
    pic.BackColor = MaskColor
    
    temp = BitBlt(dib.CompDC, 0, 0, pic.Width, pic.Height, pic.hDC, 0, 0, &HCC0020)
    
    'Picnum = num
    If num >= 0 Then
        Call genPicData(SMapPic(num), dib.addr, pic.Width, pic.Height, 0, 0)
    End If
        ' 复制到dib上
    temp = BitBlt(pic.hDC, 0, 0, pic.Width, pic.Height, dib.CompDC, 0, 0, &HCC0020)
   
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    If Me.ScaleWidth < 400 Then
        Me.Width = Me.ScaleX(400, vbPixels, vbTwips)
    End If
    pic1.Width = Me.ScaleWidth - VScrollHeight.Width - pic1.Left
    If pic1.Width Mod 2 = 1 Then          ' 宽度保持2的倍数
        pic1.Width = pic1.Width + 1
    End If
    HScrollWidth.Width = pic1.Width
    VScrollHeight.Left = pic1.Width + pic1.Left
    
    If Me.ScaleHeight < 400 Then
          Me.Height = Me.ScaleY(400, vbPixels, vbTwips)
    End If
    pic1.Height = Me.ScaleHeight - HScrollWidth.Height - pic1.Top
    If pic1.Height Mod 2 = 1 Then
        pic1.Height = pic1.Height + 1
    End If
    VScrollHeight.Height = pic1.Height
    HScrollWidth.Top = pic1.Top + pic1.Height
    picbak.Width = pic1.Width
    picbak.Height = pic1.Height
    'Call Game_Mmap_Build
    showsmap
      
End Sub

Private Sub Form_Unload(cancel As Integer)
    MDIMain.StatusBar1.Panels(1).Text = ""
    MDIMain.StatusBar1.Panels(2).Text = ""
    
End Sub

Private Sub HScrollWidth_Change()
    ScrollValue
    showsmap
End Sub

Private Sub HScrollWidth_Scroll()
    ScrollValue
End Sub

Private Sub lblMenu_Click()
    PopupMenu MDIMain.mnu_SMAPMenu
End Sub

Private Sub ListLevel_Click()
Dim i As Long
If ListLevel.ListIndex = 0 Then
    For i = 1 To ListLevel.ListCount - 1
        If ListLevel.Selected(0) = True Then
            ListLevel.Selected(i) = True
        Else
            ListLevel.Selected(i) = False
        End If
    Next
End If
showsmap
End Sub

Private Sub Option1_Click()
Label8.Caption = "增加"
End Sub

Private Sub Option2_Click()
Label8.Caption = "改为"
End Sub
Private Sub OutList_Click()
Dim Index As Long
If MouseX >= 0 And MouseX < SMapXmax And MouseY >= 0 And MouseY < SMapYmax Then
    Index = OutList.ListIndex
    Scene(ss).OutX(Index) = MouseX
    Scene(ss).OutY(Index) = MouseY
    OutList.Visible = False
End If
showsmap
End Sub

Private Sub pic1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyS And (Shift And vbAltMask) Then
    RetryStack
ElseIf KeyCode = vbKeyZ And (Shift And vbCtrlMask) Then
    Dim Smaptmp() As Integer
    ReDim Smaptmp(SMapXmax - 1, SMapYmax - 1, 6 - 1, RetryNum, Scenenum)
    If RetrySin(ss) > 0 Then
        CopyMemory SinData0(0, 0, 0, ss), SinData1(0, 0, 0, 0, ss), 6# * 2 * SMapXmax * SMapXmax
        'CopyMemory smaptmp(0, 0, 0, 0, ss), SinData1(0, 0, 0, 1, ss), 6# * 2 * SMapXmax * SMapXmax * (RetryNum - 1)
        CopyMemory SinData1(0, 0, 0, 0, ss), SinData1(0, 0, 0, 1, ss), 6# * 2 * SMapXmax * SMapXmax * (RetryNum - 1)
        RetrySin(ss) = RetrySin(ss) - 1
        showsmap
    End If
    Debug.Print ss; RetrySin(ss)
End If
End Sub
Private Sub RetryStack()
    '代码 不安全了，直接硬来吧。。
    CopyMemory SinData1(0, 0, 0, 1, ss), SinData1(0, 0, 0, 0, ss), 6# * 2 * SMapXmax * SMapXmax * (RetryNum - 1)
   '保证代码安全，还是用手动拷贝，不用块复制了
    CopyMemory SinData1(0, 0, 0, 0, ss), SinData0(0, 0, 0, ss), 6# * 2 * SMapXmax * SMapXmax
    RetrySin(ss) = RetrySin(ss) + 1
    Debug.Print ss; RetrySin(ss)
End Sub
Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long, k As Long
Dim tmplong As Long
If MouseX >= 0 And MouseX < SMapXmax And MouseY >= 0 And MouseY < SMapYmax Then
'    MsgBox SMappicnum
    If Button = vbLeftButton Then   ' 左键按下，拾取。
        Select Case ComboLevel.ListIndex
        Case 0
            If iMode = 1 Then
                BlockX1 = MouseX
                BlockY1 = MouseY
                BlockX2 = -1
                BlockY2 = -1
                SelectBlock = 1
                st = ss
            End If
        Case 1
            SelectPicNum = SinData0(MouseX, MouseY, 0, ss) / 2
            If SelectPicNum > SMappicnum Then
                MsgBox "PicNum Overflow", vbCritical
                Exit Sub
            End If
            If iMode = 1 Then
                BlockX1 = MouseX
                BlockY1 = MouseY
                BlockX2 = -1
                BlockY2 = -1
                SelectBlock = 1
                st = ss
            ElseIf iMode = 2 Then
                BlockX1 = MouseX
                BlockY1 = MouseY
                BlockX2 = -1
                BlockY2 = -1
                SelectBlock = 1
            End If

        Case 2
            SelectPicNum = SinData0(MouseX, MouseY, 1, ss) / 2
            If SelectPicNum > SMappicnum Then
                MsgBox "PicNum Overflow", vbCritical
                Exit Sub
            End If
            If iMode = 1 Then
                BlockX1 = MouseX
                BlockY1 = MouseY
                BlockX2 = -1
                BlockY2 = -1
                SelectBlock = 1
                st = ss
            ElseIf iMode = 2 Then
                BlockX1 = MouseX
                BlockY1 = MouseY
                BlockX2 = -1
                BlockY2 = -1
                SelectBlock = 1
            End If
        Case 3
            SelectPicNum = SinData0(MouseX, MouseY, 2, ss) / 2
            If SelectPicNum > SMappicnum Then
                MsgBox "PicNum Overflow", vbCritical
                Exit Sub
            End If
            If iMode = 1 Then
                BlockX1 = MouseX
                BlockY1 = MouseY
                BlockX2 = -1
                BlockY2 = -1
                SelectBlock = 1
            ElseIf iMode = 2 Then
                BlockX1 = MouseX
                BlockY1 = MouseY
                BlockX2 = -1
                BlockY2 = -1
                SelectBlock = 1
            End If
        Case 4
            If SinData0(MouseX, MouseY, 3, ss) >= 0 And iMode <> 2 Then
                FrameD_Event.Tag = SinData0(MouseX, MouseY, 3, ss)
                txtD(0).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).isGo
                txtD(1).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).id
                txtD(2).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).EventNum1
                txtD(3).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).EventNum2
                txtD(4).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).EventNum3
                txtD(5).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).picnum(0)
                txtD(6).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).picnum(1)
                txtD(7).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).picnum(2)
                txtD(8).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).PicDelay
                txtD(9).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).x
                txtD(10).Text = D_Event0(SinData0(MouseX, MouseY, 3, ss), ss).y
                txtD(9).Tag = MouseX
                txtD(10).Tag = MouseY
                cmdModifyD.Enabled = True
            End If
        Case 5
            If iMode <> 1 Then
                SinData0(MouseX, MouseY, 4, ss) = SinData0(MouseX, MouseY, 4, ss) + 1
            Else
                If SelectBlock = 1 Then
                    If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
                        If MouseX >= BlockX1 And MouseX <= BlockX2 And MouseY >= BlockY1 And MouseY <= BlockY2 Then
                            For i = BlockX1 To BlockX2
                                For j = BlockY1 To BlockY2
                                    'If SinData0(i, j, 1, ss) > 0 Then
                                     If Option1.Value = True Then
                                        SinData0(i, j, 4, ss) = SinData0(i, j, 4, ss) + Text1.Text
                                          ElseIf Option2.Value = True Then
                                        SinData0(i, j, 4, ss) = Text1.Text
                                     End If
                                    'End If
                                Next j
                            Next i
                        Else
                            SelectBlock = 0
                            BlockX1 = -1
                            BlockY1 = -1
                            BlockX2 = -1
                            BlockY2 = -1
                        End If
                    Else
                        SelectBlock = 0
                        BlockX1 = -1
                        BlockY1 = -1
                        BlockX2 = -1
                        BlockY2 = -1
                    End If
                Else
                    BlockX1 = MouseX
                    BlockY1 = MouseY
                    BlockX2 = -1
                    BlockY2 = -1
                    SelectBlock = 1
                End If
            End If
        Case 6
            If iMode <> 1 Then
                SinData0(MouseX, MouseY, 5, ss) = SinData0(MouseX, MouseY, 5, ss) + 1
            Else
                If SelectBlock = 1 Then
                    If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
                        If MouseX >= BlockX1 And MouseX <= BlockX2 And MouseY >= BlockY1 And MouseY <= BlockY2 Then
                            For i = BlockX1 To BlockX2
                                For j = BlockY1 To BlockY2
                                    If Option1.Value = True Then
                                        SinData0(i, j, 5, ss) = SinData0(i, j, 5, ss) + Val(Text1.Text)
                                    ElseIf Option2.Value = True Then
                                        SinData0(i, j, 5, ss) = Val(Text1.Text)
                                    End If
                                Next j
                            Next i
                        Else
                            SelectBlock = 0
                            BlockX1 = -1
                            BlockY1 = -1
                            BlockX2 = -1
                            BlockY2 = -1
                        End If
                    Else
                        SelectBlock = 0
                        BlockX1 = -1
                        BlockY1 = -1
                        BlockX2 = -1
                        BlockY2 = -1
                    End If
                Else
                    BlockX1 = MouseX
                    BlockY1 = MouseY
                    BlockX2 = -1
                    BlockY2 = -1
                    SelectBlock = 1
                End If
            End If
        Case 7
            Scene(ss).InX = MouseX
            Scene(ss).InY = MouseY
        End Select
        lblSelectPicNum.Caption = SelectPicNum
        

        
    ElseIf Button = vbRightButton Then
        RetryStack
        Select Case iMode
        Case 0
            Select Case ComboLevel.ListIndex
            Case 0
                
            Case 1
                SinData0(MouseX, MouseY, 0, ss) = SelectPicNum * 2
            Case 2
                SinData0(MouseX, MouseY, 1, ss) = SelectPicNum * 2
            Case 3
                SinData0(MouseX, MouseY, 2, ss) = SelectPicNum * 2
            Case 4
                If SinData0(MouseX, MouseY, 3, ss) < 0 Then
                    For i = 0 To 200 - 1
                        If D_Event0(i, ss).isGo = 0 And D_Event0(i, ss).id = 0 And D_Event0(i, ss).EventNum1 = 0 And D_Event0(i, ss).EventNum2 = 0 And D_Event0(i, ss).EventNum3 = 0 And _
                             D_Event0(i, ss).picnum(0) = 0 And D_Event0(i, ss).picnum(1) = 0 And D_Event0(i, ss).picnum(2) = 0 And D_Event0(i, ss).PicDelay = 0 And D_Event0(i, ss).x = 0 And _
                             D_Event0(i, ss).y = 0 Then
                           SinData0(MouseX, MouseY, 3, ss) = i
                           D_Event0(i, ss).id = i
                           D_Event0(i, ss).x = MouseX
                           D_Event0(i, ss).y = MouseY
                           Exit For
                        End If
                    Next i
                End If
            Case 5
                SinData0(MouseX, MouseY, 4, ss) = SinData0(MouseX, MouseY, 4, ss) - 1
            Case 6
                SinData0(MouseX, MouseY, 5, ss) = SinData0(MouseX, MouseY, 5, ss) - 1
            Case 7
                OutList.Left = x
                OutList.Top = y - OutList.Height
                OutList.Visible = True
            End Select
        Case 1
            Select Case ComboLevel.ListIndex
            Case 0
                    If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
                        For i = BlockX1 To BlockX2
                            For j = BlockY1 To BlockY2
                                If MouseX - BlockX2 + i >= 0 And MouseX - BlockX2 + i < SMapXmax And MouseY - BlockY2 + j >= 0 And MouseY - BlockY2 + j < SMapYmax Then
                                    For k = 0 To 5
                                        SinData0(MouseX - BlockX2 + i, MouseY - BlockY2 + j, k, ss) = SinData0(i, j, k, st)
                                    Next k
                                    SinData0(MouseX - BlockX2 + i, MouseY - BlockY2 + j, 3, ss) = -1
                                End If
                            Next j
                        Next i
                    End If
            Case 1
                    If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
                        For i = BlockX1 To BlockX2
                            For j = BlockY1 To BlockY2
                                If MouseX - BlockX2 + i >= 0 And MouseX - BlockX2 + i < SMapXmax And MouseY - BlockY2 + j >= 0 And MouseY - BlockY2 + j < SMapYmax Then
                                    If SinData0(i, j, 0, st) > 0 Then
                                        SinData0(MouseX - BlockX2 + i, MouseY - BlockY2 + j, 0, ss) = SinData0(i, j, 0, st)
                                    End If
                                End If
                            Next j
                        Next i
                    End If
            Case 2
                    If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
                        For i = BlockX1 To BlockX2
                            For j = BlockY1 To BlockY2
                                If MouseX - BlockX2 + i >= 0 And MouseX - BlockX2 + i < SMapXmax And MouseY - BlockY2 + j >= 0 And MouseY - BlockY2 + j < SMapYmax Then
                                    'SinData0(MouseX - BlockX2 + i, MouseY - BlockY2 + j, 1, ss) = 0
                                    'SinData0(MouseX - BlockX2 + i, MouseY - BlockY2 + j, 4, ss) = 0
                                    'If SinData0(i, j, 1, st) > 0 Then
                                        SinData0(MouseX - BlockX2 + i, MouseY - BlockY2 + j, 1, ss) = SinData0(i, j, 1, st)
                                        SinData0(MouseX - BlockX2 + i, MouseY - BlockY2 + j, 4, ss) = SinData0(i, j, 4, st)
                                    'End If
                                End If
                            Next j
                        Next i
                    End If
            Case 3
                    If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
                        For i = BlockX1 To BlockX2
                            For j = BlockY1 To BlockY2
                                If MouseX - BlockX2 + i >= 0 And MouseX - BlockX2 + i < SMapXmax And MouseY - BlockY2 + j >= 0 And MouseY - BlockY2 + j < SMapYmax Then
                                    If SinData0(i, j, 0, st) > 0 Then
                                        SinData0(MouseX - BlockX2 + i, MouseY - BlockY2 + j, 2, ss) = SinData0(i, j, 2, ss)
                                    End If
                                End If
                            Next j
                        Next i
                    End If
            Case 5
                If SelectBlock = 1 Then
                    If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
                        If MouseX >= BlockX1 And MouseX <= BlockX2 And MouseY >= BlockY1 And MouseY <= BlockY2 Then
                            For i = BlockX1 To BlockX2
                                For j = BlockY1 To BlockY2
                                    'If SinData0(i, j, 1, ss) > 0 Then
                                        SinData0(i, j, 4, ss) = SinData0(i, j, 4, ss) - 1
                                    'End If
                                Next j
                            Next i
                        Else
                            SelectBlock = 0
                            BlockX1 = -1
                            BlockY1 = -1
                            BlockX2 = -1
                            BlockY2 = -1
                        End If
                    Else
                        SelectBlock = 0
                        BlockX1 = -1
                        BlockY1 = -1
                        BlockX2 = -1
                        BlockY2 = -1
                    End If
                End If
            End Select
        Case 2
        
            Select Case ComboLevel.ListIndex
            Case 0
            
            Case 1
               SinData0(MouseX, MouseY, 0, ss) = 0
            Case 2
                SinData0(MouseX, MouseY, 1, ss) = 0
            Case 3
                SinData0(MouseX, MouseY, 2, ss) = 0
            Case 4
                If SinData0(MouseX, MouseY, 3, ss) >= 0 Then
                    tmplong = SinData0(MouseX, MouseY, 3, ss)
                    D_Event0(tmplong, ss).isGo = 0
                    D_Event0(tmplong, ss).id = 0
                    D_Event0(tmplong, ss).EventNum1 = 0
                    D_Event0(tmplong, ss).EventNum2 = 0
                    D_Event0(tmplong, ss).EventNum3 = 0
                    D_Event0(tmplong, ss).picnum(0) = 0
                    D_Event0(tmplong, ss).picnum(1) = 0
                    D_Event0(tmplong, ss).picnum(2) = 0
                    D_Event0(tmplong, ss).PicDelay = 0
                    D_Event0(tmplong, ss).x = 0
                    D_Event0(tmplong, ss).y = 0
                    SinData0(MouseX, MouseY, 3, ss) = -1
                End If
            Case 5
                SinData0(MouseX, MouseY, 4, ss) = 0
            Case 6
                SinData0(MouseX, MouseY, 5, ss) = 0

            End Select
        End Select
    End If
    showsmap
End If
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i1 As Long
    Dim j1 As Long
        i1 = ((x - pic1.Width / 2) / XSCALE + (y - pic1.Height / 2 + YSCALE) / YSCALE) / 2
        j1 = -((x - pic1.Width / 2) / XSCALE - (y - pic1.Height / 2 + YSCALE) / YSCALE) / 2
        MouseX = i1 + xx
        MouseY = j1 + yy
        
    If Button = vbRightButton And iMode = 1 Then
        'Case 0
            Select Case ComboLevel.ListIndex
            Case 0
            
            Case 1
                SinData0(MouseX, MouseY, 0, ss) = SelectPicNum * 2
            Case 2
                SinData0(MouseX, MouseY, 1, ss) = SelectPicNum * 2
            Case 3
                SinData0(MouseX, MouseY, 2, ss) = SelectPicNum * 2
        End Select
    End If
        If iMode <> 1 And iMode <> 2 Then
            If MouseX >= 0 And MouseX < SMapXmax And MouseY >= 0 And MouseY < SMapYmax Then
                Call Show_picture(PicEarth, SinData0(MouseX, MouseY, 0, ss) / 2)
                Call Show_picture(PicBiuld, SinData0(MouseX, MouseY, 1, ss) / 2)
                Call Show_picture(PicAir, SinData0(MouseX, MouseY, 2, ss) / 2)
            
                lbl1.Caption = SinData0(MouseX, MouseY, 0, ss) / 2
                lbl2.Caption = SinData0(MouseX, MouseY, 1, ss) / 2
                lbl3.Caption = SinData0(MouseX, MouseY, 2, ss) / 2
                lblEventValue.Caption = SinData0(MouseX, MouseY, 3, ss)
                lbl5.Caption = SinData0(MouseX, MouseY, 4, ss)
                lbl6.Caption = SinData0(MouseX, MouseY, 5, ss)
            End If
        Else
            If (Button And vbLeftButton) > 0 And ComboLevel.ListIndex > 0 Then
                BlockX2 = MouseX
                BlockY2 = MouseY
            End If
            
        End If
        If Rtime >= 1 Then
            Draw_smap_2
            Rtime = 0
        End If
End Sub

Private Sub pic1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim X1 As Long, Y1 As Long
Dim X2 As Long, Y2 As Long
Dim i, j As Long
        


    If iMode = 2 And (ComboLevel.ListIndex > 0 And ComboLevel.ListIndex < 4) Then
        X1 = Min_V(BlockX1, BlockX2)
        X2 = Max_V(BlockX1, BlockX2)
        Y1 = Min_V(BlockY1, BlockY2)
        Y2 = Max_V(BlockY1, BlockY2)
        BlockX1 = IIf(X1 <= 0, 0, X1)               ' 设置x1,y1为最小点，x2,y2为大点
        BlockY1 = IIf(Y1 <= 0, 0, Y1)
        BlockX2 = IIf(X2 > SMapXmax - 1, SMapXmax - 1, X2)
        BlockY2 = IIf(Y2 > SMapYmax - 1, SMapYmax - 1, Y2)
        'MsgBox 11
        SelectBlock = 0

        'del
        Debug.Print BlockX1; BlockX2; BlockY1; BlockY2
            For i = BlockX1 To BlockX2
                For j = BlockY1 To BlockY2
                    SinData0(i, j, ComboLevel.ListIndex - 1, ss) = 0
                Next j
            Next i
        'SinData0(14, 14, 0, ss) = 0
        showsmap
        'showsmap
   ElseIf iMode = 1 Then
        If BlockX2 = -1 And BlockY2 = -1 Then
            BlockX1 = -1
            BlockY1 = -1
        End If
        If ComboLevel.ListIndex <> 6 And ComboLevel.ListIndex <> 5 Then
            SelectBlock = 0
        End If

        X1 = Min_V(BlockX1, BlockX2)
        X2 = Max_V(BlockX1, BlockX2)
        Y1 = Min_V(BlockY1, BlockY2)
        Y2 = Max_V(BlockY1, BlockY2)
        
        BlockX1 = X1                   ' 设置x1,y1为最小点，x2,y2为大点
        BlockY1 = Y1
        BlockX2 = X2
        BlockY2 = Y2

        Draw_smap_2
    End If
End Sub

Private Sub pic1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmpstrArray() As String
Dim tmplong As Long
   If data.GetFormat(vbCFText) Then
       tmpstrArray = Split(data.GetData(vbCFText), ":")
       If tmpstrArray(0) = G_Var.SMAPGRP Then
'           MsgBox tmplong
           tmplong = CLng(tmpstrArray(1))
               If ComboLevel.ListIndex = 4 Then
                    txtD(5) = tmplong * 2
                    txtD(6) = tmplong * 2
                    txtD(7) = tmplong * 2
                End If
          SelectPicNum = tmplong
          lblSelectPicNum.Caption = SelectPicNum
       ElseIf tmpstrArray(0) = G_Var.SMAPGRP2 Then
           tmplong = CLng(tmpstrArray(1))
               If ComboLevel.ListIndex = 4 Then
                    txtD(5) = tmplong * 2
                    txtD(6) = tmplong * 2
                    txtD(7) = tmplong * 2
                End If
          '这里颠倒了
          SelectPicNum = -tmplong
          lblSelectPicNum.Caption = SelectPicNum
       End If
   End If
End Sub

Private Sub RT_Timer()
Rtime = Rtime + 1
End Sub

Private Sub Timer1_Timer()
Dim vv As Long
    vv = lblEventValue.Caption
    If vv < 0 Then
        Current_D_EventNum = -1
        PicEvent.Cls
        Exit Sub
    End If
    If vv <> Current_D_EventNum Then
        Current_D_EventNum = vv
        Current_D_Event_Pic = D_Event0(vv, ss).picnum(0)
    Else
        If D_Event0(vv, ss).picnum(1) > Current_D_Event_Pic Then
            Current_D_Event_Pic = Current_D_Event_Pic + 2
        Else
            Current_D_Event_Pic = D_Event0(vv, ss).picnum(0)
        End If
    End If
        
        
    Call Show_picture(PicEvent, Current_D_Event_Pic / 2)
   
    
End Sub



Private Sub txtD_OLEDragDrop(Index As Integer, data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmpstrArray() As String
Dim tmplong As Long
   If data.GetFormat(vbCFText) Then
       tmpstrArray = Split(data.GetData(vbCFText), ":")
       If tmpstrArray(0) = G_Var.SMAPGRP Then
           tmplong = CLng(tmpstrArray(1))
           txtD(Index).Text = tmplong * 2
       End If
   End If
End Sub

Private Sub VScrollHeight_Change()
    ScrollValue
    showsmap
End Sub


Public Sub ClickMenu(id As String)
    Select Case LCase(id)
    Case "grid"
        MDIMain.mnu_SMAPMenu_Grid.Checked = Not MDIMain.mnu_SMAPMenu_Grid.Checked
        isGrid = IIf(MDIMain.mnu_SMAPMenu_Grid.Checked, 1, 0)
    Case "showlevel"
        MDIMain.mnu_SMAPMenu_ShowLevel.Checked = Not MDIMain.mnu_SMAPMenu_ShowLevel.Checked
        isShowLevel = IIf(MDIMain.mnu_SMAPMenu_ShowLevel.Checked, 1, 0)
    Case "normal"
        MDIMain.mnu_SMAPMenu_Normal.Checked = True
        MDIMain.mnu_SMAPMenu_BLock.Checked = False
        MDIMain.mnu_SMAPMenu_Delete.Checked = False
        Frame2.Visible = False
        iMode = 0
        Set_Note
    Case "block"
        MDIMain.mnu_SMAPMenu_Normal.Checked = False
        MDIMain.mnu_SMAPMenu_BLock.Checked = True
        MDIMain.mnu_SMAPMenu_Delete.Checked = False
        If ComboLevel.ListIndex = 5 Or ComboLevel.ListIndex = 6 Then
            Frame2.Visible = True
        Else
            Frame2.Visible = False
        End If
        iMode = 1
        Set_Note
    Case "delete"
        MDIMain.mnu_SMAPMenu_Normal.Checked = False
        MDIMain.mnu_SMAPMenu_BLock.Checked = False
        MDIMain.mnu_SMAPMenu_Delete.Checked = True
        Frame2.Visible = False
        iMode = 2
        Set_Note
    Case "loadmap0"
        Recordnum = 0
        MDIMain.mnu_SMAPMenu_LoadMap0.Checked = True
        MDIMain.mnu_SMAPMenu_LoadMap1.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap2.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap3.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap4.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap5.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap6.Checked = False
        Frame2.Visible = False
        Load_DS
    Case "loadmap1"
        Recordnum = 1
        MDIMain.mnu_SMAPMenu_LoadMap0.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap1.Checked = True
        MDIMain.mnu_SMAPMenu_LoadMap2.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap3.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap4.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap5.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap6.Checked = False
        Frame2.Visible = False
        Load_DS
    Case "loadmap2"
        Recordnum = 2
        MDIMain.mnu_SMAPMenu_LoadMap0.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap1.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap2.Checked = True
        MDIMain.mnu_SMAPMenu_LoadMap3.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap4.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap5.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap6.Checked = False
        Frame2.Visible = False
        Load_DS
    Case "loadmap3"
        Recordnum = 3
        MDIMain.mnu_SMAPMenu_LoadMap0.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap1.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap2.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap3.Checked = True
        MDIMain.mnu_SMAPMenu_LoadMap4.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap5.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap6.Checked = False
        Frame2.Visible = False
        Load_DS
    Case "loadmap4"
        Recordnum = 4
        MDIMain.mnu_SMAPMenu_LoadMap0.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap1.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap2.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap3.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap4.Checked = True
        MDIMain.mnu_SMAPMenu_LoadMap5.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap6.Checked = False
        Frame2.Visible = False
        Load_DS
    Case "loadmap5"
        Recordnum = 5
        MDIMain.mnu_SMAPMenu_LoadMap0.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap1.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap2.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap3.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap4.Checked = True
        MDIMain.mnu_SMAPMenu_LoadMap5.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap6.Checked = False
        Frame2.Visible = False
        Load_DS
    Case "loadmap6"
        Recordnum = 6
        MDIMain.mnu_SMAPMenu_LoadMap0.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap1.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap2.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap3.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap4.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap5.Checked = False
        MDIMain.mnu_SMAPMenu_LoadMap6.Checked = True
        Frame2.Visible = False
        Load_DS
    Case "save"  ' 保存进度
        Save_DS
    Case "addmap"  ' 增加场景地图
        AddMap
    Case "deletemap"   ' 删除地图
        DeleteMap
    Case "savemap"
        SaveKGSmap
    Case "loadmap"
        LoadKGSmap
    End Select
    showsmap
End Sub

Private Sub Set_Note()
Dim str As String
    Select Case iMode
    Case 0
        Select Case ComboLevel.ListIndex
        Case 0
            str = LoadResStr(10814)
        Case 1, 2, 3
            str = LoadResStr(10709)
        Case 4
            str = LoadResStr(10820)
        Case 5, 6
            str = LoadResStr(10815)
        Case 7
            str = StrUnicode2("左键设定入口，右键设定出口")
        End Select
    Case 1
        Select Case ComboLevel.ListIndex
        Case 0
            str = LoadResStr(10814)
        Case 1, 2
            str = StrUnicode2("按下左键拖动选择操作块/右键复制块")
        Case 5
            str = StrUnicode2("按下左键拖动选择操作块/在块上左键升高海拔，右键降低海拔/块外点击取消块")
        End Select
    Case 2
        Select Case ComboLevel.ListIndex
        Case 0
            str = LoadResStr(10814)
        Case 1, 2, 3
            str = LoadResStr(10710)
        Case 4
            str = LoadResStr(10821)
        Case 5, 6
            str = LoadResStr(10816)
        End Select
    End Select
    MDIMain.StatusBar1.Panels(1).Text = str
End Sub

' 增加场景地图
Private Sub AddMap()
Dim i As Long, j As Long, k As Long
  
    SceneMapNum = SceneMapNum + 1
    ComboScene.AddItem SceneMapNum - 1
    
    ReDim Preserve SceneIDX(SceneMapNum)
    ReDim Preserve SinData0(SMapXmax - 1, SMapYmax - 1, 5, SceneMapNum - 1)
    SceneIDX(SceneMapNum) = SceneIDX(SceneMapNum - 1) + 6# * 2 * SMapXmax * SMapXmax
    
    For i = 0 To SMapXmax - 1
        For j = 0 To SMapYmax - 1
            SinData0(i, j, 3, SceneMapNum - 1) = -1
        Next j
    Next i
    
    
    ReDim Preserve D_IDX(SceneMapNum)
    ReDim Preserve D_Event0(200 - 1, SceneMapNum - 1)
    D_IDX(SceneMapNum) = D_IDX(SceneMapNum - 1) + 4400
    

    If MsgBox(StrUnicode2("是否复制当前场景到新场景？"), vbYesNo, Me.Caption) = vbYes Then
        For k = 0 To 5
            For i = 0 To SMapXmax - 1
                For j = 0 To SMapYmax - 1
                    SinData0(i, j, k, SceneMapNum - 1) = SinData0(i, j, k, ss)
                Next j
            Next i
        Next k
        
        For i = 0 To 200 - 1
            D_Event0(i, SceneMapNum - 1) = D_Event0(i, ss)
        Next i
    End If
    
    
End Sub
Private Sub LoadKGMap()

End Sub
Private Sub DeleteMap()
    If MsgBox(StrUnicode2("将要删除最后一个场景，是否继续？"), vbYesNo, Me.Caption) = vbYes Then
        SceneMapNum = SceneMapNum - 1
        
        ReDim Preserve SceneIDX(SceneMapNum)
        ReDim Preserve SinData0(SMapXmax - 1, SMapYmax - 1, 5, SceneMapNum - 1)
        
        ReDim Preserve D_IDX(SceneMapNum)
        ReDim Preserve D_Event0(200 - 1, SceneMapNum - 1)
        ComboScene.RemoveItem SceneMapNum
        ComboScene.ListIndex = 0
    End If
End Sub

Private Sub ScrollValue()
    MouseX = MouseX - xx
    MouseY = MouseY - yy
    xx = HScrollWidth.Value + VScrollHeight.Value - SMapXmax / 2
    yy = -HScrollWidth.Value + VScrollHeight.Value + SMapXmax / 2
    MouseX = MouseX + xx
    MouseY = MouseY + yy
    MDIMain.StatusBar1.Panels(2).Text = " X=" & MouseX & ",Y=" & MouseY
End Sub

Private Sub VScrollHeight_Scroll()
    ScrollValue
End Sub

Private Sub X2_Click()
On Error Resume Next
txtD(5) = txtD(5) * 2
txtD(6) = txtD(6) * 2
txtD(7) = txtD(7) * 2
End Sub
Public Function getKGoffset(filename As String, KGpicnum As Long) As Long
Dim filenum As Long
Dim tmp As Long
    filenum = OpenBin(filename, "R")
        Get filenum, 4 * KGpicnum + 12, getKGoffset
    Close (filenum)
End Function
Public Sub SaveRinS(Rnum As Long)
Dim filenum As Long, i As Long, j As Long
Dim Rlong() As Long, idnum As Long
Dim offset As Long
Dim idxr() As Long
Dim kk
On Error GoTo Rerror:
    ReDim Rlong(GetINILong("R_Modify", "TypeNumber") - 1)
    For j = 0 To GetINILong("R_Modify", "TypeNumber") - 1
        For i = 0 To GetINILong("R_Modify", "TypedataItem" & j) - 1
            kk = Split(GetINIStr("R_Modify", "data(" & j & "," & i & ")"), " ")
            Rlong(j) = Rlong(j) + Val(kk(2)) * Val(kk(0)) * Val(kk(1))
        Next i
    Next j
    
    filenum = OpenBin(G_Var.JYPath & G_Var.RIDX(Rnum), "R")
        idnum = LOF(filenum) / 4
        ReDim idxr(idnum)
        idxr(0) = 0
        For i = 1 To idnum
            Get filenum, , idxr(i)
        Next i
    Close (filenum)
    
    filenum = OpenBin(G_Var.JYPath & G_Var.RGRP(Rnum), "W")
        offset = idxr(3)
        Put filenum, offset + 1, Scene
    Close (filenum)
Exit Sub
Rerror:
End Sub
