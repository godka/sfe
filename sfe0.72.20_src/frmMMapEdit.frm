VERSION 5.00
Begin VB.Form frmMMapEdit 
   Caption         =   "MMapedit"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMMapEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   Begin VB.Frame Frame1 
      Caption         =   "当前操作图片"
      Height          =   6015
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
      Begin VB.PictureBox PicEarth 
         AutoRedraw      =   -1  'True
         Height          =   495
         Left            =   840
         ScaleHeight     =   435
         ScaleWidth      =   795
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.PictureBox PicSurface 
         AutoRedraw      =   -1  'True
         Height          =   495
         Left            =   840
         ScaleHeight     =   435
         ScaleWidth      =   795
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.PictureBox PicBuilding 
         AutoRedraw      =   -1  'True
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1395
         ScaleWidth      =   1515
         TabIndex        =   12
         Top             =   1680
         Width           =   1575
      End
      Begin VB.PictureBox PicRef 
         AutoRedraw      =   -1  'True
         Height          =   2055
         Left            =   120
         ScaleHeight     =   1995
         ScaleWidth      =   1515
         TabIndex        =   11
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "层2表面"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "层3建筑"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "引用建筑"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lbl1 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lbl2 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lbl3 
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lbl4 
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "层1地面"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.HScrollBar HScrollWidth 
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   6960
      Width           =   7095
   End
   Begin VB.VScrollBar VScrollHeight 
      Height          =   6975
      Left            =   8880
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
      Top             =   0
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
      Left            =   1080
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdSelectMap 
      Caption         =   "选择贴图"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
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
      Height          =   6915
      Left            =   1920
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.ListBox SceneList 
         Height          =   2760
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   6480
         Top             =   0
      End
   End
   Begin VB.Label lblMenu 
      Caption         =   "<快捷菜单>"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "操作层"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblSelectPicNum 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "frmMMapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private SelectPicNum As Long
Private SelectRefX As Long
Private SelectRefY As Long

Private Const MMapXmax = 480
Private Const MMapYmax = 480

Private MMapPic() As RLEPic
Private MMappicnum As Long

'retry
Private Const RetryNum = 20
Private RetrySin() As Integer

Private EarthData(MMapXmax - 1, MMapYmax - 1) As Integer ' 主地图地面贴图编号数据
Private SurfaceData(MMapXmax - 1, MMapYmax - 1) As Integer   ' 主地图地面表面贴图编号数据
Private BuildingData(MMapXmax - 1, MMapYmax - 1) As Integer  ' 主地图建筑贴图编号数据
Private BuildxData(MMapXmax - 1, MMapYmax - 1) As Integer    ' 主地图建筑占用坐标x数据
Private BuildyData(MMapXmax - 1, MMapYmax - 1) As Integer    ' 主地图建筑占用坐标y数据

Private tmpBuildX(2000) As Long
Private tmpBuildY(2000) As Long
Private tmpBuild(2000) As Long
Private numTmpBuild As Long

Dim Ctime As Long
Private xx As Long              ' 绘图中心点
Private yy As Long

Private MouseX As Long          ' 鼠标位置
Private MouseY As Long

Private BlockX1 As Long, BlockY1 As Long     ' 选择块位置
Private BlockX2 As Long, BlockY2 As Long
Private SelectBlock As Long                  ' 0 未选择块，1 选择块

Private iMode As Long                      ' 0 正常   1 块操作  2 删除

Private isGrid As Long                       ' 0 不显示网格 1 显示网格
Private isShowLevel As Long                  ' 0 全部显示   1 只显示操作层
'Private isDelete As Long                     ' 0 正常模式   1 删除模式
Private isScene As Long                      ' 0 不显示     1 显示场景

Private MouseButton As Long








Private Sub cmdSelectMap_Click()
    SelectPicNum = -1
    Load frmSelectMap
    frmSelectMap.txtIDX = G_Var.MMAPIDX
    frmSelectMap.txtGRP = G_Var.MMAPGRP
    frmSelectMap.cmdshow_Click
    frmSelectMap.Show
End Sub



Private Sub showmmap()
    pic1.Cls
    Draw_Mmap
    Draw_Mmap_2
End Sub


Private Sub ComboLevel_click()
    If ComboLevel.ListIndex = 5 Then
            MDIMain.mnu_MMAPMenu_ShowScene.Checked = True
            isScene = 1
    Else
            MDIMain.mnu_MMAPMenu_ShowScene.Checked = False
            isScene = 0
    End If
    BlockX1 = -1
    BlockY1 = -1
    BlockX2 = -1
    BlockY2 = -1
    SelectBlock = 0
    Set_Note
    showmmap
End Sub


Private Sub Form_Load()
Dim i As Long
Dim fileid As String
Dim filepic As String
    Me.Caption = LoadResStr(217)
    
    For i = 0 To Me.Controls.Count - 1
        Call SetCaption(Me.Controls(i))
    Next i

    For i = 0 To Scenenum - 1
        SceneList.AddItem Big5toUnicode(Scene(i).Name1, 10)
    Next i
    
    isGrid = 0
    isShowLevel = 0
    iMode = 0
    isScene = 0
    
    BlockX1 = -1
    BlockY1 = -1
    BlockX2 = -1
    BlockY2 = -1
    SelectBlock = 0
    
    fileid = G_Var.JYPath & G_Var.MMAPIDX
    filepic = G_Var.JYPath & G_Var.MMAPGRP
    Call LoadPicFile(fileid, filepic, MMapPic, MMappicnum)
    
    LoadEarthData
    
    
    VScrollHeight.Max = MMapXmax - 1
    VScrollHeight.LargeChange = 5
    VScrollHeight.SmallChange = 1
    VScrollHeight.Value = MMapXmax / 2
    
    HScrollWidth.Max = MMapYmax - 1
    HScrollWidth.LargeChange = 5
    HScrollWidth.SmallChange = 1
    HScrollWidth.Value = MMapXmax / 2
    
    
    
    ComboLevel.Clear
    ComboLevel.AddItem LoadResStr(10701)
    ComboLevel.AddItem LoadResStr(10702)
    ComboLevel.AddItem LoadResStr(10703)
    ComboLevel.AddItem LoadResStr(10704)
    ComboLevel.AddItem LoadResStr(10707)
    ComboLevel.AddItem StrUnicode2("设置场景入口")
    ComboLevel.ListIndex = 0
        c_Skinner.AttachSkin Me.hwnd
End Sub


'  读取主地图坐标数据
'  主地图为480*480，按照行排列，每个坐标两个字节，保存坐标的贴图等信息
Public Sub LoadEarthData()
Dim filenum As Integer
Dim i As Long, j As Long
    filenum = OpenBin(G_Var.JYPath & G_Var.MMAPStruct(0), "R")
    Get filenum, , EarthData ' 采用这样的读数据方法可以加快速度，直接顺序读到数组所有元素
    Close filenum
    
    filenum = OpenBin(G_Var.JYPath & G_Var.MMAPStruct(1), "R")
    Get filenum, , SurfaceData
    Close filenum
   
    filenum = OpenBin(G_Var.JYPath & G_Var.MMAPStruct(2), "R")
    Get filenum, , BuildingData
    Close filenum
    
    filenum = OpenBin(G_Var.JYPath & G_Var.MMAPStruct(3), "R")
    Get filenum, , BuildxData
    Close filenum
    
    filenum = OpenBin(G_Var.JYPath & G_Var.MMAPStruct(4), "R")
    Get filenum, , BuildyData
    Close filenum

End Sub

'  写主地图坐标数据
'  主地图为480*480，按照行排列，每个坐标两个字节，保存坐标的贴图等信息
Public Sub SAveEarthData()
Dim filenum As Integer
Dim i As Long, j As Long
    filenum = OpenBin(G_Var.JYPath & G_Var.MMAPStruct(0), "W")
    Put filenum, , EarthData ' 采用这样的读数据方法可以加快速度，直接顺序读到数组所有元素
    Close filenum
    
    filenum = OpenBin(G_Var.JYPath & G_Var.MMAPStruct(1), "W")
    Put filenum, , SurfaceData
    Close filenum
   
    filenum = OpenBin(G_Var.JYPath & G_Var.MMAPStruct(2), "W")
    Put filenum, , BuildingData
    Close filenum
    
    filenum = OpenBin(G_Var.JYPath & G_Var.MMAPStruct(3), "W")
    Put filenum, , BuildxData
    Close filenum
    
    filenum = OpenBin(G_Var.JYPath & G_Var.MMAPStruct(4), "W")
    Put filenum, , BuildyData
    Close filenum

    SaveRinS (0)
End Sub



' 确定各个build绘制的先后顺序
'
Private Sub Game_Mmap_Build()
Dim RangeX As Long, rangeY As Long
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim p As Long
Dim repeat As Long
Dim bakx As Long, baky As Long, bak As Long
Dim tmpx As Long, tmpy As Long, tmp As Long
Dim dy As Long
Dim Xmin As Long, Xmax As Long, ymin As Long, ymax As Long
    RangeX = 21
    rangeY = 31
    
    p = 0
            
    Xmin = xx - RangeX - 6
    RangeValue Xmin, 1, MMapXmax - 1
    Xmax = xx + RangeX + 8
    RangeValue Xmax, 1, MMapXmax - 1
    ymin = yy - rangeY
    RangeValue ymin, 1, MMapYmax - 1
    ymax = yy + rangeY + 8
    RangeValue ymax, 1, MMapYmax - 1
    
    For i = Xmin To Xmax
        dy = ymin
        For j = ymin To ymax
            If BuildxData(i, j) <> 0 And BuildyData(i, j) <> 0 Then
                repeat = 0
                For k = 0 To p - 1
                    If (tmpBuildX(k) = BuildxData(i, j)) And (tmpBuildY(k) = BuildyData(i, j)) Then
                        repeat = 1
                        If k = p - 1 Then
                            Exit For
                        End If
                        For m = j - 1 To dy Step -1
                            If BuildxData(i, m) <> 0 Or BuildyData(i, m) <> 0 Then
                                If (BuildxData(i, m) <> BuildxData(i, j)) Or (BuildyData(i, m) <> BuildyData(i, j)) Then
                                    If (BuildxData(i, m) <> tmpBuildX(k)) Or (BuildyData(i, m) <> tmpBuildY(k)) Then
                                        tmpx = tmpBuildX(p - 1)
                                        tmpy = tmpBuildY(p - 1)
                                        tmp = tmpBuild(p - 1)
                                         MoveMemory tmpBuildX(k + 1), tmpBuildX(k), (p - 2 - k + 1) * 4
                                         MoveMemory tmpBuildY(k + 1), tmpBuildY(k), (p - 2 - k + 1) * 4
                                         MoveMemory tmpBuild(k + 1), tmpBuild(k), (p - 2 - k + 1) * 4
                                        tmpBuildX(k) = tmpx
                                        tmpBuildY(k) = tmpy
                                        tmpBuild(k) = tmp
                                    End If
                                End If
                            End If
                        Next m
                        dy = j + 1
                        Exit For
                    End If
                Next k
                If repeat = 0 Then
                    tmpBuildX(p) = BuildxData(i, j)
                    tmpBuildY(p) = BuildyData(i, j)
                    tmpBuild(p) = BuildingData(BuildxData(i, j), BuildyData(i, j))
                    p = p + 1
                End If
            End If
        Next j
    Next i
    numTmpBuild = p

End Sub

' 绘大地图

Public Sub Draw_Mmap()
Dim RangeX As Long, rangeY As Long
Dim i As Long, j As Long
Dim i1 As Long, j1 As Long
Dim X1 As Long, Y1 As Long
Dim picnum As Long
Dim k As Long
Dim temp As Long
Dim dib As New clsDIB
Dim tmpStr As String

    dib.CreateDIB pic1.Width, pic1.Height
    
    RangeX = 18 + 16
    rangeY = 10 + 11
    
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
            
            If yy + j1 >= 0 And xx + i1 >= 0 And yy + j1 < MMapYmax And xx + i1 < MMapXmax Then
                picnum = EarthData(xx + i1, yy + j1) / 2
                If picnum > 0 And picnum < MMappicnum Then
                    If Not (isShowLevel = 1 And ComboLevel.ListIndex <> 1) Then
                        Call genPicData(MMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - MMapPic(picnum).x, Y1 - MMapPic(picnum).y)
                    End If
                End If
                picnum = SurfaceData(xx + i1, yy + j1) / 2
                If picnum > 0 And picnum < MMappicnum And Not (isShowLevel = 1 And ComboLevel.ListIndex <> 2) Then
                    Call genPicData(MMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - MMapPic(picnum).x, Y1 - MMapPic(picnum).y)
                End If
            End If
        Next i
    Next j
    
    For i = 0 To numTmpBuild - 1
        i1 = tmpBuildX(i) - xx
        j1 = tmpBuildY(i) - yy
        X1 = XSCALE * (i1 - j1) + pic1.Width / 2
        Y1 = YSCALE * (i1 + j1) + pic1.Height / 2
        picnum = tmpBuild(i) / 2
        If picnum > 0 And picnum < MMappicnum And ((Not (isShowLevel = 1 And ComboLevel.ListIndex <> 3)) Or ComboLevel.ListIndex = 4) Then
            Call genPicData(MMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - MMapPic(picnum).x, Y1 - MMapPic(picnum).y)
        End If
    Next i
    
    
    PicBak.Cls
     
        ' 复制到dib上
    temp = BitBlt(PicBak.hdc, 0, 0, pic1.Width, pic1.Height, dib.CompDC, 0, 0, &HCC0020)
   
    
    PicBak.ForeColor = &H808000
    
   
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
            
            If isGrid = 1 Then
                      PicBak.Line (X1, Y1)-(X1 + XSCALE, Y1 - YSCALE)
                      PicBak.Line (X1, Y1)-(X1 - XSCALE, Y1 - YSCALE)
            End If
            
        Next i
    Next j
    
     PicBak.FontSize = 7
     PicBak.ForeColor = vbRed
    
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
            
            
            If isShowLevel = 1 And ComboLevel.ListIndex = 4 Then
                 If yy + j1 >= 0 And xx + i1 >= 0 And yy + j1 < MMapYmax And xx + i1 < MMapXmax Then
                     If BuildxData(xx + i1, yy + j1) > 0 And BuildyData(xx + i1, yy + j1) > 0 Then
                        PicBak.CurrentX = X1 - XSCALE / 2
                        PicBak.CurrentY = Y1 - YSCALE - 6
                        PicBak.Print BuildxData(xx + i1, yy + j1)
                        PicBak.CurrentX = X1 - XSCALE / 2
                        PicBak.CurrentY = Y1 - YSCALE + 2
                        PicBak.Print BuildyData(xx + i1, yy + j1)
                    End If
                End If
             End If
        
        Next i
    Next j
    
    
     PicBak.FontSize = 10
     PicBak.ForeColor = vbYellow
    
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
            
            
        
            If isScene = 1 Then
                For k = 0 To Scenenum - 1
                    If (xx + i1 = Scene(k).MMapInX1 And yy + j1 = Scene(k).MMapInY1) Or (xx + i1 = Scene(k).MMapInX2 And yy + j1 = Scene(k).MMapInY2) Then
                        PicBak.CurrentX = X1 - XSCALE / 2
                        PicBak.CurrentY = Y1 - YSCALE - 4
                        tmpStr = Big5toUnicode(Scene(k).Name1, 10)
                        PicBak.Print "(" & xx + i1 & "," & yy + j1 & ")" & tmpStr
                    End If
                Next k
            End If
        
        Next i
    Next j

End Sub


Public Sub Draw_Mmap_2()
Dim RangeX As Long, rangeY As Long
Dim i As Long, j As Long
Dim i1 As Long, j1 As Long
Dim X1 As Long, Y1 As Long
Dim picnum As Long
Dim k As Long
    
   
Dim temp As Long
Dim dib As New clsDIB
    
    RangeX = 18 + 11
    rangeY = 10 + 11

    dib.CreateDIB pic1.Width, pic1.Height
    
    temp = BitBlt(dib.CompDC, 0, 0, pic1.Width, pic1.Height, PicBak.hdc, 0, 0, &HCC0020)
    
    
    
    
    
    i1 = MouseX - xx
    j1 = MouseY - yy
    
    X1 = XSCALE * (i1 - j1) + pic1.Width / 2
    Y1 = YSCALE * (i1 + j1) + pic1.Height / 2
    picnum = SelectPicNum
    If ComboLevel.ListIndex <> 4 Then
        If picnum >= 0 And picnum < MMappicnum And iMode <> 2 Then
            Call genPicData(MMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - MMapPic(picnum).x, Y1 - MMapPic(picnum).y)
       End If
    End If
     
     
   If iMode = 1 And SelectBlock = 0 Then
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
                    
                    Select Case ComboLevel.ListIndex
                    Case 0
                    Case 1
                        picnum = EarthData(BlockX2 - MouseX + i1 + xx, BlockY2 - MouseY + j1 + yy) / 2
                        If picnum > 0 And picnum < MMappicnum Then
                            Call genPicData(MMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - MMapPic(picnum).x, Y1 - MMapPic(picnum).y)
                        End If
                    Case 2
                        picnum = SurfaceData(BlockX2 - MouseX + i1 + xx, BlockY2 - MouseY + j1 + yy) / 2
                        If picnum > 0 And picnum < MMappicnum Then
                            Call genPicData(MMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - MMapPic(picnum).x, Y1 - MMapPic(picnum).y)
                        End If
                    Case 3
                        picnum = BuildingData(BlockX2 - MouseX + i1 + xx, BlockY2 - MouseY + j1 + yy) / 2
                        If picnum > 0 And picnum < MMappicnum Then
                            Call genPicData(MMapPic(picnum), dib.addr, pic1.Width, pic1.Height, X1 - MMapPic(picnum).x, Y1 - MMapPic(picnum).y)
                        End If
                    Case 4
                    End Select
                End If
               Next i
         Next j
      End If
    End If
     
     
     pic1.Cls
        ' 复制到dib上
    temp = BitBlt(pic1.hdc, 0, 0, pic1.Width, pic1.Height, dib.CompDC, 0, 0, &HCC0020)
   
   
    MDIMain.StatusBar1.Panels(2).Text = " X=" & MouseX & ",Y=" & MouseY
    
    If ComboLevel.ListIndex = 4 Then
        pic1.CurrentX = X1
        pic1.CurrentY = Y1
        pic1.ForeColor = vbYellow
        pic1.Print "(" & SelectRefX & "," & SelectRefY & ")"
    
    End If
    
   If iMode = 1 And SelectBlock = 1 And (ComboLevel.ListIndex = 1 Or ComboLevel.ListIndex = 2 Or ComboLevel.ListIndex = 3) Then
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
End Sub


Public Sub Show_picture(pic As PictureBox, ByVal num As Long)
   
Dim temp As Long
Dim dib As New clsDIB
    
    dib.CreateDIB pic.Width, pic.Height
    pic.BackColor = MaskColor
    
    temp = BitBlt(dib.CompDC, 0, 0, pic.Width, pic.Height, pic.hdc, 0, 0, &HCC0020)
    
    'Picnum = num
    If num >= 0 Then
        Call genPicData(MMapPic(num), dib.addr, pic.Width, pic.Height, 0, 0)
    End If
        ' 复制到dib上
    temp = BitBlt(pic.hdc, 0, 0, pic.Width, pic.Height, dib.CompDC, 0, 0, &HCC0020)
   
End Sub

    



Private Sub Form_Resize()

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
    pic1.Height = Me.ScaleHeight - HScrollWidth.Height - pic1.top
    If pic1.Height Mod 2 = 1 Then
        pic1.Height = pic1.Height + 1
    End If
    VScrollHeight.Height = pic1.Height
    HScrollWidth.top = pic1.top + pic1.Height
    PicBak.Width = pic1.Width
    PicBak.Height = pic1.Height
    'Call Game_Mmap_Build
    showmmap
      
End Sub

Private Sub Form_Unload(cancel As Integer)
    MDIMain.StatusBar1.Panels(1).Text = ""
    MDIMain.StatusBar1.Panels(2).Text = ""
    
End Sub

Private Sub HScrollWidth_Change()
    ScrollValue
    Call Game_Mmap_Build
    showmmap
End Sub

Private Sub HScrollWidth_Scroll()
    ScrollValue
End Sub

Private Sub lblMenu_Click()
    PopupMenu MDIMain.mnu_MMAPMenu
End Sub

Private Sub pic1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If SceneList.Visible = True Then
        SceneList.Visible = False
    End If
End If
    
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long
If MouseX >= 0 And MouseX < MMapXmax And MouseY >= 0 And MouseY < MMapYmax Then
    
    If Button = vbLeftButton Then   ' 左键按下，拾取。
        Select Case ComboLevel.ListIndex
        Case 0
        
        Case 1
            SelectPicNum = EarthData(MouseX, MouseY) / 2
        Case 2
            SelectPicNum = SurfaceData(MouseX, MouseY) / 2
        Case 3
            SelectPicNum = BuildingData(MouseX, MouseY) / 2
        Case 4
            SelectRefX = BuildxData(MouseX, MouseY)
            SelectRefY = BuildyData(MouseX, MouseY)
        Case 5
            MouseButton = 0
            SceneList.Left = x
            SceneList.top = y - SceneList.Height
            SceneList.Visible = True
        End Select
        
        If iMode = 1 Then
            BlockX1 = MouseX
            BlockY1 = MouseY
            BlockX2 = -1
            BlockY2 = -1
            SelectBlock = 1
        End If
        
    ElseIf Button = vbRightButton Then
        Select Case iMode
        Case 0
            Select Case ComboLevel.ListIndex
            Case 0
            
            Case 1
                    EarthData(MouseX, MouseY) = SelectPicNum * 2
            Case 2
                    SurfaceData(MouseX, MouseY) = SelectPicNum * 2
            Case 3
                 BuildingData(MouseX, MouseY) = SelectPicNum * 2
                 If SelectPicNum > 0 Then
                     BuildxData(MouseX, MouseY) = MouseX
                     BuildyData(MouseX, MouseY) = MouseY
                 End If
            Case 4
                 BuildxData(MouseX, MouseY) = SelectRefX
                 BuildyData(MouseX, MouseY) = SelectRefY
            Case 5
                MouseButton = 1
                SceneList.Left = x
                SceneList.top = y - SceneList.Height
                SceneList.Visible = True
            End Select
        Case 1
            Select Case ComboLevel.ListIndex
            Case 0
            
            Case 1
                    If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
                        For i = BlockX1 To BlockX2
                            For j = BlockY1 To BlockY2
                                If MouseX - BlockX2 + i >= 0 And MouseX - BlockX2 + i < MMapXmax And MouseY - BlockY2 + j >= 0 And MouseY - BlockY2 + j < MMapYmax Then
                                    If EarthData(i, j) > 0 Then
                                        EarthData(MouseX - BlockX2 + i, MouseY - BlockY2 + j) = EarthData(i, j)
                                    End If
                                End If
                            Next j
                        Next i
                    End If
            Case 2
                    If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
                        For i = BlockX1 To BlockX2
                            For j = BlockY1 To BlockY2
                                If MouseX - BlockX2 + i >= 0 And MouseX - BlockX2 + i < MMapXmax And MouseY - BlockY2 + j >= 0 And MouseY - BlockY2 + j < MMapYmax Then
                                    If SurfaceData(i, j) > 0 Then
                                        SurfaceData(MouseX - BlockX2 + i, MouseY - BlockY2 + j) = SurfaceData(i, j)
                                    End If
                                End If
                            Next j
                        Next i
                    End If
            Case 3
                     If BlockX1 >= 0 And BlockX2 >= 0 And BlockY1 >= 0 And BlockY2 >= 0 Then
                        For i = BlockX1 To BlockX2
                            For j = BlockY1 To BlockY2
                                If MouseX - BlockX2 + i >= 0 And MouseX - BlockX2 + i < MMapXmax And MouseY - BlockY2 + j >= 0 And MouseY - BlockY2 + j < MMapYmax Then
                                    If BuildingData(i, j) > 0 Then
                                        BuildingData(MouseX - BlockX2 + i, MouseY - BlockY2 + j) = BuildingData(i, j)
                                    End If
                                End If
                            Next j
                        Next i
                    End If
            End Select
        Case 2
            Select Case ComboLevel.ListIndex
            Case 0
            
            Case 1
                EarthData(MouseX, MouseY) = 0
            Case 2
                SurfaceData(MouseX, MouseY) = 0
            Case 3
                If BuildingData(MouseX, MouseY) > 0 Then
                    For i = 0 To 480 - 1
                        For j = 0 To 480 - 1
                            If BuildxData(i, j) = MouseX And BuildyData(i, j) = MouseY Then
                                BuildxData(i, j) = 0
                                BuildyData(i, j) = 0
                            End If
                        Next j
                    Next i
                    BuildingData(MouseX, MouseY) = 0
                End If
            Case 4
                BuildxData(MouseX, MouseY) = 0
                BuildyData(MouseX, MouseY) = 0
            End Select
        End Select
        Game_Mmap_Build
        showmmap
    End If
End If
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i1 As Long
Dim j1 As Long
    i1 = ((x - pic1.Width / 2) / XSCALE + (y - pic1.Height / 2 + YSCALE) / YSCALE) / 2
    j1 = -((x - pic1.Width / 2) / XSCALE - (y - pic1.Height / 2 + YSCALE) / YSCALE) / 2
    MouseX = i1 + xx
    MouseY = j1 + yy

    If iMode <> 1 Then
    
        
        If MouseX >= 0 And MouseX < MMapXmax And MouseY >= 0 And MouseY < MMapYmax Then
            Call Show_picture(PicEarth, EarthData(MouseX, MouseY) / 2)
            Call Show_picture(PicSurface, SurfaceData(MouseX, MouseY) / 2)
            Call Show_picture(PicBuilding, BuildingData(MouseX, MouseY) / 2)
            Call Show_picture(PicRef, BuildingData(BuildxData(MouseX, MouseY), BuildyData(MouseX, MouseY)) / 2)
            lbl1.Caption = EarthData(MouseX, MouseY) / 2
            lbl2.Caption = SurfaceData(MouseX, MouseY) / 2
            lbl3.Caption = BuildingData(MouseX, MouseY) / 2
            lbl4.Caption = "(" & BuildxData(MouseX, MouseY) & "," & BuildyData(MouseX, MouseY) & ")"
        End If
    Else
        If (Button And vbLeftButton) > 0 Then
            BlockX2 = MouseX
            BlockY2 = MouseY
            
        End If
    End If
    If Ctime >= 1 Then
        Draw_Mmap_2
        Ctime = 0
    End If
End Sub

Private Sub pic1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim X1 As Long, Y1 As Long
Dim X2 As Long, Y2 As Long

    If iMode = 1 Then
        If BlockX2 = -1 And BlockY2 = -1 Then
            BlockX1 = -1
            BlockY1 = -1
        End If
        SelectBlock = 0
        X1 = Min_V(BlockX1, BlockX2)
        X2 = Max_V(BlockX1, BlockX2)
        Y1 = Min_V(BlockY1, BlockY2)
        Y2 = Max_V(BlockY1, BlockY2)
        
        BlockX1 = X1                   ' 设置x1,y1为最小点，x2,y2为大点
        BlockY1 = Y1
        BlockX2 = X2
        BlockY2 = Y2
        
        Draw_Mmap_2
    End If
    
End Sub

Private Sub pic1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmpstrArray() As String
Dim tmplong As Long
   If data.GetFormat(vbCFText) Then
       tmpstrArray = Split(data.GetData(vbCFText), ":")
       If tmpstrArray(0) = G_Var.MMAPGRP Then
           tmplong = CLng(tmpstrArray(1))
           SelectPicNum = tmplong
           lblSelectPicNum.Caption = SelectPicNum
       End If
   End If

End Sub

Private Sub SceneList_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    SceneList.Visible = False
End If
End Sub

Private Sub SceneList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Index As Long
    Index = SceneList.ListIndex
    Select Case MouseButton
    Case 0
        Scene(Index).MMapInX1 = MouseX: Scene(Index).MMapInY1 = MouseY
    Case 1
        Scene(Index).MMapInX2 = MouseX: Scene(Index).MMapInY2 = MouseY
    End Select
    SceneList.Visible = False
    showmmap
End Sub

Private Sub Timer1_Timer()
Ctime = Ctime + 1
End Sub

Private Sub VScrollHeight_Change()
    ScrollValue
    Call Game_Mmap_Build
    showmmap
End Sub


Private Sub VScrollHeight_Scroll()
    ScrollValue
End Sub


Private Sub ScrollValue()
    MouseX = MouseX - xx
    MouseY = MouseY - yy
    xx = HScrollWidth.Value + VScrollHeight.Value - MMapXmax / 2
    yy = -HScrollWidth.Value + VScrollHeight.Value + MMapXmax / 2
    MouseX = MouseX + xx
    MouseY = MouseY + yy
    MDIMain.StatusBar1.Panels(2).Text = " X=" & MouseX & ",Y=" & MouseY
End Sub

Public Sub ClickMenu(id As String)
Dim b As Boolean
    Select Case LCase(id)
    Case "grid"
        MDIMain.mnu_MMAPMenu_Grid.Checked = Not MDIMain.mnu_MMAPMenu_Grid.Checked
        isGrid = IIf(MDIMain.mnu_MMAPMenu_Grid.Checked, 1, 0)
    Case "showlevel"
        MDIMain.mnu_MMAPMenu_ShowLevel.Checked = Not MDIMain.mnu_MMAPMenu_ShowLevel.Checked
        isShowLevel = IIf(MDIMain.mnu_MMAPMenu_ShowLevel.Checked, 1, 0)
    Case "normal"
        MDIMain.mnu_MMAPMenu_Normal.Checked = True
        MDIMain.mnu_MMAPMenu_Block.Checked = False
        MDIMain.mnu_MMAPMenu_Delete.Checked = False
        iMode = 0
        Set_Note
    Case "block"
        MDIMain.mnu_MMAPMenu_Normal.Checked = False
        MDIMain.mnu_MMAPMenu_Block.Checked = True
        MDIMain.mnu_MMAPMenu_Delete.Checked = False
        iMode = 1
        Set_Note
    Case "delete"
        MDIMain.mnu_MMAPMenu_Normal.Checked = False
        MDIMain.mnu_MMAPMenu_Block.Checked = False
        MDIMain.mnu_MMAPMenu_Delete.Checked = True
        iMode = 2
        Set_Note
    Case "showscene"
        MDIMain.mnu_MMAPMenu_ShowScene.Checked = Not MDIMain.mnu_MMAPMenu_ShowScene.Checked
        isScene = IIf(MDIMain.mnu_MMAPMenu_ShowScene.Checked, 1, 0)
    Case "savemap"
        SAveEarthData
    Case "loadmap"
        LoadEarthData
        Game_Mmap_Build
    End Select
    showmmap
End Sub


Private Sub Set_Note()
Dim str As String
    Select Case iMode
    Case 0
        Select Case ComboLevel.ListIndex
        Case 0
            str = ""
        Case 5
            str = StrUnicode2("左键选择入口1，右键选择入口2")
        Case Else
            str = LoadResStr(10709)
        End Select
    Case 1
        Select Case ComboLevel.ListIndex
        Case 0
            str = ""
        Case Else
                str = StrUnicode2("按下左键拖动选择操作块/右键复制块,只有层1/2能进行块操作")
        End Select
    Case 2
        Select Case ComboLevel.ListIndex
        Case 0
            str = ""
        Case Else
            str = LoadResStr(10710)
        End Select
    End Select
    MDIMain.StatusBar1.Panels(1).Text = str
End Sub
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
