VERSION 5.00
Begin VB.Form MainSelectMap 
   Caption         =   "贴图查看/编辑"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   Begin VB.CommandButton cmdShow 
      Caption         =   "贴图查看"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox pic6 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4560
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   9
      Tag             =   "左右键盘切换图片，鼠标左键选定偏移"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox pic5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   4560
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   8
      Tag             =   "左右键盘切换图片，鼠标左键选定偏移"
      Top             =   2640
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   4440
      Top             =   0
   End
   Begin VB.PictureBox Picbak 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   7560
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox PicLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00CC0020&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   7560
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   5
      ToolTipText     =   "左键拾取颜色，右键修改颜色"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7080
      Width           =   4575
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      ItemData        =   "MainSelectMap.frx":0000
      Left            =   840
      List            =   "MainSelectMap.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6615
      Left            =   4560
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Labelpic 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "每行贴图"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "MainSelectMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private X_num  As Long            ' pic中显示的每行每列贴图个数
Private Y_num  As Long

Private XWidth(10) As Long        ' pic中每行贴图最大高和每列最大宽
Private YHeight(30) As Long

Public n As Long
Dim picScale As Long

Private Current_startpic As Long   ' pic 左上角贴图编号
Private real_Current_startpic As Long
Private current_X As Long          ' 鼠标在pic中贴图位置编号
Private current_Y As Long

Private MapPic() As RLEPic         ' 贴图数据
Private Mappicnum As Long          ' 贴图个数
Private RealMappicnum As Long
Private ReadPic As Boolean

Private copyPicnum As Long         ' 复制的贴图编号
Private combotypeNum As Long, Combolistindex As Long
Private Pic5num As Long
Private WW, HH, x, y As Long
Private data() As Long
Dim Ctime As Long

Public Sub cmdshow_Click()
Dim fileid As String
Dim filepic As String
Dim TypeNum As Long, i As Long
Dim tmpStr() As String
    'If txtIDX.Text = "" Or txtGRP = "" Then Exit Sub
    
    'If GetINILong("map_add_" & ComboFilename.ListIndex, "Num") > 0 Then
    '    Combolistindex = ComboFilename.ListIndex
    '    ComboType.Clear
    '    TypeNum = GetINILong("map_add_" & ComboFilename.ListIndex, "Num")
    '    For i = 0 To TypeNum - 1
    '        'MsgBox GetINILong("map_add_" & ComboFilename.ListIndex, "0")
    '        tmpStr = Split(GetINIStr("map_add_" & ComboFilename.ListIndex, Val(i)), ",")
    '        ComboType.AddItem (tmpStr(0))
    '    Next i
    '    ComboType.ListIndex = 0
    '    ComboType.Visible = True
    'Else
    '    real_Current_startpic = 0
    '    ComboType.Visible = False
    'End If
    'MsgBox Command
    fileid = Command
    filepic = G_Var.JYPath & txtGRP
    On Error GoTo Label1:
    Call LoadPicFile(fileid, filepic, MapPic, Mappicnum)
    RealMappicnum = Mappicnum
    ReadPic = True
    Combo1.ListIndex = 0
    Combo1_Click
    copyPicnum = -1
    Exit Sub
Label1:
    MsgBox "No this file"
End Sub

Private Sub Combo1_Click()
    Select Case Combo1.ListIndex
    Case 0
        X_num = 10
        VScroll1.LargeChange = 5
    Case 1
        X_num = 5
        VScroll1.LargeChange = 3
    Case 2
    
    End Select
    Current_startpic = 0
    VScroll1.Max = Mappicnum \ X_num
    VScroll1.SmallChange = 1
    VScroll1.Value = 0
    Showpic (0)
End Sub



Private Sub ComboFilename_click()
Dim tmpStr() As String
Dim TypeNum As Long
Dim i As Long
    If ComboFilename.ListIndex = -1 Then Exit Sub
    tmpStr = Split(ComboFilename.Text, ",")
    txtIDX.Text = Trim(tmpStr(0))
    txtGRP.Text = Trim(tmpStr(1))
    'MsgBox GetINILong("map_add_" & ComboFilename.ListIndex, "Num")
    If GetINILong("map_add_" & ComboFilename.ListIndex, "Num") > 0 Then
        Combolistindex = ComboFilename.ListIndex
        ComboType.Clear
        TypeNum = GetINILong("map_add_" & ComboFilename.ListIndex, "Num")
        For i = 0 To TypeNum - 1
            'MsgBox GetINILong("map_add_" & ComboFilename.ListIndex, "0")
            tmpStr = Split(GetINIStr("map_add_" & ComboFilename.ListIndex, Val(i)), ",")
            ComboType.AddItem (tmpStr(0))
        Next i
        ComboType.ListIndex = 0
        ComboType.Visible = True
    Else
        ComboType.Visible = False
    End If
End Sub



Private Sub Xjapan2()
Dim i, j, k As Long
Dim Picbegin, Picend As Long
'Dim WW, HH As Long
Dim picstep As Long
'Dim data() As Long
Dim temp As Long
Dim dib As New clsDIB

    frmsize.Show vbModal
    If frmsize.cok = 0 Then Exit Sub
    
    picScale = frmsize.pscale '扩大两倍，=3就是扩大三倍，自己调
    Picbak.BackColor = MaskColor
    PicLarge.BackColor = MaskColor
    PicLarge.Cls
    Picbak.Cls
    Picbegin = frmsize.Pic1
    Picend = frmsize.pic2
    picstep = IIf(Picbegin <= Picend, 1, -1)
    'Picbegin = InputBox(StrUnicode2("输入起始贴图编号"))
    'Picend = InputBox(StrUnicode2("输入结束贴图编号"))
    If MsgBox(StrUnicode2("是否确定这些贴图放大" & picScale & "倍？") & "(" & Picbegin & "," & Picend & ")", vbQuestion + vbYesNo) = vbYes Then
        Me.MousePointer = vbHourglass
        For k = Picbegin To Picend Step picstep
            Labelpic.Caption = StrUnicode2("正在处理图片:") & k
            DoEvents
            WW = MapPic(k).Width
            HH = MapPic(k).Height
            x = MapPic(k).x
            y = MapPic(k).y
            
            'MsgBox WW
            If WW > 0 And HH > 0 Then
                ReDim data(WW - 1, HH - 1)
        
                Picbak.Width = WW
                Picbak.Height = HH
                
                dib.CreateDIB WW, HH
                Picbak.BackColor = MaskColor
                temp = BitBlt(dib.CompDC, 0, 0, WW, HH, Picbak.hdc, 0, 0, &HCC0020)
        
                Call genPicData(MapPic(k), dib.addr, WW, HH, 0, 0)
                ' 复制到dib上
                temp = BitBlt(Picbak.hdc, 0, 0, WW, HH, dib.CompDC, 0, 0, &HCC0020)
        
                For j = 0 To HH - 1
                    For i = 0 To WW - 1
                        data(i, j) = Picbak.Point(i, j)
                    Next i
                Next j

            

                WW = WW * picScale
                HH = HH * picScale
                PicLarge.Width = WW + 10
                PicLarge.Height = HH + 10
                PicLarge.Cls
                PicLarge.PaintPicture Picbak.Image, 0, 0, WW * picScale - 0, HH * picScale - 0, 0, 0, WW - 0 / picScale, HH - 0 / picScale


                ReDim data(WW - 1, HH - 1)
            
                For j = 0 To HH - 1
                    For i = 0 To WW - 1
                        data(i, j) = PicLarge.Point(i, j)
                    Next i
                Next j
                x = picScale * x
                y = y * picScale
                Call SaveMapPic
                MapPic(k) = g_PP
                'DoEvents
            End If
        Next k
        'MsgBox g_PP.Width
        
        Showpic (0)
        Me.MousePointer = vbDefault
        MsgBox "Done"
    End If

End Sub


Private Sub ComboType_click()
Dim tmpStr() As String
Dim i As Long
    tmpStr = Split(GetINIStr("map_add_" & Combolistindex, ComboType.ListIndex), ",")
    If tmpStr(2) <> "end" Then
        Mappicnum = Val(tmpStr(2))
    Else
        Mappicnum = RealMappicnum
    End If
    Current_startpic = Val(tmpStr(1))
    real_Current_startpic = Current_startpic
    'Mappicnum = Val(tmpstr(2))
    VScroll1.Max = Mappicnum \ X_num
    VScroll1.SmallChange = 1
    VScroll1.Value = 0
    Showpic (0)

End Sub

Private Sub Command1_Click()
Call CopyMap(1, 5, 1200)
End Sub

Private Sub Form_Load()
Dim i As Long
Dim tmpStr As String
    PicLarge.BackColor = MaskColor
    Me.Caption = StrUnicode(Me.Caption)
    For i = 0 To Me.Controls.Count - 1
        Call SetCaption(Me.Controls(i))
    Next i
    
    ReadPic = False
    
    Combo1.ListIndex = 0
    Combo1_Click
       

       
    ComboFilename.Clear
    Select Case GetINIStr("run", "style")
    Case "kys"
        For i = 0 To GetINILong("File", "Filenumber") - 1
            ComboFilename.AddItem GetINIStr("File", "File" & i)
        Next i
        tmpStr = GetINIStr("File", "FightName")
        For i = 0 To GetINILong("File", "FightNum") - 1
            ComboFilename.AddItem Replace(tmpStr, "***", Format(i, "000"))
        Next i
    Case "DOS"
        For i = 0 To GetINILong("FileDos", "Filenumber") - 1
            ComboFilename.AddItem GetINIStr("FileDos", "File" & i)
        Next i
        tmpStr = GetINIStr("FileDos", "FightName")
        For i = 0 To GetINILong("FileDos", "FightNum") - 1
            ComboFilename.AddItem Replace(tmpStr, "***", Format(i, "000"))
        Next i
    End Select
    Me.Width = Me.ScaleX(400, vbPixels, vbTwips)
    Me.Height = Me.ScaleY(400, vbPixels, vbTwips)

    MDIMain.StatusBar1.Panels(1).Text = StrUnicode2("鼠标右键激活菜单/拖放到其它窗口复制贴图")
    copyPicnum = -1
    c_Skinner.AttachSkin Me.hwnd
End Sub

' flag =0 重新计算xy
' flag =1 不计算，用于鼠标和水平滚动
Private Sub Showpic(flag As Long)
Dim i As Long
    If ReadPic = True Then
        Pic1.Cls
        If flag = 0 Then Gen_XY
        Draw_pic
    End If
End Sub


' 计算绘大地图显示贴图宽高
Public Sub Gen_XY()
Dim i As Long, j As Long
Dim picnum As Long
Dim tmpHeight As Long
Dim WidthMax As Long
    j = 0
    tmpHeight = 0
    Do             ' 计算每行图片的最大高度以及总共显示行数
        YHeight(j) = 50        ' 初始高度
        For i = 0 To X_num - 1
            picnum = j * X_num + i + Current_startpic
            If picnum >= 0 And picnum < Mappicnum Then
                If YHeight(j) < MapPic(picnum).Height Then
                    YHeight(j) = MapPic(picnum).Height
                End If
            End If
        Next i
        tmpHeight = tmpHeight + YHeight(j)
        If tmpHeight > Pic1.Height Then Exit Do
        j = j + 1
    Loop
    
    Y_num = j + 1
    
    For i = 0 To X_num - 1             ' 计算每列图片最大宽度
        XWidth(i) = Pic1.Width / X_num  ' 初始宽度
        For j = 0 To Y_num - 1
            If picnum >= 0 And picnum < Mappicnum Then
                picnum = j * X_num + i + Current_startpic
                If XWidth(i) < MapPic(picnum).Width Then
                    XWidth(i) = MapPic(picnum).Width
                End If
            End If
        Next j
    Next i
    
    WidthMax = 0
    For i = 0 To X_num - 1
       WidthMax = WidthMax + XWidth(i)
    Next i
    HScroll1.Min = 0
    If WidthMax > Pic1.Width Then
        HScroll1.Max = WidthMax - Pic1.Width
    Else
        HScroll1.Max = 0
    End If
    HScroll1.SmallChange = 1
    HScroll1.LargeChange = 5
    HScroll1.Value = 0
End Sub

Public Sub Draw_pic()
Dim RangeX As Long, rangeY As Long
Dim i As Long, j As Long
Dim i1 As Long, j1 As Long
Dim X1 As Long, Y1 As Long
Dim picnum As Long

    
Dim copmDC As Long
Dim binfo As BITMAPINFO
Dim DIBSectionHandle As Long    ' Handle to DIBSection
Dim OldCompDCBM As Long         ' Original bitmap for CompDC
Dim CompDC As Long
Dim addr As Long
Dim temp As Long
Dim lineSize As Long

    CompDC = CreateCompatibleDC(0)
    With binfo.bmiHeader
        .biSize = 40
        .biWidth = Pic1.Width
        .biHeight = -Pic1.Height
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
        lineSize = .biWidth * 4
        .biSizeImage = -lineSize * .biHeight
    End With
    
    DIBSectionHandle = CreateDIBSection(CompDC, binfo, 0, addr, 0, 0)
    OldCompDCBM = SelectObject(CompDC, DIBSectionHandle)
    
    
    Pic1.BackColor = MaskColor
    temp = BitBlt(CompDC, 0, 0, Pic1.Width, Pic1.Height, Pic1.hdc, 0, 0, &HCC0020)
    

     Y1 = 0
     For j = 0 To Y_num - 1
        X1 = 0
        For i = 0 To X_num - 1
            picnum = j * X_num + i + Current_startpic
            'picnum = j * X_num + i + 20
                If picnum >= 0 And picnum < Mappicnum Then
                    Call genPicData(MapPic(picnum), addr, binfo.bmiHeader.biWidth, -binfo.bmiHeader.biHeight, X1 - HScroll1.Value, Y1 + 10)
                End If
            X1 = X1 + XWidth(i)
        Next i
        Y1 = Y1 + YHeight(j)
    Next j
    
    
    temp = BitBlt(Pic1.hdc, 0, 0, Pic1.Width, Pic1.Height, CompDC, 0, 0, &HCC0020)
   
     Pic1.ForeColor = vbYellow
     Y1 = 0
     For j = 0 To Y_num - 1
        X1 = 0
        For i = 0 To X_num - 1
            'picnum = j * X_num + i + Current_startpic
            picnum = j * X_num + i + Current_startpic
                
            If picnum >= 0 And picnum < Mappicnum Then
                Call genPicData(MapPic(picnum), addr, binfo.bmiHeader.biWidth, -binfo.bmiHeader.biHeight, X1 - HScroll1.Value, Y1 + 10)
                Pic1.CurrentX = X1 - HScroll1.Value
                Pic1.CurrentY = Y1
                Pic1.Print picnum
                If i = current_X And j = current_Y Then
                    Pic1.Line (X1 - HScroll1.Value, Y1)-(X1 - HScroll1.Value + XWidth(i), Y1 + YHeight(j)), vbRed, B
                End If
            End If
            X1 = X1 + XWidth(i)
        Next i
        Y1 = Y1 + YHeight(j)
    Next j
   
       
    temp = GetLastError()
    temp = SelectObject(CompDC, OldCompDCBM)
    temp = DeleteDC(CompDC)
    temp = DeleteObject(DIBSectionHandle)


End Sub


Private Sub Form_Resize()
    On Error Resume Next
    If Me.ScaleWidth < 300 Then
        Me.Width = Me.ScaleX(300, vbPixels, vbTwips)
    End If
    Pic1.Width = Me.ScaleWidth - VScroll1.Width
    HScroll1.Width = Pic1.Width
    VScroll1.Left = Pic1.Width
    
    If Me.ScaleHeight < 400 Then
          Me.Height = Me.ScaleY(400, vbPixels, vbTwips)
    End If
    Pic1.Height = Me.ScaleHeight - HScroll1.Height - Pic1.top
    VScroll1.Height = Pic1.Height
    HScroll1.top = Pic1.top + Pic1.Height
    Showpic 0
      
End Sub

Private Sub HScroll1_Change()
     Showpic (1)
End Sub




Private Sub Pic1_DblClick1()


    
    
    'pic6.Cls
    

    On Error Resume Next
    drawpic5 (Current_startpic + current_Y * X_num + current_X)

End Sub
Public Sub drawpic5(ByVal n As Long)
'Dim n As Long
Dim temp As Long
Dim dib As New clsDIB
    pic6.Cls
    'MsgBox n
    WW = MapPic(n).Width
    HH = MapPic(n).Height
    x = MapPic(n).x
    y = MapPic(n).y
    pic6.Width = MapPic(n).Width + 10
    pic6.Height = MapPic(n).Height + 10
    dib.CreateDIB WW, HH
    pic5.BackColor = MaskColor
    pic6.BackColor = MaskColor
    temp = BitBlt(dib.CompDC, 0, 0, WW, HH, pic6.hdc, 0, 0, &HCC0020)
        
    Call genPicData(MapPic(n), dib.addr, WW, HH, 0, 0)
    ' 复制到dib上
    temp = BitBlt(pic6.hdc, 0, 0, WW, HH, dib.CompDC, 0, 0, &HCC0020)
    
    pic5.Cls
    pic5.PaintPicture pic6.Image, 0, 0, WW * 3, HH * 3
    pic5.Line (3 * x, 3 * (y - 10))-(3 * x, 3 * (y + 10)), vbRed
    pic5.Line (3 * (x - 10), 3 * y)-(3 * (x + 10), 3 * y), vbRed
    
    pic5.Visible = True
End Sub
Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'MsgBox Button
If Button = 2 Then
    PopupMenu MDIMain.mnu_Selectmap
'ElseIf Button = 4 Then
'    Pic1_DblClick
End If
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
Dim j As Long
Dim X1 As Long, Y1 As Long
Dim n As Long
    If Ctime < 1 Then
        Exit Sub
    Else
        Ctime = 0
    End If
    X1 = 0 - HScroll1.Value
    For i = 0 To X_num - 1
        X1 = X1 + XWidth(i)
        If X1 > x Then
            current_X = i
            Exit For
        End If
    Next i
    
    Y1 = 0
    For i = 0 To Y_num - 1
        Y1 = Y1 + YHeight(i)
        If Y1 > y Then
            current_Y = i
            Exit For
        End If
    Next i
    
    
    Showpic (1)
    ' 左按钮按下 则启动拖动
    If (Button And vbLeftButton) > 0 Then
        Pic1.OLEDrag
    End If
    n = Current_startpic + current_Y * X_num + current_X
    If n >= 0 And n < Mappicnum Then
        On Error Resume Next
        MDIMain.StatusBar1.Panels(2).Text = StrUnicode2("贴图" & n & " 宽" & MapPic(n).Width & _
                   "高" & MapPic(n).Height & "X" & MapPic(n).x & "Y" & MapPic(n).y & _
                    IIf(MapPic(n).isEmpty, "空贴图", ""))
    End If
End Sub

Private Sub Pic1_OLEStartDrag(data As DataObject, AllowedEffects As Long)
Dim n As Long
    n = Current_startpic + current_Y * X_num + current_X
    If n >= 0 And n < Mappicnum Then
        data.SetData txtGRP.Text & ":" & CStr(n), vbCFText
        AllowedEffects = vbDropEffectCopy
    End If
    
End Sub



Private Sub pic5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
    If Pic5num = 0 Then Exit Sub
    Pic5num = Pic5num - 1
    drawpic5 (Pic5num)
ElseIf KeyCode = vbKeyRight Then
    If Pic5num = Mappicnum - 1 Then Exit Sub
    Pic5num = Pic5num + 1
    drawpic5 (Pic5num)
End If
End Sub

Private Sub pic5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    pic5.Cls
    'If X >= 255 Then X = 255
    'If Y >= 255 Then Y = 255
    pic5.PaintPicture pic6.Image, 0, 0, WW * 3, HH * 3
    pic5.Line (x, y - 30)-(x, y + 30), vbRed
    pic5.Line (x - 30, y)-(x + 30, y), vbRed
    MapPic(n).x = Int(x / 3)
    MapPic(n).y = Int(y / 3)
ElseIf Button = 2 Then
    pic5.Visible = False
End If
End Sub

Private Sub Timer1_Timer()
Ctime = Ctime + 1
End Sub

Private Sub VScroll1_Change()
    Current_startpic = VScroll1.Value * X_num + real_Current_startpic
    Showpic (0)
End Sub



Public Sub ClickMenu(id As String)
'Dim n As Long
Dim i, j, k, l As Long
Dim temp As Long
Dim dib As New clsDIB
Dim Num1, Num2 As Long
Dim r2, g2, b2 As Long

    n = Current_startpic + current_Y * X_num + current_X
    Select Case LCase(id)
    'MsgBox id
    Case "edit"
        If Mappicnum <= 0 Then Exit Sub
        Load frmPicEdit
        'MsgBox n
        g_PP = MapPic(n)
        frmPicEdit.Showpic
        frmPicEdit.Show vbModal
        If frmPicEdit.YES = 1 Then
            MapPic(n) = g_PP
            Showpic (0)
        End If
    Case "switch"
        If Mappicnum <= 0 Then Exit Sub
        g_PP = MapPic(n)
        frmswitchcolor.Show vbModal
        If frmswitchcolor.YES = 1 Then
            'If MsgBox(StrUnicode2("是否修改全部图片？"), vbQuestion + vbOKCancel) = vbOK Then
            '    Num1 = 0
            '    Num2 = Mappicnum - 1
            'Else
            '    Num1 = n
            '    Num2 = n
            'End If
            
            'With frmswitchcolor
                If frmswitchcolor.Pflag = 1 Then
                    Num1 = 0
                    Num2 = Mappicnum - 1
                ElseIf frmswitchcolor.Pflag = 2 Then
                    Num1 = n
                    Num2 = n
                ElseIf frmswitchcolor.Pflag = 3 Then
                    Num1 = Val(frmswitchcolor.PBegin)
                    Num2 = Val(frmswitchcolor.PEnd)
                End If
            'End With
            
            For k = Num1 To Num2
                WW = MapPic(k).Width
                HH = MapPic(k).Height
                If WW > 0 And HH > 0 Then
                    ReDim data(WW - 1, HH - 1)
       '
                    Picbak.Cls
                    Picbak.Width = WW
                    Picbak.Height = HH
                
                    x = MapPic(k).x
                    y = MapPic(k).y
                    dib.CreateDIB WW, HH
                    Picbak.BackColor = MaskColor
                    temp = BitBlt(dib.CompDC, 0, 0, WW, HH, Picbak.hdc, 0, 0, &HCC0020)
       '
                    Call genPicData(MapPic(k), dib.addr, WW, HH, 0, 0)
                    ' 复制到dib上
                    temp = BitBlt(Picbak.hdc, 0, 0, WW, HH, dib.CompDC, 0, 0, &HCC0020)
        
                    For j = 0 To HH - 1
                        For i = 0 To WW - 1
                            If GetPixel(Picbak.hdc, i, j) <> MaskColor Then
                                For l = 0 To 9
                                    If GetPixel(Picbak.hdc, i, j) = colorA(l) Then
                                        data(i, j) = colorB(l)
                                        Exit For
                                    Else
                                        data(i, j) = GetPixel(Picbak.hdc, i, j)
                                    End If
                                Next l
                            Else
                               data(i, j) = GetPixel(Picbak.hdc, i, j)
                            End If
                        Next i
                    Next j
                    Call ReleaseDC(Me.hwnd, Picbak.hdc)
                    SaveMapPic
                    MapPic(k) = g_PP
                    
                'For i = 0 To MapPic(k).DataLong - 1
                '    For j = 0 To 9
                '        If MapPic(k).Data32(i) = colorA(j) Then
                '            MapPic(k).Data32(i) = mcolor_RGB(get256(colorB(j)))
                '            'MapPic(k).data(i) = get256(colorB(j))
                '            Exit For
                '        End If
                '    Next j
                'Next i
                
                   DoEvents
                End If
            Next k
            Showpic (0)
        End If
    Case "copy"
        copyPicnum = n
    Case "paste"
        PastePic
    Case "add"
        Mappicnum = Mappicnum + 1
        ReDim Preserve MapPic(Mappicnum - 1)
        Showpic (1)
        MapPic(Mappicnum - 1).isEmpty = True
    Case "delete"
        If Mappicnum <= 0 Then Exit Sub
        Mappicnum = Mappicnum - 1
        ReDim Preserve MapPic(Mappicnum - 1)
        Showpic (1)
    Case "save"
        SavePic
    Case "insert"
        If (MsgBox(StrUnicode2("慎用于smp,mmp,wmp，继续？"), vbQuestion + vbOKCancel, "Confirm") = vbOK) Then
            'add
            Mappicnum = Mappicnum + 1
            ReDim Preserve MapPic(Mappicnum - 1)
            MapPic(Mappicnum - 1).isEmpty = True
            'paste
            For i = (Mappicnum - 1) To (n + 1) Step -1
                MapPic(i) = MapPic(i - 1)
            Next i
            'clear pic
            MapPic(n).DataLong = 0
            MapPic(n).Height = 0
            MapPic(n).isEmpty = True
            MapPic(n).Width = 0
            MapPic(n).x = 0
            MapPic(n).y = 0
            Showpic (1)
        End If
    Case "x2"
        Xjapan2
    End Select
End Sub

Private Sub PastePic()
Dim i As Long
Dim n As Long
    If copyPicnum < 0 Or copyPicnum >= Mappicnum Then Exit Sub
    n = Current_startpic + current_Y * X_num + current_X
    If MapPic(copyPicnum).isEmpty = False Then
        MapPic(n) = MapPic(copyPicnum)
    End If
    Showpic (1)
End Sub


' 保存贴图
Private Sub SavePic()
Dim fileid As String
Dim filepic As String
Dim filenumID As Long, filenumPic As Long
Dim i As Long
Dim offset As Long
Dim Idx() As Long
    If txtIDX.Text = "" Or txtGRP = "" Then Exit Sub
    fileid = G_Var.JYPath & txtIDX
    filepic = G_Var.JYPath & txtGRP
    
    ReDim Idx(Mappicnum)
    filenumPic = OpenBin(filepic, "WN")
    
    For i = 0 To Mappicnum - 1
        Idx(i + 1) = 0
        If MapPic(i).isEmpty = False Then
            Put #filenumPic, , MapPic(i).Width
            Put #filenumPic, , MapPic(i).Height
            Put #filenumPic, , MapPic(i).x
            Put #filenumPic, , MapPic(i).y
            Put #filenumPic, , MapPic(i).data
        End If
        offset = Loc(filenumPic)
        Idx(i + 1) = offset
    Next i
    Close #filenumPic
    
    '  处理空贴图id指针指向下一个贴图。
    For i = Mappicnum To 1 Step -1
        If Idx(i) = 0 Then
            Idx(i) = Idx(i + 1)
        End If
    Next i
    
    filenumID = OpenBin(fileid, "WN")
        For i = 1 To Mappicnum
            Put #filenumID, , Idx(i)
        Next i
    Close #filenumID
End Sub
Private Sub SaveMapPic()
Dim i As Long, j As Long
Dim k As Long
Dim tmpbyte(2000) As Byte
Dim num As Long
Dim maskNum As Long
Dim solidNum As Long
Dim status As Long
Dim p As Long

    Call convertCOLOR2(mcolor_RGB(0), data(0, 0), WW, HH, Val(MaskColor))
    g_PP.Width = WW
    g_PP.Height = HH
    If WW = 0 And HH = 0 Then
        g_PP.isEmpty = True
        Exit Sub
    End If
    
    g_PP.x = x
    g_PP.y = y
    ReDim g_PP.data(0)
    p = 0
    For j = 0 To HH - 1
        num = 0
        i = 0
        Do
            maskNum = 0
            Do
               If data(i, j) <> MaskColor Then Exit Do
               i = i + 1
               maskNum = maskNum + 1
               If i >= WW Then Exit Do
            Loop
            If i >= WW Then
                Exit Do
            End If
            solidNum = 0
            tmpbyte(num) = maskNum
            Do
                If data(i, j) = MaskColor Then Exit Do
                If i >= WW Then Exit Do
                tmpbyte(num + 2 + solidNum) = get256(mcolor_RGB(0), data(i, j))
                i = i + 1
                solidNum = solidNum + 1
                If i >= WW Then Exit Do
                
            Loop
            tmpbyte(num + 1) = solidNum
            num = num + solidNum + 2
            If i >= WW Then Exit Do
        Loop
        ReDim Preserve g_PP.data(p + num)
        g_PP.data(p) = num
        For i = 0 To num - 1
            g_PP.data(p + i + 1) = tmpbyte(i)
        Next i
        p = p + num + 1
    Next j
    g_PP.DataLong = p
    Call RLEto32(g_PP)
End Sub
Private Sub cmdConvert_Click()
Dim i As Long, j As Long
Dim c As Long
Dim rr As Long, gg As Long, bb As Long
Dim yy As Long, uu As Long, vv As Long
Dim rc(255) As Long, gc(255) As Long, bc(255) As Long
Dim yc(255) As Long, uc(255) As Long, vc(255) As Long
Dim vmin As Long, v As Long
Dim nn As Long
'Dim data() As Long
    For i = 0 To 255
        rc(i) = (mcolor_RGB(i) \ 65536) And &HFF
        gc(i) = (mcolor_RGB(i) \ 256) And &HFF
        bc(i) = mcolor_RGB(i) And &HFF
        yc(i) = 0.299 * rc(i) + 0.587 * gc(i) + 0.114 * bc(i)
        uc(i) = -0.1687 * rc(i) - 0.3313 * gc(i) + 0.5 * bc(i) + 128
        vc(i) = 0.5 * rc(i) - 0.4187 * gc(i) - 0.0813 * bc(i) + 128
    Next i
    
    For i = 0 To WW - 1
        For j = 0 To HH - 1
            If data(i, j) <> MaskColor Then
                vmin = 100000#
                rr = data(i, j) And &HFF
                gg = (data(i, j) \ 256) And &HFF
                bb = (data(i, j) \ 65536) And &HFF
                yy = 0.299 * rr + 0.587 * gg + 0.114 * bb
                uu = -0.1687 * rr - 0.3313 * gg + 0.5 * bb + 128
                vv = 0.5 * rr - 0.4187 * gg - 0.0813 * bb + 128
                
                For c = 0 To 255
                    v = 2 * (yy - yc(c)) ^ 2 + (uu - uc(c)) ^ 2 + (vv - vc(c)) ^ 2
                    If v < vmin Then
                        vmin = v
                        nn = c
                    End If
                Next c
                data(i, j) = RGB(rc(nn), gc(nn), bc(nn))
            End If
        Next j
    Next i
    'ShowData
End Sub
Private Function get2562(d As Long) As Byte
Dim i As Long
Dim rr As Long, gg As Long, bb As Long
Dim r2 As Long, g2 As Long, b2 As Long
            
    b2 = (d \ 65536) And &HFF&
    g2 = (d \ 256) And &HFF
    r2 = d And &HFF
    For i = 0 To 255
        rr = (mcolor_RGB(i) \ 65536) And &HFF&
        gg = (mcolor_RGB(i) \ 256) And &HFF
        bb = mcolor_RGB(i) And &HFF
        If r2 = rr And g2 = gg And b2 = bb Then
            get2562 = i
            Exit For
        End If
    Next i
End Function
Private Sub CopyMap(ByVal X1 As Long, ByVal X2 As Long, ByVal Y1 As Long)
Dim i As Long, j As Long, tmp As Long, Y2 As Long
    Y2 = X2 - X1 + Y1
    'MsgBox y2
    If Y2 > Mappicnum Then
            'add
            ReDim Preserve MapPic(Y2)
            For i = Mappicnum To Y1
                MapPic(i).DataLong = 0
                MapPic(i).Height = 0
                MapPic(i).isEmpty = True
                MapPic(i).Width = 0
                MapPic(i).x = 0
                MapPic(i).y = 0
            Next i
            
            Mappicnum = Y2
            
    End If
            For i = X1 To X2
                MapPic(i).isEmpty = True
                    'paste
                MapPic(i - X1 + Y1) = MapPic(i)
            Next i
    Showpic (1)
End Sub
