Attribute VB_Name = "Modmain"
Option Explicit

'Public RangerNum As Long
Public Recordnum As Long                   ' 进度编号

Public Const XSCALE = 18
Public Const YSCALE = 9

Public First As Boolean

Public c_Skinner As New CSkinner

Public Const MaskColor = &H707030

Public colorA(9) As Long
Public colorB(9) As Long

'这个修改器中指令个数
Public Const KdefNum = &H48


Public Type BITMAPINFOHEADER '4? bytes
        biSize As Long

        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER

        bmiColors As RGBQUAD    ' RGB, so length here doesn't matter
End Type






Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo _
As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) _
As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) _
As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y _
As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As _
Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Public Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long


Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, addr As Byte, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Public Type SafeArrayBound
    Elements            As Long
    lLbound             As Long
End Type

Public Type SAFEARRAY2D
    Dimension           As Integer
    Features            As Integer
    Element             As Long
    Locks               As Long
    Pointer             As Long
    Bounds(1)           As SafeArrayBound
End Type

Public Type SAFEARRAY
    Dimension           As Integer
    Features            As Integer
    Element             As Long
    Locks               As Long
    Pointer             As Long
    Bounds              As SafeArrayBound
End Type

Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long



Public Type WarDataType   ' 战斗场景数据   这些数据无法放入窗体，不支持 public type
    id As Integer
    namebig5(9) As Byte
    Name As String
    mapid As Integer
    Experience As Integer
    musicid As Integer
    Warperson(5) As Integer
    SelectWarperson(5) As Integer
    personX(5) As Integer
    personY(5) As Integer
    Enemy(19) As Integer
    EnemyX(19) As Integer
    EnemyY(19) As Integer
End Type

Public WarData() As WarDataType
Public warnum As Long             ' 战斗个数


Public Type statementAttribType  ' 指令属性
    Length As Long               ' 指令长度
    isGoto As Long               ' 指令是否为条件转移，0不是 1是
    yesOffset As Long            ' 条件满足转移地址在指令中第几个字
    noOffset As Long             ' 条件不满足转移地址在指令中第几个字
    notes As String              ' 指令说明
End Type

Public StatAttrib() As statementAttribType

Public KGidxR() As Long
Public KGidx() As Long
Public pngClass As New LoadPNG

Public kdefword() As Integer   ' 保存kdef事件二进制数据
Public kdeflong As Long        ' 事件文件长度

Public KDEFIDX() As Long       ' 事件kdef索引
Public numkdef As Long         ' 保存事件个数

Public Type KdefType           ' 保存事件数据
    DataLong  As Long          ' 事件二进制数据长度
    data() As Integer          ' 事件二进制数据
    kdef As Collection         ' 事件指令集合
    numlabel As Long           ' 事件中标号数目
End Type

Public KdefInfo() As KdefType

Public KdefName As Collection       ' 指令定义的名字集合


Public ClipboardStatement As Collection  ' 复制下的指令数据

Public ClipboardKdef As Collection  ' 复制下的指令数据

Public nameidx() As Long
Public nam() As String
Public numname As Long

Public Talk() As String        ' 对话字符串
Public TalkIdx() As Long           ' 对话索引
Public numtalk As Long      ' 对话个数


Public Type PersonAttrib       ' 人物属性
    r1 As Integer
    PhotoId As Integer
    r3 As Integer
    r4 As Integer
    name1big5(9) As Byte
    Name1 As String
    name2big5(9) As Byte
    name2 As String
    
End Type

Public Person() As PersonAttrib  ' 人物属性数组
Public PersonNum As Long         ' 人物个数

Public Type ThingsAttrib         ' 物品属性
    r1 As Integer
    name1big5(19) As Byte
    Name1 As String
    name2big5(19) As Byte
    name2 As String
End Type

Public Things() As ThingsAttrib  ' 物品属性数组
Public Thingsnum As Long         ' 物品个数



Public Type SceneType
    SceneID As Integer                ' 场景代号
    Name1(9) As Byte                  ' 名字
    OutMusic As Integer               ' 出门音乐
    InMusic  As Integer               ' 进门音乐
    JumpScene As Integer              ' 跳转场景
    InCondition As Integer            ' 进入条件
    MMapInX1 As Integer               ' 主地图入口坐标
    MMapInY1 As Integer
    MMapInX2 As Integer
    MMapInY2 As Integer
    InX As Integer                    ' 进入后初始坐标
    InY As Integer
    OutX(2) As Integer                ' 三个出口坐标
    OutY(2) As Integer
    JumpX1 As Integer               ' 两个跳转口坐标
    JumpY1 As Integer
    JumpX2 As Integer               ' 两个跳转口坐标
    JumpY2 As Integer
End Type

Public Scene() As SceneType  ' 场景属性数组
Public Scenenum As Long          ' 场景个数

Public Type WuGongAttrib           ' 武功属性
    r1 As Integer
    name1big5(19) As Byte
    Name1 As String
End Type


Public WuGong() As WuGongAttrib    ' 武功属性数组
Public WuGongnum As Long           ' 武功个数


Public Type RLEPic
    isEmpty As Boolean
    Width As Integer   ' 图片宽度
    height As Integer  ' 图片高度
    x As Integer       ' 图片x偏移
    y As Integer       ' 图片y偏移
    DataLong As Long   ' 图片RLE压缩数据长度
    data() As Byte     ' 图片RLE压缩数据
    Data32() As Long   ' 图片32位压缩数据
End Type

Public HeadPic() As RLEPic  ' 人物头像数据
Public PngPic() As RLEPic  ' pic头像数据
Public Headnum As Long      ' 人物头像个数
Public NewHeadNum As Long

Public WarPic() As RLEPic  ' 战斗图片数据
Public Warpicnum As Long   ' 战斗个数

Public g_PP As RLEPic     ' 编辑贴图用，传递参数。


' D* 事件信息

Public Type D_Event_type
    isGo As Integer
    id As Integer
    EventNum1 As Integer
    EventNum2 As Integer
    EventNum3 As Integer
    picnum(2) As Integer
    PicDelay As Integer
    x As Integer
    y As Integer
End Type

Public g_DD As D_Event_type          ' 修改场景事件定义窗体使用，用来传递参数。



Public HeadtoPerson() As Collection   ' 根据头像id查人物id

Public mcolor_RGB(256) As Long  ' 颜色表


Type G_VarType
    JYPath As String
    iniFileName As String
    Palette As String
    MMAPIDX As String
    MMAPGRP As String
    MMAPStruct(5) As String
    SMAPIDX As String
    SMAPGRP As String
    SMAPIDX2 As String
    SMAPGRP2 As String
    WarMapIDX As String
    WarMapGrp As String
    WarDefine As String
    WarMapDefIDX As String
    WarMapDefGRP As String
    TalkIdx As String
    TalkGRP As String
    RIDX(6) As String
    RGRP(7) As String
    DIDX(7) As String
    DGRP(7) As String
    SIDX(7) As String
    SGRP(7) As String
    RecordNote(7) As String
    EXE As String
    KDEFIDX As String
    KDEFGRP As String
    HeadIDX As String
    HeadGRP As String
    Leave As String
    Effect As String
    Match As String
    Namegrp As String
    nameidx As String
    Exp As String
    NewHeadIDX As String
    NewHeadGRP As String
    EditMode As String
    SceneMap As String
    title As String
    Dead As String
End Type

'Public Team(0) As String
Public G_Var As G_VarType

'kgOffset
Public KGoffset() As Long
Public KGoffsetNum As Long

Public FiftyItem() As String
Public Charset As String

'Public IniFilename As String


Public Sub Main()
Dim tmpstrArray()  As String
Dim i As Long

    First = True
On Error GoTo Label1
    'MsgBox Command
    If Command = "" Then
        Charset = "GBK"
        
    ' Call LoadMB
    
        G_Var.iniFileName = App.Path & "\fishedit.ini"
    
        'ConvertBig5INI
        
        G_Var.JYPath = ""
        Charset = GetINIStr("run", "charset")
        If Charset = "" Then
            frmSelectCharset.Show vbModal
        End If
        Select Case GetINIStr("run", "style")
        Case "kys"
            G_Var.Palette = GetINIStr("File", "Palette")
            G_Var.EditMode = GetINIStr("Run", "Mode")
            G_Var.MMAPIDX = GetINIStr("File", "MMAPIDX")
            G_Var.MMAPGRP = GetINIStr("File", "MMAPGRP")
            tmpstrArray = Split(GetINIStr("File", "MMAPStruct"), ",")
            For i = 0 To 4
                G_Var.MMAPStruct(i) = tmpstrArray(i)
            Next i
    
            G_Var.SMAPIDX = GetINIStr("File", "SMAPIDX")
            G_Var.SMAPGRP = GetINIStr("File", "SMAPGRP")
            G_Var.SMAPIDX2 = GetINIStr("File", "SMAPIDX2")
            G_Var.SMAPGRP2 = GetINIStr("File", "SMAPGRP2")
            G_Var.WarMapIDX = GetINIStr("File", "WarMAPIDX")
            G_Var.WarMapGrp = GetINIStr("File", "WarMAPGRP")
            G_Var.WarDefine = GetINIStr("File", "WarDefine")
            G_Var.WarMapDefIDX = GetINIStr("File", "WarMAPDefIDX")
            G_Var.WarMapDefGRP = GetINIStr("File", "WarMAPDefGRP")
    
    
            G_Var.TalkIdx = GetINIStr("File", "TalkIDX")
            G_Var.TalkGRP = GetINIStr("File", "TalkGRP")
    
            G_Var.KDEFIDX = GetINIStr("File", "kdefIDX")
            G_Var.KDEFGRP = GetINIStr("File", "kdefGRP")
    
            If G_Var.EditMode = "classic" Then
                G_Var.HeadIDX = GetINIStr("File", "HeadIDX")
                G_Var.HeadGRP = GetINIStr("File", "HeadGRP")
            Else
                G_Var.NewHeadGRP = GetINIStr("File", "NewHeadGRP")
                G_Var.NewHeadIDX = GetINIStr("File", "NewHeadIDX")
            End If
    
            G_Var.Leave = GetINIStr("File", "Leave")
            G_Var.Effect = GetINIStr("File", "Effect")
            G_Var.Match = GetINIStr("File", "Match")
            G_Var.Exp = GetINIStr("File", "Exp")

            tmpstrArray = Split(GetINIStr("File", "RIDX"), ",")
            For i = 0 To 6
            '        MsgBox i & " " & G_Var.RIDX(i - 1)
                G_Var.RIDX(i) = tmpstrArray(i)
            Next i
    
            tmpstrArray = Split(GetINIStr("File", "RGRP"), ",")
            For i = 0 To 6
                G_Var.RGRP(i) = tmpstrArray(i)
            Next i
'    tmpstrArray = Split(GetINIStr("File", "DIDX"), ",")
'    For i = 0 To 4
'        G_Var.DIDX(i) = tmpstrArray(i)
'    Next i
            tmpstrArray = Split(GetINIStr("File", "DGRP"), ",")
            For i = 0 To 6
                G_Var.DGRP(i) = tmpstrArray(i)
            Next i
'    tmpstrArray = Split(GetINIStr("File", "SIDX"), ",")
'    For i = 0 To 4
'        G_Var.SIDX(i) = tmpstrArray(i)
'    Next i
            tmpstrArray = Split(GetINIStr("File", "SGRP"), ",")
            For i = 0 To 6
                G_Var.SGRP(i) = tmpstrArray(i)
            Next i
    
            tmpstrArray = Split(GetINIStr("File", "RecordNote"), ",")
            For i = 0 To 6
                G_Var.RecordNote(i) = tmpstrArray(i)
            Next i
    
            G_Var.Namegrp = GetINIStr("File", "NameGRP")
            G_Var.nameidx = GetINIStr("File", "NameIDX")
            
            G_Var.SceneMap = GetINIStr("File", "SceneMap")
            
            G_Var.EXE = GetINIStr("File", "EXEFilename")
            
            If Charset = "BIG5" Then
                G_Var.EXE = StrUnicode2(G_Var.EXE)
            End If
            MDIMain.mnu_z.Visible = False
        
            Case "DOS"
            G_Var.Palette = GetINIStr("FileDOS", "Palette")
            G_Var.EditMode = GetINIStr("Run", "Mode")
            G_Var.MMAPIDX = GetINIStr("FileDOS", "MMAPIDX")
            G_Var.MMAPGRP = GetINIStr("FileDOS", "MMAPGRP")
            tmpstrArray = Split(GetINIStr("FileDOS", "MMAPStruct"), ",")
            For i = 0 To 4
                G_Var.MMAPStruct(i) = tmpstrArray(i)
            Next i
            
            G_Var.SMAPIDX = GetINIStr("FileDOS", "SMAPIDX")
            G_Var.SMAPGRP = GetINIStr("FileDOS", "SMAPGRP")
            G_Var.title = GetINIStr("FileDOS", "TITLE")
            G_Var.Dead = GetINIStr("FileDOS", "DEAD")
            G_Var.WarMapIDX = GetINIStr("FileDOS", "WarMAPIDX")
            G_Var.WarMapGrp = GetINIStr("FileDOS", "WarMAPGRP")
            G_Var.WarDefine = GetINIStr("FileDOS", "WarDefine")
            G_Var.WarMapDefIDX = GetINIStr("FileDOS", "WarMAPDefIDX")
            G_Var.WarMapDefGRP = GetINIStr("FileDOS", "WarMAPDefGRP")
    
    
            G_Var.TalkIdx = GetINIStr("FileDOS", "TalkIDX")
            G_Var.TalkGRP = GetINIStr("FileDOS", "TalkGRP")
    
            G_Var.KDEFIDX = GetINIStr("FileDOS", "kdefIDX")
            G_Var.KDEFGRP = GetINIStr("FileDOS", "kdefGRP")
    
            If G_Var.EditMode = "classic" Then
                G_Var.HeadIDX = GetINIStr("FileDOS", "HeadIDX")
                G_Var.HeadGRP = GetINIStr("FileDOS", "HeadGRP")
            Else
                G_Var.NewHeadGRP = GetINIStr("FileDOS", "NewHeadGRP")
                G_Var.NewHeadIDX = GetINIStr("FileDOS", "NewHeadIDX")
            End If
            
            'G_Var.Leave = GetINIStr("File", "Leave")
            'G_Var.Effect = GetINIStr("File", "Effect")
            'G_Var.Match = GetINIStr("File", "Match")
            'G_Var.Exp = GetINIStr("File", "Exp")

            tmpstrArray = Split(GetINIStr("FileDOS", "RIDX"), ",")
            For i = 0 To 3
            '        MsgBox i & " " & G_Var.RIDX(i - 1)
                G_Var.RIDX(i) = tmpstrArray(i)
            Next i
    
            tmpstrArray = Split(GetINIStr("FileDOS", "RGRP"), ",")
            For i = 0 To 3
                G_Var.RGRP(i) = tmpstrArray(i)
            Next i

            tmpstrArray = Split(GetINIStr("FileDOS", "DIDX"), ",")
            For i = 0 To 3
                G_Var.DIDX(i) = tmpstrArray(i)
            Next i
            
            tmpstrArray = Split(GetINIStr("FileDOS", "DGRP"), ",")
            For i = 0 To 3
                G_Var.DGRP(i) = tmpstrArray(i)
            Next i
            
            tmpstrArray = Split(GetINIStr("FileDOS", "SIDX"), ",")
            For i = 0 To 3
                G_Var.SIDX(i) = tmpstrArray(i)
            Next i
            tmpstrArray = Split(GetINIStr("FileDOS", "SGRP"), ",")
            For i = 0 To 3
                G_Var.SGRP(i) = tmpstrArray(i)
            Next i
    
            tmpstrArray = Split(GetINIStr("FileDOS", "RecordNote"), ",")
            For i = 0 To 3
                G_Var.RecordNote(i) = tmpstrArray(i)
            Next i
    
            'G_Var.Namegrp = GetINIStr("File", "NameGRP")
            'G_Var.nameidx = GetINIStr("File", "NameIDX")

    
            G_Var.EXE = GetINIStr("File", "EXEFilename")
            
            If Charset = "BIG5" Then
                G_Var.EXE = StrUnicode2(G_Var.EXE)
            End If
            
            MDIMain.mnu_Team.Enabled = False
        End Select
    
    Dim bArr() As Byte
   ' bArr = LoadResData("red", "skin")
   ' SkinH_AttachRes bArr(0), UBound(bArr) + 1, "", 0, 0, 0
   ' SkinH_SetAero 1
   ' SkinH_AttachEx App.Path & "\红动中国.she", ""
        MDIMain.Show
    Else
        'MainSelectMap.Show
    End If
Exit Sub

Label1:
    MsgBox Err.Description
    If (MDIMain Is Nothing) = False Then
        Unload MDIMain
    End If

End Sub

' 读取r1文件
Public Sub ReadRR(Rnum As Long)
Dim idnum As Long
Dim filenum As Long
Dim idxr() As Long
Dim i As Long, j As Long
Dim Rlong() As Long, NameOFFset() As Long
Dim offset As Long
'Dim length As Long
'Dim result As Long
'Dim i, j As Long
Dim kk
ReDim Rlong(GetINILong("R_Modify", "TypeNumber") - 1)
ReDim NameOFFset(GetINILong("R_Modify", "TypeNumber") - 1)
For j = 0 To GetINILong("R_Modify", "TypeNumber") - 1
    For i = 0 To GetINILong("R_Modify", "TypedataItem" & j) - 1
        kk = Split(GetINIStr("R_Modify", "data(" & j & "," & i & ")"), " ")
        If Val(kk(4)) = 1 Then
            NameOFFset(j) = Rlong(j)
        End If
        Rlong(j) = Rlong(j) + Val(kk(2)) * Val(kk(0)) * Val(kk(1))

    Next i
Next j

    filenum = OpenBin(G_Var.JYPath & G_Var.RIDX(Rnum), "R")
   ' MsgBox G_Var.JYPath & G_Var.RIDX(1)
    idnum = LOF(filenum) / 4
    ReDim idxr(idnum)
    idxr(0) = 0
    For i = 1 To idnum
       Get filenum, , idxr(i)
    Next i
    Close (filenum)
    
    PersonNum = (idxr(2) - idxr(1)) / Rlong(1)
    ReDim Person(PersonNum - 1)
    filenum = OpenBin(G_Var.JYPath & G_Var.RGRP(Rnum), "R")
    offset = idxr(1)
    For i = 0 To PersonNum - 1
        Get filenum, offset + i * Rlong(1) + 1, Person(i).r1
        Get filenum, , Person(i).PhotoId
        Get filenum, , Person(i).r3
        Get filenum, , Person(i).r4
        Get filenum, offset + i * Rlong(1) + NameOFFset(1) + 1, Person(i).name1big5
        'Get filenum, , Person(i).name2big5
        Person(i).Name1 = Big5toUnicode(Person(i).name1big5, 10)
        'Person(i).name2 = Big5toUnicode(Person(i).name2big5, 10)
    Next i
    
    
    Thingsnum = (idxr(3) - idxr(2)) / Rlong(2)
    ReDim Things(Thingsnum - 1)
    offset = idxr(2)
    For i = 0 To Thingsnum - 1
        Get filenum, offset + i * Rlong(2) + 1, Things(i).r1
        Get filenum, offset + i * Rlong(2) + NameOFFset(2) + 1, Things(i).name1big5
        'Get filenum, , Things(i).name1big5
        Things(i).Name1 = Big5toUnicode(Things(i).name1big5, 20)
        'Things(i).name2 = Big5toUnicode(Things(i).name2big5, 20)
        
    Next i
 
    Scenenum = (idxr(4) - idxr(3)) / Rlong(3)
    ReDim Scene(Scenenum - 1)
    offset = idxr(3)
       
    Get filenum, offset + 1, Scene


'    WuGongnum = (idxr(5) - idxr(4)) / Rlong(4)
'    ReDim WuGong(WuGongnum - 1)
 '   offset = idxr(4)
 '   For i = 0 To WuGongnum - 1
 '       Get filenum, offset + i * Rlong(4) + 1, WuGong(i).r1
 '       Get filenum, offset + i * Rlong(4) + NameOFFset(4) + 1, WuGong(i).name1big5
 '       WuGong(i).Name1 = Big5toUnicode(WuGong(i).name1big5, 20)
 '
 '   Next i
 
    Close (filenum)
End Sub

' 读取人物照片并转化为32位rle

Public Sub LoadPicFile(fileid As String, filepic As String, picdata() As RLEPic, picdatanum As Long)

Dim filenum As Integer, filenum2 As Integer
Dim i As Long
Dim Value As Integer
Dim rr As Integer, gg As Integer, bb As Integer
Dim offset As Long
Dim picLong As Long
Dim num As Long
Dim xx As Long, yy As Long

Dim picNum2
Dim HeadIDX() As Long
    If Val(fileid) <> -2 Then
        filenum = OpenBin(fileid, "R")
        picdatanum = LOF(filenum) / 4
        ReDim HeadIDX(picdatanum)
        ReDim picdata(picdatanum)
        HeadIDX(0) = 0
        For num = 1 To picdatanum ' 地图贴图的索引个数
            Get filenum, , HeadIDX(num)
        Next num
        Close filenum
   Else
        picdatanum = FileLen(filepic) / (64 * 64 * 12)
        MsgBox picdatanum
   End If
    
    filenum = OpenBin(filepic, "R")
    For num = 0 To picdatanum - 1 ' 地图贴图的索引个数
        If HeadIDX(num) < 0 Then
            picLong = 0
        Else
            offset = HeadIDX(num)
            picLong = HeadIDX(num + 1) - HeadIDX(num)
            If (num = picdatanum - 1) And (HeadIDX(num + 1) <> LOF(filenum)) And HeadIDX(num) > 0 Then ' 最后一个idx应该等于文件长度
                picLong = LOF(filenum) - HeadIDX(num)
            End If
        End If
        If picLong > 0 Then
            picdata(num).isEmpty = False
            Get filenum, offset + 1, picdata(num).Width
            Get filenum, , picdata(num).height
            Get filenum, , picdata(num).x
            Get filenum, , picdata(num).y
            picdata(num).DataLong = picLong - 8
            ReDim picdata(num).data(picdata(num).DataLong - 1)
            Get filenum, , picdata(num).data
            Call RLEto32(picdata(num))
        Else
            picdata(num).isEmpty = True
        End If
    Next num
    Close filenum

End Sub




' 把贴图数据的8BitRLE压缩数据，转换为32Bit，方便以后处理
' RLE压缩格式：
' 第一个字节为第一行数据长度（几个字节）
' 后面一个字节为透明数据点个数，后面跟着为不透明数据点个数，然后是不透明的每个数据点8位颜色，
' 重复以上，直到第一行字节结束
' 读取下一行数据，直到没有后面数据
Public Sub RLEto32(pic As RLEPic)
Dim p As Long  ' 指向RLE数据的指针
Dim p32 As Long   ' 指向32位非压缩数据的指针
Dim i As Long, j As Long
Dim row As Byte
Dim temp As Long
Dim Start As Long
Dim maskNum As Long
Dim solidNum As Long
   
    ReDim pic.Data32(pic.DataLong)
   
    p = 0
    p32 = 0
    For i = 1 To pic.height
        row = pic.data(p)     ' 当前行数据个数
        pic.Data32(p) = row
        Start = p             ' 当前行起始位置
        p = p + 1
        If row > 0 Then
            p32 = 0
            Do
                maskNum = pic.data(p)  ' 掩码个数
                pic.Data32(p) = row
                p = p + 1
      
                p32 = p32 + maskNum
                If p32 >= pic.Width Then  ' 此掩码完成后位置指针已经指向最右端
                    Exit Do
                End If
                solidNum = pic.data(p) ' 实际点的个数
                pic.Data32(p) = solidNum
                p = p + 1
                For j = 1 To solidNum
                    temp = pic.data(p)
                    pic.Data32(p) = mcolor_RGB(temp)
                    p32 = p32 + 1
                    p = p + 1
                Next j
                If p32 >= pic.Width Then   ' 实际点完成后位置指针已经指向最右端
                    Exit Do
                End If
                If p - Start >= row Then           ' 当前行数据已经结束
                    Exit Do
                End If
            Loop
            If p + 1 >= pic.DataLong Then
                Exit For
            End If
        End If
    Next i
   
End Sub



' 读取颜色表数据
' jinyong中颜色表是按照256色，每色rgb各一个字节
Public Sub SetColor()
Dim filenum As Integer
Dim i As Long
Dim rr As Byte, gg As Byte, bb As Byte
    
    'filenum = FreeFile()
    filenum = OpenBin(G_Var.JYPath & G_Var.Palette, "R")
        For i = 0 To 255
            Get filenum, , rr
            Get filenum, , gg
            Get filenum, , bb
            rr = rr * 4           ' 颜色值需要×4
            gg = gg * 4
            bb = bb * 4
            ' 转化为32位颜色值，32位颜色值最高位为0，其余按照rgb顺序排列
            mcolor_RGB(i) = bb + gg * 256& + rr * 65536
        Next i
    Close (filenum)
End Sub


' 生成图象数据到addr指向的地址
' picnum 贴图编号
' width height addr指向的dib宽高
' x1,y1,绘图位置
Public Sub genPicData(pic As RLEPic, addr As Long, ByVal Width As Long, ByVal height As Long, ByVal X1 As Long, ByVal Y1 As Long)
Dim i As Long, j As Long
Dim x As Long, y As Long
Dim row As Byte
Dim Start As Long
Dim p As Long
Dim maskNum As Byte
Dim solidNum As Byte
Dim yoffset As Long
Dim xoffset As Long
Dim offset As Long
Dim PicWidth As Long
    PicWidth = pic.Width
   'x1 = x1 - pic.x
   'y1 = y1 - pic.y
    
    If X1 >= 0 And Y1 >= 0 And X1 + PicWidth <= Width And Y1 + pic.height <= height Then
        p = 0
        For i = 1 To pic.height
            y = i
            yoffset = (y + Y1 - 1) * Width
            
            row = pic.data(p)
            Start = p
            p = p + 1
            If row > 0 Then
                x = 0
                Do
                    x = x + pic.data(p)
                    If x >= PicWidth Then
                        Exit Do
                    End If
                    p = p + 1
                    solidNum = pic.data(p)
                    p = p + 1
                    xoffset = x + (X1)
                    offset = xoffset + yoffset
                    Call CopyMemory(ByVal (addr + offset * 4), pic.Data32(p), 4 * solidNum)
                    x = x + solidNum
                    p = p + solidNum
                    If x >= PicWidth Then
                        Exit Do
                    End If
                    If p - Start >= row Then
                        Exit Do
                    End If
                Loop
                If p + 1 >= pic.DataLong Then
                    Exit For
                End If
            End If
        Next i
    Else
        p = 0
        For i = 1 To pic.height
            y = i
            yoffset = (y + Y1 - 1) * Width
            
            row = pic.data(p)
            Start = p
            p = p + 1
            If row > 0 Then
                x = 0
                Do
                    x = x + pic.data(p)
                    If x >= pic.Width Then
                        Exit Do
                    End If
                    p = p + 1
                    solidNum = pic.data(p)
                    p = p + 1
                    xoffset = x + (X1)
                    
                    If Y1 + y - 1 >= 0 And Y1 + y < height And xoffset + solidNum >= 0 And xoffset < Width Then
                        Dim p2 As Long
                        Dim ee As Long
                        
                        If xoffset < 0 Then
                            offset = yoffset
                            p2 = p - xoffset
                            ee = solidNum + xoffset
                        Else
                            offset = xoffset + yoffset
                            p2 = p
                            ee = solidNum
                        End If
                        If xoffset + solidNum >= Width Then
                            ee = ee - (xoffset + solidNum - Width + 1)
                        End If
                        Call CopyMemory(ByVal (addr + offset * 4), pic.Data32(p2), 4 * ee)
                    End If
                    x = x + solidNum
                    p = p + solidNum
                    If x >= pic.Width Then
                        Exit Do
                    End If
                    If p - Start >= row Then
                        Exit Do
                    End If
                Loop
                If p + 1 >= pic.DataLong Then
                    Exit For
                End If
            End If
        Next i
    End If
            
End Sub


Public Sub genPngPicData(pic As RLEPic, addr As Long, ByVal Width As Long, ByVal height As Long, ByVal X1 As Long, ByVal Y1 As Long)

        If X1 >= Width Or Y1 >= height Or X1 + pic.Width <= 0 Or Y1 + pic.height <= 0 Then
            Exit Sub
        End If
        
        Dim xs As Long, xe As Long, ys As Long, ye As Long
        xs = X1
        ys = Y1
        xe = X1 + pic.Width - 1
        ye = Y1 + pic.height - 1
        
        If xs < 0 Then
            xs = 0
        End If
        If ys < 0 Then
            ys = 0
        End If
        If xe >= Width Then
            xe = Width - 1
        End If
        If ye >= height Then
            ye = height - 1
        End If
        
        Dim x_off As Long, y_off As Long, dx As Long, dy As Long
        x_off = xs - X1
        y_off = ys - Y1
        dx = xe - xs + 1
        dy = ye - ys + 1
        
   
    
    Dim i               As Long, j                      As Long
    Dim pSrc            As Long, pDesc                  As Long
    Dim Sa             As SAFEARRAY, ImageData()       As Long
    Dim PicWidth        As Long
    Dim temp            As Long
    PicWidth = pic.Width
    
    With Sa
        .Element = 4
        .Dimension = 1
        .Bounds.Elements = Width * height
        .Pointer = addr
    End With
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(Sa), 4
    
    
    For j = 0 To dy - 1
        pSrc = x_off + (y_off + j) * PicWidth
        pDesc = xs + (ys + j) * Width
        For i = 0 To dx - 1
            temp = pic.Data32(pSrc)
            If (temp And &HFF000000) <> 0 Then
                ImageData(pDesc) = (temp And &HFFFFFF)
            End If
            pSrc = pSrc + 1
            pDesc = pDesc + 1
        Next
    Next
                
    CopyMemory ByVal VarPtrArray(ImageData()), 0&, 4
    
    
End Sub


Public Sub ShowPicDIB(pic As RLEPic, hDC As Long, ByVal xoffset As Long, ByVal yoffset As Long)
 
Dim addr As Long
Dim temp As Long
Dim dib As New clsDIB
    If pic.isEmpty = True Then Exit Sub

    
    dib.CreateDIB pic.Width, pic.height
    
    
    
  
    ' 在当前坐标位置贴图
    Call genPicData(pic, dib.addr, pic.Width, pic.height, 0, 0)
    
   temp = BitBlt(hDC, xoffset - pic.x, yoffset - pic.y, pic.Width, pic.height, dib.CompDC, 0, 0, &HCC0020)


End Sub




Public Sub LoadSMap(id As Long, picdata() As RLEPic, picnum As Long)
    Call LoadPicFile(G_Var.JYPath & G_Var.SMAPIDX, G_Var.JYPath & G_Var.SMAPGRP, picdata, picnum)

End Sub




' 读kdef文件
Public Sub ReadKdef()
Dim filenum As Long
Dim i  As Long
    
    filenum = OpenBin(G_Var.JYPath & G_Var.KDEFIDX, "R")
        numkdef = LOF(filenum) / 4
        ReDim KDEFIDX(numkdef)
        KDEFIDX(0) = 0
        For i = 1 To numkdef
            Get filenum, , KDEFIDX(i)
            KDEFIDX(i) = KDEFIDX(i) / 2
        Next i
    Close (filenum)
    
Dim TmptalkNum As Integer, TmpheadNum As Integer, TmpDest As Integer
    ReDim KdefInfo(numkdef - 1)
    filenum = OpenBin(G_Var.JYPath & G_Var.KDEFGRP, "R")
        For i = 0 To numkdef - 1
            KdefInfo(i).numlabel = 0
            KdefInfo(i).DataLong = (KDEFIDX(i + 1) - KDEFIDX(i))
            ReDim KdefInfo(i).data(KdefInfo(i).DataLong - 1)
            Get filenum, KDEFIDX(i) * 2 + 1, KdefInfo(i).data
                'kdefinfo(i).data(2)=
        Next i
    Close
    
    
End Sub

' 存kdef文件
Public Sub savekdef(filename As String)
Dim filenum As Long
Dim filenum2 As Long

Dim i  As Long, j As Long

Dim Length As Long
Dim offset As Long

frmmain.pb1.Max = numkdef
    filenum = OpenBin(G_Var.JYPath & G_Var.KDEFIDX, "WN")
        filenum2 = OpenBin(G_Var.JYPath & filename, "WN")
            KDEFIDX(0) = 0
            For i = 0 To numkdef - 1
                Length = KdefInfo(i).DataLong
                KDEFIDX(i + 1) = KDEFIDX(i) + Length
                For j = 0 To Length - 1
                    Put #filenum2, , KdefInfo(i).data(j)
                Next j
                Put #filenum, , CLng(KDEFIDX(i + 1) * 2)
            frmmain.pb1.Value = i
            Next i
        Close (filenum2)
    Close (filenum)

frmmain.pb1.Value = 0
End Sub
Public Sub LoadPngPicFile(filename As String, picdata() As RLEPic, picdatanum As Long)
Dim idnum As Integer
Dim PersonNum As Long
Dim filenum As Long, filenum2 As Long
Dim i As Long
Dim cX As Long, cY As Long
Dim tmpfile As String
Dim w As Long, h As Long, num As Long
    tmpfile = App.Path & "\tmp.png"
'Dim idx() As Integer
    filenum = OpenBin(filename, "R")
        Get filenum, , KGoffsetNum
        ReDim KGoffset(KGoffsetNum)
        KGoffset(0) = KGoffsetNum * 4 + 4
        For i = 1 To KGoffsetNum
            Get filenum, , KGoffset(i)
        Next i
    Close (filenum)

    picdatanum = KGoffsetNum
    
    ReDim picdata(picdatanum - 1)
    filenum = OpenBin(filename, "R")
        For num = 0 To picdatanum - 1
               ' png文件
            picdata(num).DataLong = KGoffset(num + 1) - KGoffset(num)
            'MsgBox picdata(num).DataLong
            ReDim picdata(num).data(picdata(num).DataLong - 1)
            Get filenum, KGoffset(num) + 1, cX
            'picdata(num).X = CInt(cX)
            Get filenum, KGoffset(num) + 1 + 4, cY
            'picdata(num).Y = CInt(cY)
            Get filenum, KGoffset(num) + 1 + 12, picdata(num).data ' 裸png数据
                  
    
            filenum2 = OpenBin(tmpfile, "WN")          ' 写到临时文件
                Put filenum2, , picdata(num).data
            Close filenum2
    

            Call GetPNGInfo(tmpfile, w, h)
            picdata(num).Width = w
            picdata(num).height = h
            picdata(num).x = -w / 2
            picdata(num).y = -h / 2
               
            ReDim picdata(num).Data32(w * h - 1)
            Call GetPNGData(tmpfile, picdata(num).Data32(0))
                
        Next num
    Close (filenum)
'MsgBox NewHeadNum
End Sub
'Public Sub ShowKGPicFile(filename As String, ChooseHeadNum As Long)
'Dim i, offset As Long
'        offset = KGidxR(ChooseHeadNum)
'        MsgBox offset
'        MsgBox filename
'        Call DrawPng(filename, offset)
'End Sub
Public Function DrawPng(filename As String, offset As Long, picbox As Object, background As Object, x As Long, y As Long)
        pngClass.picbox = picbox '图片框
        pngClass.SetToBkgrnd True, x, y '是否设置背景(True 或者 false), x 和 y 坐标
        pngClass.BackgroundPicture = background '背景图
        pngClass.SetAlpha = True 'Alpha 通道透明
        pngClass.SetTrans = True '透明
        pngClass.OpenPNG filename, offset
End Function

Public Sub ReadWar()
Dim Rlong() As Long
Dim offset As Long
Dim i, j As Long
Dim kk
Dim filenum As Long
ReDim Rlong(GetINILong("W_Modify", "TypeNumber") - 1)

    For j = 0 To GetINILong("W_Modify", "TypeNumber") - 1
        For i = 0 To GetINILong("W_Modify", "TypedataItem" & j) - 1
            kk = Split(GetINIStr("W_Modify", "data(" & j & "," & i & ")"), " ")
            Rlong(j) = Rlong(j) + Val(kk(2)) * Val(kk(0)) * Val(kk(1))
        Next i
    Next j
    
    filenum = OpenBin(G_Var.JYPath & G_Var.WarDefine, "R")
        warnum = LOF(filenum) / 186
        ReDim WarData(warnum - 1)
    
        For i = 0 To warnum - 1
            Seek filenum, Rlong(0) * i + 1
            Get #filenum, , WarData(i).id
            Get #filenum, , WarData(i).namebig5
            WarData(i).Name = Big5toUnicode(WarData(i).namebig5, 10)
        Next i
       
    Close filenum
End Sub
Public Sub LoadKGPicFile(filename As String)
Dim idnum As Integer
Dim PersonNum As Long
Dim filenum As Long
Dim i As Long
'Dim idx() As Integer
    filenum = OpenBin(filename, "R")
        Get filenum, , KGoffsetNum
        ReDim KGoffset(KGoffsetNum)
        KGoffset(0) = KGoffsetNum * 4 + 4
        For i = 1 To KGoffsetNum
            Get filenum, , KGoffset(i)
        Next i
    Close (filenum)

    NewHeadNum = KGoffsetNum
End Sub
