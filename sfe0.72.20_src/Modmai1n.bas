Attribute VB_Name = "Modmain"
Option Explicit

'Public RangerNum As Long
Public Recordnum As Long                   ' ½ø¶È±àºÅ

Public Const XSCALE = 18
Public Const YSCALE = 9

Public First As Boolean

Public c_Skinner As New CSkinner

Public Const MaskColor = &H707030

Public colorA(9) As Long
Public colorB(9) As Long

'Õâ¸öÐÞ¸ÄÆ÷ÖÐÖ¸Áî¸öÊý
Public Const KdefNum = &H48


Public Type BITMAPINFOHEADER '4Ø bytes
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


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y _
As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As _
Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Public Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long


Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, addr As Byte, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long




Public Type WarDataType   ' Õ½¶·³¡¾°Êý¾Ý   ÕâÐ©Êý¾ÝÎÞ·¨·ÅÈë´°Ìå£¬²»Ö§³Ö public type
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
Public warnum As Long             ' Õ½¶·¸öÊý


Public Type statementAttribType  ' Ö¸ÁîÊôÐÔ
    Length As Long               ' Ö¸Áî³¤¶È
    isGoto As Long               ' Ö¸ÁîÊÇ·ñÎªÌõ¼þ×ªÒÆ£¬0²»ÊÇ 1ÊÇ
    yesOffset As Long            ' Ìõ¼þÂú×ã×ªÒÆµØÖ·ÔÚÖ¸ÁîÖÐµÚ¼¸¸ö×Ö
    noOffset As Long             ' Ìõ¼þ²»Âú×ã×ªÒÆµØÖ·ÔÚÖ¸ÁîÖÐµÚ¼¸¸ö×Ö
    notes As String              ' Ö¸ÁîËµÃ÷
End Type

Public StatAttrib() As statementAttribType

Public KGidxR() As Long
Public KGidx() As Long
Public pngClass As New LoadPNG

Public kdefword() As Integer   ' ±£´ækdefÊÂ¼þ¶þ½øÖÆÊý¾Ý
Public kdeflong As Long        ' ÊÂ¼þÎÄ¼þ³¤¶È

Public KDEFIDX() As Long       ' ÊÂ¼þkdefË÷Òý
Public numkdef As Long         ' ±£´æÊÂ¼þ¸öÊý

Public Type KdefType           ' ±£´æÊÂ¼þÊý¾Ý
    DataLong  As Long          ' ÊÂ¼þ¶þ½øÖÆÊý¾Ý³¤¶È
    data() As Integer          ' ÊÂ¼þ¶þ½øÖÆÊý¾Ý
    kdef As Collection         ' ÊÂ¼þÖ¸Áî¼¯ºÏ
    numlabel As Long           ' ÊÂ¼þÖÐ±êºÅÊýÄ¿
End Type

Public KdefInfo() As KdefType

Public KdefName As Collection       ' Ö¸Áî¶¨ÒåµÄÃû×Ö¼¯ºÏ


Public ClipboardStatement As Collection  ' ¸´ÖÆÏÂµÄÖ¸ÁîÊý¾Ý

Public ClipboardKdef As Collection  ' ¸´ÖÆÏÂµÄÖ¸ÁîÊý¾Ý

Public nameidx() As Long
Public nam() As String
Public numname As Long

Public Talk() As String        ' ¶Ô»°×Ö·û´®
Public TalkIdx() As Long           ' ¶Ô»°Ë÷Òý
Public numtalk As Long      ' ¶Ô»°¸öÊý


Public Type PersonAttrib       ' ÈËÎïÊôÐÔ
    r1 As Integer
    PhotoId As Integer
    r3 As Integer
    r4 As Integer
    name1big5(9) As Byte
    Name1 As String
    name2big5(9) As Byte
    name2 As String
    
End Type

Public Person() As PersonAttrib  ' ÈËÎïÊôÐÔÊý×é
Public PersonNum As Long         ' ÈËÎï¸öÊý

Public Type ThingsAttrib         ' ÎïÆ·ÊôÐÔ
    r1 As Integer
    name1big5(19) As Byte
    Name1 As String
    name2big5(19) As Byte
    name2 As String
End Type

Public Things() As ThingsAttrib  ' ÎïÆ·ÊôÐÔÊý×é
Public Thingsnum As Long         ' ÎïÆ·¸öÊý



Public Type SceneType
    SceneID As Integer                ' ³¡¾°´úºÅ
    Name1(9) As Byte                  ' Ãû×Ö
    OutMusic As Integer               ' ³öÃÅÒôÀÖ
    InMusic  As Integer               ' ½øÃÅÒôÀÖ
    JumpScene As Integer              ' Ìø×ª³¡¾°
    InCondition As Integer            ' ½øÈëÌõ¼þ
    MMapInX1 As Integer               ' Ö÷µØÍ¼Èë¿Ú×ø±ê
    MMapInY1 As Integer
    MMapInX2 As Integer
    MMapInY2 As Integer
    InX As Integer                    ' ½øÈëºó³õÊ¼×ø±ê
    InY As Integer
    OutX(2) As Integer                ' Èý¸ö³ö¿Ú×ø±ê
    OutY(2) As Integer
    JumpX1 As Integer               ' Á½¸öÌø×ª¿Ú×ø±ê
    JumpY1 As Integer
    JumpX2 As Integer               ' Á½¸öÌø×ª¿Ú×ø±ê
    JumpY2 As Integer
End Type

Public Scene() As SceneType  ' ³¡¾°ÊôÐÔÊý×é
Public Scenenum As Long          ' ³¡¾°¸öÊý

Public Type WuGongAttrib           ' Îä¹¦ÊôÐÔ
    r1 As Integer
    name1big5(19) As Byte
    Name1 As String
End Type


Public WuGong() As WuGongAttrib    ' Îä¹¦ÊôÐÔÊý×é
Public WuGongnum As Long           ' Îä¹¦¸öÊý


Public Type RLEPic
    isEmpty As Boolean
    Width As Integer   ' Í¼Æ¬¿í¶È
    Height As Integer  ' Í¼Æ¬¸ß¶È
    X As Integer       ' Í¼Æ¬xÆ«ÒÆ
    Y As Integer       ' Í¼Æ¬yÆ«ÒÆ
    DataLong As Long   ' Í¼Æ¬RLEÑ¹ËõÊý¾Ý³¤¶È
    data() As Byte     ' Í¼Æ¬RLEÑ¹ËõÊý¾Ý
    Data32() As Long   ' Í¼Æ¬32Î»Ñ¹ËõÊý¾Ý
End Type

Public HeadPic() As RLEPic  ' ÈËÎïÍ·ÏñÊý¾Ý
Public PngPic() As RLEPic  ' picÍ·ÏñÊý¾Ý
Public Headnum As Long      ' ÈËÎïÍ·Ïñ¸öÊý
Public NewHeadNum As Long

Public WarPic() As RLEPic  ' Õ½¶·Í¼Æ¬Êý¾Ý
Public Warpicnum As Long   ' Õ½¶·¸öÊý

Public g_PP As RLEPic     ' ±à¼­ÌùÍ¼ÓÃ£¬´«µÝ²ÎÊý¡£


' D* ÊÂ¼þÐÅÏ¢

Public Type D_Event_type
    isGo As Integer
    id As Integer
    EventNum1 As Integer
    EventNum2 As Integer
    EventNum3 As Integer
    picnum(2) As Integer
    PicDelay As Integer
    X As Integer
    Y As Integer
End Type

Public g_DD As D_Event_type          ' ÐÞ¸Ä³¡¾°ÊÂ¼þ¶¨Òå´°ÌåÊ¹ÓÃ£¬ÓÃÀ´´«µÝ²ÎÊý¡£



Public HeadtoPerson() As Collection   ' ¸ù¾ÝÍ·Ïñid²éÈËÎïid

Public mcolor_RGB(256) As Long  ' ÑÕÉ«±í


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
            MDIMain.mnu_z.Enabled = False
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
    
    MDIMain.Show
 
Exit Sub

Label1:
    MsgBox Err.Description
    If (MDIMain Is Nothing) = False Then
        Unload MDIMain
    End If

End Sub

' ¶ÁÈ¡r1ÎÄ¼þ
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


    WuGongnum = (idxr(5) - idxr(4)) / Rlong(4)
    ReDim WuGong(WuGongnum - 1)
    offset = idxr(4)
    For i = 0 To WuGongnum - 1
        Get filenum, offset + i * Rlong(4) + 1, WuGong(i).r1
        Get filenum, offset + i * Rlong(4) + NameOFFset(4) + 1, WuGong(i).name1big5
        WuGong(i).Name1 = Big5toUnicode(WuGong(i).name1big5, 20)
        
    Next i
 
    Close (filenum)
End Sub

' ¶ÁÈ¡ÈËÎïÕÕÆ¬²¢×ª»¯Îª32Î»rle

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
        For num = 1 To picdatanum ' µØÍ¼ÌùÍ¼µÄË÷Òý¸öÊý
            Get filenum, , HeadIDX(num)
        Next num
        Close filenum
   Else
        picdatanum = FileLen(filepic) / (64 * 64 * 12)
        MsgBox picdatanum
   End If
    
    filenum = OpenBin(filepic, "R")
    For num = 0 To picdatanum - 1 ' µØÍ¼ÌùÍ¼µÄË÷Òý¸öÊý
        If HeadIDX(num) < 0 Then
            picLong = 0
        Else
            offset = HeadIDX(num)
            picLong = HeadIDX(num + 1) - HeadIDX(num)
            If (num = picdatanum - 1) And (HeadIDX(num + 1) <> LOF(filenum)) And HeadIDX(num) > 0 Then ' ×îºóÒ»¸öidxÓ¦¸ÃµÈÓÚÎÄ¼þ³¤¶È
                picLong = LOF(filenum) - HeadIDX(num)
            End If
        End If
        If picLong > 0 Then
            picdata(num).isEmpty = False
            Get filenum, offset + 1, picdata(num).Width
            Get filenum, , picdata(num).Height
            Get filenum, , picdata(num).X
            Get filenum, , picdata(num).Y
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




' °ÑÌùÍ¼Êý¾ÝµÄ8BitRLEÑ¹ËõÊý¾Ý£¬×ª»»Îª32Bit£¬·½±ãÒÔºó´¦Àí
' RLEÑ¹Ëõ¸ñÊ½£º
' µÚÒ»¸ö×Ö½ÚÎªµÚÒ»ÐÐÊý¾Ý³¤¶È£¨¼¸¸ö×Ö½Ú£©
' ºóÃæÒ»¸ö×Ö½ÚÎªÍ¸Ã÷Êý¾Ýµã¸öÊý£¬ºóÃæ¸ú×ÅÎª²»Í¸Ã÷Êý¾Ýµã¸öÊý£¬È»ºóÊÇ²»Í¸Ã÷µÄÃ¿¸öÊý¾Ýµã8Î»ÑÕÉ«£¬
' ÖØ¸´ÒÔÉÏ£¬Ö±µ½µÚÒ»ÐÐ×Ö½Ú½áÊø
' ¶ÁÈ¡ÏÂÒ»ÐÐÊý¾Ý£¬Ö±µ½Ã»ÓÐºóÃæÊý¾Ý
Public Sub RLEto32(pic As RLEPic)
Dim p As Long  ' Ö¸ÏòRLEÊý¾ÝµÄÖ¸Õë
Dim p32 As Long   ' Ö¸Ïò32Î»·ÇÑ¹ËõÊý¾ÝµÄÖ¸Õë
Dim i As Long, j As Long
Dim row As Byte
Dim temp As Long
Dim Start As Long
Dim maskNum As Long
Dim solidNum As Long
   
    ReDim pic.Data32(pic.DataLong)
   
    p = 0
    p32 = 0
    For i = 1 To pic.Height
        row = pic.data(p)     ' µ±Ç°ÐÐÊý¾Ý¸öÊý
        pic.Data32(p) = row
        Start = p             ' µ±Ç°ÐÐÆðÊ¼Î»ÖÃ
        p = p + 1
        If row > 0 Then
            p32 = 0
            Do
                maskNum = pic.data(p)  ' ÑÚÂë¸öÊý
                pic.Data32(p) = row
                p = p + 1
      
                p32 = p32 + maskNum
                If p32 >= pic.Width Then  ' ´ËÑÚÂëÍê³ÉºóÎ»ÖÃÖ¸ÕëÒÑ¾­Ö¸Ïò×îÓÒ¶Ë
                    Exit Do
                End If
                solidNum = pic.data(p) ' Êµ¼ÊµãµÄ¸öÊý
                pic.Data32(p) = solidNum
                p = p + 1
                For j = 1 To solidNum
                    temp = pic.data(p)
                    pic.Data32(p) = mcolor_RGB(temp)
                    p32 = p32 + 1
                    p = p + 1
                Next j
                If p32 >= pic.Width Then   ' Êµ¼ÊµãÍê³ÉºóÎ»ÖÃÖ¸ÕëÒÑ¾­Ö¸Ïò×îÓÒ¶Ë
                    Exit Do
                End If
                If p - Start >= row Then           ' µ±Ç°ÐÐÊý¾ÝÒÑ¾­½áÊø
                    Exit Do
                End If
            Loop
            If p + 1 >= pic.DataLong Then
                Exit For
            End If
        End If
    Next i
   
End Sub



' ¶ÁÈ¡ÑÕÉ«±íÊý¾Ý
' jinyongÖÐÑÕÉ«±íÊÇ°´ÕÕ256É«£¬Ã¿É«rgb¸÷Ò»¸ö×Ö½Ú
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
            rr = rr * 4           ' ÑÕÉ«ÖµÐèÒª¡Á4
            gg = gg * 4
            bb = bb * 4
            ' ×ª»¯Îª32Î»ÑÕÉ«Öµ£¬32Î»ÑÕÉ«Öµ×î¸ßÎ»Îª0£¬ÆäÓà°´ÕÕrgbË³ÐòÅÅÁÐ
            mcolor_RGB(i) = bb + gg * 256& + rr * 65536
        Next i
    Close (filenum)
End Sub


' Éú³ÉÍ¼ÏóÊý¾Ýµ½addrÖ¸ÏòµÄµØÖ·
' picnum ÌùÍ¼±àºÅ
' width height addrÖ¸ÏòµÄdib¿í¸ß
' x1,y1,»æÍ¼Î»ÖÃ
Public Sub genPicData(pic As RLEPic, addr As Long, ByVal Width As Long, ByVal Height As Long, ByVal X1 As Long, ByVal Y1 As Long)
Dim i As Long, j As Long
Dim X As Long, Y As Long
Dim row As Byte
Dim Start As Long
Dim p As Long
Dim maskNum As Byte
Dim solidNum As Byte
Dim yoffset As Long
Dim xoffset As Long
Dim offset As Long
    
   'x1 = x1 - pic.x
   'y1 = y1 - pic.y
    
    If X1 >= 0 And Y1 >= 0 And X1 + pic.Width <= Width And Y1 + pic.Height <= Height Then
        p = 0
        For i = 1 To pic.Height
            Y = i
            yoffset = (Y + Y1 - 1) * Width
            
            row = pic.data(p)
            Start = p
            p = p + 1
            If row > 0 Then
                X = 0
                Do
                    X = X + pic.data(p)
                    If X >= pic.Width Then
                        Exit Do
                    End If
                    p = p + 1
                    solidNum = pic.data(p)
                    p = p + 1
                    xoffset = X + (X1)
                    offset = xoffset + yoffset
                    Call CopyMemory(ByVal (addr + offset * 4), pic.Data32(p), 4 * solidNum)
                    X = X + solidNum
                    p = p + solidNum
                    If X >= pic.Width Then
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
        For i = 1 To pic.Height
            Y = i
            yoffset = (Y + Y1 - 1) * Width
            
            row = pic.data(p)
            Start = p
            p = p + 1
            If row > 0 Then
                X = 0
                Do
                    X = X + pic.data(p)
                    If X >= pic.Width Then
                        Exit Do
                    End If
                    p = p + 1
                    solidNum = pic.data(p)
                    p = p + 1
                    xoffset = X + (X1)
                    
                    If Y1 + Y - 1 >= 0 And Y1 + Y < Height And xoffset + solidNum >= 0 And xoffset < Width Then
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
                    X = X + solidNum
                    p = p + solidNum
                    If X >= pic.Width Then
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


Public Sub genPngPicData(pic As RLEPic, addr As Long, ByVal Width As Long, ByVal Height As Long, ByVal X1 As Long, ByVal Y1 As Long)

        If X1 >= Width Or Y1 >= Height Or X1 + pic.Width <= 0 Or Y1 + pic.Height <= 0 Then
            Exit Sub
        End If
        
        Dim xs As Long, xe As Long, ys As Long, ye As Long
        xs = X1
        ys = Y1
        xe = X1 + pic.Width - 1
        ye = Y1 + pic.Height - 1
        
        If xs < 0 Then
            xs = 0
        End If
        If ys < 0 Then
            ys = 0
        End If
        If xe >= Width Then
            xe = Width - 1
        End If
        If ye >= Height Then
            ye = Height - 1
        End If
        
        Dim x_off As Long, y_off As Long, dx As Long, dy As Long
        x_off = xs - X1
        y_off = ys - Y1
        dx = xe - xs + 1
        dy = ye - ys + 1
        
        Dim psrc As Long, pDesc As Long
        
        
        Dim i As Long, j As Long
        Dim tmpdata As Long
        For j = 0 To dy - 1
            psrc = x_off + (y_off + j) * pic.Width
            pDesc = xs + (ys + j) * Width
            For i = 0 To dx - 1
                tmpdata = pic.Data32(psrc + i)
                If (tmpdata And &HFF000000) <> 0 Then      ' alpha>0 ²»Í¸Ã÷
                    tmpdata = tmpdata And &HFFFFFF
                     Call CopyMemory(ByVal (addr + (pDesc + i) * 4), tmpdata, 4)
                End If
            Next i
        Next j
End Sub


Public Sub ShowPicDIB(pic As RLEPic, hDC As Long, ByVal xoffset As Long, ByVal yoffset As Long)
 
Dim addr As Long
Dim temp As Long
Dim dib As New clsDIB
    If pic.isEmpty = True Then Exit Sub

    
    dib.CreateDIB pic.Width, pic.Height
    
    
    
  
    ' ÔÚµ±Ç°×ø±êÎ»ÖÃÌùÍ¼
    Call genPicData(pic, dib.addr, pic.Width, pic.Height, 0, 0)
    
   temp = BitBlt(hDC, xoffset - pic.X, yoffset - pic.Y, pic.Width, pic.Height, dib.CompDC, 0, 0, &HCC0020)


End Sub




Public Sub LoadSMap(id As Long, picdata() As RLEPic, picnum As Long)
    Call LoadPicFile(G_Var.JYPath & G_Var.SMAPIDX, G_Var.JYPath & G_Var.SMAPGRP, picdata, picnum)

End Sub




' ¶ÁkdefÎÄ¼þ
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
            If KdefInfo(i).data(0) = 1 Then
                KdefInfo(i).DataLong = 8
                ReDim Preserve KdefInfo(i).data(KdefInfo(i).DataLong - 1)
                TmptalkNum = KdefInfo(i).data(0)
                TmpheadNum = KdefInfo(i).data(1)
                TmpDest = KdefInfo(i).data(2)
                KdefInfo(i).data(0) = 68
                KdefInfo(i).data(1) = TmpheadNum
                KdefInfo(i).data(2) = TmptalkNum
                KdefInfo(i).data(3) = -2
                KdefInfo(i).data(4) = TmpDest Mod 2
                KdefInfo(i).data(5) = 0
                KdefInfo(i).data(6) = 28515
                KdefInfo(i).data(7) = 0
            End If
                'kdefinfo(i).data(2)=
        Next i
    Close
    
    
End Sub

' ´ækdefÎÄ¼þ
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
               ' pngÎÄ¼þ
            picdata(num).DataLong = KGoffset(num + 1) - KGoffset(num)
            'MsgBox picdata(num).DataLong
            ReDim picdata(num).data(picdata(num).DataLong - 1)
            Get filenum, KGoffset(num) + 1, cX
            'picdata(num).X = CInt(cX)
            Get filenum, KGoffset(num) + 1 + 4, cY
            'picdata(num).Y = CInt(cY)
            Get filenum, KGoffset(num) + 1 + 12, picdata(num).data ' ÂãpngÊý¾Ý
                  
    
            filenum2 = OpenBin(tmpfile, "WN")          ' Ð´µ½ÁÙÊ±ÎÄ¼þ
                Put filenum2, , picdata(num).data
            Close filenum2
    

            Call GetPNGInfo(tmpfile, w, h)
            picdata(num).Width = w
            picdata(num).Height = h
            picdata(num).X = -w / 2
            picdata(num).Y = -h / 2
               
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
Public Function DrawPng(filename As String, offset As Long, picbox As Object, background As Object, X As Long, Y As Long)
        pngClass.picbox = picbox 'Í¼Æ¬¿ò
        pngClass.SetToBkgrnd True, X, Y 'ÊÇ·ñÉèÖÃ±³¾°(True »òÕß false), x ºÍ y ×ø±ê
        pngClass.BackgroundPicture = background '±³¾°Í¼
        pngClass.SetAlpha = True 'Alpha Í¨µÀÍ¸Ã÷
        pngClass.SetTrans = True 'Í¸Ã÷
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
