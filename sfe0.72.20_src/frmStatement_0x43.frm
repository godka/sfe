VERSION 5.00
Begin VB.Form frmStatement_0x44 
   Caption         =   "新对话修改"
   ClientHeight    =   7065
   ClientLeft      =   4545
   ClientTop       =   2700
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   13665
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame5 
      Caption         =   "姓名"
      Height          =   3735
      Left            =   9840
      TabIndex        =   28
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command7 
         Caption         =   "保存修改"
         Height          =   375
         Left            =   2640
         TabIndex        =   33
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   2520
         Width           =   3495
      End
      Begin VB.ListBox List1 
         Height          =   2220
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "修改姓名"
         Height          =   375
         Left            =   1200
         TabIndex        =   30
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "新增姓名"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   12000
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   27
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "显示颜色Color"
      Height          =   4215
      Left            =   0
      TabIndex        =   12
      Top             =   2640
      Width           =   7815
      Begin VB.TextBox txt2 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1800
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txt1 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1800
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "设置颜色"
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "选择前景色"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "选择背景色"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.PictureBox PicPalette 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   3720
         ScaleHeight     =   255
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   255
         TabIndex        =   14
         ToolTipText     =   "单击选择颜色"
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox userColor 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2640
         Top             =   2400
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2640
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "对话"
      Height          =   2415
      Left            =   4920
      TabIndex        =   10
      Top             =   120
      Width           =   4815
      Begin VB.OptionButton Option4 
         Caption         =   "修改"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "增加"
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "修改对话"
         Height          =   375
         Left            =   3600
         TabIndex        =   22
         Top             =   1680
         Width           =   975
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
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   21
         Text            =   "frmStatement_0x43.frx":0000
         Top             =   600
         Width           =   4575
      End
      Begin VB.ComboBox combotxt 
         Height          =   300
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "说话人头像"
      Height          =   2295
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   1695
      Begin VB.ComboBox Comboperson 
         Height          =   300
         Left            =   120
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   360
         Width           =   1335
      End
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
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "显示属性"
      Height          =   2175
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.TextBox combo6 
         Height          =   270
         Left            =   1440
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox Comboname 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox Combo5 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "对话框颜色"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "是否显示头像"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "显示位置"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "显示名字"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label6 
      Caption         =   "对话"
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
      Left            =   5160
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmStatement_0x44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Index, i, j As Long
Dim kk As Statement
Dim picdata() As RLEPic
Dim rr As Long, gg As Long, bb As Long
Dim Color As Long
Dim picnum As Long



'Private Sub Check2_Click()
'    Command4.Caption = Check2.Caption & "对话"
'    Check1.Value = 0
'    Check2.Value = 1
'End Sub



Private Sub Comboperson_click()
Dim i, Offset As Long
       pic1.Cls
 'pic1.Cls
'Picture1.Cls
 '      pic1.Refresh
If G_Var.EditMode = "classic" Then
        Call ShowPicDIB(HeadPic(Comboperson.ListIndex), pic1.hDC, 0, pic1.ScaleHeight)
Else
         Offset = KGoffset(Comboperson.ListIndex)
        Call DrawPng(G_Var.JYPath & G_Var.NewHeadGRP, Offset + 12, frmStatement_0x44.pic1, frmStatement_0x44.Picture1, 0, 0)
 End If
 '       Call ShowKGPicFile(G_Var.JYPath & G_Var.NewHeadGRP, Comboperson.ListIndex + 5)
End Sub

Private Sub combotxt_Change()
'    MsgBox "11"
End Sub

Private Sub combotxt_Click()
    txttalk.Text = Talk(Val(combotxt.ListIndex))
End Sub

Private Sub Command2_Click()
kk.data(0) = Comboperson.Text
kk.data(1) = combotxt.ListIndex
If Comboname.ListIndex = 0 Then
    kk.data(2) = 0
ElseIf Comboname.ListIndex = 1 Then
    kk.data(2) = -2
Else
    kk.data(2) = Comboname.ListIndex - 2
End If
kk.data(3) = Combo4.ListIndex
kk.data(4) = Combo5.ListIndex
kk.data(5) = userColor
kk.data(6) = combo6
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If Option3.Value = False Then
    'If MsgBox(LoadResStr(128), vbOKCancel, Me.Caption) = vbOK Then
        frmmain.Combotalk.List(combotxt.ListIndex) = combotxt.ListIndex & ":" & txttalk
        Talk(combotxt.ListIndex) = txttalk
        Call savetalk
        'Call savekdef("kdef.grp")
    'End If
Else
    numtalk = numtalk + 1
    ReDim Preserve Talk(numtalk - 1)
    ReDim Preserve TalkIdx(numtalk)
    Talk(numtalk - 1) = txttalk
    combotxt.AddItem (numtalk - 1)
    combotxt.ListIndex = numtalk - 1
    frmmain.Combotalk.AddItem (combotxt.ListIndex & ":" & txttalk)
'    MsgBox LoadResStr(124) & numtalk - 1 & LoadResStr(125)
    'If MsgBox(LoadResStr(128), vbOKCancel, Me.Caption) = vbOK Then
        Call savetalk
        'Call savekdef("kdef.grp")
    'End If
End If
End Sub

Private Sub Command5_Click()
    numname = numname + 1
    ReDim Preserve nam(numname - 1)
    ReDim Preserve nameidx(numname)
    nam(numname - 1) = Text1
    List1.AddItem numname - 1 & ":" & nam(numname - 1)
    List1.ListIndex = numname - 1
    Comboname.AddItem numname - 1 & ":" & nam(numname - 1)
    'frmmain.Combotalk.AddItem (combotxt.ListIndex & ":" & txttalk)
End Sub

Private Sub Command6_Click()
        Comboname.List(List1.ListIndex + 2) = List1.ListIndex & ":" & Text1
        List1.List(List1.ListIndex) = List1.ListIndex & ":" & Text1
        nam(List1.ListIndex) = Text1
End Sub

Private Sub Command7_Click()
saveName
End Sub

Private Sub Form_Load()
'Call readtalkidx
'MsgBox Nam(2)

If G_Var.EditMode = "classic" Then
    For i = 0 To Headnum - 1
        Comboperson.AddItem (i)
    Next i
Else
    For i = 0 To NewHeadNum - 1
        Comboperson.AddItem (i)
    Next i
End If
Option4.Value = 1
Comboname.AddItem (0)
Comboname.AddItem (-2)
For i = 0 To numname - 1
    Comboname.AddItem (i & ":" & nam(i))
    List1.AddItem (i & ":" & nam(i))
Next i
List1.ListIndex = 0
'MsgBox Hex(67)
'    Combo1.Clear
'MsgBox kdefnum
 '   Combo6.AddItem (0)
 '   Combo6.AddItem (1)
    For i = 0 To numtalk - 1
        combotxt.AddItem i
'        combo2.AddItem i & ":" & Talk(i)
    Next i
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
'Combo3.AddItem (StrUnicode("是"))
'Combo3.AddItem (StrUnicode("否"))
    Combo4.AddItem (StrUnicode("左"))
    Combo4.AddItem (StrUnicode("右"))
    Combo5.AddItem (StrUnicode("是"))
    Combo5.AddItem (StrUnicode("否"))
 If kk.data(2) = 0 Then
    Comboname.ListIndex = 0
 ElseIf kk.data(2) = -2 Then
    Comboname.ListIndex = 1
 Else
    Comboname.ListIndex = kk.data(2) + 2
 End If
    For j = 0 To 15
        For i = 0 To 15
            rr = (mcolor_RGB(i + j * 16) \ 65536) And &HFF&
            gg = (mcolor_RGB(i + j * 16) \ 256) And &HFF
            bb = mcolor_RGB(i + j * 16) And &HFF
            
            PicPalette.Line (i * 16, j * 16)-((i + 1) * 16, (j + 1) * 16), RGB(rr, gg, bb), BF
        Next i
    Next j
'    txtid.Text = kk.id & "(" & Hex(kk.id) & ")"
Comboperson.ListIndex = kk.data(0)
    txttalk.Text = Talk(kk.data(1))
'    Combo1.ListIndex = kk.Data(0)
'    Txtperson = kk.Data(0)
'    ComboShow.ListIndex = kk.Data(2)
'    combo2.ListIndex = kk.Data(1)
    Combo4.ListIndex = kk.data(3)
    Combo5.ListIndex = kk.data(4)
    combo6 = kk.data(6)
    combotxt.ListIndex = kk.data(1)
    'TxtName = kk.Data(2)
    userColor.Text = kk.data(5)
        Color = Int2Long(userColor.Text) \ 256
        rr = (mcolor_RGB(Color) \ 65536) And &HFF&
        gg = (mcolor_RGB(Color) \ 256) And &HFF
        bb = mcolor_RGB(Color) And &HFF
        
        Shape1.FillColor = RGB(rr, gg, bb)
        txt1.Text = Color
        
        Color = Int2Long(userColor.Text) And &HFF
        rr = (mcolor_RGB(Color) \ 65536) And &HFF&
        gg = (mcolor_RGB(Color) \ 256) And &HFF
        bb = mcolor_RGB(Color) And &HFF
        
        Shape2.FillColor = RGB(rr, gg, bb)
        txt2.Text = Color
            c_Skinner.AttachSkin Me.hWnd

End Sub


Private Sub List1_Click()
Text1 = nam(List1.ListIndex)
End Sub

Private Sub Option3_Click()
Command4.Caption = StrUnicode("增加对话")
End Sub

Private Sub Option4_Click()
Command4.Caption = StrUnicode("修改对话")
End Sub

Private Sub PicPalette_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rr As Long, gg As Long, bb As Long
Dim Color As Long
Dim colorRGB As Long
    Color = (x \ 16) + (y \ 16) * 16
    rr = (mcolor_RGB(Color) \ 65536) And &HFF&
    gg = (mcolor_RGB(Color) \ 256) And &HFF
    bb = mcolor_RGB(Color) And &HFF

    colorRGB = RGB(rr, gg, bb)

    If Option1.Value = True Then
        Shape1.FillColor = colorRGB
        txt1.Text = Color
    Else
        Shape2.FillColor = colorRGB
        txt2.Text = Color
    End If
End Sub
Private Sub Command1_Click()
        userColor.Text = Long2int(txt1.Text * 256 + txt2.Text)
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
Dim Offset As Long
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
Dim Offset As Long
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
