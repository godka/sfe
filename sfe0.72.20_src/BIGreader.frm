VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBIGreader 
   AutoRedraw      =   -1  'True
   Caption         =   "Title.big修改"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   Begin VB.CommandButton Command3 
      Caption         =   "转为640*480"
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存big"
      Height          =   495
      Left            =   9120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5160
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   5520
      ScaleHeight     =   200
      ScaleMode       =   0  'User
      ScaleWidth      =   320
      TabIndex        =   3
      Top             =   1200
      Width           =   4800
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   240
      ScaleHeight     =   200
      ScaleMode       =   0  'User
      ScaleWidth      =   320
      TabIndex        =   2
      Top             =   1200
      Width           =   4800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复制到剪贴板"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton paste 
      Caption         =   "从剪贴板获取"
      Height          =   495
      Left            =   7200
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "当前游戏big图片"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmBIGreader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mcolor_RGB(256) As Long
Dim mcolor_Col(256) As Long
Dim mcolor_Win(256) As Long
Dim Color(320 - 1, 200 - 1) As Byte
Dim data() As Long

Public Sub SetColor1()
Dim filenum As Integer
Dim I As Long
Dim rr As Byte, gg As Byte, bb As Byte
    
    'filenum = FreeFile()
    filenum = OpenBin(G_Var.JYPath & G_Var.Palette, "R")
        For I = 0 To 255
            Get filenum, , rr
            Get filenum, , gg
            Get filenum, , bb
            rr = (rr * 4) Xor 0
            gg = (gg * 4) Xor 0
            bb = (bb * 4) Xor 0
'            List1.AddItem rr & " " & gg & " " & bb
            ' 转化为32位颜色值，32位颜色值最高位为0，其余按照rgb顺序排列
            mcolor_Col(I) = bb + gg * 256& + rr * 65536
            mcolor_Win(I) = RGB(rr, gg, bb)
        Next I
    Close (filenum)
    'List1.Clear
    'transColor
End Sub
' 打开二进制文件
' status = "R"   读
' status = "W"   写，备份文件
' Status = "WN"  写新文件，可以比原来文件小，备份文件
Public Function OpenBin(filename As String, status As String) As Long
   OpenBin = FreeFile()
   Select Case UCase(status)
   Case "R"
       If Dir(filename) = "" Then
           Err.Raise vbObjectError + 1, , "File " & filename & " not exist"
       End If
       Open filename For Binary Access Read As OpenBin
   Case "W"
       FileCopy filename, filename & ".bak"
       Open filename For Binary Access Write As OpenBin
   Case "WN"
       If Dir(filename & ".bak") <> "" Then
           Kill filename & ".bak"
       End If
       If Dir(filename) <> "" Then
           Name filename As filename & ".bak"
       End If
       Open filename For Binary Access Write As OpenBin
   End Select
   
End Function



Private Sub Combo1_Click()
getbig G_Var.JYPath & Combo1.Text
End Sub

Private Sub Command1_Click()
    
    convertdata G_Var.JYPath & Combo1.Text
End Sub

Private Sub Command2_Click()
    Clipboard.Clear
    Clipboard.SetData pic1.Image, vbCFBitmap
End Sub

Private Sub Command3_Click()
convertdata2 G_Var.JYPath & "huge" & Combo1.Text
End Sub

Private Sub Form_Load()
Dim filenum2 As Long
Dim filenum1 As Long
Dim tmpbyte As Byte

    Me.Caption = StrUnicode(Me.Caption)
    For I = 0 To Me.Controls.Count - 1
        Call SetCaption(Me.Controls(I))
    Next I
    Combo1.AddItem G_Var.Title
    Combo1.AddItem G_Var.Dead
    Combo1.ListIndex = 0
    
    SetColor1
    'SetColor
    getbig G_Var.JYPath & Combo1.Text
    'getbigcolor2
    c_Skinner.AttachSkin Me.hWnd
End Sub
Public Sub getbigcolor()
Dim I, j As Long
    
    For I = 0 To 320 - 1
        For j = 0 To 200 - 1
            pic1.PSet (1 * I, 1 * j), mcolor_Win(Color(I, j))
            'Me.PSet (2 * i + 1, 2 * j), mcolor_Col(color(i, j))
            'Me.PSet (2 * i + 1, 2 * j + 1), mcolor_Col(color(i, j))
            'Me.PSet (2 * i, 2 * j + 1), mcolor_Col(color(i, j))
        Next j
    Next I
End Sub
Public Sub convertdata(filename As String)
'Dim i, j As Long
Dim I As Long, j As Long
Dim c, d, e, f As Long
Dim rr As Long, gg As Long, bb As Long
Dim yy As Long, uu As Long, vv As Long
Dim rc(255) As Long, gc(255) As Long, bc(255) As Long
Dim yc(255) As Long, uc(255) As Long, vc(255) As Long
Dim vmin As Long, v As Long
Dim nn As Long


    For I = 0 To 255
        rc(I) = (mcolor_Col(I) \ 65536) And &HFF
        gc(I) = (mcolor_Col(I) \ 256) And &HFF
        bc(I) = mcolor_Col(I) And &HFF
        yc(I) = 0.299 * rc(I) + 0.587 * gc(I) + 0.114 * bc(I)
        uc(I) = -0.1687 * rc(I) - 0.3313 * gc(I) + 0.5 * bc(I) + 128
        vc(I) = 0.5 * rc(I) - 0.4187 * gc(I) - 0.0813 * bc(I) + 128
    Next I
    
    ReDim data(320 - 1, 200 - 1)
    
    For j = 0 To 200 - 1
        For I = 0 To 320 - 1
            data(I, j) = pic2.Point(I, j)
            vmin = 100000#
            rr = data(I, j) And &HFF
            gg = (data(I, j) \ 256) And &HFF
            bb = (data(I, j) \ 65536) And &HFF
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
            'pic1.PSet (i, j), RGB(rc(nn), gc(nn), bc(nn))
            Color(I, j) = nn
            pb1.Value = Int((j * 320 + I) * 100 / 64000)
            'DoEvents
        Next I
    Next j
    
    filenum = OpenBin(filename, "WN")
        Put filenum, , Color
    Close (filenum)
    
    MsgBox "pic has been saved in " & G_Var.JYPath & Combo1.Text
End Sub
Public Sub getbig(filename As String)
        filenum2 = OpenBin(filename, "R")
            Get filenum2, , Color
        Close (filenum2)
    'getbigcolor
    getbigcolor
End Sub

Private Sub paste_Click()
Dim I As Long, j As Long
Dim picdata As New StdPicture
    Set picdata = Clipboard.GetData
    'pic2.Width = ScaleX(picdata.Width, vbHimetric, vbPixels)
    'pic2.Height = ScaleY(picdata.Height, vbHimetric, vbPixels)
    If picdata = 0 Then Exit Sub
    pic2.PaintPicture picdata, 0, 0, 320, 200
    'WW = PicBak.Width
    'HH = PicBak.Height
    'txtwidth.Text = WW
    'txtHeight.Text = HH
    'ReDim data(WW - 1, HH - 1)
    'PicBak.Picture = picdata
    'For j = 0 To HH - 1
    '    For i = 0 To WW - 1
    '        data(i, j) = PicBak.Point(i, j)
    '    Next i
    'Next j
    'ShowData
End Sub
Public Sub convertdata2(filename As String)
'Dim i, j As Long
Dim I As Long, j As Long
Dim colorhuge(640 - 1, 400 - 1) As Byte
    
    For j = 0 To 200 - 1
        For I = 0 To 320 - 1
            colorhuge(2 * I, 2 * j) = Color(I, j)
            colorhuge(2 * I, 2 * j + 1) = Color(I, j)
            colorhuge(2 * I + 1, 2 * j) = Color(I, j)
            colorhuge(2 * I + 1, 2 * j + 1) = Color(I, j)
            pb1.Value = Int((j * 320 + I) * 100 / 64000)
            'DoEvents
        Next I
    Next j
    
    filenum = OpenBin(filename, "WN")
        Put filenum, , colorhuge
    Close (filenum)
    
    MsgBox "pic has been saved in " & G_Var.JYPath & Combo1.Text
End Sub
