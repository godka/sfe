VERSION 5.00
Begin VB.Form frmReplacetalk1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
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
   ScaleHeight     =   4905
   ScaleWidth      =   7980
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   3480
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      Height          =   1455
      Left            =   0
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   1680
   End
   Begin VB.CommandButton Comok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Comcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ListBox Listperson2 
      Height          =   2085
      Left            =   4560
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ComboBox Combotalkman2 
      Height          =   345
      ItemData        =   "frmReplacetalk.frx":0000
      Left            =   1440
      List            =   "frmReplacetalk.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ComboBox Combotalkman 
      Height          =   345
      ItemData        =   "frmReplacetalk.frx":0004
      Left            =   1320
      List            =   "frmReplacetalk.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox Listperson 
      Height          =   2085
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      Height          =   1455
      Left            =   0
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "˵����"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "˵����ͷ��"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "˵����"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "��˵����ͷ��"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "frmReplacetalk1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Index As Long
Dim kk As Statement



Private Sub Combotalkman_click()
Dim I As Long
    If Combotalkman.ListIndex < 0 Then Exit Sub
    pic1.Cls
    Call ShowPicDIB(HeadPic(Combotalkman.ListIndex), pic1.hDC, 0, pic1.ScaleHeight)
    pic1.Refresh
    Listperson.Clear
    
    For I = 1 To HeadtoPerson(Combotalkman.ListIndex).Count
        Listperson.AddItem HeadtoPerson(Combotalkman.ListIndex).Item(I) & Person(HeadtoPerson(Combotalkman.ListIndex).Item(I)).Name1
    Next I
End Sub

Private Sub Combotalkman_GotFocus()
    Timer1.Enabled = True
End Sub

Private Sub Combotalkman_LostFocus()
    Timer1.Enabled = False
End Sub

Private Sub Combotalkman_Scroll()
    Combotalkman_click
End Sub



Private Sub Combotalkman2_click()
Dim I As Long
    If Combotalkman2.ListIndex < 0 Then Exit Sub
    pic2.Cls
    Call ShowPicDIB(HeadPic(Combotalkman2.ListIndex), pic2.hDC, 0, pic2.ScaleHeight)
    pic2.Refresh
    Listperson2.Clear
    
    For I = 1 To HeadtoPerson(Combotalkman2.ListIndex).Count
        Listperson2.AddItem HeadtoPerson(Combotalkman2.ListIndex).Item(I) & Person(HeadtoPerson(Combotalkman2.ListIndex).Item(I)).Name1
    Next I
End Sub

Private Sub Combotalkman2_GotFocus()
    Timer2.Enabled = True
End Sub

Private Sub Combotalkman2_LostFocus()
    Timer2.Enabled = False
End Sub

Private Sub Combotalkman2_Scroll()
    Combotalkman2_click
End Sub



Private Sub Comcancel_Click()
    Unload Me
End Sub

Private Sub Comok_Click()
Dim I As Long
Dim stat As Statement
Dim num_modify As Long

    If Combotalkman.ListIndex < 0 Or Combotalkman2.ListIndex < 0 Then
        Exit Sub
    End If

    num_modify = 0
    For I = 0 To numkdef - 1
        Call DatatoKdef(I)
        
        For Each stat In KdefInfo(I).kdef
            If stat.id = 1 Then
                If stat.data(1) = Combotalkman.ListIndex Then
                    num_modify = num_modify + 1
                End If
            End If
        Next stat
    Next I
    
    If MsgBox(LoadResStr(10033) & num_modify & LoadResStr(10034), vbYesNo) = vbYes Then
        For I = 0 To numkdef - 1
            Call DatatoKdef(I)
        
            For Each stat In KdefInfo(I).kdef
                If stat.id = 1 Then
                    If stat.data(1) = Combotalkman.ListIndex Then
                        stat.data(1) = Combotalkman2.ListIndex
                    End If
                End If
            Next stat
            re_Analysis (I)
            modifykdef (I)
        Next I

    End If
    
    'Unload Me
End Sub

Private Sub Form_Load()
Dim I As Long
    Combotalkman.Clear
    For I = 0 To Headnum - 1
        Combotalkman.AddItem I
    Next I
    
    Combotalkman2.Clear
    For I = 0 To Headnum - 1
        Combotalkman2.AddItem I
    Next I
   
   
    
    Me.Caption = LoadResStr(10031)
    Label1.Caption = LoadResStr(1103)
    Label2.Caption = LoadResStr(1104)
    Label3.Caption = LoadResStr(10032)
    Label4.Caption = LoadResStr(1104)
    
   
    Comok.Caption = LoadResStr(102)
    Comcancel.Caption = LoadResStr(103)
c_Skinner.AttachSkin Me.hWnd
End Sub




Private Sub Timer1_Timer()
    Combotalkman_click
End Sub

Private Sub Timer2_Timer()
    Combotalkman2_click
End Sub
