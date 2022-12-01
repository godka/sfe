VERSION 5.00
Begin VB.Form frm50_0x43 
   Caption         =   "50指令43"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
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
   ScaleHeight     =   5820
   ScaleWidth      =   9690
   StartUpPosition =   2  '屏幕中心
   Begin sfe72.UserVar2 userX2 
      Height          =   1455
      Left            =   2400
      TabIndex        =   11
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2566
   End
   Begin sfe72.UserVar2 userX1 
      Height          =   1455
      Left            =   360
      TabIndex        =   10
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2566
   End
   Begin sfe72.userVar userN 
      Height          =   1095
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frm50_0x43.frx":0000
      Top             =   4560
      Width           =   5175
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "确定"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin sfe72.UserVar2 userX3 
      Height          =   1455
      Left            =   4440
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2566
   End
   Begin sfe72.UserVar2 userX4 
      Height          =   1455
      Left            =   6480
      TabIndex        =   13
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2566
   End
   Begin VB.Label Label5 
      Caption         =   "参数X4"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "参数X3"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "参数X2"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "调用事件编号N"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "参数X1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "frm50_0x43"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Long
Dim kk As Statement
Dim OffsetName As Collection
Dim b2, b1, tt



Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    'tt = Split(Combo1.Text, ",")
    '  If tt(1) <> "" Then
    '     userN.Text = tt(1)
    '  End If
    kk.data(1) = userN.Value + userX1.Value * 2 + userX2.Value * 4 + userX3.Value * 8 + userX4.Value * 16
    kk.data(2) = userN.Text
    kk.data(3) = userX1.Text
    kk.data(4) = userX2.Text
    kk.data(5) = userX3.Text
    kk.data(6) = userX4.Text

    
    Unload Me
    
End Sub

 




 

Private Sub Command1_Click()
define.Show
End Sub

Private Sub Combo1_Click()
'    MsgBox Combo1.Text
    If Combo1.ListIndex < 0 Then Exit Sub
    tt = Split(Combo1.Text, "=")
    If tt(0) <> "" Then userN.Text = tt(0)
    ComboTypeAdd userX1, tt(1), 0, userX1.Text
    ComboTypeAdd userX2, tt(1), 1, userX2.Text
    ComboTypeAdd userX3, tt(1), 2, userX3.Text
    ComboTypeAdd userX4, tt(1), 3, userX4.Text
End Sub

Private Sub Form_Load()
Dim num50 As Long
Dim i As Long
Dim s1 As String
'Dim kk() As String
    Call ConvertForm(Me)
    
    'userN.AddItem 1
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)

    GetINISection "50_43"
    Debug.Print UBound(FiftyItem)
    For i = 0 To UBound(FiftyItem) - 1
        Combo1.AddItem FiftyItem(i)
    Next i
    
    userX1.Clear: userX2.Clear: userX3.Clear: userX4.Clear
    userN.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    userX1.Value = IIf((kk.data(1) And &H2) > 0, 1, 0)
    userX2.Value = IIf((kk.data(1) And &H4) > 0, 1, 0)
    userX3.Value = IIf((kk.data(1) And &H8) > 0, 1, 0)
    userX4.Value = IIf((kk.data(1) And &H10) > 0, 1, 0)
    
    userN.Text = kk.data(2)
    userN.SetCombo
    userX1.Text = kk.data(3)
    userX1.SetCombo
    userX2.Text = kk.data(4)
    userX2.SetCombo
    userX3.Text = kk.data(5)
    userX3.SetCombo
    userX4.Text = kk.data(6)
    userX4.SetCombo
    
    Combo1.ListIndex = -1
    For i = 0 To UBound(FiftyItem) - 1
        tt = Split(FiftyItem(i), "=")
        'Debug.Print "50item=" & tt(0)
        'Debug.Print "userdata=" & userN.Text
        If Val(userN.Text) = Val(tt(0)) Then
            Combo1.ListIndex = i
            Exit For
        End If
    Next i

    
    Call Set50Form(Me, kk.data(0))
 
    'Combo1.AddItem ("无,")
'iniFileName = App.Path & "\txt.ini"
'b2 = GetIniS("50(43)", "number")
'   For i = 0 To b2
'       b1 = GetIniS("50(43)", i)
'       Combo1.AddItem (b1)
'   Next i
    'Combo1.ListIndex = 0
 '   tt = Split(Combo1.Text, ",")
 '   MsgBox tt(1)
    c_Skinner.AttachSkin Me.hWnd
End Sub
'增加新的判断模块，用来
Public Sub ComboTypeAdd(comboAdd As Object, ByVal ComboString As String, ComboNum As Long, ByVal ComboValue As Long)
'userN.Clear
Dim i As Long
'似乎实用了穷举
    comboAdd.Clear
    If InStr(1, ComboString, "talk(#" & ComboNum & ")", 1) > 0 Then
        For i = 0 To numtalk - 1
            comboAdd.AddItem i & ":" & Talk(i)
        Next i
        Exit Sub
    End If
    
    If InStr(1, ComboString, "name(#" & ComboNum & ")", 1) > 0 Then
        For i = 0 To numname - 1
            comboAdd.AddItem (i & ":" & nam(i))
        Next i
        Exit Sub
    End If
    
    If InStr(1, ComboString, "person(#" & ComboNum & ")", 1) > 0 Then
        For i = 0 To PersonNum - 1
            comboAdd.AddItem i & ":" & Person(i).Name1
        Next i
        Exit Sub
    End If
    
    If InStr(1, ComboString, "things(#" & ComboNum & ")", 1) > 0 Then
        For i = 0 To Thingsnum - 1
            comboAdd.AddItem i & ":" & Things(i).Name1
        Next i
        Exit Sub
    End If
    
    If InStr(1, ComboString, "scene(#" & ComboNum & ")", 1) > 0 Then
        For i = 0 To Scenenum - 1
            comboAdd.AddItem i & ":" & Big5toUnicode(Scene(i).Name1, 10)
        Next i
        Exit Sub
    End If
    
    If InStr(1, ComboString, "magic(#" & ComboNum & ")", 1) > 0 Then
        For i = 0 To WuGongnum - 1
            comboAdd.AddItem i & ":" & WuGong(i).Name1
        Next i
        Exit Sub
    End If
    
    If InStr(1, ComboString, "war(#" & ComboNum & ")", 1) > 0 Then
        For i = 0 To warnum - 1
            comboAdd.AddItem i & ":" & WarData(i).Name
        Next i
        Exit Sub
    End If
    'On Error Resume Next
'    comboAdd.ListIndex = ComboValue
End Sub
 
