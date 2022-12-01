VERSION 5.00
Begin VB.Form frmStatement_0x45 
   Caption         =   "字符替换"
   ClientHeight    =   2955
   ClientLeft      =   11160
   ClientTop       =   750
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   6390
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmStatement_0x45.frx":0000
      Top             =   1920
      Width           =   5175
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "确定"
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Txttalk 
      Height          =   270
      Left            =   840
      TabIndex        =   6
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.ComboBox Combotalk 
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1440
      Width           =   6255
   End
   Begin VB.ComboBox UserK 
      Height          =   300
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox ComboType 
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "对话"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "项目编号ID"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "属性类别"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmStatement_0x45"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index, I, j As Long
Dim kk As Statement
Dim s1 As String

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
kk.data(0) = ComboType.ListIndex
kk.data(1) = UserK.ListIndex
kk.data(2) = Combotalk.ListIndex
Unload Me
End Sub



Private Sub Combotalk_Click()
 Txttalk = Combotalk.ListIndex
End Sub

Private Sub Form_Load()
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    ComboType.Clear
    For I = 1 To 4
        s1 = GetINIStr("R_Modify", "TypeName" & I)
        ComboType.AddItem s1
    Next I
    For I = 0 To numname - 1
        Combotalk.AddItem (I & ":" & nam(I))
    Next I
    ComboType.ListIndex = kk.data(0)
    UserK.ListIndex = kk.data(1)
'    Combotalk.ListIndex = kk.Data(2) - 1
    Txttalk = kk.data(2)
    Combotalk.ListIndex = Txttalk
        c_Skinner.AttachSkin Me.hWnd

End Sub
Private Sub ComboType_click()
Dim I As Long
    UserK.Clear
    Select Case ComboType.ListIndex
    Case 0
        For I = 0 To PersonNum - 1
            UserK.AddItem I & ":" & Person(I).Name1
        Next I
        
    Case 1
        For I = 0 To Thingsnum - 1
            UserK.AddItem I & ":" & Things(I).Name1
        Next I
    
    
    Case 2
        For I = 0 To Scenenum - 1
            UserK.AddItem I & ":" & Big5toUnicode(Scene(I).Name1, 10)
        Next I
    
    Case 3
        For I = 0 To WuGongnum - 1
            UserK.AddItem I & ":" & WuGong(I).Name1
        Next I
    Case 4
         
    End Select
    
 UserK.ListIndex = 0
End Sub



Private Sub Txttalk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If IsNumeric(Txttalk) Then
        If Val(Txttalk) <= numtalk Then
            Combotalk.ListIndex = Val(Txttalk)
        End If
    End If
End If
End Sub
