VERSION 5.00
Begin VB.Form frm50_0x27 
   Caption         =   "50ָ��27"
   ClientHeight    =   4305
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
   ScaleHeight     =   4305
   ScaleWidth      =   9690
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frm50_0x27.frx":0000
      Top             =   3000
      Width           =   5175
   End
   Begin sfe72.UserVar2 UserK 
      Height          =   1095
      Left            =   2880
      TabIndex        =   9
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin sfe72.userVar userX 
      Height          =   1215
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2778
   End
   Begin VB.ComboBox ComboType 
      Height          =   345
      ItemData        =   "frm50_0x27.frx":007D
      Left            =   1560
      List            =   "frm50_0x27.frx":007F
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "="
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "�ַ�������S"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "����"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "���Ա��ID"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "�������"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm50_0x27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Long
Dim kk As Statement
Dim OffsetName As Collection



Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
 
    kk.data(1) = UserK.Value
    kk.data(2) = ComboType.ListIndex
    kk.data(3) = UserK.Text
    kk.data(4) = userX.Text
    kk.data(5) = 0
    kk.data(6) = 0

    
    Unload Me
    
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
    
    
End Sub

Private Sub Form_Load()
Dim num50 As Long
Dim I As Long
Dim s1 As String
    Call ConvertForm(Me)
    
    
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)

    ComboType.Clear
    For I = 1 To 4
        s1 = GetINIStr("R_Modify", "TypeName" & I)
        ComboType.AddItem s1
    Next I
    
    ComboType.ListIndex = kk.data(2)
    
    
    
    
    
    UserK.Text = kk.data(3)
    UserK.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    
    UserK.SetCombo

    userX.Text = kk.data(4)
    userX.Showtype = False
    userX.SetCombo

    Call Set50Form(Me, kk.data(0))
 c_Skinner.AttachSkin Me.hWnd
End Sub

 
