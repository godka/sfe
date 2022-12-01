VERSION 5.00
Begin VB.Form frm50_0x25 
   Caption         =   "50ָ��25"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
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
   ScaleHeight     =   4110
   ScaleWidth      =   9255
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox ComboMem 
      Height          =   345
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1200
      Width           =   2175
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
      TabIndex        =   7
      Text            =   "frm50_0x25.frx":0000
      Top             =   2760
      Width           =   5175
   End
   Begin VB.ComboBox comboType 
      Height          =   345
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin sfe72.userVar userX 
      Height          =   1215
      Left            =   5160
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
   End
   Begin sfe72.userVar userI 
      Height          =   1215
      Left            =   2640
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
   End
   Begin VB.Label Label2 
      Caption         =   "�����ڴ��"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "ƫ��I"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "= "
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "ֵX"
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "32λ�ڴ��ַ(16����)"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frm50_0x25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Long
Dim kk As Statement


 
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
Dim s As String
    kk.data(1) = userX.Value + userI.Value * 2
    kk.data(2) = ComboType.ListIndex
    
    s = txtAddress.Text
    If Len(s) < 8 Then
        s = String(8 - Len(s), "0") & s
    End If
    kk.data(3) = Long2int(CLng("&h" & Mid(s, 5, 4)))
    kk.data(4) = Long2int(CLng("&h" & Mid(s, 1, 4)))
    kk.data(5) = userX.Text
    kk.data(6) = userI.Text

    
    Unload Me
    
End Sub

 



Private Sub ComboMem_Click()
Dim s2 As String
Dim pp
    s2 = GetINIStr("50memory", "Mem" & ComboMem.ListIndex)
    pp = Split(s2, " ")
    txtAddress = pp(0)
End Sub

Private Sub Form_Load()
Dim num50 As Long
Dim i As Long
Dim s1, s2 As String
Dim MemNum As Long
Dim pp
Dim j
    Call ConvertForm(Me)
    
    ComboType.Clear
    ComboType.AddItem StrUnicode2("��ȡ16λ��")
    ComboType.AddItem StrUnicode2("��ȡ8λ�ֽ�")
    MemNum = GetINILong("50memory", "MemNum")
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    ComboType.ListIndex = kk.data(2)
    For i = 0 To MemNum - 1
        s2 = GetINIStr("50memory", "Mem" & i)
        ComboMem.AddItem (i & ":" & s2)
    Next i
    txtAddress.Text = HexInt(kk.data(4)) & HexInt(kk.data(3))
    
    userX.Text = kk.data(5)
    
    userX.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    userI.Value = IIf((kk.data(1) And &H2) > 0, 1, 0)
   
    userX.SetCombo
    userI.Text = kk.data(6)
    userI.SetCombo
   
    Call Set50Form(Me, kk.data(0))
c_Skinner.AttachSkin Me.hWnd
End Sub

 
