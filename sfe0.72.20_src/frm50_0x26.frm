VERSION 5.00
Begin VB.Form frm50_0x26 
   Caption         =   "50指令26"
   ClientHeight    =   4020
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
   ScaleHeight     =   4020
   ScaleWidth      =   9255
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox ComboMem 
      Height          =   345
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frm50_0x26.frx":0000
      Top             =   3000
      Width           =   5175
   End
   Begin VB.TextBox txtAddress 
      Height          =   330
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox comboType 
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "确定"
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin sfe72.userVar userX 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
   End
   Begin sfe72.userVar userI 
      Height          =   1215
      Left            =   5160
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
   End
   Begin VB.Label Label2 
      Caption         =   "辅助内存表"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "偏移I"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "= "
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "变量X"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "32位内存地址(16进制)"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frm50_0x26"
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
    
    
    kk.data(1) = userI.Value
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
Dim I As Long
Dim s1, s2 As String
Dim MemNum As Long

    Call ConvertForm(Me)
    MemNum = GetINILong("50memory", "MemNum")
    ComboType.Clear
    ComboType.AddItem StrUnicode2("存取16位字")
    ComboType.AddItem StrUnicode2("存取8位字节")
    For I = 0 To MemNum - 1
        s2 = GetINIStr("50memory", "Mem" & I)
        ComboMem.AddItem (I & ":" & s2)
   Next I
    Index = frmmain.listkdef.ListIndex
    Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
    ComboType.ListIndex = kk.data(2)
        
    txtAddress.Text = HexInt(kk.data(4)) & HexInt(kk.data(3))
    
    userX.Text = kk.data(5)
    
   
    userX.Showtype = False
    userX.SetCombo
   
    userI.Text = kk.data(6)
    userI.Value = IIf((kk.data(1) And &H1) > 0, 1, 0)
    
    userI.SetCombo
   
    Call Set50Form(Me, kk.data(0))
c_Skinner.AttachSkin Me.hWnd
End Sub




 
