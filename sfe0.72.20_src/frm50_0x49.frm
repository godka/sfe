VERSION 5.00
Begin VB.Form frm50_0x49 
   Caption         =   "50ָ������"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   10275
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frm50_0x49.frx":0000
      Top             =   2040
      Width           =   5175
   End
   Begin sfe72.userVar user5 
      Height          =   1095
      Left            =   6720
      TabIndex        =   5
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin sfe72.userVar user4 
      Height          =   1335
      Left            =   4800
      TabIndex        =   2
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2355
   End
   Begin sfe72.userVar user3 
      Height          =   1215
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2143
   End
   Begin VB.Label Label4 
      Caption         =   "����ֵ"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "��������"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "������ʼ���"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "32λ�ڴ��ַ(16����)"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frm50_0x49"
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
 
 s = txtAddress.Text
    If Len(s) < 8 Then
        s = String(8 - Len(s), "0") & s
    End If
    kk.data(1) = Long2int(CLng("&h" & Mid(s, 5, 4)))
    kk.data(2) = Long2int(CLng("&h" & Mid(s, 1, 4)))
    kk.data(3) = user3.Text
    kk.data(4) = user4.Text
    kk.data(5) = user5.Text
Unload Me
End Sub

Private Sub Form_Load()
Dim num50 As Long
Dim I As Long
Dim s1 As String
Index = frmmain.listkdef.ListIndex
Set kk = KdefInfo(frmmain.Combokdef.ListIndex).kdef.Item(Index + 1)
Call ConvertForm(Me)
'MsgBox HexInt(kk.Data(2)) & HexInt(kk.Data(1))

txtAddress = HexInt(kk.data(2)) & HexInt(kk.data(1))
'txtAddress.Text = HexInt(kk.Data(2)) & HexInt(kk.Data(1))
user3.Text = kk.data(3)
user4.Text = kk.data(4)
user5.Text = kk.data(5)
user3.Showtype = False
user4.Showtype = False
user5.Showtype = False
Call Set50Form(Me, kk.data(0))
c_Skinner.AttachSkin Me.hWnd
End Sub
