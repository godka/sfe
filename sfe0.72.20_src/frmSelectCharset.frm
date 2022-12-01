VERSION 5.00
Begin VB.Form frmSelectCharset 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "sfe0.72.20"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "黑体"
      Size            =   10.5
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmSelectCharset.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Head_Type"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3975
      Begin VB.OptionButton Option6 
         Caption         =   "PIC"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "GRP"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Style"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3975
      Begin VB.OptionButton Option4 
         Caption         =   "Kys(weyl)"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "DOS"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Language"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton Option2 
         Caption         =   "Traditional"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Simplied"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next>>"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frmSelectCharset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Sub Command1_Click()
    MsgBox "  sfe0.72.20 " & Chr(13) & Chr(13) & " Kaleo_Atras"
End Sub

'Dim E
'Dim FormWidth, FormHeight As Long
'Dim ComX, ComY As Long
'Dim StrTmp(1 To 3) As String


Private Sub Command3_Click()
If Option1.Value = True Then
    Charset = "GBK"
    Call PutINIStr("run", "charset", "GBK")
    'Unload Me
ElseIf Option2.Value = True Then
    Charset = "BIG5"
    Call PutINIStr("run", "charset", "BIG5")
    'Unload Me
End If
If Option3.Value = True Then
    Call PutINIStr("run", "style", "DOS")
ElseIf Option4.Value = True Then
    Call PutINIStr("run", "style", "kys")
End If
If Option6.Value = True Then
    Call PutINIStr("run", "Mode", "pic")
    'MsgBox "asdasd"
ElseIf Option5.Value = True Then
    Call PutINIStr("run", "Mode", "classic")
End If
'MsgBox First
If First = False Then
    MsgBox "please reset"
    End
End If
Unload Me
End Sub

Private Sub Form_Load()
'If Label1.Caption <> "请支持剧场" Or FKbutton1.Caption <> "支持" Or FKbutton2.Caption <> "不支持" Then
'    MsgBox "bs"
'    End
'End If
'FormWidth = Me.Width
'FormHeight = Me.Height
'Dim i
'Me.PaintPicture LoadResPicture(101, 1)
'    Call mciSendString("open " & loadresmp3(500) & " alias MEDIA", vbNullString, 256, 0)

'pic1.Picture = LoadResPicture(101, 0)
If GetSystemDefaultLangID = 2052 Then
    Option1.Value = True
Else
    Option2.Value = True
End If

Select Case GetINIStr("run", "Mode")
    Case ""
        Option4.Value = False
        Option5.Value = True
    Case "pic"
        Option5.Value = False
        Option6.Value = True
    Case "classic"
        Option5.Value = True
        Option6.Value = False
End Select

Select Case GetINIStr("run", "style")
    Case ""
        Option3.Value = True
        Option4.Value = False
    Case "DOS"
        Option3.Value = True
        Option4.Value = False
    Case "kys"
        Option3.Value = False
        Option4.Value = True
End Select
c_Skinner.AttachSkin Me.hwnd
'MsgBox FormHeight
'MsgBox FormWidth
End Sub


'If FKbutton1.Visible = False Then
'Else
'    MsgBox "Please choose one"
'End If

Private Sub Option1_Click()

End Sub
