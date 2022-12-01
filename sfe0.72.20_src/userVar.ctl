VERSION 5.00
Begin VB.UserControl userVar 
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   ClipBehavior    =   0  '��
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1020
   ScaleWidth      =   1860
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CheckBox chkX 
      Caption         =   "����"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox ComboX 
      Enabled         =   0   'False
      Height          =   345
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "userVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ȱʡ����ֵ:
Const m_def_Showtype = 0
'���Ա���:
Dim m_Showtype As Boolean


'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtX,txtX,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "����/���ÿؼ��а������ı���"
    Text = txtX.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtX.Text() = New_Text
    PropertyChanged "Text"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=chkX,chkX,-1,Value
Public Property Get Value() As Integer
Attribute Value.VB_Description = "����/���ö����ֵ��"
    Value = chkX.Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    chkX.Value() = New_Value
    PropertyChanged "Value"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,0
Public Property Get Showtype() As Boolean
    Showtype = m_Showtype
End Property

Public Property Let Showtype(ByVal New_Showtype As Boolean)
    m_Showtype = New_Showtype
    If m_Showtype = True Then
        chkX.Visible = True
    Else
        chkX.Value = 1
        chkX.Visible = False
    End If
    PropertyChanged "Showtype"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8
Public Function SetCombo() As Long
Dim i As Long
Dim s As String
    ComboX.Clear
    ComboX.AddItem StrUnicode2("=δ�������=")
    For i = 1 To KdefName.Count
        ComboX.AddItem KdefName(i)
    Next i
    
    If chkX.Value = 1 Then
        s = GetKdefname(txtX.Text)
        If s = "" Then
            ComboX.ListIndex = 0
        Else
            ComboX.Text = s
        End If
    End If
End Function

Private Sub chkX_Click()
    If chkX.Value = 1 Then
        ComboX.Enabled = True
    Else
        ComboX.Enabled = False
    End If
End Sub

Private Sub ComboX_click()
Dim s As String
    If ComboX.ListIndex > 0 Then
        s = ComboX.Text
        txtX.Text = CLng(Mid(s, 1, InStr(s, ":") - 1))
    End If
End Sub

Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
    ComboX.AddItem Item, Index
End Sub
Public Sub Clear()
    ComboX.Clear
End Sub
Private Sub UserControl_Initialize()
    chkX.Caption = StrUnicode2("����")
End Sub

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    m_Showtype = m_def_Showtype
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtX.Text = PropBag.ReadProperty("Text", "Text1")
    chkX.Value = PropBag.ReadProperty("Value", 0)
    m_Showtype = PropBag.ReadProperty("Showtype", m_def_Showtype)
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Text", txtX.Text, "Text1")
    Call PropBag.WriteProperty("Value", chkX.Value, 0)
    Call PropBag.WriteProperty("Showtype", m_Showtype, m_def_Showtype)
End Sub

