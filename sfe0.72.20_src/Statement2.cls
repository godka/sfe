VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Statement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ����ָ���� ��vb�����в��ܴ�Žṹ������ֻ�ܶ���һ����


Public id As Integer      ' ָ���־
Public isGoto As Long     ' �Ƿ�Ϊ��ת,0 ����ת��1 ����ת��2 ����ת
Public GotoAddress As Integer  ' ��תƫ��
Public yesOffset As Integer    ' ����תƫ��
Public noOffset As Integer     ' ����תƫ��
Public gotoLabel As String     ' ��ת�ı��
Public islabel As Boolean      ' �Ƿ�Ϊ���ָ��
Public note As String          ' ָ��˵��
Private priDataNum As Long  ' ����������ݸ���(������ָ���־)
Private priData() As Integer ' �����������

Public Property Get DataNum() As Long
    DataNum = priDataNum
End Property

Public Property Let DataNum(ByVal Value As Long)
    priDataNum = Value
    If Value > 0 Then
        ReDim priData(priDataNum - 1)
    End If
End Property

Public Property Get Data(ByVal i As Long) As Integer
'MsgBox i
 'MsgBox priDataNum
 Data = priData(i)
End Property

Public Property Let Data(ByVal i As Long, ByVal Value As Integer)
    priData(i) = Value
End Property
