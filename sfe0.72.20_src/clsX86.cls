VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsX86"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Public Str As String         ' ָ��ԭʼ�ִ�
Public Style As Long          ' ָ������  0 label 1 ��ͨ 2 fixup 3 ����ת 4 ����ת 5 ����ַ��ת 6 ����ض�λ
Public label As String         ' label�ַ���
Public Address                ' ָ����ʼ��ַ
Public num As Long             ' ָ���ֽ���
Private priData(20) As Byte     ' ����ָ��
Public Fixup As Long          ' Fixup��ֵ
Public offset As Long         ' fixup�������ת����ָ���е��ֽ���
Public PageOffset As Integer  ' fixup ��pageoffset
Public PageNum As Long        ' �ڼ���Page


Public Property Get Data(ByVal i As Long) As Integer
    Data = priData(i)
End Property

Public Property Let Data(ByVal i As Long, ByVal value As Integer)
    priData(i) = value
End Property
