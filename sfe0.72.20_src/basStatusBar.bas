Attribute VB_Name = "basStatusBar"
Option Explicit

Private Const WS_CHILD As Long = &H40000000             'WS_CHILD ��WS_VISIBLE�Ǳ��躯��
Private Const WS_VISIBLE As Long = &H10000000
Private Const WM_USER As Long = &H400
Private Const SB_SETPARTS As Long = (WM_USER + 4)       '������������VB�Դ���api��ѯ����û�У���Ҫ�ֹ����
Private Const SB_SETTEXTA As Long = (WM_USER + 1)
Private Declare Function CreateStatusWindow Lib "comctl32.dll" (ByVal style As Long, ByVal lpszText As String, ByVal hwndParent As Long, ByVal wID As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'--------------------------------------------------
'
'                       ����״̬��
'����˵����
'ParenthWnd����״̬�������ľ��
'IDC_STATBAR   ״̬����ID�ţ����ڶ�״̬���ĵ���֮��Ĳ���
'hBarWin       ��������״̬���ľ��
'szText        Ҫ��ʾ����Ϣ
'
'---------------------------------------------------
Public Function CreateStatBar(ParenthWnd As Long, IDC_STATBAR As Long, hBarWin As Long, Optional szText As String = "Demo") As Boolean
    Dim ret As Long                 '����ֵ
    Dim bar(0 To 1) As Long         '�����ĸ���λ��
    Dim szbar As Long               '��������Ŀ
   
'-------------------------------------------------------
'��������
    bar(0) = 235                    '��һ�����Ϊ245
    bar(1) = -1                     '-1��ʾ����ķ�Ϊһ��
   
'-------------------------------------------------------www.����ɱchinai tp ow er.comrK25Hny

    ret = CreateStatusWindow(WS_CHILD Or WS_VISIBLE, ByVal szText, ParenthWnd, IDC_STATBAR)     '����״̬��
    szbar = 2
    If ret = 0 Then                 '�������ʧ�����˳�����
        CreateStatBar = False
        Exit Function
    End If
    hBarWin = ret                   '����״̬���ľ��
    CreateStatBar = True            '�����ɹ�������ֵ
End Function


Public Sub SetStatBar(hbar As Long, szbar As Long, bar() As Long)
    If szbar > 1 Then               '��ΪĬ�Ͼ��Ƿ�һ�����ԣ������ж�Ϊ����1���Ƿ���
        SendMessage hbar, SB_SETPARTS, szbar, bar(0)    '����
    End If
End Sub
'----------------------------
'�ƶ�״̬��
'----------------------------
Public Sub MoveStatWindow(hbar As Long)
If hbar Then                '���״̬�������Ϊ0���ƶ�
    Call MoveWindow(hbar, 0, 0, 0, 0, True)
End If
End Sub

'------------------------------
'��ָ��������ʾ��Ϣ
'hBar Ϊ״̬���ľ��
'szbar ָ��Ҫ����һ����ʾ��Ϣ����0��ʼ�ƣ�Ҳ����˵�����������������Ҫ�ڵڶ�������ʾ��Ϣ��szbar������Ϊ1
'szText Ҫ��ʾ����Ϣ
'-------------------------------
Public Sub SetBarText(hbar As Long, szbar As Long, strText As String)
    SendMessage hbar, SB_SETTEXTA, szbar, ByVal strText
End Sub

