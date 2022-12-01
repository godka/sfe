VERSION 5.00
Begin VB.Form frmznew 
   Caption         =   "�����µ�z.dat"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6225
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "����z.dat"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmznew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private zfilenum As Long
Private zfilename As String

Dim addr_Le As Long       ' LEͷ�ļ���ʼ��ַ

Private StartData() As Byte   ' ͷ�ļ�����


Private Type LE_Head_type    ' LEͷ�ļ�
    label As String * 2
    b1 As Byte
    b2 As Byte
    formatLevel As Long
    CPUType As Integer
    OSType As Integer
    l1 As Long
    L_10_1 As Long
    PageNumber As Long
    L_18_1 As Long
    L_18_2 As Long
    L_20_1 As Long
    ESP As Long
    PageSize As Long
    LastPageSize As Long
    FixupSize As Long
    L_302 As Long
    LoadSectionSize As Long
    L_382 As Long
    ObjectTableOFfset As Long
    ObjectNumber As Long
    ObjectPageOffset As Long
    L_482 As Long
    ResourceTableOffeset As Long
    L_502 As Long
    ResidentTableOffset As Long
    EntryTableOffset As Long
    L_601 As Long
    L_602 As Long
    FixupPageOffset As Long
    FixupTableOffset As Long
    ImportTableOFfset As Long
    L_702 As Long
    ImportProcOffset As Long
    L_782 As Long
    DataPagesOffset As Long
    L_802 As Long
    L_881 As Long
    L_882 As Long
    L_901 As Long
    L_902 As Long
    L_981 As Long
    L_982 As Long
    L_a01 As Long
    L_a02 As Long
    L_A81 As Long
End Type

Dim le As LE_Head_type


Private Type ObjectTable_Type     ' Object table
    VirtualSize As Long
    RelocBaseAddr As Long
    ObjectFlags As Long
    PageTableIndex As Long
    PageTableNumber As Long
    tmp As Long
End Type

Dim ObjectTable() As ObjectTable_Type

Private Type FixupRecord_type     ' fixup table
    b1 As Byte
    b2 As Byte
    PageOffset As Integer
    Index As Byte
    OffsetLong As Long
    OffsetInt As Integer
    RealPageoffset As Long
End Type
Dim Fixup(20000) As FixupRecord_type
Dim FixupNumber As Long


Private Type ObjectPageTable_type        ' Object page table
    D1 As Integer
    D2 As Integer
End Type

Dim ObjectPageTable() As ObjectPageTable_type

Dim Resource() As Byte

Dim FixupPage() As Long

Dim LoadData() As Byte


Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
Dim i As Long
Dim offset As Long
Dim tmpstrArray() As String
Dim NewSP As Long
    If MsgBox(LoadResStr(10131), vbYesNo) = vbNo Then Exit Sub
            
    Me.MousePointer = vbHourglass
        
    ' �����һ��Fixup���˴�ԭ��Ϸ����
    zfilenum = OpenBin(App.Path & "\z_old.dat", "W")
    Put #zfilenum, CLng(GetINIStr("Newz", "LastAddr")) + 1, CLng(GetINIStr("Newz", "LastValue"))
    Close zfilenum
        
    Call Load_Fixup(App.Path & "\z_old.dat")    ' ��ȡԭʼz.dat����
        
        
    NewSP = CLng(GetINIStr("NewZ", "NewSP"))
        
    le.ESP = NewSP
    ObjectTable(1).VirtualSize = NewSP
    
    Call ChangeFixup("All")     ' �޸�ȫ��fixup
    
    Call Enlarge_Z                ' ����z
    
    Call ChangeFixup("IDX")
    Call ChangeFixup("person")
    Call ChangeFixup("Wupin")
    Call ChangeFixup("Changjing")
    Call ChangeFixup("Wugong")
    Call ChangeFixup("VRAM")
    
    Call ChangeFixup("dong")           ' ����ɽ��
    Call ChangeFixup("Peihe")          ' �����书�������
    Call ChangeFixup("Donghua")        ' �����书����
    Call ChangeFixup("Exp")            ' ���Ӿ����б�
    
    'zfilenum = OpenBin(App.Path & "\z_new.dat", "WN")
    zfilenum = OpenBin(G_Var.JYPath & "z.dat", "WN")
    Put #zfilenum, , StartData
    Put #zfilenum, , le
    Put #zfilenum, le.ObjectTableOFfset + 1 + addr_Le, ObjectTable
    Put #zfilenum, , ObjectPageTable
    Put #zfilenum, , Resource
    Put #zfilenum, , FixupPage
    
    For i = 0 To FixupNumber - 1
            Put #zfilenum, , Fixup(i).b1
            Put #zfilenum, , Fixup(i).b2
            If Fixup(i).b2 = &H10 Then
                Put #zfilenum, , Fixup(i).PageOffset
                Put #zfilenum, , Fixup(i).Index
                Put #zfilenum, , Fixup(i).OffsetLong
            Else
                Put #zfilenum, , Fixup(i).PageOffset
                Put #zfilenum, , Fixup(i).Index
                Put #zfilenum, , Fixup(i).OffsetInt
            End If
    Next i
    
    Put #zfilenum, le.DataPagesOffset + 1, LoadData
    
    
    Close #zfilenum
   
    Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Dim i As Long
    For i = 0 To Me.Controls.Count - 1
         Call SetCaption(Me.Controls(i))
    Next i

    
End Sub



' ��ȡ��Ϸz.dat �ļ���¼
Private Sub Load_Fixup(filename As String)
Dim i As Long, j As Long
Dim fixoffset As Long
Dim PageStart As Long, pageEnd As Long
Dim fileout As Long
Dim offset As Long
Dim tmplong As Long
    zfilenum = OpenBin(filename, "R")
    
    Get #zfilenum, &H3C + 1, addr_Le         ' LE��ʼ��ַ
    ' ��ȡLE֮ǰ������
    ReDim StartData(addr_Le - 1)
    Get #zfilenum, 1, StartData
    ' ��ȡLE
    Get #zfilenum, addr_Le + 1, le
    ' ��ȡ�����
    ReDim ObjectTable(le.ObjectNumber - 1)
    Get #zfilenum, addr_Le + 1 + le.ObjectTableOFfset, ObjectTable
    ' ��ȡ����ҳ��
     ReDim ObjectPageTable(le.PageNumber - 1)
    Get #zfilenum, addr_Le + 1 + le.ObjectPageOffset, ObjectPageTable
    ' ��ȡ��Դ��
    ReDim Resource(le.FixupPageOffset - le.ResourceTableOffeset - 1)
    Get #zfilenum, addr_Le + 1 + le.ResourceTableOffeset, Resource
    ' ��ȡfixupҳ��
    ReDim FixupPage(le.PageNumber)
    Get #zfilenum, addr_Le + 1 + le.FixupPageOffset, FixupPage
    
    
'    ��ȡfixup��¼
    
    i = 0
    For j = 0 To le.PageNumber - 1
        PageStart = FixupPage(j)
        pageEnd = FixupPage(j + 1)
        offset = PageStart
        Do
            If offset >= pageEnd Then Exit Do
            Get #zfilenum, offset + addr_Le + 1 + le.FixupTableOffset, Fixup(i).b1
            Get #zfilenum, , Fixup(i).b2
            If Fixup(i).b2 = &H10 Then
                Get #zfilenum, , Fixup(i).PageOffset
                Get #zfilenum, , Fixup(i).Index
                Get #zfilenum, , Fixup(i).OffsetLong
                offset = offset + 9
            Else
                Get #zfilenum, , Fixup(i).PageOffset
                Get #zfilenum, , Fixup(i).Index
                Get #zfilenum, , Fixup(i).OffsetInt
                offset = offset + 7
            End If
            Fixup(i).RealPageoffset = Fixup(i).PageOffset + le.DataPagesOffset + j * le.PageSize
            i = i + 1
            Get #zfilenum, Fixup(i).RealPageoffset + 1, tmplong
        Loop
    Next j
    FixupNumber = i
   
    ' ��ȡʵ����������
    ReDim LoadData((le.PageNumber - 1) * le.PageSize + le.LastPageSize - 1)
    Get #zfilenum, le.DataPagesOffset + 1, LoadData
    
    Close #zfilenum
    
End Sub


' �޸�fixup��¼
Private Sub ChangeFixup(Str1 As String)
Dim tmpstr() As String
Dim MinAddr As Long
Dim MaxAddr As Long
Dim newAddr As Long
Dim i As Long
Dim k As Long
Dim addr As Long
Dim BaseAddr As Long
Dim filenum As Long
Dim is32 As Long
    tmpstr = Split(GetINIStr("NewZ", Str1 & "AddrRange"), ",")
    newAddr = CLng(GetINIStr("NewZ", Str1 & "NewAddr"))
    MinAddr = CLng(tmpstr(0))
    MaxAddr = CLng(tmpstr(1))

    filenum = FreeFile
    For i = 0 To FixupNumber
        If Fixup(i).OffsetLong >= MinAddr And Fixup(i).OffsetLong < MaxAddr Then
            Fixup(i).OffsetLong = Fixup(i).OffsetLong - MinAddr + newAddr
        End If
    Next i
End Sub

' ����z.dat�����������ӻ��ָ�����������
Private Sub Enlarge_Z()
Dim i As Long, j As Long
Dim PageAdd As Long
Dim tmpstrArray() As String
Dim fixNUM As Long
Dim casm As Collection
Dim newPageFixupnum() As Long
Dim currentFixup As Long
Dim currentPage As Long

    PageAdd = GetINILong("newZ", "PageAdd")     ' ���ӵ�ҳ��
   
    le.PageNumber = le.PageNumber + PageAdd      ' ҳ������
    le.LastPageSize = le.PageSize                ' ���ҳ������ҳһ��
    
    ObjectTable(1).PageTableNumber = ObjectTable(1).PageTableNumber + PageAdd     ' �ڶ���Object��ҳ������
    
    ReDim Preserve ObjectPageTable(le.PageNumber - 1)     ' �ض���ObjectPage��
    
    For i = 0 To PageAdd - 1    ' ObjectPage������
        ObjectPageTable(le.PageNumber - 1 - i).D1 = 0
        ObjectPageTable(le.PageNumber - 1 - i).D2 = le.PageNumber - i
    Next i
    
    ' ��������ƫ�ƣ���������ObjectPage
    le.ResourceTableOffeset = le.ResourceTableOffeset + 4 * PageAdd
    le.ResidentTableOffset = le.ResidentTableOffset + 4 * PageAdd
    le.EntryTableOffset = le.EntryTableOffset + 4 * PageAdd
    le.FixupPageOffset = le.FixupPageOffset + 4 * PageAdd
    
    ' ��������ƫ�ƣ�����������FixupPageOffset
    le.FixupTableOffset = le.FixupTableOffset + 4 * PageAdd + 4 * PageAdd
    
    ReDim Preserve FixupPage(le.PageNumber)
    
    fixNUM = 200
    For i = 0 To PageAdd - 1   ' ����fixupҳ
        FixupPage(le.PageNumber - PageAdd + i + 1) = FixupPage(le.PageNumber - PageAdd) + 9 * fixNUM * (i + 1)
    Next i
    FixupNumber = FixupNumber + fixNUM * PageAdd
    'ReDim Fixup(FixupNumber - 1)
    
    For i = 0 To PageAdd * fixNUM - 1   ' �µ�fix��ʼֵ
        Fixup(FixupNumber - fixNUM * PageAdd + i).b1 = &H7
        Fixup(FixupNumber - fixNUM * PageAdd + i).b2 = &H10
        Fixup(FixupNumber - fixNUM * PageAdd + i).PageOffset = 0
        Fixup(FixupNumber - fixNUM * PageAdd + i).Index = 2
        Fixup(FixupNumber - fixNUM * PageAdd + i).OffsetLong = 0
    Next i
    
    
    le.ImportProcOffset = le.ImportProcOffset + 4 * PageAdd + 4 * PageAdd + 9 * fixNUM * PageAdd
    le.ImportTableOFfset = le.ImportProcOffset
    le.DataPagesOffset = &H1F000
    le.FixupSize = le.ImportProcOffset - le.FixupPageOffset + 1
    le.LoadSectionSize = le.ImportProcOffset - le.ObjectTableOFfset + 1
        
    ReDim Preserve LoadData((le.PageNumber - 1) * le.PageSize + le.LastPageSize - 1)
    
    
    
    Call ReadAsm("addkdef.txt", le.DataPagesOffset, le.PageSize, casm)
    
    
    If PageAdd > 0 Then
        ReDim newPageFixupnum(PageAdd - 1)
        
        For i = 0 To PageAdd - 1
            newPageFixupnum(i) = 0
        Next i
        For i = 1 To casm.Count
            If casm(i).Style = 2 Or casm(i).Style = 6 Then
                currentPage = casm(i).PageNum - le.PageNumber + PageAdd - 1
                currentFixup = FixupNumber - fixNUM * PageAdd + currentPage * fixNUM + newPageFixupnum(currentPage)
                Fixup(currentFixup).PageOffset = casm(i).PageOffset
                Fixup(currentFixup).OffsetLong = casm(i).Fixup
                newPageFixupnum(currentPage) = newPageFixupnum(currentPage) + 1
                If newPageFixupnum(currentPage) > fixNUM Then
                    MsgBox "fixup > fixupmax"
                End If
            End If
        Next i
    End If
    
    For i = 1 To casm.Count
        If casm(i).Style <> 0 Then
            For j = 0 To casm(i).num - 1
                LoadData(casm(i).Address + j - CLng(&H1F000)) = casm(i).Data(j)
            Next j
        End If
    Next i
    
    
    
End Sub

