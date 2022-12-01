Attribute VB_Name = "basZ"
Option Explicit




Private Type LE_Head_type
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




Private Type ObjectTable_Type
    VirtualSize As Long
    RelocBaseAddr As Long
    ObjectFlags As Long
    PageTableIndex As Long
    PageTableNumber As Long
    tmp As Long
End Type


Private Type FixupRecord_type
    b1 As Byte
    b2 As Byte
    PageOffset As Integer
    Index As Byte
    OffsetLong As Long
    OffsetInt As Integer
    RealPageoffset As Long
End Type

Private Type ObjectPageTable_type
    D1 As Integer
    D2 As Integer
End Type



' �õ��������ʼ��ַ
Public Function GetZStart() As Long
Dim zfilenum As Long
Dim addr_Le As Long
Dim le As LE_Head_type
    zfilenum = OpenBin(G_Var.JYPath & G_Var.EXE, "R")
    
    Get #zfilenum, &H3C + 1, addr_Le
    Get #zfilenum, addr_Le + 1, le
    GetZStart = le.DataPagesOffset + le.PageSize
    Close #zfilenum
End Function



' �ı����
' fileid �򿪵�z.dat�ļ����
' Base Ҫд�Ļ���ַ   ��������ʼ�ε�ַ�������λ����Ϣ�ǰ���load��ĵ�ַ20000,
' INIsection Inistr Ini��Ϣ
'   ��ʽΪ�� ��һ���ַ���Ϊλ����Ϣ����������ΪҪ�޸ĵ��ֽ����ݡ�ע�ⶼΪ16���ƣ�ǰ��û��ǰ����

Public Sub ChangeZCode(fileid As Long, base As Long, INISection As String, INIstr As String)
Dim tmpstrArray() As String
Dim i As Long
        tmpstrArray = Split(GetINIStr(INISection, INIstr), ",")
        For i = 0 To UBound(tmpstrArray, 1)
            tmpstrArray(i) = "&h" & Trim(tmpstrArray(i))
        Next i
        For i = 1 To UBound(tmpstrArray, 1)
            Put #fileid, CLng(tmpstrArray(0)) - &H20000 + base + i, CByte(tmpstrArray(i))
        Next i
End Sub



Public Function Get16(base As Long, Str As String) As Long
    Get16 = CLng("&h" & Trim(Str)) + base - &H20000
End Function


 

Public Function ReadZValue(fileid As Long, INISection As String, INIstr As String) As Long
Dim tmpstrArray() As String
Dim tmpbyte As Byte
Dim tmpInt As Integer
Dim tmplong As Long
    tmpstrArray = Split(GetINIStr(INISection, INIstr), ",")
    Select Case tmpstrArray(1)
    Case 1
        Get #fileid, CLng("&h" & tmpstrArray(0)) + 1, tmpbyte
        ReadZValue = tmpbyte
    Case 2
        Get #fileid, CLng("&h" & tmpstrArray(0)) + 1, tmpInt
        ReadZValue = tmplong
    Case 4
        Get #fileid, CLng("&h" & tmpstrArray(0)) + 1, tmplong
        ReadZValue = tmplong
    End Select
End Function





' �޷�������ת��Ϊ�޷��ŵ�long
Public Function Int2Long(X As Integer) As Long
If X >= 0 Then
    Int2Long = X
Else
    Int2Long = 65536 + X
End If
End Function

' long��ת��Ϊ�޷���int
Public Function Long2int(X As Long) As Integer
If X < 32768 Then
    Long2int = X
Else
    Long2int = X - 65536
End If
End Function



' ��ȡasmָ�������Ϊָ������
' filename asm�ļ���
' Startaddr z.dat ָ������ʼ��ַ
' SectionSize z.datÿ��section�Ĵ�С
' casm ���صļ���

Public Sub ReadAsm(ByVal filename As String, ByVal StartAddr As Long, ByVal SectionSize As Long, casm As Collection)
Dim fileid As Long
Dim tmpinput As String
Dim i As Long, j As Long
Dim currentAddr As Long
Dim currentSection As Long
Dim tmparray() As String
Dim sectionNum As Long
Dim asm As clsX86
Dim tmpHex As String
    Set casm = New Collection
    
    fileid = FreeFile()
    Open App.Path & "\" & filename For Input Access Read As #fileid
    currentAddr = 0                            ' ��ǰ��ַ
    Do
        Line Input #fileid, tmpinput           ' ����һ��
        tmpinput = Trim(tmpinput)              ' ȥ���ո�
        i = InStr(1, tmpinput, ";")            ' ";"��λ��
        If i = 0 Then
            tmpinput = tmpinput                ' û�зֺ�
        ElseIf i = 1 Then
            tmpinput = ""                      ' ��һ���ַ��Ƿֺ�
        Else
            tmpinput = Mid(tmpinput, 1, i - 1) ' �зֺ�
        End If
        tmpinput = LCase(tmpinput)             ' ���Сд
        tmpinput = SubSpace(tmpinput)              ' ȥ���ո�
        If tmpinput <> "" Then
            If tmpinput = "end" Then
                Exit Do
            End If
            tmparray = Split(tmpinput)
            If tmparray(0) = "section" Then         ' Section��ʼ�����¼�����ʼƫ��
                sectionNum = CLng(tmparray(1))
                currentAddr = StartAddr + (sectionNum - 1) * SectionSize
                currentSection = sectionNum
            ElseIf tmparray(0) = "start" Then
                currentAddr = CLng("&h" & tmparray(1))
                currentSection = (currentAddr - StartAddr) / SectionSize + 1
            Else
                Set asm = New clsX86
                asm.Str = tmpinput
                asm.Address = currentAddr
                asm.PageNum = currentSection
                If Mid(tmparray(0), 1, 1) = ":" Then    ' label ���
                    asm.Style = 0
                    asm.label = Mid(tmparray(0), 2)
                Else
                    asm.num = 0
                    asm.Style = 1
                    For i = 0 To UBound(tmparray)
                        Select Case Mid(tmparray(i), 1, 1)
                        Case "*"    'fixup
                            asm.Style = 2
                            asm.offset = i
                            asm.Fixup = CLng("&h" & Mid(tmparray(i), 2)) - CLng(&H20000)
                            asm.PageOffset = asm.Address + asm.offset - StartAddr - (sectionNum - 1) * SectionSize
                            tmpHex = String(8 - Len(tmparray(i)) + 1, "0") & Mid(tmparray(i), 2)
                            asm.Data(asm.num) = CLng("&h" & Mid(tmpHex, 7, 2))
                            asm.num = asm.num + 1
                            asm.Data(asm.num) = CLng("&h" & Mid(tmpHex, 5, 2))
                            asm.num = asm.num + 1
                            asm.Data(asm.num) = CLng("&h" & Mid(tmpHex, 3, 2))
                            asm.num = asm.num + 1
                            asm.Data(asm.num) = CLng("&h" & Mid(tmpHex, 1, 2))
                            asm.num = asm.num + 1
                        Case "'"   ' ����ת
                            asm.Style = 3
                            asm.offset = i
                            asm.Data(asm.num) = 0
                            asm.num = asm.num + 1
                            asm.label = Mid(tmparray(i), 2)
                        Case """"    ' ����ת
                            asm.Style = 4
                            asm.offset = i
                            asm.Data(asm.num) = 0
                            asm.num = asm.num + 4
                            asm.label = Mid(tmparray(i), 2)
                        Case "&"        ' ����ַ��ת
                            asm.Style = 5
                            asm.offset = i
                            tmpHex = Hex(CLng("&h" & Mid(tmparray(i), 2)) - (asm.Address + asm.num + 4))
                            tmpHex = String(8 - Len(tmpHex), "0") & tmpHex
                            asm.Data(asm.num) = CLng("&h" & Mid(tmpHex, 7, 2))
                            asm.num = asm.num + 1
                            asm.Data(asm.num) = CLng("&h" & Mid(tmpHex, 5, 2))
                            asm.num = asm.num + 1
                            asm.Data(asm.num) = CLng("&h" & Mid(tmpHex, 3, 2))
                            asm.num = asm.num + 1
                            asm.Data(asm.num) = CLng("&h" & Mid(tmpHex, 1, 2))
                            asm.num = asm.num + 1
                        Case "#"    '���fixup
                            asm.Style = 6
                            asm.offset = i
                            asm.label = Mid(tmparray(i), 2)
                            asm.PageOffset = asm.Address + asm.offset - StartAddr - (sectionNum - 1) * SectionSize
                            asm.num = asm.num + 4
                        Case Else        ' ��ָͨ��
                            asm.Data(asm.num) = CLng("&h" & tmparray(i))
                            asm.num = asm.num + 1
                        End Select
                    Next i
                    currentAddr = currentAddr + asm.num
                End If
                casm.Add asm
            End If
        End If
    Loop

    Close #fileid


   For i = 1 To casm.Count
       Select Case casm(i).Style
       Case 3  '����ת
           For j = 1 To casm.Count
               If casm(j).Style = 0 Then
                   If casm(j).label = casm(i).label Then
                       tmpHex = Hex(CInt((casm(j).Address - casm(i).offset - 1 - casm(i).Address) * 256))
                       tmpHex = String(4 - Len(tmpHex), "0") & tmpHex
                       casm(i).Data(casm(i).offset) = CLng("&h" & Mid(tmpHex, 1, 2))
                       Exit For
                   End If
               End If
           Next j
       Case 4   '����ת
           For j = 1 To casm.Count
               If casm(j).Style = 0 Then
                   If casm(j).label = casm(i).label Then
                       tmpHex = Hex(casm(j).Address - casm(i).offset - 4 - casm(i).Address)
                       tmpHex = String(8 - Len(tmpHex), "0") & tmpHex
                       casm(i).Data(casm(i).offset) = CLng("&h" & Mid(tmpHex, 7, 2))
                       casm(i).Data(casm(i).offset + 1) = CLng("&h" & Mid(tmpHex, 5, 2))
                       casm(i).Data(casm(i).offset + 2) = CLng("&h" & Mid(tmpHex, 3, 2))
                       casm(i).Data(casm(i).offset + 3) = CLng("&h" & Mid(tmpHex, 1, 2))
                       Exit For
                   End If
               End If
           Next j
       Case 6   '���fixup
           For j = 1 To casm.Count
               If casm(j).Style = 0 Then
                   If casm(j).label = casm(i).label Then
                       casm(i).Fixup = casm(j).Address - CLng(&H20000)
                       tmpHex = Hex(casm(i).Fixup)
                       tmpHex = String(8 - Len(tmpHex), "0") & tmpHex
                       casm(i).Data(casm(i).offset) = CLng("&h" & Mid(tmpHex, 7, 2))
                       casm(i).Data(casm(i).offset + 1) = CLng("&h" & Mid(tmpHex, 5, 2))
                       casm(i).Data(casm(i).offset + 2) = CLng("&h" & Mid(tmpHex, 3, 2))
                       casm(i).Data(casm(i).offset + 3) = CLng("&h" & Mid(tmpHex, 1, 2))
                       Exit For
                   End If
               End If
           Next j
       End Select

   Next i





End Sub


' ��ȡasmָ�������Ϊָ������
' filename asm�ļ���
' Startaddr z.dat ָ������ʼ��ַ
' SectionSize z.datÿ��section�Ĵ�С
' casm ���صļ���

Public Sub ReadZmodify(ByVal filename As String, casm As Collection)
Dim fileid As Long
Dim tmpinput As String
Dim i As Long, j As Long
Dim currentAddr As Long
Dim tmparray() As String
Dim asm As clsX86
Dim tmpHex As String
    Set casm = New Collection
    
    fileid = FreeFile()
    Open filename For Input Access Read As #fileid
    currentAddr = 0                            ' ��ǰ��ַ
    Do
        Line Input #fileid, tmpinput           ' ����һ��
        tmpinput = Trim(tmpinput)              ' ȥ���ո�
        i = InStr(1, tmpinput, ";")            ' ";"��λ��
        If i = 0 Then
            tmpinput = tmpinput                ' û�зֺ�
        ElseIf i = 1 Then
            tmpinput = ""                      ' ��һ���ַ��Ƿֺ�
        Else
            tmpinput = Mid(tmpinput, 1, i - 1) ' �зֺ�
        End If
        tmpinput = LCase(tmpinput)             ' ���Сд
        tmpinput = SubSpace(tmpinput)              ' ȥ���ո�
        If tmpinput <> "" Then
            If tmpinput = "end" Then
                Exit Do
            End If
            tmparray = Split(tmpinput)
            If tmparray(0) = "start" Then
                currentAddr = CLng("&h" & tmparray(1))
            Else
                Set asm = New clsX86
                asm.Str = tmpinput
                asm.Address = currentAddr
                asm.num = 0
                asm.Style = 1
                For i = 0 To UBound(tmparray)
                    asm.Data(asm.num) = CLng("&h" & tmparray(i))
                    asm.num = asm.num + 1
                Next i
                currentAddr = currentAddr + asm.num
                casm.Add asm
            End If
        End If
    Loop

    Close #fileid


End Sub




