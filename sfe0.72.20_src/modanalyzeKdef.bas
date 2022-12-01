Attribute VB_Name = "modanalyzeKdef"
Option Explicit

Public Sub modifykdef(ii As Long)
Dim tmpdatalong As Long
Dim i As Long, j As Long
Dim tmpKdef As Collection
Dim p As Long
    If ii < 0 Then Exit Sub
    
    Set tmpKdef = KdefInfo(ii).kdef
    Call HandleLabel(tmpKdef)
    
    tmpdatalong = 0
    For i = 1 To tmpKdef.Count
        If tmpKdef.Item(i).islabel = False Then
            tmpdatalong = tmpdatalong + 1 + tmpKdef.Item(i).DataNum
        End If
    Next i
    If tmpdatalong <> KdefInfo(ii).DataLong Then
      '  MsgBox "错误cmdmodifykdef_Click"
      '  End
    End If
    
    KdefInfo(ii).DataLong = tmpdatalong
    ReDim KdefInfo(ii).data(tmpdatalong - 1)
    p = 0
    For i = 1 To tmpKdef.Count
        If tmpKdef.Item(i).islabel = False Then
            KdefInfo(ii).data(p) = tmpKdef.Item(i).id
            p = p + 1
            For j = 0 To tmpKdef.Item(i).DataNum - 1
                KdefInfo(ii).data(p) = tmpKdef.Item(i).data(j)
                p = p + 1
            Next j
        End If
    Next i
    
    
End Sub



' 处理指令集合中的label，重新定位语句中的偏移量
Public Sub HandleLabel(kdef As Collection)
Dim tmpstat As Statement
Dim i As Long, j As Long
Dim offset As Long
    For i = 1 To kdef.Count
        If kdef.Item(i).isGoto > 0 Then
            If kdef.Item(i).data(kdef.Item(i).yesOffset - 1) > 0 Or kdef.Item(i).data(kdef.Item(i).noOffset - 1) > 0 Then
                offset = 0
                For j = i + 1 To kdef.Count
                    If kdef.Item(j).islabel = False Then
                        offset = offset + 1 + kdef.Item(j).DataNum
                    Else
                        If kdef.Item(i).gotoLabel = kdef.Item(j).note Then
                            Exit For
                        End If
                    End If
                Next j
                If kdef.Item(i).isGoto = 1 Then
                    kdef.Item(i).data(kdef.Item(i).yesOffset - 1) = offset
                    kdef.Item(i).data(kdef.Item(i).noOffset - 1) = 0
                Else
                    kdef.Item(i).data(kdef.Item(i).yesOffset - 1) = 0
                    kdef.Item(i).data(kdef.Item(i).noOffset - 1) = offset
                End If
            Else
                offset = 0
                For j = i - 1 To 1 Step -1
                    If kdef.Item(j).islabel = False Then
                        offset = offset + 1 + kdef.Item(j).DataNum
                    Else
                        If kdef.Item(i).gotoLabel = kdef.Item(j).note Then
                            Exit For
                        End If
                    End If
                Next j
                If kdef.Item(i).isGoto = 1 Then
                    kdef.Item(i).data(kdef.Item(i).yesOffset - 1) = -offset - kdef.Item(i).DataNum - 1
                    kdef.Item(i).data(kdef.Item(i).noOffset - 1) = 0
                Else
                    kdef.Item(i).data(kdef.Item(i).yesOffset - 1) = 0
                    kdef.Item(i).data(kdef.Item(i).noOffset - 1) = -offset - kdef.Item(i).DataNum - 1
                End If
            End If
        End If
    Next i
End Sub


' 根据偏移量生成指令集合中的label，标记跳转位置
Public Sub GenLabel(kdef As Collection, numlabel As Long)
Dim tmpstat As Statement
Dim i As Long, j As Long
Dim offset As Long
    i = 1
    numlabel = 0
    Do
        If kdef.Item(i).isGoto > 0 Then
            If kdef.Item(i).data(kdef.Item(i).yesOffset - 1) > 0 Or kdef.Item(i).data(kdef.Item(i).noOffset - 1) > 0 Then
                offset = 0
                For j = i + 1 To kdef.Count
                    If kdef.Item(j).islabel = False Then
                        offset = offset + 1 + kdef.Item(j).DataNum
                        If offset = kdef.Item(i).GotoAddress Then
                            Set tmpstat = New Statement
                            tmpstat.islabel = True
                            kdef.Item(i).gotoLabel = ":Label" & numlabel
                            numlabel = numlabel + 1
                            tmpstat.note = kdef.Item(i).gotoLabel
                            kdef.Add tmpstat, , , j
                            Exit For
                        End If
                    End If
                Next j
            Else
                offset = 0
                For j = i To 1 Step -1
                    If kdef.Item(j).islabel = False Then
                        offset = offset + 1 + kdef.Item(j).DataNum
                        If offset = -kdef.Item(i).GotoAddress Then
                            Set tmpstat = New Statement
                            tmpstat.islabel = True
                            kdef.Item(i).gotoLabel = ":Label" & numlabel
                            numlabel = numlabel + 1
                            tmpstat.note = kdef.Item(i).gotoLabel
                            kdef.Add tmpstat, , j
                            i = i + 1
                            Exit For
                        End If
                    End If
                Next j

            End If
        End If
        i = i + 1
        If i > kdef.Count Then Exit Do
    Loop
    
End Sub



' 根据指令集合重新生成指令集合
Public Sub re_Analysis(ii As Long)
Dim tmpdatalong As Long
Dim tmpdata() As Integer
Dim i As Long, j As Long
Dim tmpKdef As Collection
Dim p As Long
Dim stat As Statement
    Set tmpKdef = KdefInfo(ii).kdef
    Call HandleLabel(tmpKdef)
    tmpdatalong = 0
    For i = 1 To tmpKdef.Count
        If tmpKdef.Item(i).islabel = False Then
             tmpdatalong = tmpdatalong + 1 + tmpKdef.Item(i).DataNum
        End If
    Next i
    If tmpdatalong <> KdefInfo(ii).DataLong Then
    '    MsgBox "错误re_Analysis"
    '    End
    End If
    
    ReDim tmpdata(tmpdatalong - 1)
    p = 0
    For i = 1 To tmpKdef.Count
        If tmpKdef.Item(i).islabel = False Then
            tmpdata(p) = tmpKdef.Item(i).id
            p = p + 1
            For j = 0 To tmpKdef.Item(i).DataNum - 1
                tmpdata(p) = tmpKdef.Item(i).data(j)
                p = p + 1
            Next j
        End If
    Next i
    
    
    While tmpKdef.Count > 0
        tmpKdef.Remove 1
    Wend
    KdefInfo(ii).numlabel = 0
    Call GetStatement(tmpdata, 0, tmpdatalong, tmpKdef, KdefInfo(ii).numlabel)
    Call GenLabel(tmpKdef, KdefInfo(ii).numlabel)
    Call GetVarName(ii)
    For Each stat In KdefInfo(ii).kdef
        Call getStatementnote(stat)
    Next
    
    
End Sub


' 把data中的二进制翻译成指令保存在集合
Public Sub DatatoKdef(i As Long)
Dim stat As Statement
    While KdefInfo(i).kdef.Count > 0
        KdefInfo(i).kdef.Remove 1
    Wend
        
    Call GetStatement(KdefInfo(i).data, 0, KdefInfo(i).DataLong, KdefInfo(i).kdef, KdefInfo(i).numlabel)
    Call GenLabel(KdefInfo(i).kdef, KdefInfo(i).numlabel)
    Call GetVarName(i)
    For Each stat In KdefInfo(i).kdef
        Call getStatementnote(stat)
    Next
End Sub

' 产生名字集合

Public Sub GetVarName(i As Long)
Dim stat As Statement
Dim s As String
Dim id As Long
Dim j As Long
    On Error Resume Next
    Set KdefName = Nothing
    Set KdefName = New Collection
    For Each stat In KdefInfo(i).kdef
        If stat.id = 50 Then
            If stat.data(0) = 6 Then
                s = ""
                For j = 2 To 6
                    s = s & Chr(stat.data(j))
                Next j
                's = GetXchar(stat.Data(2), stat.Data(3), stat.Data(4), stat.Data(5), stat.Data(6))
                id = stat.data(1)
                KdefName.Add id & ":" & s, "ID" & id
            End If
        End If
    Next

    

End Sub


Public Sub GetStatement(kdefdata() As Integer, ByVal startpos As Long, ByVal endpos As Long, kdef As Collection, numlabel As Long)
Dim k As Integer
Dim newstat As Statement
Dim labelstat As Statement
Dim i As Long
Dim notestr As String
Dim Length As Long
Dim tempstr As String
Dim tmp1 As Long, tmp2 As Long, tmp3 As Long, tmp4 As Long
On Error Resume Next
    If startpos = endpos Then Exit Sub
      
    k = kdefdata(startpos)
    
    If k = &HFFFF Then
        Length = 1
        kdef.Add setstatement(kdefdata, startpos, k, 0, Length - 1)
        Call GetStatement(kdefdata, startpos + Length, endpos, kdef, numlabel)
        Exit Sub
    ElseIf k <= KdefNum Then
        Length = StatAttrib(k).Length
        If k = 50 Then
            If kdefdata(startpos + 1) = 43 Then
                If kdefdata(startpos + 3) = 201 Then
                    tmp1 = kdefdata(startpos + 4)
                    tmp2 = kdefdata(startpos + 5)
                    tmp3 = kdefdata(startpos + 6)
                    tmp4 = kdefdata(startpos + 7)
                    
                'NewTalk(e3, e4, e5, e6 mod 100, (e6 mod 100) div 10, e6 div 100, 0);
                    kdefdata(startpos) = 68
                    kdefdata(startpos + 1) = tmp1
                    kdefdata(startpos + 2) = tmp2
                    kdefdata(startpos + 3) = tmp3
                    kdefdata(startpos + 4) = tmp4 Mod 100
                    kdefdata(startpos + 5) = (tmp4 Mod 100) / 10
                    kdefdata(startpos + 6) = tmp4 / 100
                    kdefdata(startpos + 7) = 0
                    k = kdefdata(startpos)
                End If
            End If
        End If
        If check_it(startpos, endpos, Length) = True Then
            If StatAttrib(k).isGoto = 0 Then
                Set newstat = setstatement(kdefdata, startpos, k, 0, Length - 1)
                newstat.isGoto = 0
                kdef.Add newstat
                Call GetStatement(kdefdata, startpos + Length, endpos, kdef, numlabel)
                Exit Sub
            Else
                Set newstat = setstatement(kdefdata, startpos, k, 0, Length - 1)
                newstat.gotoLabel = ":Label" & numlabel
                numlabel = numlabel + 1
                newstat.yesOffset = StatAttrib(k).yesOffset
                newstat.noOffset = StatAttrib(k).noOffset
                
                If kdefdata(startpos + newstat.yesOffset) <> 0 And kdefdata(startpos + newstat.noOffset) = 0 Then
                    newstat.GotoAddress = kdefdata(startpos + newstat.yesOffset)
                    newstat.note = newstat.note & LoadResStr(501) & newstat.gotoLabel   ' 是则跳转
                    newstat.isGoto = 1
                ElseIf kdefdata(startpos + newstat.yesOffset) = 0 And kdefdata(startpos + newstat.noOffset) <> 0 Then
                    newstat.GotoAddress = kdefdata(startpos + newstat.noOffset)
                    newstat.note = newstat.note & LoadResStr(502) & newstat.gotoLabel '否则跳转
                    newstat.isGoto = 2
                Else
                    newstat.gotoLabel = ""
                    numlabel = numlabel - 1
                    newstat.isGoto = 0
                    kdef.Add newstat
                    newstat.note = newstat.note & LoadResStr(503) '"不规范的条件转移"
                    Call GetStatement(kdefdata, startpos + Length, endpos, kdef, numlabel)
                    Exit Sub
                End If
                kdef.Add newstat
                
                Call GetStatement(kdefdata, startpos + Length, endpos, kdef, numlabel)
                Exit Sub
                
                
'                Call GetStatement(kdefdata, startpos + length, startpos + length + newstat.GotoAddress, kdef, numlabel)
'                Set labelstat = New Statement
'                labelstat.islabel = True
'                labelstat.note = newstat.gotoLabel
'                kdef.Add labelstat
'                Call GetStatement(kdefdata, startpos + length + newstat.GotoAddress, endpos, kdef, numlabel)
'                Exit Sub
            End If
        End If
    End If
        
    Length = 1
    If check_it(startpos, endpos, Length) = True Then
        MsgBox LoadResStr(504)  '  "存在未知语句"
            kdef.Add setstatement(kdefdata, startpos, k, 0, Length - 1)
            Call GetStatement(kdefdata, startpos + Length, endpos, kdef, numlabel)
        Exit Sub
    End If
    
End Sub



' 生成一个新语句集合
Private Function setstatement(kdefdata() As Integer, ByVal startpos As Long, ByVal KeyWord As Long, ByVal isGoto As Long, ByVal DataNum As Long) As Statement
Dim i As Long
Dim newstat As Statement
    Set newstat = New Statement
    newstat.id = KeyWord
    newstat.isGoto = isGoto
    newstat.DataNum = DataNum
    For i = 0 To newstat.DataNum - 1
        newstat.data(i) = kdefdata(startpos + 1 + i)
    Next i
    Set setstatement = newstat
End Function



' 检查指令长度是否符合
Private Function check_it(startpos As Long, endpos As Long, Length As Long)
Dim i As Long
Dim datastr As String
Dim auto As Long
    auto = 1
    check_it = False
    If endpos - startpos >= Length Then
        If auto = 0 Then
'            For i = 1 To length
'                datastr = datastr & Chr(13) & Hex(kdefword(startpos + i - 1)) & "-"
'            Next i
'            If MsgBox(datastr, vbOKCancel) = vbOK Then
'                check_it = True
'            End If
        Else
            check_it = True
        End If
    End If
End Function

' 读取指令属性
Public Sub ReadStatementAttrib()
Dim i As Long
Dim tmpStr() As String

    ReDim StatAttrib(KdefNum)

    For i = 0 To KdefNum
        tmpStr = Split(GetINIStr("KdefAttrib", "attrib" & i), " ")
        StatAttrib(i).Length = tmpStr(1)
        StatAttrib(i).isGoto = tmpStr(2)
        StatAttrib(i).yesOffset = tmpStr(3)
        StatAttrib(i).noOffset = tmpStr(4)
        StatAttrib(i).notes = tmpStr(5)
    Next i
End Sub


' 翻译指令含义
Public Sub getStatementnote(stat As Statement)
Dim k As Integer
Dim tempstr, tmps, tmps2 As String
Dim tmpstat As Long
    On Error Resume Next
    If stat.islabel = True Then Exit Sub
    
    k = stat.id
    
    Select Case k
    
    Case &H0
        stat.note = ":" & LoadResStr(505) '" 空语句"
    Case &H1
        tempstr = Person(HeadtoPerson(stat.data(1)).Item(1)).Name1
        If stat.data(0) < 0 Then
            tmpstat = 1 - stat.data(0)
        Else
            tmpstat = stat.data(0)
        End If
        stat.note = "[" & tempstr & "]" & LoadResStr(506) & Talk(tmpstat) ' 说话
    Case &H2
        stat.note = LoadResStr(507) & "[" & Things(stat.data(0)).Name1 & "][" & stat.data(1) & "]" '得到物品
    Case &H3
        stat.note = LoadResStr(508) '"修改事件定义"
        If stat.data(0) = &HFFFE Then
            stat.note = stat.note & ":" & LoadResStr(509) ' "当前场景"
        Else
            stat.note = stat.note & ":" & LoadResStr(510) & "[" & Big5toUnicode(Scene(stat.data(0)).Name1, 10) & "]" '场景"
        End If
        If stat.data(1) = &HFFFE Then
            stat.note = stat.note & ":" & LoadResStr(511) '"：当前场景事件编号"
        Else
            stat.note = stat.note & ":" & LoadResStr(512) & " [" & stat.data(1) & "]"
        End If
    Case &H4
        stat.note = LoadResStr(513) & "[" & Things(stat.data(0)).Name1 & "]？" '是否使用物品
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel '是则跳转到
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel '"否则跳转到"
        End If
    Case &H5
        stat.note = LoadResStr(516) '"是否选择战斗？"
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H6
        If stat.data(0) < warnum Then
            tmps = WarData(stat.data(0)).Name
            stat.note = LoadResStr(517) & "[" & tmps & "]" '"战斗"
        Else
            stat.note = LoadResStr(517) & "[" & stat.data(0) & "]" '"战斗"
        End If
        
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H7
        stat.note = LoadResStr(518) ' "无条件返回"
    Case &H8
        stat.note = LoadResStr(519) '"切换大地图音乐"
    Case &H9
        stat.note = LoadResStr(520) '"是否要求加入?"
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &HA
        stat.note = LoadResStr(521) & "[" & Person(stat.data(0)).Name1 & "]" '"加入人物"
    Case &HB
        stat.note = LoadResStr(522)  '"是否住宿"
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &HC
        stat.note = LoadResStr(523) '"住宿休息"
    Case &HD
        stat.note = LoadResStr(524) '"重新显示场景"
    Case &HE
        stat.note = LoadResStr(525) '"场景变黑"
    Case &HF
        stat.note = LoadResStr(526)  '"战斗失败，死亡"
    Case &H10
        stat.note = LoadResStr(527) & "[" & Person(stat.data(0)).Name1 & "]" ' "队伍是否有"
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H11
    
        stat.note = LoadResStr(528) '"修改场景贴图"
        If stat.data(0) = &HFFFE Then
            stat.note = stat.note & ":" & LoadResStr(509) ' "当前场景"
        Else
            stat.note = stat.note & ":" & LoadResStr(510) & "[" & Big5toUnicode(Scene(stat.data(0)).Name1, 10) & "]"
        End If
    
       stat.note = stat.note & LoadResStr(529) & _
                           Hex(stat.data(1)) & LoadResStr(530) & Hex(stat.data(2)) & "-" & _
                           Hex(stat.data(3))
    
    Case &H12
        stat.note = LoadResStr(531) & "[" & Things(stat.data(0)).Name1 & "]" '"是否有物品"
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H13
        stat.note = LoadResStr(532) & Hex(stat.data(0)) & "-" & Hex(stat.data(1)) '"主角移动至"
    Case &H14
        stat.note = LoadResStr(533) '"队伍是否满？"
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H15
        stat.note = "[" & Person(stat.data(0)).Name1 & "]" & LoadResStr(534) '"离队"
    Case &H16
        stat.note = LoadResStr(535) '"内力降为0"
    Case &H17
        stat.note = LoadResStr(536) & "[" & Person(stat.data(0)).Name1 & "]" & LoadResStr(537) & "[" & stat.data(1) & "]" '设置人物用毒
    Case &H19
        stat.note = LoadResStr(538) & stat.data(0) & "-" & stat.data(1) & "--" & stat.data(2) & "-" & stat.data(3)  '"场景移动从"
    Case &H1A
        stat.note = LoadResStr(539) '"增加场景事件编号的三个触发事件编号"
    Case &H1B
        stat.note = LoadResStr(540) '"显示动画"
    Case &H1C
        stat.note = LoadResStr(541) & Person(stat.data(0)).Name1 & LoadResStr(542) & _
                            stat.data(1) & "-" & stat.data(2)   '判断品德
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H1D
        stat.note = LoadResStr(541) & Person(stat.data(0)).Name1 & LoadResStr(543) & _
                            stat.data(1) & "-" & stat.data(2) '武力
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H1E
        stat.note = LoadResStr(544) & stat.data(0) & "-" & stat.data(1) & _
                            "--" & stat.data(2) & "-" & stat.data(3)  '主角走动
    Case &H1F
        stat.note = LoadResStr(545) & stat.data(0) '判断银子是否够
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H20
        stat.note = LoadResStr(546) & "[" & Things(stat.data(0)).Name1 & "]" & "+" & "[" & stat.data(1) & "]"  '物品
    Case &H21
        stat.note = Person(stat.data(0)).Name1 & LoadResStr(547) & WuGong(stat.data(1)).Name1 '学会武功"
    Case &H22
        stat.note = Person(stat.data(0)).Name1 & LoadResStr(548) & stat.data(1) '增加资质
    Case &H23
        stat.note = LoadResStr(549) & Person(stat.data(0)).Name1 & LoadResStr(550) & stat.data(1) & ":" & _
                            WuGong(stat.data(2)).Name1 & LoadResStr(551) & stat.data(3)  '设置武功
    Case &H24
        If stat.data(0) < 256 Then
            stat.note = LoadResStr(552) & stat.data(0) '"主角性别是否为"
        Else
            stat.note = StrUnicode2("跳转变量JMP是否为0")
        End If
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H25
        stat.note = LoadResStr(553) & stat.data(0)  '"增加道德"
    Case &H26
        stat.note = LoadResStr(554) '"修改图层贴图"
    Case &H27
        stat.note = LoadResStr(555) & Big5toUnicode(Scene(stat.data(0)).Name1, 10) '"打开场景"
    Case &H28
        stat.note = LoadResStr(556) & stat.data(0) '"主角面向方向"
    Case &H29
        stat.note = Person(stat.data(0)).Name1 & LoadResStr(507) & _
                           Things(stat.data(1)).Name1 & stat.data(2)  ' 得到物品
    Case &H2A
        stat.note = LoadResStr(557) '"队伍中是否有女性"
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H2B
        stat.note = LoadResStr(558) & Things(stat.data(0)).Name1 '"是否有物品"
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H2C
        stat.note = LoadResStr(540) '"播放动画"
    Case &H2D
        stat.note = Person(stat.data(0)).Name1 & LoadResStr(559) & stat.data(1)  '"增加轻功"
    Case &H2E
        stat.note = Person(stat.data(0)).Name1 & LoadResStr(560) & stat.data(1) '增加内力"
    Case &H2F
        stat.note = Person(stat.data(0)).Name1 & LoadResStr(561) & stat.data(1) '"增加武力"
    Case &H30
        stat.note = Person(stat.data(0)).Name1 & LoadResStr(562) & stat.data(1) '增加生命
    Case &H31
        stat.note = Person(stat.data(0)).Name1 & LoadResStr(585) & stat.data(1) '"内力属性"
    Case &H32
        stat.note = GetNote50(stat.data(0), stat.data(1), stat.data(2), stat.data(3), stat.data(4), stat.data(5), stat.data(6))
        
    Case &H33
        stat.note = LoadResStr(563) '‘"问软体娃娃"
    Case &H34
        stat.note = LoadResStr(564) '"看道德"
    Case &H35
        stat.note = LoadResStr(565) ' "看声望"
    Case &H36
        stat.note = LoadResStr(566) '"开放其它场景"
    Case &H37
        stat.note = LoadResStr(567) & stat.data(0) & LoadResStr(568) & stat.data(1) '事件d*编号是否为
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H38
        stat.note = LoadResStr(569) & stat.data(0) '"提高声望值"
    Case &H39
        stat.note = LoadResStr(570) ' "高昌迷宫劈门"
    Case &H3A
        stat.note = LoadResStr(571) '"武道大会比武"
    Case &H3B
        stat.note = LoadResStr(572) '"全体队员离队并找不着"
    Case &H3C
        stat.note = LoadResStr(573) & stat.data(0) & LoadResStr(574) & stat.data(1) & _
                               LoadResStr(575) & stat.data(2) '判断场景事件位置是否有贴图
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H3D
        stat.note = LoadResStr(576) '"判断是否放完14本书"
        If stat.isGoto = 1 Then
            stat.note = stat.note & LoadResStr(514) & stat.gotoLabel
        Else
            stat.note = stat.note & LoadResStr(515) & stat.gotoLabel
        End If
    Case &H3E
        stat.note = LoadResStr(577) ' "播放进入时空机动画"
    Case &H3F
        stat.note = LoadResStr(549) & Person(stat.data(0)).Name1 & _
                            LoadResStr(578) & stat.data(1)  '设置性别
    Case &H40
        stat.note = LoadResStr(579) '"韦小宝卖东西"
    Case &H41
        stat.note = LoadResStr(580) ' "韦小宝去其它客栈"
    Case &H42
        stat.note = LoadResStr(581) & stat.data(0) '"播放音乐"
    Case &H43
        stat.note = LoadResStr(582) & stat.data(0)  '"播放音效"
    Case &H44
    'MsgBox NewHeadNum
    If stat.data(0) < PersonNum Then
        If stat.data(2) = 0 Then
            stat.note = "[" & Person(stat.data(0)).Name1 & "(False)" & "]:" & Talk(stat.data(1)) & "(" & StrUnicode2("背景") & stat.data(6) & ")"
        ElseIf stat.data(2) = -2 Then
            stat.note = "[" & Person(stat.data(0)).Name1 & "]:" & Talk(stat.data(1)) & "(" & StrUnicode2("背景") & stat.data(6) & ")"
        Else
'            stat.note = "[" & Nam(stat.Data(2)) & "(" & Person(stat.Data(0)).Name1 & ")" & "]:" & Talk(stat.Data(1))
            'MsgBox nam(1)
            stat.note = "[" & nam(stat.data(2)) & "(" & stat.data(0) & ")" & "]:" & Talk(stat.data(1)) & "(" & StrUnicode2("背景") & stat.data(6) & ")"
        End If
    Else
            stat.note = "[" & stat.data(0) & "(False)" & "]:" & Talk(stat.data(1)) & "(" & StrUnicode2("背景") & stat.data(6) & ")"
    End If
    Case &H45
        Select Case stat.data(0)
            Case 0
                tmps = StrUnicode2("人物")
                tmps2 = Person(stat.data(1)).Name1
           Case 1
                tmps = StrUnicode2("物品")
                tmps2 = Things(stat.data(1)).Name1
            Case 2
                tmps = StrUnicode2("场景")
                tmps2 = Big5toUnicode(Scene(stat.data(0)).Name1, 10)
            Case 3
                tmps = StrUnicode2("武功")
                tmps2 = WuGong(stat.data(0)).Name1
        End Select
        stat.note = StrUnicode2("替换名称") & "[" & tmps & " " & tmps2 & "(" & stat.data(1) & ")" & "]" & "=" & StrUnicode2("name") & "[" & stat.data(2) & "]"
    Case &H46
        stat.note = StrUnicode2("显示字幕[talk(") & stat.data(0) & ")]" & StrUnicode2(" 背景颜色") & stat.data(1)
    Case &H47
        tmps2 = Big5toUnicode(Scene(stat.data(0)).Name1, 10)
        stat.note = StrUnicode2("场景移动:") & tmps2 & "(" & stat.data(1) & "," & stat.data(2) & ")"
    Case &H48
    Select Case stat.data(1)
        Case 0
            tmps = StrUnicode2("主角")
        Case 1
            tmps = StrUnicode2("最高")
        Case 2
            tmps = StrUnicode2("最低")
        Case 3
            tmps = StrUnicode2("平均")
        End Select
        If stat.data(2) > 0 Then
            tmps2 = "+" & stat.data(2)
        Else
            tmps2 = stat.data(2)
        End If
        stat.note = StrUnicode2("设置人物:") & Person(stat.data(0)).Name1 & "(" & stat.data(0) & ")_" & tmps & StrUnicode2("等级") & tmps2 & "(min=" & stat.data(3) & ",max=" & stat.data(4) & ")"
    Case &HFFFF
        stat.note = LoadResStr(583) '"事件结束"
    Case Else
        stat.note = LoadResStr(584) '"未知指令"
    End Select
    stat.note = "  " & k & "(" & Hex(k) & ")" & ":" & stat.note
End Sub

' 翻译指令50的含义
' 参数为指令50的7个参数
Private Function GetNote50(ByVal x0 As Long, ByVal X1 As Long, ByVal X2 As Long, ByVal x3 As Long, ByVal x4 As Long, ByVal x5 As Long, ByVal x6 As Long) As String
Dim s1 As String, s2 As String
Dim tmps As String
Dim i As Long
Dim kk() As String
Dim MemNum As Long

    If (GetINILong("50_43", Val(X2)) <> -65536) And (x0 = 43) Then
        GetNote50 = GetINIStr("50_43", Val(X2))
        On Error Resume Next
        'MsgBox ValueX(X1, Val(0), "#0", X3)
        GetNote50 = Replace(GetNote50, "#0", ValueX(X1, Val(1), "#0", x3))
        GetNote50 = Replace(GetNote50, "#1", ValueX(X1, Val(2), "#1", x4))
        GetNote50 = Replace(GetNote50, "#2", ValueX(X1, Val(3), "#2", x5))
        GetNote50 = Replace(GetNote50, "#3", ValueX(X1, Val(4), "#3", x6))
        GetNote50 = extra(GetNote50, Val(0), x3)
        GetNote50 = extra(GetNote50, Val(1), x4)
        GetNote50 = extra(GetNote50, Val(2), x5)
        GetNote50 = extra(GetNote50, Val(3), x6)
        GetNote50 = Replace(GetNote50, "#0", x3)
        GetNote50 = Replace(GetNote50, "#1", x4)
        GetNote50 = Replace(GetNote50, "#2", x5)
        GetNote50 = Replace(GetNote50, "#3", x6)
        Exit Function
    End If
    s1 = GetINIStr("Kdef50", "sub" & x0)
    If s1 = "" Then
        s1 = GetINIStr("Kdef50", "Other")
    End If
    Select Case x0
    Case 0
        s2 = StrVar(X1) & "=" & X2
    Case 1
        s2 = StrUnicode2("数组") & x3 & "(" & VarX(X1, 0, x4) & ")=" & VarX(X1, 1, x5) & IIf(X2, "(Byte)", "(Int)")
    Case 2
        s2 = StrVar(x5) & StrUnicode2("=数组") & x3 & "(" & VarX(X1, 0, x4) & ")" & IIf(X2, "(Byte)", "(Int)")
    Case 3
        Select Case X2
        Case 0
            tmps = "+"
        Case 1
            tmps = "-"
        Case 2
            tmps = "*"
        Case 3
            tmps = "/"
        End Select
        s2 = StrVar(x3) & "=" & StrVar(x4) & tmps & VarX(X1, 0, x5)
    
    Case 4
        Select Case X2
        Case 0
            tmps = "If" & StrVar(x3) & "<" & VarX(X1, 0, x4) & " then JMP=0 else JMP=1"
        Case 1
            tmps = "If" & StrVar(x3) & "<=" & VarX(X1, 0, x4) & " then JMP=0 else JMP=1"
        Case 2
            tmps = "If" & StrVar(x3) & "=" & VarX(X1, 0, x4) & "  then JMP=0 else JMP=1"
        Case 3
            tmps = "If" & StrVar(x3) & "<>" & VarX(X1, 0, x4) & " then JMP=0 else JMP=1"
        Case 4
            tmps = "If" & StrVar(x3) & ">=" & VarX(X1, 0, x4) & " then JMP=0 else JMP=1"
        Case 5
            tmps = "If" & StrVar(x3) & ">" & VarX(X1, 0, x4) & " then JMP=0 else JMP=1"
        Case 6
            tmps = " JMP=0"
        Case 7
            tmps = " JMP=1"
        End Select
        s2 = tmps
    Case 6
        s2 = "[" & "X" & X1 & "]:" & GetXchar(X2, x3, x4, x5, x6)
        's2 = "[" & "X" & X1 & "]:" & tmps
    Case 8
        If X1 = 0 Then
           s2 = " Str" & StrVar(x3) & "=talk(" & VarX(X1, 0, X2) & ")" & " [" & Talk(X2) & "]"
           Else
           s2 = " Str" & StrVar(x3) & "=talk(" & VarX(X1, 0, X2) & ")"
        End If
    Case 9
        s2 = " Str" & StrVar(X2) & "=CStr(" & VarX(X1, 0, x4) & ") FORMAT=Str" & StrVar(x3)
    Case 10
        s2 = StrVar(X2) & "=Len(Str" & StrVar(X1) & " )"
    Case 11
        s2 = "Str" & StrVar(X1) & "=Str" & StrVar(X2) & "+Str" & StrVar(x3)
    Case 12
        s2 = "Str" & StrVar(X2) & "=" & VarX(X1, 0, x3) & StrUnicode2("个空格")
    Case 16
        Select Case X2
        Case 0
            tmps = StrUnicode2("人物")
        Case 1
            tmps = StrUnicode2("物品")
        Case 2
            tmps = StrUnicode2("场景")
        Case 3
            tmps = StrUnicode2("武功")
        Case 4
            tmps = StrUnicode2("小宝商店")
        End Select
        s2 = tmps & VarX(X1, 0, x3) & StrUnicode2("属性偏移") & VarX(X1, 1, x4) & "=" & VarX(X1, 2, x5)
    Case 17
        Select Case X2
        Case 0
            tmps = StrUnicode2("人物")
        Case 1
            tmps = StrUnicode2("物品")
        Case 2
            tmps = StrUnicode2("场景")
        Case 3
            tmps = StrUnicode2("武功")
        Case 4
            tmps = StrUnicode2("小宝商店")
        End Select
        s2 = StrVar(x5) & "=" & tmps & VarX(X1, 0, x3) & StrUnicode2("属性偏移") & VarX(X1, 1, x4)
    Case 18
        s2 = StrUnicode2("队友") & VarX(X1, 0, X2) & "=" & VarX(X1, 1, x3)
    Case 19
        s2 = StrVar(x3) & "=" & StrUnicode2("队友") & VarX(X1, 0, X2)
    Case 20
        s2 = StrVar(x3) & "=" & StrUnicode2("物品编号") & VarX(X1, 0, X2) & StrUnicode2("数量")
    Case 21
        s2 = StrUnicode2("场景") & VarX(X1, 0, X2) & StrUnicode2("场景事件") & VarX(X1, 1, x3) & StrUnicode2("属性") & VarX(X1, 2, x4) & "=" & VarX(X1, 3, x5)
    Case 22
        s2 = StrVar(x5) & "=" & StrUnicode2("场景") & VarX(X1, 0, X2) & StrUnicode2("场景事件") & VarX(X1, 1, x3) & StrUnicode2("属性") & VarX(X1, 2, x4)
      
    Case 23
        s2 = StrUnicode2("场景") & VarX(X1, 0, X2) & StrUnicode2("层") & VarX(X1, 1, x3) & StrUnicode2("坐标(") & VarX(X1, 2, x4) & "," & VarX(X1, 3, x5) & ")=" & VarX(X1, 4, x6)
    Case 24
        s2 = StrVar(x6) & "=" & StrUnicode2("场景") & VarX(X1, 0, X2) & StrUnicode2("层") & VarX(X1, 1, x3) & StrUnicode2("坐标(") & VarX(X1, 2, x4) & "," & VarX(X1, 3, x5) & ")"
        
    Case 25
        
        s2 = "[" & Hex(Long2int(x4)) & "-" & Hex(Long2int(x3)) & "+" & VarX(X1, 1, x6) & "]=" & VarX(X1, 0, x5) & IIf(X2, "(Byte)", "(Int)")

        MemNum = GetINILong("50memory", "MemNum")
        For i = 0 To MemNum - 1
            tmps = GetINIStr("50memory", "Mem" & i)
            kk = Split(tmps, " ")
            If StrComp(GetStrLng(CStr(Hex(Long2int(x4))), 4) & GetStrLng(CStr(Hex(Long2int(x3))), 4), GetStrLng(kk(0), 8)) = 0 Then
                s2 = "[" & kk(1) & "+" & VarX(X1, 1, x6) & "]=" & VarX(X1, 0, x5) & IIf(X2, "(Byte)", "(Int)")
                Exit For
            End If
        Next i
    Case 26
        s2 = StrVar(x5) & "=[" & Hex(Long2int(x4)) & "-" & Hex(Long2int(x3)) & "+" & VarX(X1, 0, x6) & "]" & IIf(X2, "(Byte)", "(Int)")
        MemNum = GetINILong("50memory", "MemNum")
        For i = 0 To MemNum - 1
            tmps = GetINIStr("50memory", "Mem" & i)
            kk = Split(tmps, " ")
            If StrComp(GetStrLng(CStr(Hex(Long2int(x4))), 4) & GetStrLng(CStr(Hex(Long2int(x3))), 4), GetStrLng(kk(0), 8)) = 0 Then
                s2 = "[" & kk(1) & "+" & VarX(X1, 1, x6) & "]=" & VarX(X1, 0, x5) & IIf(X2, "(Byte)", "(Int)")
                Exit For
            End If
        Next i
    Case 27
        Select Case X2
        Case 0
            tmps = StrUnicode2("人物")
        Case 1
            tmps = StrUnicode2("物品")
        Case 2
            tmps = StrUnicode2("场景")
        Case 3
            tmps = StrUnicode2("武功")
        End Select
        s2 = "Str" & x4 & "=" & tmps & VarX(X1, 0, x3)
     'sub28=取当前人物战斗序号
    Case 28
        s2 = "=" & StrVar(X1)
    'sub29=选择攻击目标
    Case 29
        If x5 = 0 Then
           s2 = StrUnicode2("战斗序号") & VarX(X1, 0, X2) & StrUnicode2("步数") & VarX(X1, 1, x3) & StrUnicode2("返回值") & StrVar(x4) & "" & StrUnicode2("显示")
         Else
           s2 = StrUnicode2("战斗序号") & VarX(X1, 0, X2) & StrUnicode2("步数") & VarX(X1, 1, x3) & StrUnicode2("返回值") & StrVar(x4) & "" & StrUnicode2("不显示")
        End If
    'sub30=读取人物战斗属性
    Case 30
       s2 = StrUnicode2("战斗序号") & VarX(X1, 0, X2) & StrUnicode2("偏移") & VarX(X1, 1, x3) & "=" & StrVar(x4)
     '写入人物战斗属性
    Case 31
       s2 = StrUnicode2("战斗序号") & VarX(X1, 0, X2) & StrUnicode2("偏移") & VarX(X1, 1, x3) & "=" & VarX(X1, 2, x4)
    Case 32
        s2 = StrUnicode2("下一条指令参数") & VarX(X1, 0, x3) & "=" & StrVar(X2)
    Case 33
        s2 = StrUnicode2("显示字符串") & X2 & " (" & VarX(X1, 0, x3) & "," & VarX(X1, 1, x4) & ")" & "Color" & VarX(X1, 2, x5)
    Case 34
        s2 = StrUnicode2("位置(") & VarX(X1, 0, X2) & "," & VarX(X1, 1, x3) & ") Width=" & VarX(X1, 2, x4) & " Height=" & VarX(X1, 3, x5) & " Transparency=" & VarX(X1, 4, x6)
    Case 35
        'Debug.Print X1
        'If Val(X1) <> Val(-1) Then
        '   s2 = StrVar(X1) & StrUnicode2("=键值")
        'Else
            s2 = StrUnicode2("键值") & StrVar(X1) & StrUnicode2(" 鼠标坐标(") & StrVar(X2) & "," & StrVar(x3) & ")"
        'End If
    Case 36
        s2 = StrUnicode2("显示字符串") & X2 & " (" & VarX(X1, 0, x3) & "," & VarX(X1, 1, x4) & ")" & "Color" & VarX(X1, 2, x5)
    Case 37
        s2 = StrUnicode2("延时") & VarX(X1, 0, X2)
    Case 38
        s2 = StrVar(x3) & "=Random(" & VarX(X1, 0, X2) & ")"
    Case 39
        s2 = StrUnicode2("菜单个数") & VarX(X1, 0, X2) & StrVar(x3) & "()" & StrUnicode2("返回=") & StrVar(x4) & StrUnicode2("显示位置(") & VarX(X1, 1, x5) & "," & VarX(X1, 2, x6) & ")"
    Case 40
        s2 = StrUnicode2("菜单个数") & VarX(X1, 0, X2) & StrVar(x3) & "()" & StrUnicode2("返回=") & StrVar(x4) & StrUnicode2("显示位置(") & VarX(X1, 1, x5) & "," & VarX(X1, 2, x6) & ")"
    Case 41
        Select Case X2
        Case 0
            tmps = StrUnicode2("场景图片")
        Case 1
            tmps = StrUnicode2("头像图片")
        Case 2
            tmps = StrUnicode2("物品图片")
        End Select
        s2 = tmps & VarX(X1, 2, x5) & StrUnicode2(" 位置(") & VarX(X1, 0, x3) & "," & VarX(X1, 1, x4) & ")"
    Case 42
        s2 = "(" & VarX(X1, 0, X2) & "," & VarX(X1, 1, x3) & ")"
    Case 43
        s2 = "Call Sub" & VarX(X1, 0, X2) & "(" & VarX(X1, 1, x3) & "," & VarX(X1, 2, x4) & "," & VarX(X1, 3, x5) & "," & VarX(X1, 4, x6) & ")"
    '44 指令：播放效果
    Case 44
        s2 = StrUnicode2("序号") & VarX(X1, 0, X2) & StrUnicode2("动作类型") & VarX(X1, 1, x3) & StrUnicode2("效果编号") & VarX(X1, 2, x4)
    '45 指令：显示数字
    Case 45
    If x3 = 0 Then
        s2 = StrUnicode2("数字颜色") & VarX(X1, 0, X2) & StrUnicode2("闪烁") & " " & StrUnicode2("闪烁颜色") & VarX(X1, 1, x4)
      Else
        s2 = StrUnicode2("数字颜色") & VarX(X1, 0, X2) & StrUnicode2("不闪烁") & " " & StrUnicode2("闪烁颜色") & VarX(X1, 1, x4)
    End If
    '46 指令：设定效果层
    Case 46
    If x6 = 0 Then
        s2 = StrUnicode2("起始点") & "(" & VarX(X1, 0, X2) & "," & VarX(X1, 1, x3) & ")" & StrUnicode2("长度") & "(" & VarX(X1, 2, x4) & "," & VarX(X1, 3, x5) & ")" & StrUnicode2("无效果")
      Else
        s2 = StrUnicode2("起始点") & "(" & VarX(X1, 0, X2) & "," & VarX(X1, 1, x3) & ")" & StrUnicode2("长度") & "(" & VarX(X1, 2, x4) & "," & VarX(X1, 3, x5) & ")" & StrUnicode2("有效果")
    End If
     '47重置战场贴图
    Case 47
        s2 = StrUnicode2("战斗序号") & "=" & VarX(X1, 0, X2)
    Case 48
        s2 = StrVar(X1) & " -- " & StrVar(X1 + X2 - 1)
    Case 49
        s2 = "[" & Hex(Long2int(X2)) & "-" & Hex(Long2int(X1)) & "]" & StrUnicode2("起始编号") & StrVar(x3) & " " & StrUnicode2("个数") & x4 & "=" & StrVar(x5)
    End Select
l4:
       
    GetNote50 = s1 & " " & s2
    
End Function

' x的第n位为0 返回s1 为1 返回s2
' x 指示器
' n x的第n位
Private Function VarValue(ByVal x As Long, ByVal n As Long, s1 As String, s2 As String) As String
Dim i As Long
    i = 2 ^ n
    If (x And i) = 0 Then
        VarValue = s1
    Else
        VarValue = s2
    End If
End Function


Private Function VarX(ByVal x As Long, ByVal n As Long, ByVal v As Long) As String
    VarX = VarValue(x, n, CStr(v), StrVar(v))
End Function
'new,getvalue in note
Private Function ValueX(ByVal x As Long, ByVal n As Long, tmpStr As String, v As Long) As String
Dim i As Long
    'get boooooooooolean string
    i = 2 ^ n
    If (x And i) = 0 Then
        ValueX = tmpStr
    Else
        ValueX = "[" & "X" & v & "]"
    End If
    'MsgBox ValueX
End Function

Private Function StrVar(ByVal x As Long) As String
Dim s As String
    On Error GoTo Label1:
    s = KdefName.Item("ID" & x)
    StrVar = s
    Exit Function
Label1:
    StrVar = "[" & "X" & x & "]"
End Function

'extra,for more expression
Public Function extra(tmpStr As String, ByVal index1 As Long, ByVal index2 As Long)
Attribute extra.VB_Description = "增加的部分啦"
On Error Resume Next
tmpStr = Replace(tmpStr, "name(#" & index1 & ")", nam(index2))
tmpStr = Replace(tmpStr, "talk(#" & index1 & ")", Talk(index2))
tmpStr = Replace(tmpStr, "person(#" & index1 & ")", Person(index2).Name1)
tmpStr = Replace(tmpStr, "things(#" & index1 & ")", Things(index2).Name1)
tmpStr = Replace(tmpStr, "scene(#" & index1 & ")", Big5toUnicode(Scene(index2).Name1, 10))
tmpStr = Replace(tmpStr, "magic(#" & index1 & ")", WuGong(index2).Name1)
tmpStr = Replace(tmpStr, "war(#" & index1 & ")", WarData(index2).Name)
extra = tmpStr
End Function
Public Function GetStrLng(StrOld As String, StrLong As String) As String

    If Len(StrOld) < StrLong Then
        GetStrLng = String(StrLong - Len(StrOld), "0") & StrOld
    End If
End Function
