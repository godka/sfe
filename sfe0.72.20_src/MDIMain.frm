VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7830
   ClientLeft      =   3765
   ClientTop       =   2115
   ClientWidth     =   10485
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7575
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7532
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_file 
      Caption         =   "文件"
      Begin VB.Menu mnu_setpath 
         Caption         =   "设置游戏路径"
      End
      Begin VB.Menu mnu_setconfig 
         Caption         =   "设置修改器属性"
      End
      Begin VB.Menu bb 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu mnu_GameEdit 
      Caption         =   "游戏编辑"
      Begin VB.Menu mnu_R_Edit 
         Caption         =   "基本进度R*编辑"
      End
      Begin VB.Menu mnu_SEdit 
         Caption         =   "场景编辑"
      End
      Begin VB.Menu mnu_KdefEdit 
         Caption         =   "对话/事件编辑"
      End
      Begin VB.Menu mnu_War 
         Caption         =   "战斗编辑"
      End
      Begin VB.Menu mnu_Team 
         Caption         =   "复刻版属性编辑"
      End
   End
   Begin VB.Menu mnu_KdefPopup 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnu_kdef_copy 
         Caption         =   "复制指令"
      End
      Begin VB.Menu mnu_kdef_copyAll 
         Caption         =   "复制整个事件"
      End
      Begin VB.Menu mnu_kdef_Paste 
         Caption         =   "粘贴指令"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_kdef_PasteALL 
         Caption         =   "粘贴整个事件"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_kdef_txt2grp 
         Caption         =   "导入剧本"
      End
      Begin VB.Menu mnu_Kdef_ReplaceAllTalk 
         Caption         =   "修改全部说话人头像编号"
      End
   End
   Begin VB.Menu mnu_z 
      Caption         =   "z.dat编辑"
      Begin VB.Menu mnu_z_Edit 
         Caption         =   "z.dat编辑"
      End
      Begin VB.Menu mnu_z_HighEdit 
         Caption         =   "z.dat高级编辑"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_z_New 
         Caption         =   "新游戏属性编辑"
      End
      Begin VB.Menu mnu_z_Big 
         Caption         =   "Title.big修改"
      End
      Begin VB.Menu mnu_z_Crypt 
         Caption         =   "游戏加密"
      End
      Begin VB.Menu mnu_new_z 
         Caption         =   "生成新的z.dat"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_Map 
      Caption         =   "地图贴图编辑"
      Begin VB.Menu mnu_ShowPic 
         Caption         =   "贴图编辑"
      End
      Begin VB.Menu mnu_MMapEdit 
         Caption         =   "主地图编辑"
      End
      Begin VB.Menu mnu_WarMapEdit 
         Caption         =   "战斗地图编辑"
      End
   End
   Begin VB.Menu mnu_window 
      Caption         =   "窗口"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnu_help 
      Caption         =   "help"
      Visible         =   0   'False
   End
   Begin VB.Menu mnu_MMAPMenu 
      Caption         =   "MMAPMenu"
      Visible         =   0   'False
      Begin VB.Menu mnu_MMAPMenu_Grid 
         Caption         =   "显示网格"
      End
      Begin VB.Menu mnu_MMAPMenu_ShowLevel 
         Caption         =   "只显示操作层"
      End
      Begin VB.Menu mnu_MMAPMenu_Mode 
         Caption         =   "操作模式"
         Begin VB.Menu mnu_MMAPMenu_Normal 
            Caption         =   "正常模式"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_MMAPMenu_Block 
            Caption         =   "块操作模式"
         End
         Begin VB.Menu mnu_MMAPMenu_Delete 
            Caption         =   "删除模式"
         End
      End
      Begin VB.Menu mnu_MMAPMenu_ShowScene 
         Caption         =   "显示场景入口"
      End
      Begin VB.Menu mnu_MMAPMenu_LoadMap 
         Caption         =   "读取地图"
      End
      Begin VB.Menu mnu_MMAPMenu_SaveMap 
         Caption         =   "保存地图"
      End
   End
   Begin VB.Menu mnu_SMAPMenu 
      Caption         =   "SMAPMenu"
      Visible         =   0   'False
      Begin VB.Menu mnu_SMAPMenu_Grid 
         Caption         =   "显示网格"
      End
      Begin VB.Menu mnu_SMAPMenu_ShowLevel 
         Caption         =   "只显示操作层"
      End
      Begin VB.Menu mnu_SMAPMenu_Mode 
         Caption         =   "操作模式"
         Begin VB.Menu mnu_SMAPMenu_Normal 
            Caption         =   "正常模式"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_SMAPMenu_BLock 
            Caption         =   "块操作模式"
         End
         Begin VB.Menu mnu_SMAPMenu_Delete 
            Caption         =   "删除模式"
         End
      End
      Begin VB.Menu mnu_SMAPMenu_SLKGSmap 
         Caption         =   "导入/导出场景"
         Begin VB.Menu mnu_SMAPMenu_LoadKGsmap 
            Caption         =   "导入场景"
         End
         Begin VB.Menu mnu_SMAPMenu_SaveKGsmap 
            Caption         =   "导出场景"
         End
      End
      Begin VB.Menu mnu_SMAPMenu_AddMap 
         Caption         =   "增加场景地图"
      End
      Begin VB.Menu mnu_SMAPMenu_DeleteMap 
         Caption         =   "删除场景地图"
      End
      Begin VB.Menu mnu_SMAPMenu_LoadMap 
         Caption         =   "读取进度"
         Begin VB.Menu mnu_SMAPMenu_LoadMap0 
            Caption         =   "新游戏进度"
         End
         Begin VB.Menu mnu_SMAPMenu_LoadMap1 
            Caption         =   "进度一"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_SMAPMenu_LoadMap2 
            Caption         =   "进度二"
         End
         Begin VB.Menu mnu_SMAPMenu_LoadMap3 
            Caption         =   "进度三"
         End
         Begin VB.Menu mnu_SMAPMenu_LoadMap4 
            Caption         =   "进度四"
         End
         Begin VB.Menu mnu_SMAPMenu_LoadMap5 
            Caption         =   "进度五"
         End
         Begin VB.Menu mnu_SMAPMenu_LoadMap6 
            Caption         =   "自动挡"
         End
      End
      Begin VB.Menu mnu_SMAPMenu_Save 
         Caption         =   "保存当前场景进度"
      End
   End
   Begin VB.Menu mnu_Selectmap 
      Caption         =   "Selectmap"
      Visible         =   0   'False
      Begin VB.Menu mnu_Selectmap_Edit 
         Caption         =   "编辑贴图"
      End
      Begin VB.Menu mnu_Selectmap_switch 
         Caption         =   "颜色转换"
      End
      Begin VB.Menu mnu_Selectmap_copy 
         Caption         =   "复制贴图"
      End
      Begin VB.Menu mnu_Selectmap_Paste 
         Caption         =   "粘贴贴图"
      End
      Begin VB.Menu mnu_Selectmap_Insert 
         Caption         =   "插入贴图"
      End
      Begin VB.Menu mnu_Selectmap_Size 
         Caption         =   "贴图放大"
      End
      Begin VB.Menu mnu_Selectmap_Add 
         Caption         =   "添加贴图到最后"
      End
      Begin VB.Menu mnu_Selectmap_Delete 
         Caption         =   "删除最后一个贴图"
      End
      Begin VB.Menu mnu_Selectmap_Save 
         Caption         =   "保存贴图"
      End
   End
   Begin VB.Menu mnu_WarMAPMenu 
      Caption         =   "WarMAPMenu"
      Visible         =   0   'False
      Begin VB.Menu mnu_WarMAPMenu_Grid 
         Caption         =   "显示网格"
      End
      Begin VB.Menu mnu_WarMAPMenu_ShowLevel 
         Caption         =   "只显示操作层"
      End
      Begin VB.Menu mnu_WarMAPMenu_Mode 
         Caption         =   "操作模式"
         Begin VB.Menu mnu_WarMAPMenu_Normal 
            Caption         =   "正常模式"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_WarMAPMenu_BLock 
            Caption         =   "块操作模式"
         End
         Begin VB.Menu mnu_WarMAPMenu_Delete 
            Caption         =   "删除模式"
         End
      End
      Begin VB.Menu mnu_WarMAPMenu_SLKGSmap 
         Caption         =   "导入/导出战斗场景"
         Begin VB.Menu mnu_WarMAPMenu_LoadWmap 
            Caption         =   "导入场景"
         End
         Begin VB.Menu mnu_WarMAPMenu_LoadSmap 
            Caption         =   "导出场景"
         End
      End
      Begin VB.Menu mnu_WarMAPMenu_AddMap 
         Caption         =   "增加战斗地图"
      End
      Begin VB.Menu mnu_WarMAPMenu_DeleteMap 
         Caption         =   "删除战斗地图"
      End
      Begin VB.Menu mnu_WarMAPMenu_Loadmap 
         Caption         =   "读取地图"
      End
      Begin VB.Menu mnu_WarMAPMenu_SaveMap 
         Caption         =   "保存地图"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Type BrowseInfo
   hwndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias _
"lstrcatA" (ByVal lpString1 As String, ByVal lpString2 _
As String) As Long
Const BIF_RETURNONLYFSDIRS = 100
Const BIF_DONTGOBELOWDOMAIN = 0
Const MAX_PATH = 260
Public hWndBar As Long      '状态栏句柄

Private Sub big52_Click()

End Sub

Private Sub MDIForm_Load()
Dim i As Long
Dim tmplong(2) As Long
Dim tmpstr As String
    First = False
    SkinH_SetAero 1
 '   mnu_KdefEdit.Visible = False
    If GetINILong("run", "debug") <> 99 Then
        Me.Caption = "sFishedit0.72.20(KYSsp1)"
    Else
        Me.Caption = "sFishedit0.72.20(KYSsp1)---debug"
    End If
    G_Var.iniFileName = App.Path & "\fishedit.ini"
    For i = 0 To Me.Controls.Count - 1
        Call SetCaption(Me.Controls(i))
    Next i
    
    Set_menu
    
    If GetINIStr("run", "debug") = "BigGodIWantToEnlarge" Then
        MDIMain.mnu_Selectmap_Size.Enabled = True
    Else
        MDIMain.mnu_Selectmap_Size.Enabled = False
    End If
     
    
    tmpstr = GetINIStr("run", "gamepath")
    If tmpstr <> "" Then
        G_Var.JYPath = tmpstr
        Setpath
    Else
            'Getgamepath.Show vbModal
        Dim lpIDList As Long 'Declare Varibles
        Dim sBuffer As String
        Dim szTitle As String
        Dim tBrowseInfo As BrowseInfo
  
        szTitle = "Select JYpath"
        With tBrowseInfo
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
        End With
  
        lpIDList = SHBrowseForFolder(tBrowseInfo)
  
        If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            G_Var.JYPath = sBuffer
        End If
        G_Var.JYPath = IIf(right(G_Var.JYPath, 1) = "\", G_Var.JYPath, G_Var.JYPath & "\")
        If G_Var.JYPath = "\" Then G_Var.JYPath = ""
        'Call PutINIStr("run", "gamepath", G_Var.JYPath)
        'Setpath
        Call PutINIStr("run", "gamepath", G_Var.JYPath)
        Setpath
        'MsgBox StrUnicode2("良知帝万岁！"), vbInformation, "Long Life Chairman Liang!"
    End If
    
End Sub



Private Sub mnu_DS_Edit_Click()
    frmSMapEdit.Show
End Sub

Private Sub MDIForm_Unload(cancel As Integer)
Set c_Skinner = Nothing
'SkinH_Detach
Unload Me
End
End Sub

'Private Sub mnu_Add_Click()
'    Form1.Show vbNormal
'End Sub

Private Sub mnu_Exit_Click()
    Unload Me
End Sub

Private Sub mnu_kdef_copy_Click()
    frmmain.mnu_copy_Click
End Sub

Private Sub mnu_kdef_copyAll_Click()
    frmmain.mnu_copyall_Click
End Sub

Private Sub mnu_kdef_Paste_Click()
    frmmain.mnu_paste_Click
End Sub

Private Sub mnu_kdef_PasteALL_Click()
    frmmain.mnu_pasteall_Click
End Sub
Private Sub mnu_kdef_txt2grp_Click()
    frmmain.mnu_kdef_txt2grp_Click
End Sub
Private Sub mnu_Kdef_ReplaceAllTalk_Click()
    frmmain.mnu_replacetalk_Click
End Sub

Private Sub mnu_KdefEdit_Click()
    frmmain.Show
End Sub


Private Sub mnu_MMapEdit_Click()
    frmMMapEdit.Show
End Sub

Private Sub mnu_MMAPMenu_Block_Click()
    frmMMapEdit.ClickMenu "Block"
End Sub

Private Sub mnu_MMAPMenu_Delete_Click()
    frmMMapEdit.ClickMenu "Delete"
End Sub

Private Sub mnu_MMAPMenu_Grid_Click()
    frmMMapEdit.ClickMenu "Grid"
End Sub

Private Sub mnu_MMAPMenu_LoadMap_Click()
    frmMMapEdit.ClickMenu "Loadmap"
End Sub

Private Sub mnu_MMAPMenu_Normal_Click()
    frmMMapEdit.ClickMenu "Normal"
End Sub

Private Sub mnu_MMAPMenu_SaveMap_Click()
    frmMMapEdit.ClickMenu "Savemap"
End Sub

Private Sub mnu_MMAPMenu_ShowLevel_Click()
    frmMMapEdit.ClickMenu "ShowLevel"
End Sub

Private Sub mnu_MMAPMenu_ShowScene_Click()
    frmMMapEdit.ClickMenu "ShowScene"
End Sub



Private Sub mnu_R_Edit_Click()
    frmREdit.Show
End Sub

Private Sub mnu_ReadR1_Click()
    Call ReadRR(0)
End Sub

Private Sub mnu_SEdit_Click()
    frmSMapEdit.Show
End Sub

Private Sub mnu_Selectmap_Add_Click()
    frmSelectMap.ClickMenu "Add"
End Sub
Private Sub mnu_Selectmap_switch_Click()
    frmSelectMap.ClickMenu "switch"
End Sub

Private Sub mnu_Selectmap_copy_Click()
    frmSelectMap.ClickMenu "Copy"
End Sub

Private Sub mnu_Selectmap_Insert_click()
    frmSelectMap.ClickMenu "insert"
End Sub
Private Sub mnu_Selectmap_Delete_Click()
    frmSelectMap.ClickMenu "Delete"
End Sub

Private Sub mnu_Selectmap_Edit_Click()
    frmSelectMap.ClickMenu "Edit"
End Sub

Private Sub mnu_Selectmap_Paste_Click()
    frmSelectMap.ClickMenu "Paste"
End Sub
Private Sub mnu_Selectmap_Size_Click()
    frmSelectMap.ClickMenu "x2"
End Sub
Private Sub mnu_Selectmap_Save_Click()
    frmSelectMap.ClickMenu "Save"
End Sub

Private Sub mnu_setconfig_Click()
    frmSelectCharset.Show vbNormal
End Sub

Private Sub mnu_setpath_Click()
    'Getgamepath.Show vbModal
  Dim lpIDList As Long 'Declare Varibles
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo
  
  szTitle = "Select JYpath"
  With tBrowseInfo
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    G_Var.JYPath = sBuffer
  End If
    
    G_Var.JYPath = IIf(right(G_Var.JYPath, 1) = "\", G_Var.JYPath, G_Var.JYPath & "\")
    If G_Var.JYPath = "\" Then G_Var.JYPath = ""
    Call PutINIStr("run", "gamepath", G_Var.JYPath)
    Setpath
End Sub
Private Sub Setpath()
    If GetINILong("run", "debug") <> 99 Then
        Me.Caption = "sFishedit0.72.20" & "-" & G_Var.JYPath
    Else
        Me.Caption = "sFishedit0.72.20" & "-" & G_Var.JYPath & "-debug"
    End If
    'On Error GoTo Label1
    If G_Var.JYPath <> "" Then
'On Error GoTo Label2:
        SetColor       ' 读取颜色表
'On Error GoTo Label3:
        Call ReadRR(0)         ' 读取进度1
       ' Call ReadWar
        Set_menu
    End If
'    mnu_readfile.Enabled = True
    Exit Sub
'Label1:
    G_Var.JYPath = ""
    Exit Sub
'Label2:
    MsgBox StrUnicode2("错误代号：01") & Chr(13) & StrUnicode2("错误内容：游戏目录设置错误，请重新检查你的设置"), vbCritical, StrUnicode2("错误代号：01")
    Call PutINIStr("run", "gamepath", "")
    Call PutINIStr("run", "Mode", "")
    Call PutINIStr("run", "Charset", "")
    Call PutINIStr("run", "style", "")
    End
    Exit Sub
'Label3:
    MsgBox StrUnicode2("错误代号：02") & Chr(13) & StrUnicode2("错误内容：存档文件读取出错，可能与ini文件不匹配"), vbCritical, StrUnicode2("错误代号：02")
    Exit Sub
End Sub


Private Sub Set_menu()
Dim b As Boolean
    If G_Var.JYPath = "" Then
        b = False
    Else
        b = True
    End If
    
    mnu_z.Enabled = b
    mnu_Map.Enabled = b
    mnu_GameEdit.Enabled = b
'    mnu_ReadR1.Enabled = b

End Sub



Private Sub mnu_ShowPic_Click()
    frmSelectMap.Show
End Sub

Private Sub mnu_SMAPMenu_AddMap_Click()
    frmSMapEdit.ClickMenu "Addmap"
End Sub

Private Sub mnu_SMAPMenu_BLock_Click()
    frmSMapEdit.ClickMenu "Block"
End Sub

Private Sub mnu_SMAPMenu_Delete_Click()
    frmSMapEdit.ClickMenu "Delete"
End Sub

Private Sub mnu_SMAPMenu_DeleteMap_Click()
    frmSMapEdit.ClickMenu "Deletemap"
End Sub

Private Sub mnu_SMAPMenu_Grid_Click()
    frmSMapEdit.ClickMenu "Grid"
End Sub

Private Sub mnu_SMAPMenu_LoadKGsmap_Click()
    frmSMapEdit.ClickMenu "loadmap"
End Sub

Private Sub mnu_SMAPMenu_LoadMap0_Click()
    frmSMapEdit.ClickMenu "Loadmap0"
End Sub

Private Sub mnu_SMAPMenu_LoadMap1_Click()
    frmSMapEdit.ClickMenu "Loadmap1"
End Sub

Private Sub mnu_SMAPMenu_LoadMap2_Click()
    frmSMapEdit.ClickMenu "Loadmap2"
End Sub

Private Sub mnu_SMAPMenu_LoadMap3_Click()
    frmSMapEdit.ClickMenu "Loadmap3"
End Sub
Private Sub mnu_SMAPMenu_LoadMap4_Click()
    frmSMapEdit.ClickMenu "Loadmap4"
End Sub
Private Sub mnu_SMAPMenu_LoadMap5_Click()
    frmSMapEdit.ClickMenu "Loadmap5"
End Sub
Private Sub mnu_SMAPMenu_LoadMap6_Click()
    frmSMapEdit.ClickMenu "Loadmap6"
End Sub

Private Sub mnu_SMAPMenu_Normal_Click()
    frmSMapEdit.ClickMenu "Normal"
End Sub

Private Sub mnu_SMAPMenu_Save_Click()
    frmSMapEdit.ClickMenu "Save"
End Sub

Private Sub mnu_SMAPMenu_SaveKGsmap_Click()
    frmSMapEdit.ClickMenu "savemap"
End Sub

Private Sub mnu_SMAPMenu_ShowLevel_Click()
    frmSMapEdit.ClickMenu "Showlevel"
End Sub










Private Sub mnu_Team_Click()
    frmTeam.Show
End Sub

Private Sub mnu_WarMapEdit_Click()
    frmWMapEdit.Show
End Sub

Private Sub mnu_WarMAPMenu_AddMap_Click()
    frmWMapEdit.ClickMenu "Addmap"
End Sub

Private Sub mnu_warMAPMenu_BLock_Click()
    frmWMapEdit.ClickMenu "Block"
End Sub

Private Sub mnu_warMAPMenu_Delete_Click()
    frmWMapEdit.ClickMenu "Delete"
End Sub

Private Sub mnu_warMAPMenu_DeleteMap_Click()
    frmWMapEdit.ClickMenu "Deletemap"
End Sub

Private Sub mnu_warMAPMenu_Grid_Click()
    frmWMapEdit.ClickMenu "Grid"
End Sub

Private Sub mnu_WarMAPMenu_LoadMap_Click()
    frmWMapEdit.ClickMenu "Loadmap"
End Sub

Private Sub mnu_WarMAPMenu_LoadWmap_Click()
    frmWMapEdit.ClickMenu "loadwmap"
    'MsgBox 1
End Sub

Private Sub mnu_WarMAPMenu_SaveWmap_Click()
    frmWMapEdit.ClickMenu "savewmap"
End Sub

Private Sub mnu_WarMAPMenu_Normal_Click()
    frmWMapEdit.ClickMenu "Normal"
End Sub

Private Sub mnu_WarMAPMenu_SaveMAP_Click()
    frmWMapEdit.ClickMenu "Save"
End Sub

Private Sub mnu_WarMAPMenu_ShowLevel_Click()
    frmWMapEdit.ClickMenu "Showlevel"
End Sub


Private Sub mnu_z_Big_Click()
    frmBIGreader.Show
End Sub

Private Sub mnu_z_Crypt_Click()
    frmZCrypt.Show vbModal
End Sub

Private Sub mnu_z_Edit_Click()
'    frmzModify.Show
    frmznewedit.Show
End Sub

Private Sub mnu_z_HighEdit_Click()
   ' frmZHighModify.Show
End Sub

Private Sub mnu_z_New_Click()
    frmInitEdit.Show
End Sub





Private Sub mnu_War_Click()
    frmWarEditNew.Show
    'FrmWarEdit.Show
End Sub




