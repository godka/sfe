VERSION 5.00
Begin VB.Form frmZCrypt 
   Caption         =   "��Ϸ����"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
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
   ScaleHeight     =   2880
   ScaleWidth      =   8205
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chkCryptKdef 
      Caption         =   "Kdef����"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
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
      Left            =   6600
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
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
      Left            =   6600
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "ע��: ʹ��ʱ�Ѽ��ܺ��.new�ļ����ƻ�ȥ���ɡ�ע�Ᵽ��ԭ���ļ������ܺ��޷���ԭ��"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Label Label1 
      Caption         =   "���ɼ��ܺ��kdef.idx.new, kdef.grp.new��z.dat.new"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmZCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Const Varkdef = &H5B7BE
'Const StartKey = &H5B7C3
'Const StrKdefOffset = &H5B7DB
'Const KeyIDX = &H5B7EB
'Const keyGRP = &H5B82B
Const Varkdef = &H5B548
Const StartKey = &H5B54D
Const StrKdefOffset = &H5B565
Const KeyIDX = &H5B575
Const keyGRP = &H5B5B5

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
Dim I As Long
Dim zfilenum As Long
Dim tmpbyte As Byte
    
    

    ' д����
    If chkCryptKdef.Value = 1 Then
        FileCopy G_Var.JYPath & "z.dat", G_Var.JYPath & "z.dat.new"
        zfilenum = OpenBin(G_Var.JYPath & "z.dat.new", "W")
        Put #zfilenum, Varkdef + 1, CLng(1)
        Randomize
        For I = 0 To 160
            tmpbyte = CByte(RndLong(256))
            Put #zfilenum, StartKey + 1 + I, tmpbyte
        Next I
        Close zfilenum
        Call CryptKdef
    End If

End Sub

Private Sub Form_Load()
Dim I As Long
    Me.Caption = StrUnicode(Me.Caption)
 
    For I = 0 To Me.Controls.Count - 1
         Call SetCaption(Me.Controls(I))
    Next I
    c_Skinner.AttachSkin Me.hWnd

End Sub


Private Sub CryptKdef()
Dim id As Long
Dim idxByte() As Byte
Dim tmpidx() As Long
Dim grpByte() As Byte
Dim grplong As Long
Dim idxlong As Long
Dim I As Long, j As Long, k As Long
Dim offset As Byte
Dim cryptStr(20) As Byte
Dim cryptStrInt(20) As Integer
Dim flag As Long
Dim Length As Long
    
    idxlong = FileLen(G_Var.JYPath & "kdef.idx")
    
    ReDim idxByte(idxlong - 1)
    id = OpenBin(G_Var.JYPath & "kdef.idx", "R")
        Get #id, , idxByte
    Close id
    
    id = OpenBin(G_Var.JYPath & "z.dat.new", "R")
        Get #id, StrKdefOffset + 1, offset
        offset = offset And 15
        For I = 0 To 20
            Get #id, KeyIDX + 1 + I + offset, cryptStr(I)
        Next I
    Close id
    
    flag = 0
    I = 0
 
    Do While flag = 0
  
        For j = 0 To 14
            idxByte(I) = idxByte(I) Xor cryptStr(j)
            I = I + 1
            If I >= idxlong Then
                flag = 1
                Exit For
            End If
        Next j
    Loop
    
    id = OpenBin(G_Var.JYPath & "kdef.idx.new", "WN")
    Put #id, , idxByte
    Close id
    
    ' ����kdefgrp
    
    id = OpenBin(G_Var.JYPath & "z.dat.new", "R")
    Get #id, StrKdefOffset + 1, offset
    offset = offset And 15
    For I = 0 To 20
        Get #id, keyGRP + 1 + I + offset, cryptStr(I)
    Next I
    Close id
    
    
    ReDim tmpidx(idxlong / 4)
    id = OpenBin(G_Var.JYPath & "kdef.idx", "R")
    For I = 1 To idxlong / 4
        Get #id, , tmpidx(I)
    Next I
    tmpidx(0) = 0
    Close id
    
    grplong = FileLen(G_Var.JYPath & "kdef.grp")
    ReDim grpByte(grplong - 1)
    id = OpenBin(G_Var.JYPath & "kdef.grp", "R")
    Get #id, , grpByte
    Close id
    
    
    
    
    For k = 0 To idxlong / 4 - 1
        Length = tmpidx(k + 1) - tmpidx(k)
       flag = 0
       I = 0
    
       Do While flag = 0
     
           For j = 0 To 13
               grpByte(I + tmpidx(k)) = grpByte(I + tmpidx(k)) Xor cryptStr(j)
               I = I + 1
               If I >= Length Then
                   flag = 1
                   Exit For
               End If
           Next j
       Loop
    
    Next k
    
    id = OpenBin(G_Var.JYPath & "kdef.grp.new", "WN")
    Put #id, , grpByte
    Close id
    
    
End Sub



