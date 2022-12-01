Attribute VB_Name = "SkinH"
Public Declare Function SkinH_Attach Lib "SkinH_VB6.dll" () As Long
Public Declare Function SkinH_AttachEx Lib "SkinH_VB6.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long
Public Declare Function SkinH_AttachExt Lib "SkinH_VB6.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String, ByVal nHue As Integer, ByVal nSat As Integer, ByVal nBri As Integer) As Long
Public Declare Function SkinH_AttachRes Lib "SkinH_VB6.dll" (lpRes As Any, ByVal nSize As Long, ByVal lpPasswd As String, ByVal nHue As Integer, ByVal nSat As Integer, ByVal nBri As Integer) As Long
Public Declare Function SkinH_AdjustHSV Lib "SkinH_VB6.dll" (ByVal nHue As Integer, ByVal nSat As Integer, ByVal nBri As Integer) As Long
Public Declare Function SkinH_Detach Lib "SkinH_VB6.dll" () As Long
Public Declare Function SkinH_DetachEx Lib "SkinH_VB6.dll" (ByVal hWnd As Long) As Long
Public Declare Function SkinH_SetAero Lib "SkinH_VB6.dll" (ByVal hWnd As Long) As Long
Public Declare Function SkinH_SetWindowAlpha Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nAlpha As Integer) As Long
Public Declare Function SkinH_SetMenuAlpha Lib "SkinH_VB6.dll" (ByVal nAlpha As Integer) As Long
Public Declare Function SkinH_GetColor Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nPosX As Integer, ByVal nPosY As Integer) As Long
Public Declare Function SkinH_Map Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nType As Integer) As Long
Public Declare Function SkinH_LockUpdate Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nLocked As Integer) As Long
Public Declare Function SkinH_SetBackColor Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long
Public Declare Function SkinH_SetForeColor Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long
Public Declare Function SkinH_SetWindowMovable Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal bMove As Integer) As Long
Public Declare Function SkinH_AdjustAero Lib "SkinH_VB6.dll" (ByVal nAlpha As Integer, ByVal nShwDark As Integer, ByVal nShwSharp As Integer, ByVal nShwSize As Integer, ByVal nX As Integer, ByVal nY As Integer, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long
Public Declare Function SkinH_NineBlt Lib "SkinH_VB6.dll" (ByVal hDtDC As Long, ByVal left As Integer, ByVal top As Integer, ByVal right As Integer, ByVal bottom As Integer, ByVal nMRect As Integer) As Long
Public Declare Function SkinH_SetTitleMenuBar Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal bEnable As Integer, ByVal nMenuY As Integer, ByVal nTopOffs As Integer, ByVal nRightOffs As Integer) As Long
Public Declare Function SkinH_SetFont Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal hFont As Long) As Long
Public Declare Function SkinH_SetFontEx Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal szFace As String, ByVal nHeight As Integer, ByVal nWidth As Integer, ByVal nWeight As Integer, ByVal nItalic As Integer, ByVal nUnderline As Integer, ByVal nStrikeOut As Integer) As Long
Public Declare Function SkinH_VerifySign Lib "SkinH_VB6.dll" () As Long

