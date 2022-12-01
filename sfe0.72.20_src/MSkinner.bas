Attribute VB_Name = "MSkinner"
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'■                                                              ■
'■      模块创建者: gvu                                         ■
'■      网      址: http://hi.baidu.com/googlevipuser           ■
'■      E-Mail地址: owhao@126.com                               ■
'■                                                              ■
'■      你可以用于任何用途，转载请保留作者信息。                ■
'■                                                              ■
'■      可能不稳定，请慎用，谢谢。                              ■
'■                                                              ■
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type RECT
        Left        As Long
        Top         As Long
        Right       As Long
        Bottom      As Long
End Type

Private Type RECTW
        Left                As Long
        Top                 As Long
        Right               As Long
        Bottom              As Long
        Width               As Long
        Height              As Long
End Type

Private Type PAINTSTRUCT
        hDC                 As Long
        fErase              As Long
        rcPaint             As RECT
        fRestore            As Long
        fIncUpdate          As Long
        rgbReserved(32)     As Byte
End Type

Private Type HDITEM
        mask                As Long
        cxy                 As Long
        pszText             As String
        hbm                 As Long
        cchTextMax          As Long
        fmt                 As Long
        IntPtr              As Long
End Type

Private Type TRACKMOUSEEVENTTYPE
    cbSize      As Long
    dwFlags     As Long
    hwndTrack   As Long
    dwHoverTime As Long
End Type

Private Type WINDOWPOS
   hWnd                     As Long
   hWndInsertAfter          As Long
   X                        As Long
   Y                        As Long
   cX                       As Long
   cY                       As Long
   flags                    As Long
End Type

Private Type NCCALCSIZE_PARAMS
   rgrc(0 To 2)             As RECT
   lppos                    As Long
End Type

Private Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hDC As Long
        rcItem As RECT
        itemData As Long
End Type

Private Enum DTSTYLE
    DT_LEFT = &H0
    DT_TOP = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_FORE_ELLIPSIS = &H8000
    DT_END_ELLIPSIS = &H8000&
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
End Enum

Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

Private Const WM_ACTIVATE                   As Long = &H6
Private Const WM_ACTIVATEAPP                As Long = &H1C
Private Const WM_ASKCBFORMATNAME            As Long = &H30C
Private Const WM_CANCELJOURNAL              As Long = &H4B
Private Const WM_CANCELMODE                 As Long = &H1F
Private Const WM_CHANGECBCHAIN              As Long = &H30D
Private Const WM_CHAR                       As Long = &H102
Private Const WM_CHARTOITEM                 As Long = &H2F
Private Const WM_CHILDACTIVATE              As Long = &H22
Private Const WM_CLEAR                      As Long = &H303
Private Const WM_CLOSE                      As Long = &H10
Private Const WM_COMMAND                    As Long = &H111
Private Const WM_COMMNOTIFY                 As Long = &H44
Private Const WM_COMPACTING                 As Long = &H41
Private Const WM_COMPAREITEM                As Long = &H39
Private Const WM_CONVERTREQUESTEX           As Long = &H108
Private Const WM_COPY                       As Long = &H301
Private Const WM_COPYDATA                   As Long = &H4A
Private Const WM_CREATE                     As Long = &H1
Private Const WM_CUT                        As Long = &H300
Private Const WM_DEADCHAR                   As Long = &H103
Private Const WM_DELETEITEM                 As Long = &H2D
Private Const WM_DESTROY                    As Long = &H2
Private Const WM_DESTROYCLIPBOARD           As Long = &H307
Private Const WM_DEVMODECHANGE              As Long = &H1B
Private Const WM_DRAWCLIPBOARD              As Long = &H308
Private Const WM_DRAWITEM                   As Long = &H2B
Private Const WM_DROPFILES                  As Long = &H233
Private Const WM_ENABLE                     As Long = &HA
Private Const WM_ENDSESSION                 As Long = &H16
Private Const WM_ENTERIDLE                  As Long = &H121
Private Const WM_ENTERMENULOOP              As Long = &H211
Private Const WM_ERASEBKGND                 As Long = &H14
Private Const WM_EXITMENULOOP               As Long = &H212
Private Const WM_FONTCHANGE                 As Long = &H1D
Private Const WM_GETDLGCODE                 As Long = &H87
Private Const WM_GETFONT                    As Long = &H31
Private Const WM_GETHOTKEY                  As Long = &H33
Private Const WM_GETMINMAXINFO              As Long = &H24
Private Const WM_GETTEXT                    As Long = &HD
Private Const WM_GETTEXTLENGTH              As Long = &HE
Private Const WM_HOTKEY                     As Long = &H312
Private Const WM_HSCROLL                    As Long = &H114
Private Const WM_HSCROLLCLIPBOARD           As Long = &H30E
Private Const WM_ICONERASEBKGND             As Long = &H27
Private Const WM_IME_CHAR                   As Long = &H286
Private Const WM_IME_COMPOSITION            As Long = &H10F
Private Const WM_IME_COMPOSITIONFULL        As Long = &H284
Private Const WM_IME_CONTROL                As Long = &H283
Private Const WM_IME_ENDCOMPOSITION         As Long = &H10E
Private Const WM_IME_KEYDOWN                As Long = &H290
Private Const WM_IME_KEYLAST                As Long = &H10F
Private Const WM_IME_KEYUP                  As Long = &H291
Private Const WM_IME_NOTIFY                 As Long = &H282
Private Const WM_IME_SELECT                 As Long = &H285
Private Const WM_IME_SETCONTEXT             As Long = &H281
Private Const WM_IME_STARTCOMPOSITION       As Long = &H10D
Private Const WM_INITDIALOG                 As Long = &H110
Private Const WM_INITMENU                   As Long = &H116
Private Const WM_INITMENUPOPUP              As Long = &H117
Private Const WM_KEYDOWN                    As Long = &H100
Private Const WM_KEYFIRST                   As Long = &H100
Private Const WM_KEYLAST                    As Long = &H108
Private Const WM_KEYUP                      As Long = &H101
Private Const WM_KILLFOCUS                  As Long = &H8
Private Const WM_LBUTTONDBLCLK              As Long = &H203
Private Const WM_LBUTTONDOWN                As Long = &H201
Private Const WM_LBUTTONUP                  As Long = &H202
Private Const WM_MBUTTONDBLCLK              As Long = &H209
Private Const WM_MBUTTONDOWN                As Long = &H207
Private Const WM_MBUTTONUP                  As Long = &H208
Private Const WM_MDIACTIVATE                As Long = &H222
Private Const WM_MDICASCADE                 As Long = &H227
Private Const WM_MDICREATE                  As Long = &H220
Private Const WM_MDIDESTROY                 As Long = &H221
Private Const WM_MDIGETACTIVE               As Long = &H229
Private Const WM_MDIICONARRANGE             As Long = &H228
Private Const WM_MDIMAXIMIZE                As Long = &H225
Private Const WM_MDINEXT                    As Long = &H224
Private Const WM_MDIREFRESHMENU             As Long = &H234
Private Const WM_MDIRESTORE                 As Long = &H223
Private Const WM_MDISETMENU                 As Long = &H230
Private Const WM_MDITILE                    As Long = &H226
Private Const WM_MEASUREITEM                As Long = &H2C
Private Const WM_MENUCHAR                   As Long = &H120
Private Const WM_MENUSELECT                 As Long = &H11F
Private Const WM_MOUSEACTIVATE              As Long = &H21
Private Const WM_MOUSEFIRST                 As Long = &H200
Private Const WM_MOUSELAST                  As Long = &H209
Private Const WM_MOUSELEAVE                 As Long = &H2A3
Private Const WM_MOUSEWHEEL                 As Long = &H20A
Private Const WM_MOUSEMOVE                  As Long = &H200
Private Const WM_MOVE                       As Long = &H3
Private Const WM_NEXTDLGCTL                 As Long = &H28
Private Const WM_NULL                       As Long = &H0
Private Const WM_OTHERWINDOWCREATED         As Long = &H42               '  no longer suported
Private Const WM_OTHERWINDOWDESTROYED       As Long = &H43             '  no longer suported
Private Const WM_PAINT                      As Long = &HF
Private Const WM_PAINTCLIPBOARD             As Long = &H309
Private Const WM_PAINTICON                  As Long = &H26
Private Const WM_PALETTECHANGED             As Long = &H311
Private Const WM_PALETTEISCHANGING          As Long = &H310
Private Const WM_PARENTNOTIFY               As Long = &H210
Private Const WM_PASTE                      As Long = &H302
Private Const WM_PENWINFIRST                As Long = &H380
Private Const WM_PENWINLAST                 As Long = &H38F
Private Const WM_POWER                      As Long = &H48
Private Const WM_USER                       As Long = &H400
Private Const WM_PSD_ENVSTAMPRECT           As Long = (WM_USER + 5)
Private Const WM_PSD_FULLPAGERECT           As Long = (WM_USER + 1)
Private Const WM_PSD_GREEKTEXTRECT          As Long = (WM_USER + 4)
Private Const WM_PSD_MARGINRECT             As Long = (WM_USER + 3)
Private Const WM_PSD_MINMARGINRECT          As Long = (WM_USER + 2)
Private Const WM_PSD_PAGESETUPDLG           As Long = (WM_USER)
Private Const WM_PSD_YAFULLPAGERECT         As Long = (WM_USER + 6)
Private Const WM_QUERYDRAGICON              As Long = &H37
Private Const WM_QUERYENDSESSION            As Long = &H11
Private Const WM_QUERYNEWPALETTE            As Long = &H30F
Private Const WM_QUERYOPEN                  As Long = &H13
Private Const WM_QUEUESYNC                  As Long = &H23
Private Const WM_QUIT                       As Long = &H12
Private Const WM_RBUTTONDBLCLK              As Long = &H206
Private Const WM_RBUTTONDOWN                As Long = &H204
Private Const WM_RBUTTONUP                  As Long = &H205
Private Const WM_RENDERALLFORMATS           As Long = &H306
Private Const WM_RENDERFORMAT               As Long = &H305
Private Const WM_SETCURSOR                  As Long = &H20
Private Const WM_SETFOCUS                   As Long = &H7
Private Const WM_SETFONT                    As Long = &H30
Private Const WM_SETHOTKEY                  As Long = &H32
Private Const WM_SETREDRAW                  As Long = &HB
Private Const WM_SETTEXT                    As Long = &HC
Private Const WM_SHOWWINDOW                 As Long = &H18
Private Const WM_SIZE                       As Long = &H5
Private Const WM_SIZECLIPBOARD              As Long = &H30B
Private Const WM_SPOOLERSTATUS              As Long = &H2A
Private Const WM_SYSCHAR                    As Long = &H106
Private Const WM_SYSCOLORCHANGE             As Long = &H15
Private Const WM_SYSCOMMAND                 As Long = &H112
Private Const WM_SYSDEADCHAR                As Long = &H107
Private Const WM_SYSKEYDOWN                 As Long = &H104
Private Const WM_SYSKEYUP                   As Long = &H105
Private Const WM_STYLECHANGED               As Long = &H7D
Private Const WM_TIMECHANGE                 As Long = &H1E
Private Const WM_TIMER                      As Long = &H113
Private Const WM_UNDO                       As Long = &H304
Private Const WM_VKEYTOITEM                 As Long = &H2E
Private Const WM_VSCROLL                    As Long = &H115
Private Const WM_VSCROLLCLIPBOARD           As Long = &H30A
Private Const WM_WINDOWPOSCHANGED           As Long = &H47
Private Const WM_WINDOWPOSCHANGING          As Long = &H46
Private Const WM_WININICHANGE               As Long = &H1A
Private Const WM_CHOOSEFONT_GETLOGFONT      As Long = (WM_USER + 1)
Private Const WM_CHOOSEFONT_SETFLAGS        As Long = (WM_USER + 102)
Private Const WM_CHOOSEFONT_SETLOGFONT      As Long = (WM_USER + 101)
Private Const WM_DDE_FIRST                  As Long = &H3E0
Private Const WM_DDE_ACK                    As Long = (WM_DDE_FIRST + 4)
Private Const WM_DDE_ADVISE                 As Long = (WM_DDE_FIRST + 2)
Private Const WM_DDE_DATA                   As Long = (WM_DDE_FIRST + 5)
Private Const WM_DDE_EXECUTE                As Long = (WM_DDE_FIRST + 8)
Private Const WM_DDE_INITIATE               As Long = (WM_DDE_FIRST)
Private Const WM_DDE_LAST                   As Long = (WM_DDE_FIRST + 8)
Private Const WM_DDE_POKE                   As Long = (WM_DDE_FIRST + 7)
Private Const WM_DDE_REQUEST                As Long = (WM_DDE_FIRST + 6)
Private Const WM_DDE_TERMINATE              As Long = (WM_DDE_FIRST + 1)
Private Const WM_DDE_UNADVISE               As Long = (WM_DDE_FIRST + 3)

Private Const WM_NCACTIVATE                 As Long = &H86
Private Const WM_NCCALCSIZE                 As Long = &H83
Private Const WM_NCCREATE                   As Long = &H81
Private Const WM_NCDESTROY                  As Long = &H82
Private Const WM_NCHITTEST                  As Long = &H84
Private Const WM_NCLBUTTONDBLCLK            As Long = &HA3
Private Const WM_NCLBUTTONDOWN              As Long = &HA1
Private Const WM_NCLBUTTONUP                As Long = &HA2
Private Const WM_NCMBUTTONDBLCLK            As Long = &HA9
Private Const WM_NCMBUTTONDOWN              As Long = &HA7
Private Const WM_NCMBUTTONUP                As Long = &HA8
Private Const WM_NCMOUSEMOVE                As Long = &HA0
Private Const WM_NCPAINT                    As Long = &H85
Private Const WM_NCRBUTTONDBLCLK            As Long = &HA6
Private Const WM_NCRBUTTONDOWN              As Long = &HA4
Private Const WM_NCRBUTTONUP                As Long = &HA5
Private Const WM_NCPOPUPMENU                As Long = &HAE

Private Const WM_CTLCOLOR                   As Long = &H19
Private Const WM_CTLCOLORBTN                As Long = &H135
Private Const WM_CTLCOLORDLG                As Long = &H136
Private Const WM_CTLCOLOREDIT               As Long = &H133
Private Const WM_CTLCOLORLISTBOX            As Long = &H134
Private Const WM_CTLCOLORMSGBOX             As Long = &H132
Private Const WM_CTLCOLORSCROLLBAR          As Long = &H137
Private Const WM_CTLCOLORSTATIC             As Long = &H138

Private Const HDM_FIRST                     As Long = &H1200
Private Const HDM_CLEARFILTER               As Long = (HDM_FIRST + 24)
Private Const HDM_CREATEDRAGIMAGE           As Long = (HDM_FIRST + 16)
Private Const HDM_DELETEITEM                As Long = (HDM_FIRST + 2)
Private Const HDM_EDITFILTER                As Long = (HDM_FIRST + 23)
Private Const HDM_GETBITMAPMARGIN           As Long = (HDM_FIRST + 21)
Private Const HDM_GETIMAGELIST              As Long = (HDM_FIRST + 9)
Private Const HDM_GETITEMA                  As Long = (HDM_FIRST + 3)
Private Const HDM_GETITEMCOUNT              As Long = (HDM_FIRST + 0)
Private Const HDM_GETITEMRECT               As Long = (HDM_FIRST + 7)
Private Const HDM_GETITEMW                  As Long = (HDM_FIRST + 11)
Private Const HDM_GETORDERARRAY             As Long = (HDM_FIRST + 17)
Private Const HDM_HITTEST                   As Long = (HDM_FIRST + 6)
Private Const HDM_INSERTITEMA               As Long = (HDM_FIRST + 1)
Private Const HDM_INSERTITEMW               As Long = (HDM_FIRST + 10)
Private Const HDM_LAYOUT                    As Long = (HDM_FIRST + 5)
Private Const HDM_ORDERTOINDEX              As Long = (HDM_FIRST + 15)
Private Const HDM_SETBITMAPMARGIN           As Long = (HDM_FIRST + 20)
Private Const HDM_SETFILTERCHANGETIMEOUT    As Long = (HDM_FIRST + 22)
Private Const HDM_SETHOTDIVIDER             As Long = (HDM_FIRST + 19)
Private Const HDM_SETIMAGELIST              As Long = (HDM_FIRST + 8)
Private Const HDM_SETITEMA                  As Long = (HDM_FIRST + 4)
Private Const HDM_SETITEMW                  As Long = (HDM_FIRST + 12)
Private Const HDM_SETORDERARRAY             As Long = (HDM_FIRST + 18)
Private Const HDM_SETUNICODEFORMAT          As Long = &H2005
Private Const HDM_GETUNICODEFORMAT          As Long = &H2006

Private Const HDI_BITMAP                    As Long = &H10
Private Const HDI_DI_SETITEM                As Long = &H40
Private Const HDI_FILTER                    As Long = &H100
Private Const HDI_FORMAT                    As Long = &H4
Private Const HDI_WIDTH                     As Long = &H1
Private Const HDI_HEIGHT                    As Long = HDI_WIDTH
Private Const HDI_HIDDEN                    As Long = (&H1)
Private Const HDI_IMAGE                     As Long = &H20
Private Const HDI_LPARAM                    As Long = &H8
Private Const HDI_ORDER                     As Long = &H80
Private Const HDI_TEXT                      As Long = &H2

Private Const MK_LBUTTON                    As Long = &H1
Private Const MK_MBUTTON                    As Long = &H10
Private Const MK_RBUTTON                    As Long = &H2

Private Const WS_BORDER                     As Long = &H800000
Private Const WS_VSCROLL                    As Long = &H200000
Private Const WS_HSCROLL                    As Long = &H100000
Private Const WS_EX_CLIENTEDGE              As Long = &H200&

Private Const SM_CXVSCROLL                  As Long = &H2
Private Const SM_CYVSCROLL                  As Long = &H14
Private Const SM_CXHSCROLL                  As Long = &H15
Private Const SM_CYHSCROLL                  As Long = &H3
Private Const SM_CXDLGFRAME                 As Long = &H7
Private Const SM_CYDLGFRAME                 As Long = &H8
Private Const SM_CXCHECKBOX                 As Long = &H47
Private Const SM_CYCHECKBOX                 As Long = &H48

Private Const CB_SHOWDROPDOWN               As Long = &H14F
Private Const CB_GETDROPPEDSTATE            As Long = &H157

Private Const PBM_GETPOS                    As Long = (WM_USER + 8)
Private Const PBM_SETBARCOLOR               As Long = (WM_USER + 9)
Private Const PBS_SMOOTH                    As Long = &H1
Private Const PBS_VERTICAL                  As Long = &H4

Private Const BM_GETCHECK                   As Long = &HF0
Private Const BM_SETCHECK                   As Long = &HF1
Private Const BM_GETSTATE                   As Long = &HF2
Private Const BM_SETSTYLE                   As Long = &HF4

Private Const BS_NULL                       As Long = 1
Private Const BS_3STATE                     As Long = &H5&
Private Const BS_AUTO3STATE                 As Long = &H6&
Private Const BS_AUTOCHECKBOX               As Long = &H3&
Private Const BS_AUTORADIOBUTTON            As Long = &H9&
Private Const BS_CHECKBOX                   As Long = &H2&
Private Const BS_DEFPUSHBUTTON              As Long = &H1&
Private Const BS_DIBPATTERN                 As Long = 5
Private Const BS_DIBPATTERN8X8              As Long = 8
Private Const BS_DIBPATTERNPT               As Long = 6
Private Const BS_GROUPBOX                   As Long = &H7&
Private Const BS_HATCHED                    As Long = 2
Private Const BS_HOLLOW                     As Long = BS_NULL
Private Const BS_INDEXED                    As Long = 4
Private Const BS_LEFTTEXT                   As Long = &H20&
Private Const BS_OWNERDRAW                  As Long = &HB&
Private Const BS_PATTERN                    As Long = 3
Private Const BS_PATTERN8X8                 As Long = 7
Private Const BS_PUSHBUTTON                 As Long = &H0&
Private Const BS_RADIOBUTTON                As Long = &H4&
Private Const BS_SOLID                      As Long = 0
Private Const BS_USERBUTTON                 As Long = &H8&

Private Const LBS_OWNERDRAWVARIABLE = &H20&
Private Const CBS_OWNERDRAWVARIABLE = &H20&

Private Const ODT_BUTTON = 4
Private Const ODT_COMBOBOX = 3
Private Const ODT_HEADER = 100
Private Const ODT_LISTBOX = 2
Private Const ODT_LISTVIEW = 102
Private Const ODT_MENU = 1
Private Const ODT_STATIC = 5
Private Const ODT_TAB = 101

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long ' Win98 or later
Private Declare Function TrackMouseEvent2 Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long ' Win95 w/ IE 3.0
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private m_Init              As Boolean  '保存是否已经初始化
Private m_hBtnSrcDC         As Long
Private m_hCbbSrcDC         As Long
Private m_hCkbSrcDC         As Long
Private m_hOpbSrcDC         As Long
Private m_hHdbSrcDC         As Long
Private m_bTrackHandler32   As Boolean
Private m_SubclassCount     As Long     '保存子类化的个数，以便在销毁所有窗口和按钮之后可以释放资源

Public Function Attach(ByVal hWnd As Long) As Long
    If m_Init = False Then  '如果没有初始化，则初始化
        m_Init = True
        m_bTrackHandler32 = IsFunctionSupported("TrackMouseEvent", "User32")
        Call pInit
    End If
    Attach = pAttach(hWnd)
End Function

Public Function Detach(ByVal hWnd As Long) As Long
    Detach = pDetach(hWnd)
End Function

Private Function pAttach(ByVal hWnd As Long) As Long
If hWnd = 0 Then Exit Function
    If GetProp(hWnd, "PROCADDR") Then Exit Function
    Dim sClassName  As String
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "#32770", "thunderformdc", "thunderrt6formdc", "form"
            Call EnumChildWindows(hWnd, AddressOf pEnumChildProc, ByVal 0&)
            
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            Dim i           As Long
            Dim m_hDC       As Long
            Dim m_mDC(3)    As Long
            Dim m_BMP(3)    As Long
            Dim m_wRect     As RECTW
            Dim m_dwStyle   As Long
            m_hDC = GetWindowDC(hWnd)
            pGetWindowRectW hWnd, m_wRect
            For i = 0 To 3
                m_mDC(i) = CreateCompatibleDC(m_hDC)
                m_BMP(i) = CreateCompatibleBitmap(m_hDC, m_wRect.Width, m_wRect.Height)
                DeleteObject SelectObject(m_mDC(i), m_BMP(i))
                SetProp hWnd, "HDC" & CStr(i), m_mDC(i)
                SetProp hWnd, "BMP" & CStr(i), m_BMP(i)
            Next
            Call pDrawMemDC(hWnd)
            ReleaseDC hWnd, m_hDC
            m_dwStyle = GetWindowLong(hWnd, GWL_STYLE)
            If (m_dwStyle And BS_CHECKBOX) Or (m_dwStyle And BS_RADIOBUTTON) Then
            Else
                SendMessage hWnd, BM_SETSTYLE, BS_OWNERDRAW, ByVal True
            End If
            SetProp hWnd, "OLDSTYLE", m_dwStyle         '保存按钮旧的风格,以便再取消皮肤的时候恢复原来的风格
            SetProp hWnd, "MOUSEFLAG", 0
            SetProp hWnd, "TIMERID", 0
            SetProp hWnd, "OLDSTATE", IIf(IsWindowEnabled(hWnd), 0, 3)
            SetProp hWnd, "ALPHALEVEL", 0
            SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, m_wRect.Width + 1, m_wRect.Height + 1, 3, 3), True
            
        '=====================================================================================
        Case "thundercombobox", "thunderrt6combobox", "combo", "combobox", "thunderdrivelistbox", "thunderrt6drivelistbox", _
             "thundercheckbox", "thunderrt6checkbox", "thunderoptionbutton", "thunderrt6optionbutton"
            SetProp hWnd, "MOUSEFLAG", 0
            SetProp hWnd, "OLDSTATE", 0
        
        '=====================================================================================
        Case "progressbar20wndclass", "progressbarwndclass"
            'Call pGetWindowRectW(hWnd, m_wRect)
            'SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, m_wRect.Width + 1, m_wRect.Height + 1, 3, 3), False
        
        '=====================================================================================
        Case "msvb_lib_header", "sysheader32"
            SetProp hWnd, "MOUSEFLAG", 0
            SetProp hWnd, "HDINDEX", -1
            SetProp hWnd, "HMINDEX", -1
            
        '=====================================================================================
        Case Else
    
    End Select
    m_SubclassCount = m_SubclassCount + 1
    SetProp hWnd, "PROCADDR", SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
    SendMessage hWnd, WM_NCPAINT, 1&, 0&
    RedrawWindow hWnd, ByVal 0&, ByVal 0&, &H1 Or &H2
    pAttach = 1
End Function

Private Function pDetach(ByVal hWnd As Long) As Long
If hWnd = 0 Then Exit Function
    Dim OrigProc As Long
    OrigProc = GetProp(hWnd, "PROCADDR")
    If OrigProc = 0 Then Exit Function
    Dim sClassName  As String
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "#32770", "thunderformdc", "thunderrt6formdc", "form"
            Call EnumChildWindows(hWnd, AddressOf pDeEnumChildProc, ByVal 0&)
        
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            Dim m_mDC(3)    As Long
            Dim m_BMP(3)    As Long
            Dim i As Long
            For i = 0 To 3
                m_mDC(i) = GetProp(hWnd, "HDC" & CStr(i))
                m_BMP(i) = GetProp(hWnd, "BMP" & CStr(i))
                DeleteObject m_mDC(i)
                DeleteDC m_BMP(i)
                RemoveProp hWnd, "HDC" & CStr(i)
                RemoveProp hWnd, "BMP" & CStr(i)
            Next
            Call pKillTimer(hWnd)
            SetWindowLong hWnd, -16, GetProp(hWnd, "OLDSTYLE")
            RemoveProp hWnd, "OLDSTYLE"
            RemoveProp hWnd, "MOUSEFLAG"
            RemoveProp hWnd, "TIMERID"
            RemoveProp hWnd, "OLDSTATE"
            RemoveProp hWnd, "ALPHALEVEL"
            SetWindowRgn hWnd, 0&, ByVal True
        '=====================================================================================
        Case "thundercombobox", "thunderrt6combobox", "combo", "combobox", "thunderdrivelistbox", "thunderrt6drivelistbox", _
             "thundercheckbox", "thunderrt6checkbox", "thunderoptionbutton", "thunderrt6optionbutton"
            RemoveProp hWnd, "MOUSEFLAG"
            RemoveProp hWnd, "OLDSTATE"
        
        Case "msvb_lib_header", "sysheader32"
            RemoveProp hWnd, "MOUSEFLAG"
            RemoveProp hWnd, "HDINDEX"
            RemoveProp hWnd, "HMINDEX"
            
        '=====================================================================================
        Case "progressbar20wndclass", "progressbarwndclass"
            'SetWindowRgn hWnd, 0&, ByVal True
                    
        '=====================================================================================
        Case "datalistwndclass", "dblistwndclass"
                                
        '=====================================================================================
        Case Else
    
    End Select
    RemoveProp hWnd, "PROCADDR"
    Call SetWindowLong(hWnd, GWL_WNDPROC, OrigProc)
    SendMessage hWnd, WM_NCPAINT, 1&, 0&
    RedrawWindow hWnd, ByVal 0&, ByVal 0&, &H1 Or &H2
    m_SubclassCount = m_SubclassCount - 1
    If m_SubclassCount <= 0 Then
        m_SubclassCount = 0
        DeleteDC m_hBtnSrcDC
        DeleteDC m_hCbbSrcDC
        DeleteDC m_hCkbSrcDC
        DeleteDC m_hOpbSrcDC
        DeleteDC m_hHdbSrcDC
        m_Init = False
    End If
    pDetach = 1
End Function

Private Sub pDrawComboBox(ByVal hWnd As Long, ByVal hDC As Long, State As Long, Optional ByVal Redraw As Boolean = False)
    Dim mOldState As Long
    Dim bDrop     As Long
    bDrop = SendMessage(hWnd, CB_GETDROPPEDSTATE, 0&, 0&)
    mOldState = GetProp(hWnd, "OLDSTATE")
    If bDrop Then State = 2
    If mOldState = State And Redraw = False Then Exit Sub
    If Not GetWindowLong(hWnd, GWL_STYLE) And &H2 Then Exit Sub
    Call SetProp(hWnd, "OLDSTATE", State)
    Dim m_BtSize    As Long
    Dim m_hDC       As Long
    Dim TmpDC       As Long
    Dim TmpBMP      As Long
    Dim m_wRect     As RECTW
    Call pGetWindowRectW(hWnd, m_wRect)
    m_BtSize = GetSystemMetrics(SM_CXVSCROLL) + 1
    TmpDC = pCreateDC(m_BtSize, m_wRect.Height - 2)
    Select Case State
            Case 0
                Call pFillRectL(TmpDC, 0, 0, m_BtSize, m_wRect.Height - 2, &HFFFFFF)
                                    
            Case 1
                Call GridBlt(TmpDC, 0, 0, m_BtSize, m_wRect.Height - 2, m_hCbbSrcDC, 0, 0, 4, 18, 2, 1, 1, 1)
                
            Case 2
                Call GridBlt(TmpDC, 0, 0, m_BtSize, m_wRect.Height - 2, m_hCbbSrcDC, 4, 0, 4, 18, 2, 1, 1, 1)
                                    
    End Select
    If IsWindowEnabled(hWnd) Then
        Call TransBlt(TmpDC, m_BtSize - 7 - (m_BtSize - 7) / 2, (m_wRect.Height - 6) / 2, 7, 4, m_hCbbSrcDC, 8, 0)
    Else
        Call TransBlt(TmpDC, m_BtSize - 7 - (m_BtSize - 7) / 2, (m_wRect.Height - 6) / 2, 7, 4, m_hCbbSrcDC, 8, 4)
    End If
    If hDC = 0 Then
        m_hDC = GetWindowDC(hWnd)
    Else
        m_hDC = hDC
    End If
    BitBlt m_hDC, m_wRect.Width - m_BtSize - 1, 1, m_BtSize, m_wRect.Height - 2, TmpDC, 0, 0, vbSrcCopy
    DeleteDC TmpDC
    DeleteObject TmpBMP
    If hDC = 0 Then Call ReleaseDC(hWnd, m_hDC)
End Sub

Private Function pDrawButton(ByVal hWnd As Long, ByVal hDC As Long) As Long
    Dim m_Style As Long
    Dim m_State As Long
    Dim m_OldSt As Long
    Dim m_SrcDC As Long
    Dim m_DstDC As Long
    Dim m_Level As Long
    Dim m_wRect As RECTW
    If IsWindowEnabled(hWnd) = 0 Then Call SetProp(hWnd, "OLDSTATE", 3)
    m_Style = GetProp(hWnd, "OLDSTYLE")
    If (m_Style And BS_CHECKBOX) Or (m_Style And BS_RADIOBUTTON) Then Exit Function
    Call pGetWindowRectW(hWnd, m_wRect)
    m_OldSt = GetProp(hWnd, "OLDSTATE")
    m_Level = GetProp(hWnd, "ALPHALEVEL")
    m_SrcDC = GetProp(hWnd, "HDC" & CStr(m_OldSt))
    m_DstDC = IIf(hDC = 0, GetWindowDC(hWnd), hDC)
    AlphaBlend m_DstDC, 0, 0, m_wRect.Width, m_wRect.Height, m_SrcDC, 0, 0, m_wRect.Width, m_wRect.Height, m_Level * &H10000
    If hDC = 0 Then Call ReleaseDC(hWnd, m_DstDC)
End Function

Private Function pDrawCheckBox(ByVal hWnd As Long, ByVal State As Long, Optional ByVal Redraw As Boolean = False) As Long
    Dim mOldState As Long
    mOldState = GetProp(hWnd, "OLDSTATE")
    If mOldState = State And Redraw = False Then Exit Function
    Call SetProp(hWnd, "OLDSTATE", State)
    Dim m_hDC       As Long
    Dim TmpDC       As Long
    Dim m_wRect     As RECTW
    Dim m_cX        As Long
    Dim m_cY        As Long
    Dim mValue      As Long
    m_cX = GetSystemMetrics(SM_CXCHECKBOX)
    m_cY = GetSystemMetrics(SM_CYCHECKBOX)
    Call pGetWindowRectW(hWnd, m_wRect)
    mValue = SendMessage(hWnd, BM_GETCHECK, 0&, 0&)
    TmpDC = pCreateDC(m_cX, m_cY)
    m_hDC = GetWindowDC(hWnd)
    Call pFillRectL(TmpDC, 0, 0, m_cX, m_cY, &HFFFFFF)
    If IsWindowEnabled(hWnd) Then
        If State = 2 Then
            Call pFrameRect(TmpDC, 0, 0, m_cX, m_cY, &HC48639)
        Else
            Call pFrameRect(TmpDC, 0, 0, m_cX, m_cY, &HD5A554)
        End If
        If State = 1 Then Call StretchBlt(TmpDC, 1, 1, m_cX - 2, m_cY - 2, m_hOpbSrcDC, 1, 17, 11, 5, vbSrcCopy)
        If State = 2 Then Call StretchBlt(TmpDC, 1, 1, m_cX - 2, m_cY - 2, m_hOpbSrcDC, 1, 30, 11, 5, vbSrcCopy)
        If mValue = 1 Then Call TransBlt(TmpDC, (m_cX - 9) / 2, (m_cY - 8) / 2, 9, 8, m_hCkbSrcDC, 0, 0)
        If mValue = 2 Then Call TransBlt(TmpDC, (m_cX - 7) / 2, (m_cY - 7) / 2, 7, 7, m_hCkbSrcDC, 1, 9)
    Else
        Call pFrameRect(TmpDC, 0, 0, m_cX, m_cY, &HE9CFA4)
        If mValue = 1 Then Call TransBlt(TmpDC, (m_cX - 9) / 2, (m_cY - 8) / 2, 9, 8, m_hCkbSrcDC, 9, 0)
        If mValue = 2 Then Call TransBlt(TmpDC, (m_cX - 7) / 2, (m_cY - 7) / 2, 7, 7, m_hCkbSrcDC, 10, 9)
    End If
    BitBlt m_hDC, 0, (m_wRect.Height - m_cY) / 2, m_cX, m_cY, TmpDC, 0, 0, vbSrcCopy
    Call ReleaseDC(hWnd, m_hDC)
    DeleteDC TmpDC
    pDrawCheckBox = 1
End Function

Private Function pDrawRadioBox(ByVal hWnd As Long, ByVal State As Long, Optional ByVal Redraw As Boolean = False) As Long
    Dim mOldState As Long
    mOldState = GetProp(hWnd, "OLDSTATE")
    If mOldState = State And Redraw = False Then Exit Function
    Call SetProp(hWnd, "OLDSTATE", State)
    Dim m_hDC       As Long
    Dim TmpDC       As Long
    Dim m_wRect     As RECTW
    Dim m_cX        As Long
    Dim m_cY        As Long
    Dim m_dX        As Long
    Dim m_dY        As Long
    Dim mValue      As Long
    Call pGetWindowRectW(hWnd, m_wRect)
    m_cX = GetSystemMetrics(SM_CXCHECKBOX)
    m_cY = GetSystemMetrics(SM_CYCHECKBOX)
    m_dX = 0
    m_dY = (m_wRect.Height - m_cY) / 2
    mValue = SendMessage(hWnd, BM_GETCHECK, 0&, 0&)
    m_hDC = GetWindowDC(hWnd)
    Call pFillRectL(TmpDC, 0, 0, m_cX, m_cY, &HFFFFFF)
    If IsWindowEnabled(hWnd) Then
        Select Case State
            Case 0
                Call GridBlt(m_hDC, m_dX, m_dY, m_cX, m_cY, m_hOpbSrcDC, 0, 0, 13, 13, 5, 5, 5, 5, RGB(255, 0, 255))
                
            Case 1
                Call GridBlt(m_hDC, m_dX, m_dY, m_cX, m_cY, m_hOpbSrcDC, 0, 13, 13, 13, 5, 5, 5, 5, RGB(255, 0, 255))
                
            Case 2
                Call GridBlt(m_hDC, m_dX, m_dY, m_cX, m_cY, m_hOpbSrcDC, 0, 26, 13, 13, 5, 5, 5, 5, RGB(255, 0, 255))
                
        End Select
        If mValue = 1 Then Call TransBlt(m_hDC, m_dX + (m_cX - 5) / 2, m_dY + (m_cY - 5) / 2, 5, 5, m_hCkbSrcDC, 2, 18)
    Else
        Call GridBlt(m_hDC, m_dX, m_dY, m_cX, m_cY, m_hOpbSrcDC, 0, 39, 13, 13, 5, 5, 5, 5, RGB(255, 0, 255))
        If mValue = 1 Then Call TransBlt(m_hDC, m_dX + (m_cX - 5) / 2, m_dY + (m_cY - 5) / 2, 5, 5, m_hCkbSrcDC, 11, 18)
    End If
    Call ReleaseDC(hWnd, m_hDC)
    DeleteDC TmpDC
    pDrawRadioBox = 1
End Function

Private Function pDrawHeader(ByVal hWnd As Long, ByVal hDC As Long, ByVal Index As Long, ByVal State As Long) As Long
If Index < 0 Then Exit Function
    Dim m_hDC       As Long
    Dim m_wRect     As RECTW
    Dim m_iRect     As RECT
    Dim m_iWidth    As Long
    Dim m_iHeight   As Long
    Dim TmpDC       As Long
    Dim m_iText     As String
    Dim m_iItem     As HDITEM
    Call pGetWindowRectW(hWnd, m_wRect)
    Call SendMessage(hWnd, HDM_GETITEMRECT, Index, m_iRect)
    m_iWidth = m_iRect.Right - m_iRect.Left
    m_iHeight = m_iRect.Bottom - m_iRect.Top
    TmpDC = pCreateDC(m_iWidth, m_wRect.Height)
    SelectObject TmpDC, SendMessage(hWnd, WM_GETFONT, 0, 0)
    SetBkMode TmpDC, 1
    Select Case State
        Case 1
            Call GridBlt(TmpDC, 0, 0, m_iWidth, m_wRect.Height, m_hHdbSrcDC, 3, 0, 3, 20, 1, 1, 1, 1)
        
        Case 2
            Call GridBlt(TmpDC, 0, 0, m_iWidth, m_wRect.Height, m_hHdbSrcDC, 6, 0, 3, 20, 1, 1, 1, 1)
        
        Case Else
            Call GridBlt(TmpDC, 0, 0, m_iWidth, m_wRect.Height, m_hHdbSrcDC, 0, 0, 3, 20, 1, 1, 1, 1)
            
    End Select

    With m_iItem
        .mask = HDI_TEXT
        .cchTextMax = 256
        .pszText = String$(255, 0)
    End With
    Call SendMessage(hWnd, HDM_GETITEMA, Index, m_iItem)
    m_iText = Replace$(m_iItem.pszText, Chr(0), vbNullString)
    If Len(m_iText) > 0 Then
        Call pDrawTextL(TmpDC, m_iText, 0, 0, m_iWidth, m_wRect.Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS)
    End If
    m_hDC = IIf(hDC = 0, GetWindowDC(hWnd), hDC)
        BitBlt m_hDC, m_iRect.Left, 0, m_iWidth, m_iHeight, TmpDC, 0, 0, vbSrcCopy
    If hDC = 0 Then Call ReleaseDC(hWnd, m_hDC)
    DeleteDC TmpDC
    pDrawHeader = 0
End Function

Private Function OnActivate(OrigProc, hWnd, uMsg, wParam, lParam) As Long
    Dim sClassName  As String
    If lParam Then
        If GetParent(lParam) = hWnd Then
            pAttach lParam
        Else
            sClassName = LCase(pGetClassName(lParam))
            Select Case sClassName
                   '=====================================================================================
                   Case "#32770", "thunderformdc", "thunderrt6formdc", "form"
                        pAttach lParam
                        
            End Select
        End If
    End If
    OnActivate = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnPaint(OrigProc As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
Dim sClassName  As String
Dim TmpDC       As Long
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
                Dim m_Style As Long
                m_Style = GetProp(hWnd, "OLDSTYLE")
                If (m_Style And BS_CHECKBOX) Or (m_Style And BS_RADIOBUTTON) Then
                    OnPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
                Else
                    Dim PS As PAINTSTRUCT
                    Call pKillTimer(hWnd)
                    Call SetProp(hWnd, "ALPHALEVEL", 255)
                    Call BeginPaint(hWnd, PS)
                    Call pDrawButton(hWnd, PS.hDC)
                    Call EndPaint(hWnd, PS)
                    OnPaint = False
                End If
                Exit Function
        
        '=====================================================================================
        Case "thundercombobox", "thunderrt6combobox", "combo", "combobox"
            Dim m_BtSize  As Long
            Dim mOldState As Long
            m_BtSize = GetSystemMetrics(SM_CXVSCROLL)
            mOldState = GetProp(hWnd, "OLDSTATE")
            Call pGetWindowRectW(hWnd, m_wRect)
            If GetWindowLong(hWnd, GWL_STYLE) And &H1 Then
                OnPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
                m_hDC = GetWindowDC(hWnd)
                If IsWindowEnabled(hWnd) Then
                    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
                Else
                    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HE9CFA4)
                End If
                Call pFrameRect(m_hDC, 1, 1, m_wRect.Width - 2, m_wRect.Height - 2, &HFFFFFF)
                Call pDrawComboBox(hWnd, m_hDC, mOldState, True)
                Call ReleaseDC(hWnd, m_hDC)
           Else
                OnPaint = False
                Call BeginPaint(hWnd, PS)
                Call pFillRectL(PS.hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HFFFFFF)
                If IsWindowEnabled(hWnd) Then
                    Call pFrameRect(PS.hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
                Else
                    Call pFrameRect(PS.hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HE9CFA4)
                End If
                Call pDrawComboBox(hWnd, PS.hDC, mOldState, True)
                Call EndPaint(hWnd, PS)
            End If
            Exit Function
        
        Case "thunderdrivelistbox", "thunderrt6drivelistbox"
                OnPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
                m_BtSize = GetSystemMetrics(SM_CXVSCROLL)
                mOldState = GetProp(hWnd, "OLDSTATE")
                Call pGetWindowRectW(hWnd, m_wRect)
                m_hDC = GetWindowDC(hWnd)
                If IsWindowEnabled(hWnd) Then
                    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
                Else
                    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HE9CFA4)
                End If
                Call pFrameRect(m_hDC, 1, 1, m_wRect.Width - 2, m_wRect.Height - 2, &HFFFFFF)
                Call pDrawComboBox(hWnd, m_hDC, mOldState, True)
                Call ReleaseDC(hWnd, m_hDC)
                Exit Function
                
        '=====================================================================================
        Case "thundercheckbox", "thunderrt6checkbox"
            OnPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pDrawCheckBox(hWnd, 0, True)
            Exit Function
        
        '=====================================================================================
        Case "thunderoptionbutton", "thunderrt6optionbutton"
            OnPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pDrawRadioBox(hWnd, 0, True)
            Exit Function
            
        '=====================================================================================
        Case "datalistwndclass", "dblistwndclass"
            OnPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pGetWindowRectW(hWnd, m_wRect)
            m_hDC = GetWindowDC(hWnd)
            If IsWindowEnabled(hWnd) Then
                Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
            Else
                Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HE9CFA4)
            End If
            Call pFrameRect(m_hDC, 1, 1, m_wRect.Width - 2, m_wRect.Height - 2, &HFFFFFF)
            Call ReleaseDC(hWnd, m_hDC)
            Exit Function
        
        Case "combolbox"
            OnPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pGetWindowRectW(hWnd, m_wRect)
            m_hDC = GetWindowDC(hWnd)
            If IsWindowEnabled(hWnd) Then
                Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
            Else
                Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HE9CFA4)
            End If
            Call pFrameRect(m_hDC, 1, 1, m_wRect.Width - 2, m_wRect.Height - 2, &HFFFFFF)
            Call ReleaseDC(hWnd, m_hDC)
            Exit Function
        
        '=====================================================================================
        Case "progressbar20wndclass", "progressbarwndclass"
            OnPaint = False
            Dim dwStyle As Long
            Dim lPos    As Single
            Dim lW      As Long
            Dim lH      As Long
            Dim c_wRect As RECT
            dwStyle = GetWindowLong(hWnd, GWL_STYLE)
            Call GetClientRect(hWnd, c_wRect)
            lW = c_wRect.Right - c_wRect.Left
            lH = c_wRect.Bottom - c_wRect.Top
            TmpDC = pCreateDC(lW, lH)
            lPos = SendMessage(hWnd, PBM_GETPOS, 0, 0) / &H7530
            Call pFillRectL(TmpDC, 0, 0, lW, lH, &HFFFFFF)
            If lPos Then
                If dwStyle And PBS_VERTICAL Then
                    Call pFillRectL(TmpDC, 0, lH - lH * lPos, lW, lH * lPos, &HF6D64C)
                Else
                    Call pFillRectL(TmpDC, 0, 0, lW * lPos, lH, &HF6D64C)
                End If
            End If
            Call BeginPaint(hWnd, PS)
                BitBlt PS.hDC, 0, 0, lW, lH, TmpDC, 0, 0, vbSrcCopy
            Call EndPaint(hWnd, PS)
            DeleteDC TmpDC
            Exit Function
        
        '=====================================================================================
        Case "msvb_lib_header", "sysheader32"
            OnPaint = False
            Dim m_iCount As Long
            Call pGetWindowRectW(hWnd, m_wRect)
            TmpDC = pCreateDC(m_wRect.Width, m_wRect.Height)
            StretchBlt TmpDC, 0, 0, m_wRect.Width, m_wRect.Height, m_hHdbSrcDC, 1, 0, 1, 20, vbSrcCopy
            m_iCount = SendMessage(hWnd, HDM_GETITEMCOUNT, 0&, 0&)
            If m_iCount > 0 Then
                Dim m_iIndex As Long
                Dim m_mIndex As Long
                Dim i As Long
                m_iIndex = GetProp(hWnd, "HDINDEX")
                m_mIndex = GetProp(hWnd, "HMINDEX")
                For i = 0 To m_iCount - 1
                    Dim m_iText As String
                    Dim m_iRect As RECT
                    Dim m_iItem  As HDITEM
                    With m_iItem
                        .mask = HDI_TEXT
                        .cchTextMax = 256
                        .pszText = String$(255, 0)
                    End With
                    Call SendMessage(hWnd, HDM_GETITEMRECT, i, m_iRect)
                    Call SendMessage(hWnd, HDM_GETITEMA, i, m_iItem)
                    m_iText = Replace$(m_iItem.pszText, Chr(0), vbNullString)
                    If i = m_iIndex Then
                        Call GridBlt(TmpDC, m_iRect.Left, 0, m_iRect.Right - m_iRect.Left, m_wRect.Height, m_hHdbSrcDC, 6, 0, 3, 20, 1, 1, 1, 1)
                    ElseIf i = m_mIndex Then
                        Call GridBlt(TmpDC, m_iRect.Left, 0, m_iRect.Right - m_iRect.Left, m_wRect.Height, m_hHdbSrcDC, 3, 0, 3, 20, 1, 1, 1, 1)
                    Else
                        Call GridBlt(TmpDC, m_iRect.Left, 0, m_iRect.Right - m_iRect.Left, m_wRect.Height, m_hHdbSrcDC, 0, 0, 3, 20, 1, 1, 1, 1)
                    End If
                    If Len(m_iText) > 0 Then
                        SelectObject TmpDC, SendMessage(hWnd, WM_GETFONT, 0, 0)
                        SetBkMode TmpDC, 1
                        Call pDrawTextL(TmpDC, m_iText, m_iRect.Left, 0, m_iRect.Right - m_iRect.Left, m_iRect.Bottom, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS)
                    End If
                Next
            End If
            Call BeginPaint(hWnd, PS)
                BitBlt PS.hDC, 0, 0, m_wRect.Width, m_wRect.Height, TmpDC, 0, 0, vbSrcCopy
            Call EndPaint(hWnd, PS)
            Exit Function
            
        Case Else
            If sClassName Like "apexgrid*" Then
                OnPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
                dwStyle = GetWindowLong(hWnd, GWL_STYLE)
                If dwStyle And &H800000 Then
                    Call pGetWindowRectW(hWnd, m_wRect)
                    m_hDC = GetWindowDC(hWnd)
                        If IsWindowEnabled(hWnd) Then
                            Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
                        Else
                            Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HE9CFA4)
                        End If
                    Call ReleaseDC(hWnd, m_hDC)
                End If
                Exit Function
            End If
                          
    End Select
    OnPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnNcPaint(OrigProc, hWnd, uMsg, wParam, lParam) As Long
Dim sClassName  As String
Dim TmpDC       As Long
Dim m_hDC       As Long
Dim m_wRect     As RECTW
Dim cX          As Long
Dim cY          As Long
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "progressbar20wndclass", "progressbarwndclass"
            Dim dwStyle As Long
            OnNcPaint = False
            Dim lPos    As Single
            Dim lW      As Long
            Dim lH      As Long
            Call pGetBorderSize(hWnd, cX, cY)
            Dim c_wRect As RECT
            Call pGetWindowRectW(hWnd, m_wRect)
            Call GetClientRect(hWnd, c_wRect)
            Call pGetBorderSize(hWnd, cX, cY)
            lW = c_wRect.Right - c_wRect.Left
            lH = c_wRect.Bottom - c_wRect.Top
            TmpDC = pCreateDC(m_wRect.Width, m_wRect.Height)
            lPos = SendMessage(hWnd, PBM_GETPOS, 0, 0) / &H7530
            dwStyle = GetWindowLong(hWnd, GWL_STYLE)
            Call pFillRectL(TmpDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
            Call pFillRectL(TmpDC, 1, 1, m_wRect.Width - 2, m_wRect.Height - 2, &HFFFFFF)
            If lPos Then
                If dwStyle And PBS_VERTICAL Then
                    If dwStyle And &H1 Then Call pFillRectL(TmpDC, cX, cY, lW, lH, &HF6D64C)
                    Call pFillRectL(TmpDC, cX, cY, lW, lH - lH * lPos, &HFFFFFF)
                Else
                    Call pFillRectL(TmpDC, cX, cY, lW * lPos, lH, &HF6D64C)
                End If
            End If
            m_hDC = GetWindowDC(hWnd)
                BitBlt m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, TmpDC, 0, 0, vbSrcCopy
            Call ReleaseDC(hWnd, m_hDC)
            DeleteDC TmpDC
            Exit Function
        '=====================================================================================
        Case "datalistwndclass", "dblistwndclass"
        
        '=====================================================================================
        Case Else
            OnNcPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_CLIENTEDGE Then
                Dim m_cDC       As Long
                Dim m_Width     As Long
                Dim m_Height    As Long
                Dim i           As Long
                Call pGetBorderSize(hWnd, cX, cY)
                Call pGetWindowRectW(hWnd, m_wRect)
                m_hDC = GetWindowDC(hWnd)
                m_cDC = GetDC(hWnd)
                If IsWindowEnabled(hWnd) Then
                    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
                Else
                    Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HE9CFA4)
                End If
                If cX > 0 Then
                    For i = 1 To cX
                        Call pFrameRect(m_hDC, i, i, m_wRect.Width - 1 - i, m_wRect.Height - 1 - i, GetBkColor(m_cDC))
                    Next
                End If
                ReleaseDC hWnd, m_cDC
                ReleaseDC hWnd, m_hDC
            End If
            Exit Function
            
    End Select
    OnNcPaint = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnSize(OrigProc As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            OnSize = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Dim i As Long
            Dim m_mDC(3)    As Long
            Dim m_BMP(3)    As Long
            m_hDC = GetWindowDC(hWnd)
            Call pGetWindowRectW(hWnd, m_wRect)
            For i = 0 To 3
                m_mDC(i) = GetProp(hWnd, "HDC" & CStr(i))
                m_BMP(i) = CreateCompatibleBitmap(m_hDC, m_wRect.Width, m_wRect.Height)
                DeleteObject SelectObject(m_mDC(i), m_BMP(i))
            Next
            Call pDrawMemDC(hWnd)
            ReleaseDC hWnd, m_hDC
            SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, m_wRect.Width + 1, m_wRect.Height + 1, 3, 3), True
            Exit Function
        
        '=====================================================================================
        Case "progressbar20wndclass", "progressbarwndclass"
            'OnSize = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            'Call pGetWindowRectW(hWnd, m_wRect)
            'SetWindowRgn hWnd, CreateRoundRectRgn(0, 0, m_wRect.Width + 1, m_wRect.Height + 1, 8, 8), True
            'Exit Function
            
        '=====================================================================================
        Case Else
                            
    End Select
    OnSize = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnLButtonDBClick(OrigProc As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            OnLButtonDBClick = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pKillTimer(hWnd)
            Call SetProp(hWnd, "OLDSTATE", 2)
            Call SetProp(hWnd, "ALPHALEVEL", 255)
            Call pDrawButton(hWnd, 0)
            Call SetProp(hWnd, "ALPHALEVEL", 0)
            Exit Function
            
        '=====================================================================================
        Case "msvb_lib_header", "sysheader32"
            OnLButtonDBClick = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Dim m_iCount As Long
            m_iCount = SendMessage(hWnd, HDM_GETITEMCOUNT, 0&, 0&)
            If m_iCount > 0 Then
                Dim iIndex  As Long
                Dim mX      As Long
                Dim i       As Long
                mX = LWORD(lParam)
                iIndex = -1
                For i = 0 To m_iCount - 1
                    Dim m_iRect As RECT
                    Call SendMessage(hWnd, HDM_GETITEMRECT, i, m_iRect)
                    If mX >= m_iRect.Left And mX <= m_iRect.Right Then
                        iIndex = i
                        Exit For
                    End If
                Next
                SetProp hWnd, "HDINDEX", iIndex
                If iIndex >= 0 Then Call pDrawHeader(hWnd, 0, iIndex, 2)
            End If
            Exit Function
            
    End Select
    OnLButtonDBClick = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnLButtonDown(OrigProc As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            OnLButtonDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pKillTimer(hWnd)
            Call SetProp(hWnd, "OLDSTATE", 2)
            Call SetProp(hWnd, "ALPHALEVEL", 255)
            Call pDrawButton(hWnd, 0)
            Call SetProp(hWnd, "ALPHALEVEL", 0)
            Exit Function
        '=====================================================================================
        Case "#32770", "thunderformdc", "thunderrt6formdc", "form"
                                
        '=====================================================================================
        Case "thundercombobox", "thunderrt6combobox", "combo", "combobox"
            OnLButtonDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Dim mX          As Long
            Dim mY          As Long
            Dim m_BtSize    As Long
            mX = LWORD(lParam)
            mY = HWORD(lParam)
            Call pGetWindowRectW(hWnd, m_wRect)
            m_BtSize = GetSystemMetrics(SM_CXVSCROLL)
            If GetWindowLong(hWnd, GWL_STYLE) And &H1 Then
                pDrawComboBox hWnd, 0, 2
            ElseIf mX >= m_wRect.Width - m_BtSize - 2 And mX <= m_wRect.Width - 2 And mY >= 2 And mY <= m_wRect.Height - 2 Then
                pDrawComboBox hWnd, 0, 2
            End If
            Exit Function
            
         '=====================================================================================
         Case "thunderdrivelistbox", "thunderrt6drivelistbox"
            OnLButtonDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            pDrawComboBox hWnd, 0, 2
            Exit Function
            
        '=====================================================================================
         Case "thundercheckbox", "thunderrt6checkbox"
            OnLButtonDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pDrawCheckBox(hWnd, 2)
            Exit Function
                
        '=====================================================================================
         Case "thunderoptionbutton", "thunderrt6optionbutton"
            OnLButtonDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pDrawRadioBox(hWnd, 2)
            Exit Function
            
        '=====================================================================================
        Case "msvb_lib_header", "sysheader32"
            OnLButtonDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Dim m_iCount As Long
            m_iCount = SendMessage(hWnd, HDM_GETITEMCOUNT, 0&, 0&)
            If m_iCount > 0 Then
                Dim iIndex  As Long
                Dim i       As Long
                mX = LWORD(lParam)
                iIndex = -1
                For i = 0 To m_iCount - 1
                    Dim m_iRect As RECT
                    Call SendMessage(hWnd, HDM_GETITEMRECT, i, m_iRect)
                    If mX >= m_iRect.Left And mX <= m_iRect.Right Then
                        iIndex = i
                        Exit For
                    End If
                Next
                SetProp hWnd, "HDINDEX", iIndex
                If iIndex >= 0 Then Call pDrawHeader(hWnd, 0, iIndex, 2)
            End If
            Exit Function
            
    End Select
    OnLButtonDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnLButtonUp(OrigProc As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            OnLButtonUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call SetProp(hWnd, "MOUSEFLAG", 0)
            Call SetProp(hWnd, "OLDSTATE", 0)
            Call SetProp(hWnd, "ALPHALEVEL", 60)
            Call pSetTimer(hWnd)
            Exit Function
        
        '=====================================================================================
        Case "thundercombobox", "thunderrt6combobox", "combo", "combobox"
            OnLButtonUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Dim Pt As POINTAPI
            Call GetCursorPos(Pt)
            If WindowFromPoint(Pt.X, Pt.Y) = hWnd Then
                pDrawComboBox hWnd, 0, 1
            Else
                pDrawComboBox hWnd, 0, 0
            End If
            Exit Function
            
         '=====================================================================================
         Case "thunderdrivelistbox", "thunderrt6drivelistbox"
            OnLButtonUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call GetCursorPos(Pt)
            If WindowFromPoint(Pt.X, Pt.Y) = hWnd Then
                pDrawComboBox hWnd, 0, 1
            Else
                pDrawComboBox hWnd, 0, 0
            End If
            Exit Function
            
        '=====================================================================================
         Case "thundercheckbox", "thunderrt6checkbox"
            OnLButtonUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pDrawCheckBox(hWnd, 0)
            Exit Function
            
        '=====================================================================================
         Case "thunderoptionbutton", "thunderrt6optionbutton"
            OnLButtonUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pDrawRadioBox(hWnd, 0)
            Exit Function
            
        '=====================================================================================
        Case "msvb_lib_header", "sysheader32"
            OnLButtonUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Dim m_oIndex As Long
            m_oIndex = GetProp(hWnd, "HDINDEX")
            SetProp hWnd, "HDINDEX", -1
            If m_oIndex >= 0 Then
                Call GetCursorPos(Pt)
                If WindowFromPoint(Pt.X, Pt.Y) = hWnd Then
                    Call pDrawHeader(hWnd, 0, m_oIndex, 0)
                Else
                    Call pDrawHeader(hWnd, 0, m_oIndex, 1)
                End If
            End If
            Exit Function
            
    End Select
    OnLButtonUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnKeyDown(OrigProc As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            OnKeyDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If wParam = 32 Then
                Call pKillTimer(hWnd)
                Call SetProp(hWnd, "OLDSTATE", 2)
                Call SetProp(hWnd, "ALPHALEVEL", 255)
                Call pDrawButton(hWnd, 0)
                Call SetProp(hWnd, "ALPHALEVEL", 0)
            End If
            Exit Function
        
        '=====================================================================================
         Case "thundercheckbox", "thunderrt6checkbox"
            OnKeyDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If wParam = 32 Then Call pDrawCheckBox(hWnd, 2)
            Exit Function
            
        '=====================================================================================
         Case "thunderoptionbutton", "thunderrt6optionbutton"
            OnKeyDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If wParam = 32 Then Call pDrawRadioBox(hWnd, 2)
            Exit Function
    
    End Select
    OnKeyDown = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnKeyUp(OrigProc As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            If wParam = 32 Then
                OnKeyUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
                Call SetProp(hWnd, "MOUSEFLAG", 0)
                Call SetProp(hWnd, "ALPHALEVEL", 0)
                Call SetProp(hWnd, "OLDSTATE", 0)
                Call pSetTimer(hWnd)
                Exit Function
            End If
        
        '=====================================================================================
         Case "thundercheckbox", "thunderrt6checkbox"
            OnKeyUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If wParam = 32 Then Call pDrawCheckBox(hWnd, 0)
            Exit Function
        
         Case "thunderoptionbutton", "thunderrt6optionbutton"
            OnKeyUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If wParam = 32 Then Call pDrawRadioBox(hWnd, 0)
            Exit Function
    
    End Select
    OnKeyUp = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnMouseMove(OrigProc As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            OnMouseMove = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If GetProp(hWnd, "MOUSEFLAG") = 0 Then
                Call SetProp(hWnd, "MOUSEFLAG", 1)
                Call pTrackMouseTracking(hWnd)
                Call SetProp(hWnd, "OLDSTATE", 1)
                Call SetProp(hWnd, "ALPHALEVEL", 0)
                Call pSetTimer(hWnd)
            Else
                Dim Pt As POINTAPI
                Call GetCursorPos(Pt)
                If Not WindowFromPoint(Pt.X, Pt.Y) = hWnd Then
                    Call pKillTimer(hWnd)
                    Call SetProp(hWnd, "ALPHALEVEL", 255)
                    Call pDrawButton(hWnd, 0)
                End If
            End If
            Exit Function
        
        '=====================================================================================
        Case "thundercombobox", "thunderrt6combobox", "combo", "combobox"
            OnMouseMove = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Dim mX          As Long
            Dim mY          As Long
            Dim m_BtSize    As Long
            If GetProp(hWnd, "MOUSEFLAG") = 0 Then
                Call SetProp(hWnd, "MOUSEFLAG", 1)
                Call pTrackMouseTracking(hWnd)
            End If
            mX = LWORD(lParam)
            mY = HWORD(lParam)
            Call pGetWindowRectW(hWnd, m_wRect)
            m_BtSize = GetSystemMetrics(SM_CXVSCROLL)
            If GetWindowLong(hWnd, GWL_STYLE) And &H1 Then
                If wParam And MK_LBUTTON Then
                    pDrawComboBox hWnd, 0, 2
                Else
                    pDrawComboBox hWnd, 0, 1
                End If
            ElseIf mX >= m_wRect.Width - m_BtSize - 2 And mX <= m_wRect.Width - 2 And mY >= 2 And mY <= m_wRect.Height - 2 Then
                If wParam And MK_LBUTTON Then
                    pDrawComboBox hWnd, 0, 2
                Else
                    pDrawComboBox hWnd, 0, 1
                End If
            Else
                pDrawComboBox hWnd, 0, 0
            End If
            Exit Function
            
         '=====================================================================================
         Case "thunderdrivelistbox", "thunderrt6drivelistbox"
            OnMouseMove = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If GetProp(hWnd, "MOUSEFLAG") = 0 Then
                Call SetProp(hWnd, "MOUSEFLAG", 1)
                Call pTrackMouseTracking(hWnd)
            End If
            If wParam And MK_LBUTTON Then
                pDrawComboBox hWnd, 0, 2
            Else
                pDrawComboBox hWnd, 0, 1
            End If
           Exit Function
         
        '=====================================================================================
         Case "thundercheckbox", "thunderrt6checkbox"
            OnMouseMove = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If GetProp(hWnd, "MOUSEFLAG") = 0 Then
                Call SetProp(hWnd, "MOUSEFLAG", 1)
                Call pTrackMouseTracking(hWnd)
            End If
            If wParam And MK_LBUTTON Then
                Call pDrawCheckBox(hWnd, 2, True)
            Else
                Call pDrawCheckBox(hWnd, 1, True)
            End If
            Exit Function
            
         Case "thunderoptionbutton", "thunderrt6optionbutton"
            OnMouseMove = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If GetProp(hWnd, "MOUSEFLAG") = 0 Then
                Call SetProp(hWnd, "MOUSEFLAG", 1)
                Call pTrackMouseTracking(hWnd)
            End If
            If wParam And MK_LBUTTON Then
                Call pDrawRadioBox(hWnd, 2, True)
            Else
                Call pDrawRadioBox(hWnd, 1, True)
            End If
            Exit Function
            
        '=====================================================================================
        Case "msvb_lib_header", "sysheader32"
            OnMouseMove = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If GetProp(hWnd, "MOUSEFLAG") = 0 Then
                Call SetProp(hWnd, "MOUSEFLAG", 1)
                Call pTrackMouseTracking(hWnd)
            End If
            Dim m_iCount As Long
            m_iCount = SendMessage(hWnd, HDM_GETITEMCOUNT, 0&, 0&)
            Dim m_iIndex  As Long
            Dim m_oIndex  As Long
            Dim m_iRect   As RECT
            mX = LWORD(lParam)
            m_iIndex = -1
            m_oIndex = GetProp(hWnd, "HMINDEX")
            If m_iCount > 0 Then
                Dim i       As Long
                For i = 0 To m_iCount - 1
                    Call SendMessage(hWnd, HDM_GETITEMRECT, i, m_iRect)
                    If mX >= m_iRect.Left And mX <= m_iRect.Right Then
                        m_iIndex = i
                        Exit For
                    End If
                Next
            End If
            Dim m_dIndex As Long
            m_dIndex = GetProp(hWnd, "HDINDEX")
            If Not m_oIndex = m_iIndex Then
                If m_oIndex >= 0 And Not m_oIndex = m_dIndex Then Call pDrawHeader(hWnd, 0, m_oIndex, 0)
                If wParam And MK_LBUTTON Then
                    If m_dIndex = m_iIndex Then Call pDrawHeader(hWnd, 0, m_iIndex, 2)
                Else
                    SetProp hWnd, "HMINDEX", m_iIndex
                    Call pDrawHeader(hWnd, 0, m_iIndex, 1)
                End If
            End If
            Exit Function
            
    End Select
    OnMouseMove = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnMouseLeave(OrigProc As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            OnMouseLeave = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call SetProp(hWnd, "MOUSEFLAG", 0)
            Call SetProp(hWnd, "OLDSTATE", 1)
            Call SetProp(hWnd, "ALPHALEVEL", 140)
            Call pDrawButton(hWnd, 0)
            Call SetProp(hWnd, "OLDSTATE", 0)
            Call SetProp(hWnd, "ALPHALEVEL", 0)
            Call pSetTimer(hWnd)
            Exit Function
            
        '=====================================================================================
        Case "#32770", "thunderformdc", "thunderrt6formdc", "form"
                                
        '=====================================================================================
        Case "thundercombobox", "thunderrt6combobox", "combo", "combobox", "thunderdrivelistbox", "thunderrt6drivelistbox"
            OnMouseLeave = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call SetProp(hWnd, "MOUSEFLAG", 0)
            pDrawComboBox hWnd, 0, 0
            Exit Function
        
        '=====================================================================================
         Case "thundercheckbox", "thunderrt6checkbox"
            OnMouseLeave = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call SetProp(hWnd, "MOUSEFLAG", 0)
            Call pDrawCheckBox(hWnd, 0)
            Exit Function
        
        
        '=====================================================================================
         Case "thunderoptionbutton", "thunderrt6optionbutton"
            OnMouseLeave = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call SetProp(hWnd, "MOUSEFLAG", 0)
            Call pDrawRadioBox(hWnd, 0)
            Exit Function
        
        '=====================================================================================
        Case "msvb_lib_header", "sysheader32"
            OnMouseLeave = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Dim m_oIndex As Long
            m_oIndex = GetProp(hWnd, "HMINDEX")
            Call SetProp(hWnd, "MOUSEFLAG", 0)
            Call SetProp(hWnd, "HDINDEX", -1)
            Call SetProp(hWnd, "HMINDEX", -1)
            If m_oIndex >= 0 Then Call pDrawHeader(hWnd, 0, m_oIndex, 0)
            Exit Function
                            
    End Select
    OnMouseLeave = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnTimer(OrigProc As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            Dim m_Level As Long
            m_Level = GetProp(hWnd, "ALPHALEVEL")
            Call pDrawButton(hWnd, 0)
            m_Level = m_Level + 3       '这里的+3是速度，可以改变Timer的时间和+的数值以改变速度
            If m_Level > 255 Then
                Call pKillTimer(hWnd)
                Call SetProp(hWnd, "ALPHALEVEL", 0)
            Else
                Call SetProp(hWnd, "ALPHALEVEL", m_Level)
            End If
                            
    End Select
    OnTimer = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnEnable(OrigProc, hWnd, uMsg, wParam, lParam) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            OnEnable = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            If wParam Then
                Call SetProp(hWnd, "OLDSTATE", 0)
            Else
                Call SetProp(hWnd, "OLDSTATE", 3)
            End If
            Call pDrawButton(hWnd, 0)
            Exit Function
    
    End Select
    OnEnable = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnSetFocus(OrigProc, hWnd, uMsg, wParam, lParam) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            OnSetFocus = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pDrawMemDC(hWnd)
            Call SetProp(hWnd, "ALPHALEVEL", 0)
            Call pSetTimer(hWnd)
            Exit Function
        
        '=====================================================================================
        Case "thunderoptionbutton", "thunderrt6optionbutton"
            OnSetFocus = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pDrawRadioBox(hWnd, 0, True)
            Exit Function
    
    End Select
    OnSetFocus = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnKillFocus(OrigProc, hWnd, uMsg, wParam, lParam) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercommandbutton", "thunderrt6commandbutton", "button"
            OnKillFocus = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pDrawMemDC(hWnd)
            Call SetProp(hWnd, "ALPHALEVEL", 0)
            Call pSetTimer(hWnd)
            Exit Function
        
        '=====================================================================================
        Case "thunderoptionbutton", "thunderrt6optionbutton"
            OnKillFocus = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            Call pDrawRadioBox(hWnd, 0, True)
            Exit Function
    
    End Select
    OnKillFocus = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function OnWindowposChanging(OrigProc, hWnd, uMsg, wParam, lParam) As Long
Dim sClassName  As String
Dim m_hDC       As Long
Dim m_wRect     As RECTW
    sClassName = LCase(pGetClassName(hWnd))
    Select Case sClassName
        '=====================================================================================
        Case "thundercombobox", "thunderrt6combobox", "combo", "combobox", "thunderdrivelistbox", "thunderrt6drivelistbox"
            OnWindowposChanging = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            pDrawComboBox hWnd, 0, 0
            Exit Function
    
    End Select
    OnWindowposChanging = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim OrigProc    As Long
If hWnd = 0 Then Exit Function
    OrigProc = GetProp(hWnd, "PROCADDR")
    If Not OrigProc = 0 Then
        If uMsg = WM_DESTROY Then
            Call pDetach(hWnd)
            WindowProc = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
        Else
            Select Case uMsg
                    Case WM_ACTIVATE
                        WindowProc = OnActivate(OrigProc, hWnd, uMsg, wParam, lParam)
                    
                    Case WM_PAINT
                        WindowProc = OnPaint(OrigProc, hWnd, uMsg, wParam, lParam)
                    
                    Case WM_NCPAINT
                        WindowProc = OnNcPaint(OrigProc, hWnd, uMsg, wParam, lParam)
                        
                    Case WM_SIZE
                        WindowProc = OnSize(OrigProc, hWnd, uMsg, wParam, lParam)
                    
                    Case WM_LBUTTONDBLCLK
                        WindowProc = OnLButtonDBClick(OrigProc, hWnd, uMsg, wParam, lParam)
                        
                    Case WM_LBUTTONDOWN
                        WindowProc = OnLButtonDown(OrigProc, hWnd, uMsg, wParam, lParam)
                        
                    Case WM_LBUTTONUP
                        WindowProc = OnLButtonUp(OrigProc, hWnd, uMsg, wParam, lParam)
                        Dim hParent As Long
                        hParent = GetParent(hWnd)
                        If hParent = 0 Then hParent = hWnd
                        Call EnumChildWindows(hParent, AddressOf pEnumRedrawProc, 0&)
                        
                    Case WM_KEYDOWN
                        WindowProc = OnKeyDown(OrigProc, hWnd, uMsg, wParam, lParam)
                        
                    Case WM_KEYUP
                        WindowProc = OnKeyUp(OrigProc, hWnd, uMsg, wParam, lParam)
                        
                    Case WM_MOUSEMOVE
                        WindowProc = OnMouseMove(OrigProc, hWnd, uMsg, wParam, lParam)
                        
                    Case WM_MOUSELEAVE
                        WindowProc = OnMouseLeave(OrigProc, hWnd, uMsg, wParam, lParam)
                        
                    Case WM_TIMER
                        WindowProc = OnTimer(OrigProc, hWnd, uMsg, wParam, lParam)
                    
                    Case WM_ENABLE
                        WindowProc = OnEnable(OrigProc, hWnd, uMsg, wParam, lParam)
                    
                    Case WM_SETFOCUS
                        WindowProc = OnSetFocus(OrigProc, hWnd, uMsg, wParam, lParam)
                        
                    Case WM_KILLFOCUS
                        WindowProc = OnKillFocus(OrigProc, hWnd, uMsg, wParam, lParam)
                    
                    Case WM_WINDOWPOSCHANGING
                        WindowProc = OnWindowposChanging(OrigProc, hWnd, uMsg, wParam, lParam)
                    
                    Case WM_CTLCOLORLISTBOX
                        WindowProc = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
                        If lParam Then pAttach (lParam)
                                           
                    Case WM_ENTERIDLE
                        WindowProc = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
                        If lParam Then pAttach lParam
                                                                
                    Case Else
                        'Debug.Print uMsg & "|" & wParam & "|" & lParam & "|" & Now
                        WindowProc = CallWindowProc(OrigProc, hWnd, uMsg, wParam, lParam)
            
            End Select
        End If
    Else
        WindowProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function

Private Function pEnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Call pAttach(hWnd)
    pEnumChildProc = 1
End Function

Private Function pEnumRedrawProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    If GetProp(hWnd, "PROCADDR") Then
        Dim sClassName As String
        sClassName = LCase(pGetClassName(hWnd))
        Select Case sClassName
            Case "thundercheckbox", "thunderrt6checkbox"
                Call pDrawCheckBox(hWnd, 0, True)
                
            Case "thunderoptionbutton", "thunderrt6optionbutton"
                Call pDrawRadioBox(hWnd, 0, True)
                
        End Select
    End If
    pEnumRedrawProc = 1
End Function

Private Function pDeEnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Call pDetach(hWnd)
    pDeEnumChildProc = 1
End Function

Private Function pGetClassName(ByVal hWnd As Long) As String
On Error Resume Next
    Dim BuffStr     As String
    Dim BuffStrLen  As Long
    Dim Rtn         As Long
    BuffStr = String$(255, Chr(0))
    BuffStrLen = Len(BuffStr)
    Rtn = GetClassName(hWnd, ByVal BuffStr, BuffStrLen)
    If Not Rtn = 0 Then
        Dim iPos As Long
        iPos = InStr(1, BuffStr, Chr(0)) - 1
        If iPos < Len(BuffStr) Then
            pGetClassName = Left$(BuffStr, iPos)
        Else
            pGetClassName = BuffStr
        End If
    End If
End Function

Private Function pGetWindowText(ByVal hWnd As Long) As String
    Dim BuffStr     As String
    Dim BuffStrLen  As Long
    BuffStrLen = GetWindowTextLength(hWnd)
    BuffStr = String(BuffStrLen, Chr(0))
    Call GetWindowText(hWnd, ByVal BuffStr, BuffStrLen + 1)
    pGetWindowText = BuffStr
End Function

Private Function pGetText(ByVal hWnd As Long) As String
    Dim BuffStr As String, BuffStrLen As Long, Rtn As Long
    BuffStrLen = GetWindowTextLength(hWnd)
    BuffStr = String(BuffStrLen, Chr(0))
    Rtn = SendMessage(hWnd, WM_GETTEXT, BuffStrLen + 1, ByVal BuffStr)
    pGetText = BuffStr
End Function

Private Function pDrawText(ByVal hDC As Long, ByVal Text As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal lpFlag As DTSTYLE) As Long
    Dim TmpRect As RECT
    With TmpRect
        .Left = X1
        .Top = Y1
        .Right = X2
        .Bottom = Y2
    End With
    pDrawText = DrawText(hDC, Text, -1, TmpRect, lpFlag)
End Function

Private Function pDrawTextL(ByVal hDC As Long, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal lpFlag As DTSTYLE) As Long
    Dim TmpRect As RECT
    With TmpRect
        .Left = X
        .Top = Y
        .Right = X + Width
        .Bottom = Y + Height
    End With
    pDrawTextL = DrawText(hDC, Text, -1, TmpRect, lpFlag)
End Function

Private Function pFillRect(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long) As Long
    Dim TmpRect     As RECT
    Dim TmpBrush    As Long
    With TmpRect
        .Left = X1
        .Top = Y1
        .Right = X2
        .Bottom = Y2
    End With
    TmpBrush = CreateSolidBrush(Color)
    pFillRect = FillRect(hDC, TmpRect, TmpBrush)
    DeleteObject TmpBrush
End Function

Private Function pFillRectL(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long) As Long
    Dim TmpRect     As RECT
    Dim TmpBrush    As Long
    With TmpRect
        .Left = X
        .Top = Y
        .Right = X + Width
        .Bottom = Y + Height
    End With
    TmpBrush = CreateSolidBrush(Color)
    pFillRectL = FillRect(hDC, TmpRect, TmpBrush)
    DeleteObject TmpBrush
End Function

Private Function pGetWindowRectW(ByVal hWnd As Long, lpRectW As RECTW) As Long
    Dim TmpRect As RECT
    Dim Rtn     As Long
    Rtn = GetWindowRect(hWnd, TmpRect)
    With lpRectW
        .Left = TmpRect.Left
        .Top = TmpRect.Top
        .Right = TmpRect.Right
        .Bottom = TmpRect.Bottom
        .Width = TmpRect.Right - TmpRect.Left
        .Height = TmpRect.Bottom - TmpRect.Top
    End With
    pGetWindowRectW = Rtn
End Function

Private Sub pSetRectW(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, lpRectW As RECTW)
    With lpRectW
        .Left = X1
        .Right = X2
        .Top = Y1
        .Bottom = Y2
        .Width = X2 - X1
        .Height = Y2 - Y1
    End With
End Sub

Private Function pDrawFocusRect(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
    Dim TmpRect As RECT
    With TmpRect
        .Left = X
        .Top = Y
        .Right = X + Width
        .Bottom = Y + Height
    End With
    pDrawFocusRect = DrawFocusRect(hDC, TmpRect)
End Function

Private Function pFrameRect(ByVal hDC As Long, ByVal X As Long, Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long) As Long
    Dim TmpRect     As RECT
    Dim m_hBrush    As Long
    With TmpRect
        .Left = X
        .Top = Y
        .Right = X + Width
        .Bottom = Y + Height
    End With
    m_hBrush = CreateSolidBrush(Color)
    pFrameRect = FrameRect(hDC, TmpRect, m_hBrush)
    DeleteObject m_hBrush
End Function

Private Function pDrawBorderLine(ByVal hWnd As Long, ByVal State As Long) As Long
    Dim m_wRect As RECTW
    Dim m_hDC   As Long
    If GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_CLIENTEDGE Then
        Call pGetWindowRectW(hWnd, m_wRect)
        m_hDC = GetWindowDC(hWnd)
        If State = 0 Then
            Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HD5A554)
            Call pFrameRect(m_hDC, 1, 1, m_wRect.Width - 2, m_wRect.Height - 2, &HF4E7D3)
        Else
            Call pFrameRect(m_hDC, 0, 0, m_wRect.Width, m_wRect.Height, &HF4E7D3)
            Call pFrameRect(m_hDC, 1, 1, m_wRect.Width - 2, m_wRect.Height - 2, &HD5A554)
        End If
        ReleaseDC hWnd, m_hDC
        pDrawBorderLine = 1
    End If
End Function

Private Function pSetTimer(ByVal hWnd As Long) As Long
    Dim m_TimerID       As Long
    m_TimerID = GetProp(hWnd, "TIMERID")
    If m_TimerID Then Exit Function
    m_TimerID = SetTimer(hWnd, 1, 15, 0&)
    Call SetProp(hWnd, "TIMERID", m_TimerID)
    pSetTimer = m_TimerID
End Function

Private Function pKillTimer(ByVal hWnd As Long) As Long
    Dim m_TimerID       As Long
    m_TimerID = GetProp(hWnd, "TIMERID")
    If m_TimerID = 0 Then Exit Function
    pKillTimer = KillTimer(hWnd, m_TimerID)
    Call SetProp(hWnd, "TIMERID", 0)
End Function

Private Function pGetBorderSize(ByVal hWnd As Long, cX As Long, cY As Long) As Long
    Dim wRect   As RECT
    Dim cRect   As RECT
    Dim lW      As Long
    Dim lH      As Long
    Dim sX      As Long
    Dim sY      As Long
    Dim lStyle  As Long
    Call GetWindowRect(hWnd, wRect)
    Call GetClientRect(hWnd, cRect)
    sX = GetSystemMetrics(SM_CXVSCROLL)
    sY = GetSystemMetrics(SM_CYHSCROLL)
    lW = (wRect.Right - wRect.Left) - (cRect.Right - cRect.Left)
    lH = (wRect.Bottom - wRect.Top) - (cRect.Bottom - cRect.Top)
    lStyle = GetWindowLong(hWnd, GWL_STYLE)
    If lStyle And WS_VSCROLL Then lW = lW - sX
    If lStyle And WS_VSCROLL Then lH = lH - sY
    cX = CLng(lW / 2)
    cY = CLng(lH / 2)
    pGetBorderSize = 1
End Function

Private Function IsFunctionSupported(sFunction As String, sModule As String) As Boolean
    Dim hModule As Long
        hModule = GetModuleHandleA(sModule)
        
    If (hModule = 0) Then
        hModule = LoadLibrary(sModule)
    End If
    
    If (hModule) Then
        If (GetProcAddress(hModule, sFunction)) Then
            IsFunctionSupported = True
        End If
        FreeLibrary hModule
    End If
End Function

Private Sub pTrackMouseTracking(hWnd As Long)
    Dim lpEventTrack As TRACKMOUSEEVENTTYPE
    With lpEventTrack
        .cbSize = Len(lpEventTrack)
        .dwFlags = &H2
        .hwndTrack = hWnd
    End With
    
    If (m_bTrackHandler32) Then
        TrackMouseEvent lpEventTrack
    Else
        TrackMouseEvent2 lpEventTrack
    End If
End Sub

Private Sub pDrawMemDC(ByVal hWnd As Long)
    Dim m_wRect       As RECTW
    Dim m_wText       As String
    Dim i           As Long
    Dim m_hDC(3)    As Long
    Call pGetWindowRectW(hWnd, m_wRect)
    m_wText = pGetWindowText(hWnd)
    For i = 0 To 3
        m_hDC(i) = GetProp(hWnd, "HDC" & CStr(i))
        SelectObject m_hDC(i), SendMessage(hWnd, WM_GETFONT, 0&, 0&)
        SetBkMode m_hDC(i), 1
        Call GridBlt(m_hDC(i), 0, 0, m_wRect.Width, m_wRect.Height, m_hBtnSrcDC, 0, i * 21, 7, 21, 3, 3, 3, 3)
        SetTextColor m_hDC(i), IIf(i = 3, &H968A79, 0&)
        pDrawTextL m_hDC(i), m_wText, 2, 2, m_wRect.Width - 4, m_wRect.Height - 4, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
        If GetFocus = hWnd Then pDrawFocusRect m_hDC(i), 2, 2, m_wRect.Width - 4, m_wRect.Height - 4
    Next
End Sub

Private Function pCreateDC(ByVal Width As Long, ByVal Height As Long) As Long
    Dim TmpDC   As Long
    Dim rDC     As Long
    Dim rBmp    As Long
    TmpDC = CreateDC("DISPLAY", "", "", ByVal 0&)
    If TmpDC Then
        rDC = CreateCompatibleDC(TmpDC)
        If rDC Then
            rBmp = CreateCompatibleBitmap(TmpDC, Width, Height)
            If rBmp Then
                DeleteObject SelectObject(rDC, rBmp)
                pCreateDC = rDC
                DeleteObject rBmp
            Else
                DeleteDC rDC
            End If
        End If
        DeleteDC TmpDC
    End If
End Function

Private Function pCreateDCFromHandle(Handle As Long) As Long
    If Handle = 0 Then Exit Function
    Dim TmpDC   As Long
    TmpDC = pCreateDC(1, 1)
    DeleteObject SelectObject(TmpDC, Handle)
    pCreateDCFromHandle = TmpDC
End Function

Private Function GridBlt(ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, _
                        ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcwidth As Long, ByVal srcheight As Long, _
                        Optional ByVal gX1 As Long = 0, Optional ByVal gY1 As Long = 0, Optional ByVal gX2 As Long = 0, Optional ByVal gY2 As Long = 0, _
                        Optional ByVal MaskColor As Variant) As Long
    If dstWidth = 0 Or dstHeight = 0 Or srcwidth = 0 Or srcheight = 0 Then Exit Function
    Dim TmpDC As Long
    TmpDC = pCreateDC(dstWidth, dstHeight)
    If gX1 <= 0 And gX2 <= 0 And gY1 <= 0 And gY2 <= 0 Then
        StretchBlt TmpDC, 0, 0, dstWidth, dstHeight, hSrcDC, srcX + gX1, srcY + gY1, srcwidth - gX2, srcheight - gY2, vbSrcCopy
    Else
        If gX1 > 0 And gY1 > 0 Then '左上角
            BitBlt TmpDC, 0, 0, gX1, gY1, hSrcDC, srcX, srcY, vbSrcCopy
        End If
        If gX2 > 0 And gY1 > 0 Then '右上角
            BitBlt TmpDC, dstWidth - gX2, 0, gX2, gY1, hSrcDC, srcX + srcwidth - gX2, srcY, vbSrcCopy
        End If
        If gX1 > 0 And gY2 > 0 Then '左下角
            BitBlt TmpDC, 0, dstHeight - gY2, gX1, gY2, hSrcDC, srcX, srcY + srcheight - gY2, vbSrcCopy
        End If
        If gX2 > 0 And gY2 > 0 Then '右下角
            BitBlt TmpDC, dstWidth - gX2, dstHeight - gY2, gX2, gY2, hSrcDC, srcX + srcwidth - gX2, srcY + srcheight - gY2, vbSrcCopy
        End If
        If gX1 > 0 Then '左边框
            StretchBlt TmpDC, 0, gY1, gX1, dstHeight - gY1 - gY2, hSrcDC, srcX, srcY + gY1, gX1, srcheight - gY1 - gY2, vbSrcCopy
        End If
        If gX2 > 0 Then '右边框
            StretchBlt TmpDC, dstWidth - gX2, gY1, gX2, dstHeight - gY1 - gY2, hSrcDC, srcX + srcwidth - gX2, srcY + gY1, gX2, srcheight - gY1 - gY2, vbSrcCopy
        End If
        If gY1 > 0 Then '上边框
            StretchBlt TmpDC, gX1, 0, dstWidth - gX1 - gX2, gY1, hSrcDC, srcX + gX1, srcY, srcwidth - gX1 - gX2, gY1, vbSrcCopy
        End If
        If gY2 > 0 Then '下边框
            StretchBlt TmpDC, gX1, dstHeight - gY2, dstWidth - gX1 - gX2, gY2, hSrcDC, srcX + gX1, srcY + srcheight - gY2, srcwidth - gX1 - gX2, gY2, vbSrcCopy
        End If
        '中间的伸展部分
        StretchBlt TmpDC, gX1, gY1, dstWidth - gX1 - gX2, dstHeight - gY1 - gY2, hSrcDC, srcX + gX1, srcY + gY1, srcwidth - gX1 - gX2, srcheight - gY1 - gY2, vbSrcCopy
    End If
    If IsMissing(MaskColor) Then
        GridBlt = BitBlt(hDestDC, dstX, dstY, dstWidth, dstHeight, TmpDC, 0, 0, vbSrcCopy)
    Else
        Call TransBlt(hDestDC, dstX, dstY, dstWidth, dstHeight, TmpDC, 0, 0, CLng(Val(MaskColor)))
        GridBlt = 1
    End If
    Call DeleteDC(TmpDC)
End Function

Private Sub TransBlt(ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, _
            ByVal hSrcDC As Long, Optional ByVal srcX As Long = 0, Optional ByVal srcY As Long = 0, Optional ByVal MaskColor As OLE_COLOR = vbMagenta, _
            Optional ByVal hpal As Long = 0)
    Dim hdcMask             As Long
    Dim hdcColor            As Long
    Dim hbmMask             As Long
    Dim hbmColor            As Long
    Dim hbmColorOld         As Long
    Dim hbmMaskOld          As Long
    Dim hpalOld             As Long
    Dim hTmpDC              As Long
    Dim hdcScnBuffer        As Long
    Dim hbmScnBuffer        As Long
    Dim hbmScnBufferOld     As Long
    Dim hPalBufferOld       As Long
    Dim lMaskColor          As Long
    Dim hpalHalftone        As Long

    hTmpDC = CreateDC("DISPLAY", "", "", ByVal 0&)
    If hpal = 0 Then
        hpalHalftone = CreateHalftonePalette(hTmpDC)
        hpal = hpalHalftone
    End If
    OleTranslateColor MaskColor, hpal, lMaskColor
    lMaskColor = lMaskColor And &HFFFFFF
    hbmScnBuffer = CreateCompatibleBitmap(hTmpDC, dstWidth, dstHeight)
    hdcScnBuffer = CreateCompatibleDC(hTmpDC)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hpal, True)
    RealizePalette hdcScnBuffer
    BitBlt hdcScnBuffer, 0, 0, dstWidth, dstHeight, hDestDC, dstX, dstY, vbSrcCopy
    hbmColor = CreateCompatibleBitmap(hTmpDC, dstWidth, dstHeight)
    hbmMask = CreateBitmap(dstWidth, dstHeight, 1, 1, ByVal 0&)
    hdcColor = CreateCompatibleDC(hTmpDC)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hpalOld = SelectPalette(hdcColor, hpal, True)
    RealizePalette hdcColor
    Call SetBkColor(hdcColor, GetBkColor(hSrcDC))
    Call SetTextColor(hdcColor, GetTextColor(hSrcDC))
    Call BitBlt(hdcColor, 0, 0, dstWidth, dstHeight, hSrcDC, srcX, srcY, vbSrcCopy)
    hdcMask = CreateCompatibleDC(hTmpDC)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)
    Call SetBkColor(hdcColor, lMaskColor)
    Call SetTextColor(hdcColor, vbWhite)
    Call BitBlt(hdcMask, 0, 0, dstWidth, dstHeight, hdcColor, 0, 0, vbSrcCopy)
    Call SetTextColor(hdcColor, vbBlack)
    Call SetBkColor(hdcColor, vbWhite)
    Call BitBlt(hdcColor, 0, 0, dstWidth, dstHeight, hdcMask, 0, 0, &H220326)
    Call BitBlt(hdcScnBuffer, 0, 0, dstWidth, dstHeight, hdcMask, 0, 0, vbSrcAnd)
    Call BitBlt(hdcScnBuffer, 0, 0, dstWidth, dstHeight, hdcColor, 0, 0, vbSrcPaint)
    Call BitBlt(hDestDC, dstX, dstY, dstWidth, dstHeight, hdcScnBuffer, 0, 0, vbSrcCopy)
    Call DeleteObject(SelectObject(hdcColor, hbmColorOld))
    Call SelectPalette(hdcColor, hpalOld, True)
    Call RealizePalette(hdcColor)
    Call DeleteDC(hdcColor)
    Call DeleteObject(SelectObject(hdcScnBuffer, hbmScnBufferOld))
    Call SelectPalette(hdcScnBuffer, hPalBufferOld, 0)
    Call RealizePalette(hdcScnBuffer)
    Call DeleteDC(hdcScnBuffer)
    Call DeleteObject(SelectObject(hdcMask, hbmMaskOld))
    Call DeleteDC(hdcMask)
    If hpalHalftone <> 0 Then
        Call DeleteObject(hpalHalftone)
    End If
End Sub

Private Function LWORD(Param As Long) As Long
    Dim intRet As Integer
    CopyMemory intRet, Param, 2
    LWORD = intRet
End Function

Private Function HWORD(Param As Long) As Long
    Dim intRet As Integer
    CopyMemory intRet, ByVal VarPtr(Param) + 2, 2
    HWORD = intRet
End Function

Private Sub pInit()
    m_hBtnSrcDC = pCreateDC(7, 84)
    m_hCbbSrcDC = pCreateDC(15, 18)
    m_hCkbSrcDC = pCreateDC(18, 24)
    m_hOpbSrcDC = pCreateDC(13, 52)
    m_hHdbSrcDC = pCreateDC(9, 20)
    
    SetPixel m_hBtnSrcDC, 0, 0, 16376255
    SetPixel m_hBtnSrcDC, 1, 0, 14922603
    SetPixel m_hBtnSrcDC, 2, 0, 14194476
    SetPixel m_hBtnSrcDC, 3, 0, 13995803
    SetPixel m_hBtnSrcDC, 4, 0, 14194476
    SetPixel m_hBtnSrcDC, 5, 0, 14922603
    SetPixel m_hBtnSrcDC, 6, 0, 16376255
    SetPixel m_hBtnSrcDC, 0, 1, 14856809
    SetPixel m_hBtnSrcDC, 1, 1, 15188094
    SetPixel m_hBtnSrcDC, 2, 1, 16511975
    SetPixel m_hBtnSrcDC, 3, 1, 16777215
    SetPixel m_hBtnSrcDC, 4, 1, 16511975
    SetPixel m_hBtnSrcDC, 5, 1, 15188094
    SetPixel m_hBtnSrcDC, 6, 1, 14856809
    SetPixel m_hBtnSrcDC, 0, 2, 14128425
    SetPixel m_hBtnSrcDC, 1, 2, 16578284
    SetPixel m_hBtnSrcDC, 2, 2, 16777215
    SetPixel m_hBtnSrcDC, 3, 2, 16777215
    SetPixel m_hBtnSrcDC, 4, 2, 16777215
    SetPixel m_hBtnSrcDC, 5, 2, 16578284
    SetPixel m_hBtnSrcDC, 6, 2, 14128425
    SetPixel m_hBtnSrcDC, 0, 3, 13995803
    SetPixel m_hBtnSrcDC, 1, 3, 16644853
    SetPixel m_hBtnSrcDC, 2, 3, 16578801
    SetPixel m_hBtnSrcDC, 3, 3, 16578801
    SetPixel m_hBtnSrcDC, 4, 3, 16578801
    SetPixel m_hBtnSrcDC, 5, 3, 16644853
    SetPixel m_hBtnSrcDC, 6, 3, 13995803
    SetPixel m_hBtnSrcDC, 0, 4, 13995803
    SetPixel m_hBtnSrcDC, 1, 4, 16579059
    SetPixel m_hBtnSrcDC, 2, 4, 16512750
    SetPixel m_hBtnSrcDC, 3, 4, 16512750
    SetPixel m_hBtnSrcDC, 4, 4, 16512750
    SetPixel m_hBtnSrcDC, 5, 4, 16579059
    SetPixel m_hBtnSrcDC, 6, 4, 13995803
    SetPixel m_hBtnSrcDC, 0, 5, 13995803
    SetPixel m_hBtnSrcDC, 1, 5, 16578544
    SetPixel m_hBtnSrcDC, 2, 5, 16512234
    SetPixel m_hBtnSrcDC, 3, 5, 16512234
    SetPixel m_hBtnSrcDC, 4, 5, 16512234
    SetPixel m_hBtnSrcDC, 5, 5, 16578544
    SetPixel m_hBtnSrcDC, 6, 5, 13995803
    SetPixel m_hBtnSrcDC, 0, 6, 13995803
    SetPixel m_hBtnSrcDC, 1, 6, 16578286
    SetPixel m_hBtnSrcDC, 2, 6, 16511718
    SetPixel m_hBtnSrcDC, 3, 6, 16511718
    SetPixel m_hBtnSrcDC, 4, 6, 16511718
    SetPixel m_hBtnSrcDC, 5, 6, 16578286
    SetPixel m_hBtnSrcDC, 6, 6, 13995803
    SetPixel m_hBtnSrcDC, 0, 7, 13995803
    SetPixel m_hBtnSrcDC, 1, 7, 16578027
    SetPixel m_hBtnSrcDC, 2, 7, 16445666
    SetPixel m_hBtnSrcDC, 3, 7, 16445666
    SetPixel m_hBtnSrcDC, 4, 7, 16445666
    SetPixel m_hBtnSrcDC, 5, 7, 16578027
    SetPixel m_hBtnSrcDC, 6, 7, 13995803
    SetPixel m_hBtnSrcDC, 0, 8, 13995803
    SetPixel m_hBtnSrcDC, 1, 8, 16577512
    SetPixel m_hBtnSrcDC, 2, 8, 16445150
    SetPixel m_hBtnSrcDC, 3, 8, 16445150
    SetPixel m_hBtnSrcDC, 4, 8, 16445150
    SetPixel m_hBtnSrcDC, 5, 8, 16577512
    SetPixel m_hBtnSrcDC, 6, 8, 13995803
    SetPixel m_hBtnSrcDC, 0, 9, 13995803
    SetPixel m_hBtnSrcDC, 1, 9, 16511717
    SetPixel m_hBtnSrcDC, 2, 9, 16379098
    SetPixel m_hBtnSrcDC, 3, 9, 16379098
    SetPixel m_hBtnSrcDC, 4, 9, 16379098
    SetPixel m_hBtnSrcDC, 5, 9, 16511717
    SetPixel m_hBtnSrcDC, 6, 9, 13995803
    SetPixel m_hBtnSrcDC, 0, 10, 13995803
    SetPixel m_hBtnSrcDC, 1, 10, 16511203
    SetPixel m_hBtnSrcDC, 2, 10, 16378583
    SetPixel m_hBtnSrcDC, 3, 10, 16378583
    SetPixel m_hBtnSrcDC, 4, 10, 16378583
    SetPixel m_hBtnSrcDC, 5, 10, 16511203
    SetPixel m_hBtnSrcDC, 6, 10, 13995803
    SetPixel m_hBtnSrcDC, 0, 11, 13995803
    SetPixel m_hBtnSrcDC, 1, 11, 16313307
    SetPixel m_hBtnSrcDC, 2, 11, 16114380
    SetPixel m_hBtnSrcDC, 3, 11, 16114380
    SetPixel m_hBtnSrcDC, 4, 11, 16114380
    SetPixel m_hBtnSrcDC, 5, 11, 16313307
    SetPixel m_hBtnSrcDC, 6, 11, 13995803
    SetPixel m_hBtnSrcDC, 0, 12, 13995803
    SetPixel m_hBtnSrcDC, 1, 12, 16247257
    SetPixel m_hBtnSrcDC, 2, 12, 15982536
    SetPixel m_hBtnSrcDC, 3, 12, 15982536
    SetPixel m_hBtnSrcDC, 4, 12, 15982536
    SetPixel m_hBtnSrcDC, 5, 12, 16247257
    SetPixel m_hBtnSrcDC, 6, 12, 13995803
    SetPixel m_hBtnSrcDC, 0, 13, 13995803
    SetPixel m_hBtnSrcDC, 1, 13, 16115669
    SetPixel m_hBtnSrcDC, 2, 13, 15850691
    SetPixel m_hBtnSrcDC, 3, 13, 15850691
    SetPixel m_hBtnSrcDC, 4, 13, 15850691
    SetPixel m_hBtnSrcDC, 5, 13, 16115669
    SetPixel m_hBtnSrcDC, 6, 13, 13995803
    SetPixel m_hBtnSrcDC, 0, 14, 13995803
    SetPixel m_hBtnSrcDC, 1, 14, 16049362
    SetPixel m_hBtnSrcDC, 2, 14, 15718590
    SetPixel m_hBtnSrcDC, 3, 14, 15718590
    SetPixel m_hBtnSrcDC, 4, 14, 15718590
    SetPixel m_hBtnSrcDC, 5, 14, 16049362
    SetPixel m_hBtnSrcDC, 6, 14, 13995803
    SetPixel m_hBtnSrcDC, 0, 15, 13995803
    SetPixel m_hBtnSrcDC, 1, 15, 15917773
    SetPixel m_hBtnSrcDC, 2, 15, 15586744
    SetPixel m_hBtnSrcDC, 3, 15, 15586744
    SetPixel m_hBtnSrcDC, 4, 15, 15586744
    SetPixel m_hBtnSrcDC, 5, 15, 15917773
    SetPixel m_hBtnSrcDC, 6, 15, 13995803
    SetPixel m_hBtnSrcDC, 0, 16, 13995803
    SetPixel m_hBtnSrcDC, 1, 16, 15851723
    SetPixel m_hBtnSrcDC, 2, 16, 15454900
    SetPixel m_hBtnSrcDC, 3, 16, 15454900
    SetPixel m_hBtnSrcDC, 4, 16, 15454900
    SetPixel m_hBtnSrcDC, 5, 16, 15851723
    SetPixel m_hBtnSrcDC, 6, 16, 13995803
    SetPixel m_hBtnSrcDC, 0, 17, 13995803
    SetPixel m_hBtnSrcDC, 1, 17, 15785673
    SetPixel m_hBtnSrcDC, 2, 17, 15388849
    SetPixel m_hBtnSrcDC, 3, 17, 15388849
    SetPixel m_hBtnSrcDC, 4, 17, 15388849
    SetPixel m_hBtnSrcDC, 5, 17, 15785673
    SetPixel m_hBtnSrcDC, 6, 17, 13995803
    SetPixel m_hBtnSrcDC, 0, 18, 14128424
    SetPixel m_hBtnSrcDC, 1, 18, 15652794
    SetPixel m_hBtnSrcDC, 2, 18, 15521467
    SetPixel m_hBtnSrcDC, 3, 18, 15323056
    SetPixel m_hBtnSrcDC, 4, 18, 15521211
    SetPixel m_hBtnSrcDC, 5, 18, 15652794
    SetPixel m_hBtnSrcDC, 6, 18, 14128424
    SetPixel m_hBtnSrcDC, 0, 19, 14856292
    SetPixel m_hBtnSrcDC, 1, 19, 14791272
    SetPixel m_hBtnSrcDC, 2, 19, 15653051
    SetPixel m_hBtnSrcDC, 3, 19, 15851724
    SetPixel m_hBtnSrcDC, 4, 19, 15653051
    SetPixel m_hBtnSrcDC, 5, 19, 14791272
    SetPixel m_hBtnSrcDC, 6, 19, 14790498
    SetPixel m_hBtnSrcDC, 0, 20, 16508874
    SetPixel m_hBtnSrcDC, 1, 20, 14790240
    SetPixel m_hBtnSrcDC, 2, 20, 14194218
    SetPixel m_hBtnSrcDC, 3, 20, 13995803
    SetPixel m_hBtnSrcDC, 4, 20, 14194218
    SetPixel m_hBtnSrcDC, 5, 20, 14790240
    SetPixel m_hBtnSrcDC, 6, 20, 16508874
    SetPixel m_hBtnSrcDC, 0, 21, 16376255
    SetPixel m_hBtnSrcDC, 1, 21, 14922603
    SetPixel m_hBtnSrcDC, 2, 21, 14194476
    SetPixel m_hBtnSrcDC, 3, 21, 13995803
    SetPixel m_hBtnSrcDC, 4, 21, 14194476
    SetPixel m_hBtnSrcDC, 5, 21, 14922603
    SetPixel m_hBtnSrcDC, 6, 21, 16376255
    SetPixel m_hBtnSrcDC, 0, 22, 14856809
    SetPixel m_hBtnSrcDC, 1, 22, 15115048
    SetPixel m_hBtnSrcDC, 2, 22, 16300860
    SetPixel m_hBtnSrcDC, 3, 22, 16499532
    SetPixel m_hBtnSrcDC, 4, 22, 16300860
    SetPixel m_hBtnSrcDC, 5, 22, 15115048
    SetPixel m_hBtnSrcDC, 6, 22, 14856809
    SetPixel m_hBtnSrcDC, 0, 23, 14128425
    SetPixel m_hBtnSrcDC, 1, 23, 16366653
    SetPixel m_hBtnSrcDC, 2, 23, 16702093
    SetPixel m_hBtnSrcDC, 3, 23, 16772550
    SetPixel m_hBtnSrcDC, 4, 23, 16702093
    SetPixel m_hBtnSrcDC, 5, 23, 16366653
    SetPixel m_hBtnSrcDC, 6, 23, 14128425
    SetPixel m_hBtnSrcDC, 0, 24, 13995803
    SetPixel m_hBtnSrcDC, 1, 24, 16499014
    SetPixel m_hBtnSrcDC, 2, 24, 16638894
    SetPixel m_hBtnSrcDC, 3, 24, 16576210
    SetPixel m_hBtnSrcDC, 4, 24, 16704688
    SetPixel m_hBtnSrcDC, 5, 24, 16499014
    SetPixel m_hBtnSrcDC, 6, 24, 13995803
    SetPixel m_hBtnSrcDC, 0, 25, 13995803
    SetPixel m_hBtnSrcDC, 1, 25, 16499274
    SetPixel m_hBtnSrcDC, 2, 25, 16639930
    SetPixel m_hBtnSrcDC, 3, 25, 16577247
    SetPixel m_hBtnSrcDC, 4, 25, 16639931
    SetPixel m_hBtnSrcDC, 5, 25, 16499274
    SetPixel m_hBtnSrcDC, 6, 25, 13995803
    SetPixel m_hBtnSrcDC, 0, 26, 13995803
    SetPixel m_hBtnSrcDC, 1, 26, 16499532
    SetPixel m_hBtnSrcDC, 2, 26, 16640192
    SetPixel m_hBtnSrcDC, 3, 26, 16577252
    SetPixel m_hBtnSrcDC, 4, 26, 16640192
    SetPixel m_hBtnSrcDC, 5, 26, 16499532
    SetPixel m_hBtnSrcDC, 6, 26, 13995803
    SetPixel m_hBtnSrcDC, 0, 27, 13995803
    SetPixel m_hBtnSrcDC, 1, 27, 16499532
    SetPixel m_hBtnSrcDC, 2, 27, 16639936
    SetPixel m_hBtnSrcDC, 3, 27, 16577250
    SetPixel m_hBtnSrcDC, 4, 27, 16639936
    SetPixel m_hBtnSrcDC, 5, 27, 16499532
    SetPixel m_hBtnSrcDC, 6, 27, 13995803
    SetPixel m_hBtnSrcDC, 0, 28, 13995803
    SetPixel m_hBtnSrcDC, 1, 28, 16499532
    SetPixel m_hBtnSrcDC, 2, 28, 16574399
    SetPixel m_hBtnSrcDC, 3, 28, 16511456
    SetPixel m_hBtnSrcDC, 4, 28, 16574399
    SetPixel m_hBtnSrcDC, 5, 28, 16499532
    SetPixel m_hBtnSrcDC, 6, 28, 13995803
    SetPixel m_hBtnSrcDC, 0, 29, 13995803
    SetPixel m_hBtnSrcDC, 1, 29, 16499275
    SetPixel m_hBtnSrcDC, 2, 29, 16574140
    SetPixel m_hBtnSrcDC, 3, 29, 16511197
    SetPixel m_hBtnSrcDC, 4, 29, 16574140
    SetPixel m_hBtnSrcDC, 5, 29, 16499275
    SetPixel m_hBtnSrcDC, 6, 29, 13995803
    SetPixel m_hBtnSrcDC, 0, 30, 13995803
    SetPixel m_hBtnSrcDC, 1, 30, 16499274
    SetPixel m_hBtnSrcDC, 2, 30, 16573882
    SetPixel m_hBtnSrcDC, 3, 30, 16510683
    SetPixel m_hBtnSrcDC, 4, 30, 16573882
    SetPixel m_hBtnSrcDC, 5, 30, 16499274
    SetPixel m_hBtnSrcDC, 6, 30, 13995803
    SetPixel m_hBtnSrcDC, 0, 31, 13995803
    SetPixel m_hBtnSrcDC, 1, 31, 16499274
    SetPixel m_hBtnSrcDC, 2, 31, 16573625
    SetPixel m_hBtnSrcDC, 3, 31, 16510425
    SetPixel m_hBtnSrcDC, 4, 31, 16573625
    SetPixel m_hBtnSrcDC, 5, 31, 16499274
    SetPixel m_hBtnSrcDC, 6, 31, 13995803
    SetPixel m_hBtnSrcDC, 0, 32, 13995803
    SetPixel m_hBtnSrcDC, 1, 32, 16432966
    SetPixel m_hBtnSrcDC, 2, 32, 16375214
    SetPixel m_hBtnSrcDC, 3, 32, 16245963
    SetPixel m_hBtnSrcDC, 4, 32, 16375214
    SetPixel m_hBtnSrcDC, 5, 32, 16432966
    SetPixel m_hBtnSrcDC, 6, 32, 13995803
    SetPixel m_hBtnSrcDC, 0, 33, 13995803
    SetPixel m_hBtnSrcDC, 1, 33, 16432965
    SetPixel m_hBtnSrcDC, 2, 33, 16243370
    SetPixel m_hBtnSrcDC, 3, 33, 16114119
    SetPixel m_hBtnSrcDC, 4, 33, 16243370
    SetPixel m_hBtnSrcDC, 5, 33, 16432965
    SetPixel m_hBtnSrcDC, 6, 33, 13995803
    SetPixel m_hBtnSrcDC, 0, 34, 13995803
    SetPixel m_hBtnSrcDC, 1, 34, 16367172
    SetPixel m_hBtnSrcDC, 2, 34, 16177319
    SetPixel m_hBtnSrcDC, 3, 34, 15982274
    SetPixel m_hBtnSrcDC, 4, 34, 16177319
    SetPixel m_hBtnSrcDC, 5, 34, 16367172
    SetPixel m_hBtnSrcDC, 6, 34, 13995803
    SetPixel m_hBtnSrcDC, 0, 35, 13995803
    SetPixel m_hBtnSrcDC, 1, 35, 16301378
    SetPixel m_hBtnSrcDC, 2, 35, 16045730
    SetPixel m_hBtnSrcDC, 3, 35, 15915964
    SetPixel m_hBtnSrcDC, 4, 35, 16045730
    SetPixel m_hBtnSrcDC, 5, 35, 16301378
    SetPixel m_hBtnSrcDC, 6, 35, 13995803
    SetPixel m_hBtnSrcDC, 0, 36, 13995803
    SetPixel m_hBtnSrcDC, 1, 36, 16301121
    SetPixel m_hBtnSrcDC, 2, 36, 15979165
    SetPixel m_hBtnSrcDC, 3, 36, 15849399
    SetPixel m_hBtnSrcDC, 4, 36, 15979165
    SetPixel m_hBtnSrcDC, 5, 36, 16301121
    SetPixel m_hBtnSrcDC, 6, 36, 13995803
    SetPixel m_hBtnSrcDC, 0, 37, 13995803
    SetPixel m_hBtnSrcDC, 1, 37, 16300861
    SetPixel m_hBtnSrcDC, 2, 37, 15978646
    SetPixel m_hBtnSrcDC, 3, 37, 15717807
    SetPixel m_hBtnSrcDC, 4, 37, 15978644
    SetPixel m_hBtnSrcDC, 5, 37, 16300862
    SetPixel m_hBtnSrcDC, 6, 37, 13995803
    SetPixel m_hBtnSrcDC, 0, 38, 13995803
    SetPixel m_hBtnSrcDC, 1, 38, 16300601
    SetPixel m_hBtnSrcDC, 2, 38, 16109711
    SetPixel m_hBtnSrcDC, 3, 38, 15914151
    SetPixel m_hBtnSrcDC, 4, 38, 15977608
    SetPixel m_hBtnSrcDC, 5, 38, 16300601
    SetPixel m_hBtnSrcDC, 6, 38, 13995803
    SetPixel m_hBtnSrcDC, 0, 39, 14128424
    SetPixel m_hBtnSrcDC, 1, 39, 16168499
    SetPixel m_hBtnSrcDC, 2, 39, 16238960
    SetPixel m_hBtnSrcDC, 3, 39, 15979161
    SetPixel m_hBtnSrcDC, 4, 39, 16239217
    SetPixel m_hBtnSrcDC, 5, 39, 16168501
    SetPixel m_hBtnSrcDC, 6, 39, 14128424
    SetPixel m_hBtnSrcDC, 0, 40, 14856292
    SetPixel m_hBtnSrcDC, 1, 40, 15048997
    SetPixel m_hBtnSrcDC, 2, 40, 16168243
    SetPixel m_hBtnSrcDC, 3, 40, 16300863
    SetPixel m_hBtnSrcDC, 4, 40, 16168499
    SetPixel m_hBtnSrcDC, 5, 40, 15048996
    SetPixel m_hBtnSrcDC, 6, 40, 14790498
    SetPixel m_hBtnSrcDC, 0, 41, 16508874
    SetPixel m_hBtnSrcDC, 1, 41, 14790240
    SetPixel m_hBtnSrcDC, 2, 41, 14194218
    SetPixel m_hBtnSrcDC, 3, 41, 13995803
    SetPixel m_hBtnSrcDC, 4, 41, 14194218
    SetPixel m_hBtnSrcDC, 5, 41, 14790240
    SetPixel m_hBtnSrcDC, 6, 41, 16508874
    SetPixel m_hBtnSrcDC, 0, 42, 16376255
    SetPixel m_hBtnSrcDC, 1, 42, 14922603
    SetPixel m_hBtnSrcDC, 2, 42, 14194476
    SetPixel m_hBtnSrcDC, 3, 42, 13995803
    SetPixel m_hBtnSrcDC, 4, 42, 14194476
    SetPixel m_hBtnSrcDC, 5, 42, 14922603
    SetPixel m_hBtnSrcDC, 6, 42, 16376255
    SetPixel m_hBtnSrcDC, 0, 43, 14856809
    SetPixel m_hBtnSrcDC, 1, 43, 15114530
    SetPixel m_hBtnSrcDC, 2, 43, 16299312
    SetPixel m_hBtnSrcDC, 3, 43, 16497721
    SetPixel m_hBtnSrcDC, 4, 43, 16233775
    SetPixel m_hBtnSrcDC, 5, 43, 15114531
    SetPixel m_hBtnSrcDC, 6, 43, 14856809
    SetPixel m_hBtnSrcDC, 0, 44, 14128425
    SetPixel m_hBtnSrcDC, 1, 44, 16299566
    SetPixel m_hBtnSrcDC, 2, 44, 16304495
    SetPixel m_hBtnSrcDC, 3, 44, 15977606
    SetPixel m_hBtnSrcDC, 4, 44, 16304237
    SetPixel m_hBtnSrcDC, 5, 44, 16233772
    SetPixel m_hBtnSrcDC, 6, 44, 14128425
    SetPixel m_hBtnSrcDC, 0, 45, 13995803
    SetPixel m_hBtnSrcDC, 1, 45, 16365361
    SetPixel m_hBtnSrcDC, 2, 45, 15911555
    SetPixel m_hBtnSrcDC, 3, 45, 15979679
    SetPixel m_hBtnSrcDC, 4, 45, 16043658
    SetPixel m_hBtnSrcDC, 5, 45, 16365361
    SetPixel m_hBtnSrcDC, 6, 45, 13995803
    SetPixel m_hBtnSrcDC, 0, 46, 13995803
    SetPixel m_hBtnSrcDC, 1, 46, 16365621
    SetPixel m_hBtnSrcDC, 2, 46, 15912077
    SetPixel m_hBtnSrcDC, 3, 46, 15980459
    SetPixel m_hBtnSrcDC, 4, 46, 15912335
    SetPixel m_hBtnSrcDC, 5, 46, 16365621
    SetPixel m_hBtnSrcDC, 6, 46, 13995803
    SetPixel m_hBtnSrcDC, 0, 47, 13995803
    SetPixel m_hBtnSrcDC, 1, 47, 16365879
    SetPixel m_hBtnSrcDC, 2, 47, 15847061
    SetPixel m_hBtnSrcDC, 3, 47, 15980979
    SetPixel m_hBtnSrcDC, 4, 47, 15912597
    SetPixel m_hBtnSrcDC, 5, 47, 16365879
    SetPixel m_hBtnSrcDC, 6, 47, 13995803
    SetPixel m_hBtnSrcDC, 0, 48, 13995803
    SetPixel m_hBtnSrcDC, 1, 48, 16365882
    SetPixel m_hBtnSrcDC, 2, 48, 15913114
    SetPixel m_hBtnSrcDC, 3, 48, 15981495
    SetPixel m_hBtnSrcDC, 4, 48, 15913114
    SetPixel m_hBtnSrcDC, 5, 48, 16365882
    SetPixel m_hBtnSrcDC, 6, 48, 13995803
    SetPixel m_hBtnSrcDC, 0, 49, 13995803
    SetPixel m_hBtnSrcDC, 1, 49, 16366139
    SetPixel m_hBtnSrcDC, 2, 49, 15979423
    SetPixel m_hBtnSrcDC, 3, 49, 16113084
    SetPixel m_hBtnSrcDC, 4, 49, 15979423
    SetPixel m_hBtnSrcDC, 5, 49, 16366139
    SetPixel m_hBtnSrcDC, 6, 49, 13995803
    SetPixel m_hBtnSrcDC, 0, 50, 13995803
    SetPixel m_hBtnSrcDC, 1, 50, 16431932
    SetPixel m_hBtnSrcDC, 2, 50, 16111267
    SetPixel m_hBtnSrcDC, 3, 50, 16179647
    SetPixel m_hBtnSrcDC, 4, 50, 16111267
    SetPixel m_hBtnSrcDC, 5, 50, 16431932
    SetPixel m_hBtnSrcDC, 6, 50, 13995803
    SetPixel m_hBtnSrcDC, 0, 51, 13995803
    SetPixel m_hBtnSrcDC, 1, 51, 16432189
    SetPixel m_hBtnSrcDC, 2, 51, 16177319
    SetPixel m_hBtnSrcDC, 3, 51, 16310979
    SetPixel m_hBtnSrcDC, 4, 51, 16177319
    SetPixel m_hBtnSrcDC, 5, 51, 16432189
    SetPixel m_hBtnSrcDC, 6, 51, 13995803
    SetPixel m_hBtnSrcDC, 0, 52, 13995803
    SetPixel m_hBtnSrcDC, 1, 52, 16497983
    SetPixel m_hBtnSrcDC, 2, 52, 16308906
    SetPixel m_hBtnSrcDC, 3, 52, 16377029
    SetPixel m_hBtnSrcDC, 4, 52, 16308906
    SetPixel m_hBtnSrcDC, 5, 52, 16497983
    SetPixel m_hBtnSrcDC, 6, 52, 13995803
    SetPixel m_hBtnSrcDC, 0, 53, 13995803
    SetPixel m_hBtnSrcDC, 1, 53, 16564034
    SetPixel m_hBtnSrcDC, 2, 53, 16507573
    SetPixel m_hBtnSrcDC, 3, 53, 16509903
    SetPixel m_hBtnSrcDC, 4, 53, 16507573
    SetPixel m_hBtnSrcDC, 5, 53, 16564034
    SetPixel m_hBtnSrcDC, 6, 53, 13995803
    SetPixel m_hBtnSrcDC, 0, 54, 13995803
    SetPixel m_hBtnSrcDC, 1, 54, 16564291
    SetPixel m_hBtnSrcDC, 2, 54, 16507831
    SetPixel m_hBtnSrcDC, 3, 54, 16510161
    SetPixel m_hBtnSrcDC, 4, 54, 16507831
    SetPixel m_hBtnSrcDC, 5, 54, 16564291
    SetPixel m_hBtnSrcDC, 6, 54, 13995803
    SetPixel m_hBtnSrcDC, 0, 55, 13995803
    SetPixel m_hBtnSrcDC, 1, 55, 16564292
    SetPixel m_hBtnSrcDC, 2, 55, 16573625
    SetPixel m_hBtnSrcDC, 3, 55, 16575956
    SetPixel m_hBtnSrcDC, 4, 55, 16573625
    SetPixel m_hBtnSrcDC, 5, 55, 16564292
    SetPixel m_hBtnSrcDC, 6, 55, 13995803
    SetPixel m_hBtnSrcDC, 0, 56, 13995803
    SetPixel m_hBtnSrcDC, 1, 56, 16564549
    SetPixel m_hBtnSrcDC, 2, 56, 16573884
    SetPixel m_hBtnSrcDC, 3, 56, 16576214
    SetPixel m_hBtnSrcDC, 4, 56, 16573884
    SetPixel m_hBtnSrcDC, 5, 56, 16564549
    SetPixel m_hBtnSrcDC, 6, 56, 13995803
    SetPixel m_hBtnSrcDC, 0, 57, 13995803
    SetPixel m_hBtnSrcDC, 1, 57, 16564548
    SetPixel m_hBtnSrcDC, 2, 57, 16574141
    SetPixel m_hBtnSrcDC, 3, 57, 16576470
    SetPixel m_hBtnSrcDC, 4, 57, 16574141
    SetPixel m_hBtnSrcDC, 5, 57, 16564548
    SetPixel m_hBtnSrcDC, 6, 57, 13995803
    SetPixel m_hBtnSrcDC, 0, 58, 13995803
    SetPixel m_hBtnSrcDC, 1, 58, 16564290
    SetPixel m_hBtnSrcDC, 2, 58, 16574136
    SetPixel m_hBtnSrcDC, 3, 58, 16576209
    SetPixel m_hBtnSrcDC, 4, 58, 16574136
    SetPixel m_hBtnSrcDC, 5, 58, 16564290
    SetPixel m_hBtnSrcDC, 6, 58, 13995803
    SetPixel m_hBtnSrcDC, 0, 59, 13995803
    SetPixel m_hBtnSrcDC, 1, 59, 16564030
    SetPixel m_hBtnSrcDC, 2, 59, 16704686
    SetPixel m_hBtnSrcDC, 3, 59, 16640707
    SetPixel m_hBtnSrcDC, 4, 59, 16638636
    SetPixel m_hBtnSrcDC, 5, 59, 16564030
    SetPixel m_hBtnSrcDC, 6, 59, 13995803
    SetPixel m_hBtnSrcDC, 0, 60, 14128424
    SetPixel m_hBtnSrcDC, 1, 60, 16431414
    SetPixel m_hBtnSrcDC, 2, 60, 16701318
    SetPixel m_hBtnSrcDC, 3, 60, 16639150
    SetPixel m_hBtnSrcDC, 4, 60, 16701320
    SetPixel m_hBtnSrcDC, 5, 60, 16431670
    SetPixel m_hBtnSrcDC, 6, 60, 14128424
    SetPixel m_hBtnSrcDC, 0, 61, 14856292
    SetPixel m_hBtnSrcDC, 1, 61, 15115045
    SetPixel m_hBtnSrcDC, 2, 61, 16366135
    SetPixel m_hBtnSrcDC, 3, 61, 16630080
    SetPixel m_hBtnSrcDC, 4, 61, 16366135
    SetPixel m_hBtnSrcDC, 5, 61, 15115045
    SetPixel m_hBtnSrcDC, 6, 61, 14790498
    SetPixel m_hBtnSrcDC, 0, 62, 16508874
    SetPixel m_hBtnSrcDC, 1, 62, 14790240
    SetPixel m_hBtnSrcDC, 2, 62, 14194218
    SetPixel m_hBtnSrcDC, 3, 62, 13995803
    SetPixel m_hBtnSrcDC, 4, 62, 14194218
    SetPixel m_hBtnSrcDC, 5, 62, 14790240
    SetPixel m_hBtnSrcDC, 6, 62, 16508874
    SetPixel m_hBtnSrcDC, 0, 63, 16513010
    SetPixel m_hBtnSrcDC, 1, 63, 15717807
    SetPixel m_hBtnSrcDC, 2, 63, 15386511
    SetPixel m_hBtnSrcDC, 3, 63, 15254407
    SetPixel m_hBtnSrcDC, 4, 63, 15386511
    SetPixel m_hBtnSrcDC, 5, 63, 15717807
    SetPixel m_hBtnSrcDC, 6, 63, 16513010
    SetPixel m_hBtnSrcDC, 0, 64, 15717806
    SetPixel m_hBtnSrcDC, 1, 64, 15850680
    SetPixel m_hBtnSrcDC, 2, 64, 16512493
    SetPixel m_hBtnSrcDC, 3, 64, 16645112
    SetPixel m_hBtnSrcDC, 4, 64, 16512493
    SetPixel m_hBtnSrcDC, 5, 64, 15850680
    SetPixel m_hBtnSrcDC, 6, 64, 15717806
    SetPixel m_hBtnSrcDC, 0, 65, 15320718
    SetPixel m_hBtnSrcDC, 1, 65, 16513007
    SetPixel m_hBtnSrcDC, 2, 65, 16645112
    SetPixel m_hBtnSrcDC, 3, 65, 16645112
    SetPixel m_hBtnSrcDC, 4, 65, 16645112
    SetPixel m_hBtnSrcDC, 5, 65, 16513007
    SetPixel m_hBtnSrcDC, 6, 65, 15320718
    SetPixel m_hBtnSrcDC, 0, 66, 15254407
    SetPixel m_hBtnSrcDC, 1, 66, 16578803
    SetPixel m_hBtnSrcDC, 2, 66, 16513010
    SetPixel m_hBtnSrcDC, 3, 66, 16513010
    SetPixel m_hBtnSrcDC, 4, 66, 16513010
    SetPixel m_hBtnSrcDC, 5, 66, 16578803
    SetPixel m_hBtnSrcDC, 6, 66, 15254407
    SetPixel m_hBtnSrcDC, 0, 67, 15254407
    SetPixel m_hBtnSrcDC, 1, 67, 16513266
    SetPixel m_hBtnSrcDC, 2, 67, 16513008
    SetPixel m_hBtnSrcDC, 3, 67, 16513008
    SetPixel m_hBtnSrcDC, 4, 67, 16513008
    SetPixel m_hBtnSrcDC, 5, 67, 16513266
    SetPixel m_hBtnSrcDC, 6, 67, 15254407
    SetPixel m_hBtnSrcDC, 0, 68, 15254407
    SetPixel m_hBtnSrcDC, 1, 68, 16513009
    SetPixel m_hBtnSrcDC, 2, 68, 16512750
    SetPixel m_hBtnSrcDC, 3, 68, 16512750
    SetPixel m_hBtnSrcDC, 4, 68, 16512750
    SetPixel m_hBtnSrcDC, 5, 68, 16513009
    SetPixel m_hBtnSrcDC, 6, 68, 15254407
    SetPixel m_hBtnSrcDC, 0, 69, 15254407
    SetPixel m_hBtnSrcDC, 1, 69, 16513008
    SetPixel m_hBtnSrcDC, 2, 69, 16512492
    SetPixel m_hBtnSrcDC, 3, 69, 16512492
    SetPixel m_hBtnSrcDC, 4, 69, 16512492
    SetPixel m_hBtnSrcDC, 5, 69, 16513008
    SetPixel m_hBtnSrcDC, 6, 69, 15254407
    SetPixel m_hBtnSrcDC, 0, 70, 15254407
    SetPixel m_hBtnSrcDC, 1, 70, 16512751
    SetPixel m_hBtnSrcDC, 2, 70, 16512234
    SetPixel m_hBtnSrcDC, 3, 70, 16512234
    SetPixel m_hBtnSrcDC, 4, 70, 16512234
    SetPixel m_hBtnSrcDC, 5, 70, 16512751
    SetPixel m_hBtnSrcDC, 6, 70, 15254407
    SetPixel m_hBtnSrcDC, 0, 71, 15254407
    SetPixel m_hBtnSrcDC, 1, 71, 16512493
    SetPixel m_hBtnSrcDC, 2, 71, 16511976
    SetPixel m_hBtnSrcDC, 3, 71, 16511976
    SetPixel m_hBtnSrcDC, 4, 71, 16511976
    SetPixel m_hBtnSrcDC, 5, 71, 16512493
    SetPixel m_hBtnSrcDC, 6, 71, 15254407
    SetPixel m_hBtnSrcDC, 0, 72, 15254407
    SetPixel m_hBtnSrcDC, 1, 72, 16512492
    SetPixel m_hBtnSrcDC, 2, 72, 16446182
    SetPixel m_hBtnSrcDC, 3, 72, 16446182
    SetPixel m_hBtnSrcDC, 4, 72, 16446182
    SetPixel m_hBtnSrcDC, 5, 72, 16512492
    SetPixel m_hBtnSrcDC, 6, 72, 15254407
    SetPixel m_hBtnSrcDC, 0, 73, 15254407
    SetPixel m_hBtnSrcDC, 1, 73, 16512235
    SetPixel m_hBtnSrcDC, 2, 73, 16445925
    SetPixel m_hBtnSrcDC, 3, 73, 16445925
    SetPixel m_hBtnSrcDC, 4, 73, 16445925
    SetPixel m_hBtnSrcDC, 5, 73, 16512235
    SetPixel m_hBtnSrcDC, 6, 73, 15254407
    SetPixel m_hBtnSrcDC, 0, 74, 15254407
    SetPixel m_hBtnSrcDC, 1, 74, 16445927
    SetPixel m_hBtnSrcDC, 2, 74, 16313823
    SetPixel m_hBtnSrcDC, 3, 74, 16313823
    SetPixel m_hBtnSrcDC, 4, 74, 16313823
    SetPixel m_hBtnSrcDC, 5, 74, 16445927
    SetPixel m_hBtnSrcDC, 6, 74, 15254407
    SetPixel m_hBtnSrcDC, 0, 75, 15254407
    SetPixel m_hBtnSrcDC, 1, 75, 16380134
    SetPixel m_hBtnSrcDC, 2, 75, 16247773
    SetPixel m_hBtnSrcDC, 3, 75, 16247773
    SetPixel m_hBtnSrcDC, 4, 75, 16247773
    SetPixel m_hBtnSrcDC, 5, 75, 16380134
    SetPixel m_hBtnSrcDC, 6, 75, 15254407
    SetPixel m_hBtnSrcDC, 0, 76, 15254407
    SetPixel m_hBtnSrcDC, 1, 76, 16314340
    SetPixel m_hBtnSrcDC, 2, 76, 16181979
    SetPixel m_hBtnSrcDC, 3, 76, 16181979
    SetPixel m_hBtnSrcDC, 4, 76, 16181979
    SetPixel m_hBtnSrcDC, 5, 76, 16314340
    SetPixel m_hBtnSrcDC, 6, 76, 15254407
    SetPixel m_hBtnSrcDC, 0, 77, 15254407
    SetPixel m_hBtnSrcDC, 1, 77, 16313826
    SetPixel m_hBtnSrcDC, 2, 77, 16115672
    SetPixel m_hBtnSrcDC, 3, 77, 16115672
    SetPixel m_hBtnSrcDC, 4, 77, 16115672
    SetPixel m_hBtnSrcDC, 5, 77, 16313826
    SetPixel m_hBtnSrcDC, 6, 77, 15254407
    SetPixel m_hBtnSrcDC, 0, 78, 15254407
    SetPixel m_hBtnSrcDC, 1, 78, 16248032
    SetPixel m_hBtnSrcDC, 2, 78, 16049877
    SetPixel m_hBtnSrcDC, 3, 78, 16049877
    SetPixel m_hBtnSrcDC, 4, 78, 16049877
    SetPixel m_hBtnSrcDC, 5, 78, 16248032
    SetPixel m_hBtnSrcDC, 6, 78, 15254407
    SetPixel m_hBtnSrcDC, 0, 79, 15254407
    SetPixel m_hBtnSrcDC, 1, 79, 16182239
    SetPixel m_hBtnSrcDC, 2, 79, 15983827
    SetPixel m_hBtnSrcDC, 3, 79, 15983827
    SetPixel m_hBtnSrcDC, 4, 79, 15983827
    SetPixel m_hBtnSrcDC, 5, 79, 16182239
    SetPixel m_hBtnSrcDC, 6, 79, 15254407
    SetPixel m_hBtnSrcDC, 0, 80, 15254407
    SetPixel m_hBtnSrcDC, 1, 80, 16181982
    SetPixel m_hBtnSrcDC, 2, 80, 15983570
    SetPixel m_hBtnSrcDC, 3, 80, 15983570
    SetPixel m_hBtnSrcDC, 4, 80, 15983570
    SetPixel m_hBtnSrcDC, 5, 80, 16181982
    SetPixel m_hBtnSrcDC, 6, 80, 15254407
    SetPixel m_hBtnSrcDC, 0, 81, 15320717
    SetPixel m_hBtnSrcDC, 1, 81, 16115670
    SetPixel m_hBtnSrcDC, 2, 81, 16049879
    SetPixel m_hBtnSrcDC, 3, 81, 15918033
    SetPixel m_hBtnSrcDC, 4, 81, 16049879
    SetPixel m_hBtnSrcDC, 5, 81, 16115670
    SetPixel m_hBtnSrcDC, 6, 81, 15320717
    SetPixel m_hBtnSrcDC, 0, 82, 15717291
    SetPixel m_hBtnSrcDC, 1, 82, 15652013
    SetPixel m_hBtnSrcDC, 2, 82, 16115671
    SetPixel m_hBtnSrcDC, 3, 82, 16182239
    SetPixel m_hBtnSrcDC, 4, 82, 16115671
    SetPixel m_hBtnSrcDC, 5, 82, 15652013
    SetPixel m_hBtnSrcDC, 6, 82, 15651754
    SetPixel m_hBtnSrcDC, 0, 83, 16512754
    SetPixel m_hBtnSrcDC, 1, 83, 15651497
    SetPixel m_hBtnSrcDC, 2, 83, 15386254
    SetPixel m_hBtnSrcDC, 3, 83, 15254407
    SetPixel m_hBtnSrcDC, 4, 83, 15386254
    SetPixel m_hBtnSrcDC, 5, 83, 15651497
    SetPixel m_hBtnSrcDC, 6, 83, 16512754
    
    SetPixel m_hCbbSrcDC, 0, 0, 13738062
    SetPixel m_hCbbSrcDC, 1, 0, 16775656
    SetPixel m_hCbbSrcDC, 2, 0, 16775656
    SetPixel m_hCbbSrcDC, 3, 0, 16775656
    SetPixel m_hCbbSrcDC, 4, 0, 13738062
    SetPixel m_hCbbSrcDC, 5, 0, 15983038
    SetPixel m_hCbbSrcDC, 6, 0, 15983038
    SetPixel m_hCbbSrcDC, 7, 0, 15983038
    SetPixel m_hCbbSrcDC, 8, 0, 6247705
    SetPixel m_hCbbSrcDC, 9, 0, 6247705
    SetPixel m_hCbbSrcDC, 10, 0, 6247705
    SetPixel m_hCbbSrcDC, 11, 0, 6247705
    SetPixel m_hCbbSrcDC, 12, 0, 6247705
    SetPixel m_hCbbSrcDC, 13, 0, 6247705
    SetPixel m_hCbbSrcDC, 14, 0, 6247705
    SetPixel m_hCbbSrcDC, 0, 1, 13738062
    SetPixel m_hCbbSrcDC, 1, 1, 16709862
    SetPixel m_hCbbSrcDC, 2, 1, 16641729
    SetPixel m_hCbbSrcDC, 3, 1, 16709862
    SetPixel m_hCbbSrcDC, 4, 1, 13738062
    SetPixel m_hCbbSrcDC, 5, 1, 15983040
    SetPixel m_hCbbSrcDC, 6, 1, 14857569
    SetPixel m_hCbbSrcDC, 7, 1, 15983040
    SetPixel m_hCbbSrcDC, 8, 1, 16711935
    SetPixel m_hCbbSrcDC, 9, 1, 6247705
    SetPixel m_hCbbSrcDC, 10, 1, 6247705
    SetPixel m_hCbbSrcDC, 11, 1, 6247705
    SetPixel m_hCbbSrcDC, 12, 1, 6247705
    SetPixel m_hCbbSrcDC, 13, 1, 6247705
    SetPixel m_hCbbSrcDC, 14, 1, 16711935
    SetPixel m_hCbbSrcDC, 0, 2, 13738062
    SetPixel m_hCbbSrcDC, 1, 2, 16709604
    SetPixel m_hCbbSrcDC, 2, 2, 16575420
    SetPixel m_hCbbSrcDC, 3, 2, 16709604
    SetPixel m_hCbbSrcDC, 4, 2, 13738062
    SetPixel m_hCbbSrcDC, 5, 2, 16049089
    SetPixel m_hCbbSrcDC, 6, 2, 14923877
    SetPixel m_hCbbSrcDC, 7, 2, 16049089
    SetPixel m_hCbbSrcDC, 8, 2, 16711935
    SetPixel m_hCbbSrcDC, 9, 2, 16711935
    SetPixel m_hCbbSrcDC, 10, 2, 6247705
    SetPixel m_hCbbSrcDC, 11, 2, 6247705
    SetPixel m_hCbbSrcDC, 12, 2, 6247705
    SetPixel m_hCbbSrcDC, 13, 2, 16711935
    SetPixel m_hCbbSrcDC, 14, 2, 16711935
    SetPixel m_hCbbSrcDC, 0, 3, 13738062
    SetPixel m_hCbbSrcDC, 1, 3, 16643810
    SetPixel m_hCbbSrcDC, 2, 3, 16443575
    SetPixel m_hCbbSrcDC, 3, 3, 16643810
    SetPixel m_hCbbSrcDC, 4, 3, 13738062
    SetPixel m_hCbbSrcDC, 5, 3, 16114883
    SetPixel m_hCbbSrcDC, 6, 3, 15055722
    SetPixel m_hCbbSrcDC, 7, 3, 16114883
    SetPixel m_hCbbSrcDC, 8, 3, 16711935
    SetPixel m_hCbbSrcDC, 9, 3, 16711935
    SetPixel m_hCbbSrcDC, 10, 3, 16711935
    SetPixel m_hCbbSrcDC, 11, 3, 6247705
    SetPixel m_hCbbSrcDC, 12, 3, 16711935
    SetPixel m_hCbbSrcDC, 13, 3, 16711935
    SetPixel m_hCbbSrcDC, 14, 3, 16711935
    SetPixel m_hCbbSrcDC, 0, 4, 13738062
    SetPixel m_hCbbSrcDC, 1, 4, 16643552
    SetPixel m_hCbbSrcDC, 2, 4, 16377266
    SetPixel m_hCbbSrcDC, 3, 4, 16643552
    SetPixel m_hCbbSrcDC, 4, 4, 13738062
    SetPixel m_hCbbSrcDC, 5, 4, 16115142
    SetPixel m_hCbbSrcDC, 6, 4, 15122032
    SetPixel m_hCbbSrcDC, 7, 4, 16115142
    SetPixel m_hCbbSrcDC, 8, 4, 11314564
    SetPixel m_hCbbSrcDC, 9, 4, 11314564
    SetPixel m_hCbbSrcDC, 10, 4, 11314564
    SetPixel m_hCbbSrcDC, 11, 4, 11314564
    SetPixel m_hCbbSrcDC, 12, 4, 11314564
    SetPixel m_hCbbSrcDC, 13, 4, 11314564
    SetPixel m_hCbbSrcDC, 14, 4, 11314564
    SetPixel m_hCbbSrcDC, 0, 5, 13738062
    SetPixel m_hCbbSrcDC, 1, 5, 16577502
    SetPixel m_hCbbSrcDC, 2, 5, 16245164
    SetPixel m_hCbbSrcDC, 3, 5, 16577502
    SetPixel m_hCbbSrcDC, 4, 5, 13738062
    SetPixel m_hCbbSrcDC, 5, 5, 16181192
    SetPixel m_hCbbSrcDC, 6, 5, 15254134
    SetPixel m_hCbbSrcDC, 7, 5, 16181192
    SetPixel m_hCbbSrcDC, 8, 5, 16711935
    SetPixel m_hCbbSrcDC, 9, 5, 11314564
    SetPixel m_hCbbSrcDC, 10, 5, 11314564
    SetPixel m_hCbbSrcDC, 11, 5, 11314564
    SetPixel m_hCbbSrcDC, 12, 5, 11314564
    SetPixel m_hCbbSrcDC, 13, 5, 11314564
    SetPixel m_hCbbSrcDC, 14, 5, 16711935
    SetPixel m_hCbbSrcDC, 0, 6, 13738062
    SetPixel m_hCbbSrcDC, 1, 6, 16511451
    SetPixel m_hCbbSrcDC, 2, 6, 16113061
    SetPixel m_hCbbSrcDC, 3, 6, 16511451
    SetPixel m_hCbbSrcDC, 4, 6, 13738062
    SetPixel m_hCbbSrcDC, 5, 6, 16246987
    SetPixel m_hCbbSrcDC, 6, 6, 15385980
    SetPixel m_hCbbSrcDC, 7, 6, 16246987
    SetPixel m_hCbbSrcDC, 8, 6, 16711935
    SetPixel m_hCbbSrcDC, 9, 6, 16711935
    SetPixel m_hCbbSrcDC, 10, 6, 11314564
    SetPixel m_hCbbSrcDC, 11, 6, 11314564
    SetPixel m_hCbbSrcDC, 12, 6, 11314564
    SetPixel m_hCbbSrcDC, 13, 6, 16711935
    SetPixel m_hCbbSrcDC, 14, 6, 16711935
    SetPixel m_hCbbSrcDC, 0, 7, 13738062
    SetPixel m_hCbbSrcDC, 1, 7, 16445657
    SetPixel m_hCbbSrcDC, 2, 7, 15980959
    SetPixel m_hCbbSrcDC, 3, 7, 16445657
    SetPixel m_hCbbSrcDC, 4, 7, 13738062
    SetPixel m_hCbbSrcDC, 5, 7, 16247245
    SetPixel m_hCbbSrcDC, 6, 7, 15518083
    SetPixel m_hCbbSrcDC, 7, 7, 16247245
    SetPixel m_hCbbSrcDC, 8, 7, 16711935
    SetPixel m_hCbbSrcDC, 9, 7, 16711935
    SetPixel m_hCbbSrcDC, 10, 7, 16711935
    SetPixel m_hCbbSrcDC, 11, 7, 11314564
    SetPixel m_hCbbSrcDC, 12, 7, 16711935
    SetPixel m_hCbbSrcDC, 13, 7, 16711935
    SetPixel m_hCbbSrcDC, 14, 7, 16711935
    SetPixel m_hCbbSrcDC, 0, 8, 13738062
    SetPixel m_hCbbSrcDC, 1, 8, 16379606
    SetPixel m_hCbbSrcDC, 2, 8, 15848856
    SetPixel m_hCbbSrcDC, 3, 8, 16379606
    SetPixel m_hCbbSrcDC, 4, 8, 13738062
    SetPixel m_hCbbSrcDC, 5, 8, 16313296
    SetPixel m_hCbbSrcDC, 6, 8, 15650186
    SetPixel m_hCbbSrcDC, 7, 8, 16313296
    SetPixel m_hCbbSrcDC, 8, 8, 16711935
    SetPixel m_hCbbSrcDC, 9, 8, 16711935
    SetPixel m_hCbbSrcDC, 10, 8, 16711935
    SetPixel m_hCbbSrcDC, 11, 8, 16711935
    SetPixel m_hCbbSrcDC, 12, 8, 16711935
    SetPixel m_hCbbSrcDC, 13, 8, 16711935
    SetPixel m_hCbbSrcDC, 14, 8, 16711935
    SetPixel m_hCbbSrcDC, 0, 9, 13738062
    SetPixel m_hCbbSrcDC, 1, 9, 16379347
    SetPixel m_hCbbSrcDC, 2, 9, 15716753
    SetPixel m_hCbbSrcDC, 3, 9, 16379347
    SetPixel m_hCbbSrcDC, 4, 9, 13738062
    SetPixel m_hCbbSrcDC, 5, 9, 16379347
    SetPixel m_hCbbSrcDC, 6, 9, 15716753
    SetPixel m_hCbbSrcDC, 7, 9, 16379347
    SetPixel m_hCbbSrcDC, 8, 9, 16711935
    SetPixel m_hCbbSrcDC, 9, 9, 16711935
    SetPixel m_hCbbSrcDC, 10, 9, 16711935
    SetPixel m_hCbbSrcDC, 11, 9, 16711935
    SetPixel m_hCbbSrcDC, 12, 9, 16711935
    SetPixel m_hCbbSrcDC, 13, 9, 16711935
    SetPixel m_hCbbSrcDC, 14, 9, 16711935
    SetPixel m_hCbbSrcDC, 0, 10, 13738062
    SetPixel m_hCbbSrcDC, 1, 10, 16313296
    SetPixel m_hCbbSrcDC, 2, 10, 15650186
    SetPixel m_hCbbSrcDC, 3, 10, 16313296
    SetPixel m_hCbbSrcDC, 4, 10, 13738062
    SetPixel m_hCbbSrcDC, 5, 10, 16379606
    SetPixel m_hCbbSrcDC, 6, 10, 15848856
    SetPixel m_hCbbSrcDC, 7, 10, 16379606
    SetPixel m_hCbbSrcDC, 8, 10, 16711935
    SetPixel m_hCbbSrcDC, 9, 10, 16711935
    SetPixel m_hCbbSrcDC, 10, 10, 16711935
    SetPixel m_hCbbSrcDC, 11, 10, 16711935
    SetPixel m_hCbbSrcDC, 12, 10, 16711935
    SetPixel m_hCbbSrcDC, 13, 10, 16711935
    SetPixel m_hCbbSrcDC, 14, 10, 16711935
    SetPixel m_hCbbSrcDC, 0, 11, 13738062
    SetPixel m_hCbbSrcDC, 1, 11, 16247245
    SetPixel m_hCbbSrcDC, 2, 11, 15518083
    SetPixel m_hCbbSrcDC, 3, 11, 16247245
    SetPixel m_hCbbSrcDC, 4, 11, 13738062
    SetPixel m_hCbbSrcDC, 5, 11, 16445657
    SetPixel m_hCbbSrcDC, 6, 11, 15980959
    SetPixel m_hCbbSrcDC, 7, 11, 16445657
    SetPixel m_hCbbSrcDC, 8, 11, 16711935
    SetPixel m_hCbbSrcDC, 9, 11, 16711935
    SetPixel m_hCbbSrcDC, 10, 11, 16711935
    SetPixel m_hCbbSrcDC, 11, 11, 16711935
    SetPixel m_hCbbSrcDC, 12, 11, 16711935
    SetPixel m_hCbbSrcDC, 13, 11, 16711935
    SetPixel m_hCbbSrcDC, 14, 11, 16711935
    SetPixel m_hCbbSrcDC, 0, 12, 13738062
    SetPixel m_hCbbSrcDC, 1, 12, 16246987
    SetPixel m_hCbbSrcDC, 2, 12, 15385980
    SetPixel m_hCbbSrcDC, 3, 12, 16246987
    SetPixel m_hCbbSrcDC, 4, 12, 13738062
    SetPixel m_hCbbSrcDC, 5, 12, 16511451
    SetPixel m_hCbbSrcDC, 6, 12, 16113061
    SetPixel m_hCbbSrcDC, 7, 12, 16511451
    SetPixel m_hCbbSrcDC, 8, 12, 16711935
    SetPixel m_hCbbSrcDC, 9, 12, 16711935
    SetPixel m_hCbbSrcDC, 10, 12, 16711935
    SetPixel m_hCbbSrcDC, 11, 12, 16711935
    SetPixel m_hCbbSrcDC, 12, 12, 16711935
    SetPixel m_hCbbSrcDC, 13, 12, 16711935
    SetPixel m_hCbbSrcDC, 14, 12, 16711935
    SetPixel m_hCbbSrcDC, 0, 13, 13738062
    SetPixel m_hCbbSrcDC, 1, 13, 16181192
    SetPixel m_hCbbSrcDC, 2, 13, 15254134
    SetPixel m_hCbbSrcDC, 3, 13, 16181192
    SetPixel m_hCbbSrcDC, 4, 13, 13738062
    SetPixel m_hCbbSrcDC, 5, 13, 16577502
    SetPixel m_hCbbSrcDC, 6, 13, 16245164
    SetPixel m_hCbbSrcDC, 7, 13, 16577502
    SetPixel m_hCbbSrcDC, 8, 13, 16711935
    SetPixel m_hCbbSrcDC, 9, 13, 16711935
    SetPixel m_hCbbSrcDC, 10, 13, 16711935
    SetPixel m_hCbbSrcDC, 11, 13, 16711935
    SetPixel m_hCbbSrcDC, 12, 13, 16711935
    SetPixel m_hCbbSrcDC, 13, 13, 16711935
    SetPixel m_hCbbSrcDC, 14, 13, 16711935
    SetPixel m_hCbbSrcDC, 0, 14, 13738062
    SetPixel m_hCbbSrcDC, 1, 14, 16115142
    SetPixel m_hCbbSrcDC, 2, 14, 15122032
    SetPixel m_hCbbSrcDC, 3, 14, 16115142
    SetPixel m_hCbbSrcDC, 4, 14, 13738062
    SetPixel m_hCbbSrcDC, 5, 14, 16643552
    SetPixel m_hCbbSrcDC, 6, 14, 16377266
    SetPixel m_hCbbSrcDC, 7, 14, 16643552
    SetPixel m_hCbbSrcDC, 8, 14, 16711935
    SetPixel m_hCbbSrcDC, 9, 14, 16711935
    SetPixel m_hCbbSrcDC, 10, 14, 16711935
    SetPixel m_hCbbSrcDC, 11, 14, 16711935
    SetPixel m_hCbbSrcDC, 12, 14, 16711935
    SetPixel m_hCbbSrcDC, 13, 14, 16711935
    SetPixel m_hCbbSrcDC, 14, 14, 16711935
    SetPixel m_hCbbSrcDC, 0, 15, 13738062
    SetPixel m_hCbbSrcDC, 1, 15, 16114883
    SetPixel m_hCbbSrcDC, 2, 15, 15055722
    SetPixel m_hCbbSrcDC, 3, 15, 16114883
    SetPixel m_hCbbSrcDC, 4, 15, 13738062
    SetPixel m_hCbbSrcDC, 5, 15, 16643810
    SetPixel m_hCbbSrcDC, 6, 15, 16443575
    SetPixel m_hCbbSrcDC, 7, 15, 16643810
    SetPixel m_hCbbSrcDC, 8, 15, 16711935
    SetPixel m_hCbbSrcDC, 9, 15, 16711935
    SetPixel m_hCbbSrcDC, 10, 15, 16711935
    SetPixel m_hCbbSrcDC, 11, 15, 16711935
    SetPixel m_hCbbSrcDC, 12, 15, 16711935
    SetPixel m_hCbbSrcDC, 13, 15, 16711935
    SetPixel m_hCbbSrcDC, 14, 15, 16711935
    SetPixel m_hCbbSrcDC, 0, 16, 13738062
    SetPixel m_hCbbSrcDC, 1, 16, 16049089
    SetPixel m_hCbbSrcDC, 2, 16, 14923877
    SetPixel m_hCbbSrcDC, 3, 16, 16049089
    SetPixel m_hCbbSrcDC, 4, 16, 13738062
    SetPixel m_hCbbSrcDC, 5, 16, 16709604
    SetPixel m_hCbbSrcDC, 6, 16, 16575420
    SetPixel m_hCbbSrcDC, 7, 16, 16709604
    SetPixel m_hCbbSrcDC, 8, 16, 16711935
    SetPixel m_hCbbSrcDC, 9, 16, 16711935
    SetPixel m_hCbbSrcDC, 10, 16, 16711935
    SetPixel m_hCbbSrcDC, 11, 16, 16711935
    SetPixel m_hCbbSrcDC, 12, 16, 16711935
    SetPixel m_hCbbSrcDC, 13, 16, 16711935
    SetPixel m_hCbbSrcDC, 14, 16, 16711935
    SetPixel m_hCbbSrcDC, 0, 17, 13738062
    SetPixel m_hCbbSrcDC, 1, 17, 15983040
    SetPixel m_hCbbSrcDC, 2, 17, 15983040
    SetPixel m_hCbbSrcDC, 3, 17, 15983040
    SetPixel m_hCbbSrcDC, 4, 17, 13738062
    SetPixel m_hCbbSrcDC, 5, 17, 16709862
    SetPixel m_hCbbSrcDC, 6, 17, 16709862
    SetPixel m_hCbbSrcDC, 7, 17, 16709862
    SetPixel m_hCbbSrcDC, 8, 17, 16711935
    SetPixel m_hCbbSrcDC, 9, 17, 16711935
    SetPixel m_hCbbSrcDC, 10, 17, 16711935
    SetPixel m_hCbbSrcDC, 11, 17, 16711935
    SetPixel m_hCbbSrcDC, 12, 17, 16711935
    SetPixel m_hCbbSrcDC, 13, 17, 16711935
    SetPixel m_hCbbSrcDC, 14, 17, 16711935
    
    SetPixel m_hCkbSrcDC, 0, 0, 16711935
    SetPixel m_hCkbSrcDC, 1, 0, 16711935
    SetPixel m_hCkbSrcDC, 2, 0, 16711935
    SetPixel m_hCkbSrcDC, 3, 0, 16711935
    SetPixel m_hCkbSrcDC, 4, 0, 16711935
    SetPixel m_hCkbSrcDC, 5, 0, 16711935
    SetPixel m_hCkbSrcDC, 6, 0, 16711935
    SetPixel m_hCkbSrcDC, 7, 0, 16711935
    SetPixel m_hCkbSrcDC, 8, 0, 1947476
    SetPixel m_hCkbSrcDC, 9, 0, 16711935
    SetPixel m_hCkbSrcDC, 10, 0, 16711935
    SetPixel m_hCkbSrcDC, 11, 0, 16711935
    SetPixel m_hCkbSrcDC, 12, 0, 16711935
    SetPixel m_hCkbSrcDC, 13, 0, 16711935
    SetPixel m_hCkbSrcDC, 14, 0, 16711935
    SetPixel m_hCkbSrcDC, 15, 0, 16711935
    SetPixel m_hCkbSrcDC, 16, 0, 16711935
    SetPixel m_hCkbSrcDC, 17, 0, 9231016
    SetPixel m_hCkbSrcDC, 0, 1, 16711935
    SetPixel m_hCkbSrcDC, 1, 1, 16711935
    SetPixel m_hCkbSrcDC, 2, 1, 16711935
    SetPixel m_hCkbSrcDC, 3, 1, 16711935
    SetPixel m_hCkbSrcDC, 4, 1, 16711935
    SetPixel m_hCkbSrcDC, 5, 1, 16711935
    SetPixel m_hCkbSrcDC, 6, 1, 16711935
    SetPixel m_hCkbSrcDC, 7, 1, 1815633
    SetPixel m_hCkbSrcDC, 8, 1, 1815633
    SetPixel m_hCkbSrcDC, 9, 1, 16711935
    SetPixel m_hCkbSrcDC, 10, 1, 16711935
    SetPixel m_hCkbSrcDC, 11, 1, 16711935
    SetPixel m_hCkbSrcDC, 12, 1, 16711935
    SetPixel m_hCkbSrcDC, 13, 1, 16711935
    SetPixel m_hCkbSrcDC, 14, 1, 16711935
    SetPixel m_hCkbSrcDC, 15, 1, 16711935
    SetPixel m_hCkbSrcDC, 16, 1, 9165222
    SetPixel m_hCkbSrcDC, 17, 1, 9165222
    SetPixel m_hCkbSrcDC, 0, 2, 16711935
    SetPixel m_hCkbSrcDC, 1, 2, 16711935
    SetPixel m_hCkbSrcDC, 2, 2, 16711935
    SetPixel m_hCkbSrcDC, 3, 2, 16711935
    SetPixel m_hCkbSrcDC, 4, 2, 16711935
    SetPixel m_hCkbSrcDC, 5, 2, 16711935
    SetPixel m_hCkbSrcDC, 6, 2, 1552205
    SetPixel m_hCkbSrcDC, 7, 2, 1552205
    SetPixel m_hCkbSrcDC, 8, 2, 1552205
    SetPixel m_hCkbSrcDC, 9, 2, 16711935
    SetPixel m_hCkbSrcDC, 10, 2, 16711935
    SetPixel m_hCkbSrcDC, 11, 2, 16711935
    SetPixel m_hCkbSrcDC, 12, 2, 16711935
    SetPixel m_hCkbSrcDC, 13, 2, 16711935
    SetPixel m_hCkbSrcDC, 14, 2, 16711935
    SetPixel m_hCkbSrcDC, 15, 2, 9033380
    SetPixel m_hCkbSrcDC, 16, 2, 9033380
    SetPixel m_hCkbSrcDC, 17, 2, 9033380
    SetPixel m_hCkbSrcDC, 0, 3, 1288520
    SetPixel m_hCkbSrcDC, 1, 3, 1288520
    SetPixel m_hCkbSrcDC, 2, 3, 16711935
    SetPixel m_hCkbSrcDC, 3, 3, 16711935
    SetPixel m_hCkbSrcDC, 4, 3, 16711935
    SetPixel m_hCkbSrcDC, 5, 3, 1288520
    SetPixel m_hCkbSrcDC, 6, 3, 1288520
    SetPixel m_hCkbSrcDC, 7, 3, 1288520
    SetPixel m_hCkbSrcDC, 8, 3, 16711935
    SetPixel m_hCkbSrcDC, 9, 3, 8901538
    SetPixel m_hCkbSrcDC, 10, 3, 8901538
    SetPixel m_hCkbSrcDC, 11, 3, 16711935
    SetPixel m_hCkbSrcDC, 12, 3, 16711935
    SetPixel m_hCkbSrcDC, 13, 3, 16711935
    SetPixel m_hCkbSrcDC, 14, 3, 8901538
    SetPixel m_hCkbSrcDC, 15, 3, 8901538
    SetPixel m_hCkbSrcDC, 16, 3, 8901538
    SetPixel m_hCkbSrcDC, 17, 3, 16711935
    SetPixel m_hCkbSrcDC, 0, 4, 1024579
    SetPixel m_hCkbSrcDC, 1, 4, 1024579
    SetPixel m_hCkbSrcDC, 2, 4, 1024579
    SetPixel m_hCkbSrcDC, 3, 4, 16711935
    SetPixel m_hCkbSrcDC, 4, 4, 1024579
    SetPixel m_hCkbSrcDC, 5, 4, 1024579
    SetPixel m_hCkbSrcDC, 6, 4, 1024579
    SetPixel m_hCkbSrcDC, 7, 4, 16711935
    SetPixel m_hCkbSrcDC, 8, 4, 16711935
    SetPixel m_hCkbSrcDC, 9, 4, 8769695
    SetPixel m_hCkbSrcDC, 10, 4, 8769695
    SetPixel m_hCkbSrcDC, 11, 4, 8769695
    SetPixel m_hCkbSrcDC, 12, 4, 16711935
    SetPixel m_hCkbSrcDC, 13, 4, 8769695
    SetPixel m_hCkbSrcDC, 14, 4, 8769695
    SetPixel m_hCkbSrcDC, 15, 4, 8769695
    SetPixel m_hCkbSrcDC, 16, 4, 16711935
    SetPixel m_hCkbSrcDC, 17, 4, 16711935
    SetPixel m_hCkbSrcDC, 0, 5, 16711935
    SetPixel m_hCkbSrcDC, 1, 5, 695358
    SetPixel m_hCkbSrcDC, 2, 5, 695358
    SetPixel m_hCkbSrcDC, 3, 5, 695358
    SetPixel m_hCkbSrcDC, 4, 5, 695358
    SetPixel m_hCkbSrcDC, 5, 5, 695358
    SetPixel m_hCkbSrcDC, 6, 5, 16711935
    SetPixel m_hCkbSrcDC, 7, 5, 16711935
    SetPixel m_hCkbSrcDC, 8, 5, 16711935
    SetPixel m_hCkbSrcDC, 9, 5, 16711935
    SetPixel m_hCkbSrcDC, 10, 5, 8572317
    SetPixel m_hCkbSrcDC, 11, 5, 8572317
    SetPixel m_hCkbSrcDC, 12, 5, 8572317
    SetPixel m_hCkbSrcDC, 13, 5, 8572317
    SetPixel m_hCkbSrcDC, 14, 5, 8572317
    SetPixel m_hCkbSrcDC, 15, 5, 16711935
    SetPixel m_hCkbSrcDC, 16, 5, 16711935
    SetPixel m_hCkbSrcDC, 17, 5, 16711935
    SetPixel m_hCkbSrcDC, 0, 6, 16711935
    SetPixel m_hCkbSrcDC, 1, 6, 16711935
    SetPixel m_hCkbSrcDC, 2, 6, 431673
    SetPixel m_hCkbSrcDC, 3, 6, 431673
    SetPixel m_hCkbSrcDC, 4, 6, 431673
    SetPixel m_hCkbSrcDC, 5, 6, 16711935
    SetPixel m_hCkbSrcDC, 6, 6, 16711935
    SetPixel m_hCkbSrcDC, 7, 6, 16711935
    SetPixel m_hCkbSrcDC, 8, 6, 16711935
    SetPixel m_hCkbSrcDC, 9, 6, 16711935
    SetPixel m_hCkbSrcDC, 10, 6, 16711935
    SetPixel m_hCkbSrcDC, 11, 6, 8440218
    SetPixel m_hCkbSrcDC, 12, 6, 8440218
    SetPixel m_hCkbSrcDC, 13, 6, 8440218
    SetPixel m_hCkbSrcDC, 14, 6, 16711935
    SetPixel m_hCkbSrcDC, 15, 6, 16711935
    SetPixel m_hCkbSrcDC, 16, 6, 16711935
    SetPixel m_hCkbSrcDC, 17, 6, 16711935
    SetPixel m_hCkbSrcDC, 0, 7, 16711935
    SetPixel m_hCkbSrcDC, 1, 7, 16711935
    SetPixel m_hCkbSrcDC, 2, 7, 16711935
    SetPixel m_hCkbSrcDC, 3, 7, 168245
    SetPixel m_hCkbSrcDC, 4, 7, 16711935
    SetPixel m_hCkbSrcDC, 5, 7, 16711935
    SetPixel m_hCkbSrcDC, 6, 7, 16711935
    SetPixel m_hCkbSrcDC, 7, 7, 16711935
    SetPixel m_hCkbSrcDC, 8, 7, 16711935
    SetPixel m_hCkbSrcDC, 9, 7, 16711935
    SetPixel m_hCkbSrcDC, 10, 7, 16711935
    SetPixel m_hCkbSrcDC, 11, 7, 16711935
    SetPixel m_hCkbSrcDC, 12, 7, 8308632
    SetPixel m_hCkbSrcDC, 13, 7, 16711935
    SetPixel m_hCkbSrcDC, 14, 7, 16711935
    SetPixel m_hCkbSrcDC, 15, 7, 16711935
    SetPixel m_hCkbSrcDC, 16, 7, 16711935
    SetPixel m_hCkbSrcDC, 17, 7, 16711935
    SetPixel m_hCkbSrcDC, 0, 8, 16711935
    SetPixel m_hCkbSrcDC, 1, 8, 16711935
    SetPixel m_hCkbSrcDC, 2, 8, 16711935
    SetPixel m_hCkbSrcDC, 3, 8, 16711935
    SetPixel m_hCkbSrcDC, 4, 8, 16711935
    SetPixel m_hCkbSrcDC, 5, 8, 16711935
    SetPixel m_hCkbSrcDC, 6, 8, 16711935
    SetPixel m_hCkbSrcDC, 7, 8, 16711935
    SetPixel m_hCkbSrcDC, 8, 8, 16711935
    SetPixel m_hCkbSrcDC, 9, 8, 16711935
    SetPixel m_hCkbSrcDC, 10, 8, 16711935
    SetPixel m_hCkbSrcDC, 11, 8, 16711935
    SetPixel m_hCkbSrcDC, 12, 8, 16711935
    SetPixel m_hCkbSrcDC, 13, 8, 16711935
    SetPixel m_hCkbSrcDC, 14, 8, 16711935
    SetPixel m_hCkbSrcDC, 15, 8, 16711935
    SetPixel m_hCkbSrcDC, 16, 8, 16711935
    SetPixel m_hCkbSrcDC, 17, 8, 16711935
    SetPixel m_hCkbSrcDC, 0, 9, 16711935
    SetPixel m_hCkbSrcDC, 1, 9, 1815633
    SetPixel m_hCkbSrcDC, 2, 9, 1815633
    SetPixel m_hCkbSrcDC, 3, 9, 1815633
    SetPixel m_hCkbSrcDC, 4, 9, 1815633
    SetPixel m_hCkbSrcDC, 5, 9, 1815633
    SetPixel m_hCkbSrcDC, 6, 9, 1815633
    SetPixel m_hCkbSrcDC, 7, 9, 1815633
    SetPixel m_hCkbSrcDC, 8, 9, 16711935
    SetPixel m_hCkbSrcDC, 9, 9, 16711935
    SetPixel m_hCkbSrcDC, 10, 9, 9165222
    SetPixel m_hCkbSrcDC, 11, 9, 9165222
    SetPixel m_hCkbSrcDC, 12, 9, 9165222
    SetPixel m_hCkbSrcDC, 13, 9, 9165222
    SetPixel m_hCkbSrcDC, 14, 9, 9165222
    SetPixel m_hCkbSrcDC, 15, 9, 9165222
    SetPixel m_hCkbSrcDC, 16, 9, 9165222
    SetPixel m_hCkbSrcDC, 17, 9, 16711935
    SetPixel m_hCkbSrcDC, 0, 10, 16711935
    SetPixel m_hCkbSrcDC, 1, 10, 1617998
    SetPixel m_hCkbSrcDC, 2, 10, 1617998
    SetPixel m_hCkbSrcDC, 3, 10, 1617998
    SetPixel m_hCkbSrcDC, 4, 10, 1617998
    SetPixel m_hCkbSrcDC, 5, 10, 1617998
    SetPixel m_hCkbSrcDC, 6, 10, 1617998
    SetPixel m_hCkbSrcDC, 7, 10, 1617998
    SetPixel m_hCkbSrcDC, 8, 10, 16711935
    SetPixel m_hCkbSrcDC, 9, 10, 16711935
    SetPixel m_hCkbSrcDC, 10, 10, 9033380
    SetPixel m_hCkbSrcDC, 11, 10, 9033380
    SetPixel m_hCkbSrcDC, 12, 10, 9033380
    SetPixel m_hCkbSrcDC, 13, 10, 9033380
    SetPixel m_hCkbSrcDC, 14, 10, 9033380
    SetPixel m_hCkbSrcDC, 15, 10, 9033380
    SetPixel m_hCkbSrcDC, 16, 10, 9033380
    SetPixel m_hCkbSrcDC, 17, 10, 16711935
    SetPixel m_hCkbSrcDC, 0, 11, 16711935
    SetPixel m_hCkbSrcDC, 1, 11, 1354569
    SetPixel m_hCkbSrcDC, 2, 11, 1354569
    SetPixel m_hCkbSrcDC, 3, 11, 1354569
    SetPixel m_hCkbSrcDC, 4, 11, 1354569
    SetPixel m_hCkbSrcDC, 5, 11, 1354569
    SetPixel m_hCkbSrcDC, 6, 11, 1354569
    SetPixel m_hCkbSrcDC, 7, 11, 1354569
    SetPixel m_hCkbSrcDC, 8, 11, 16711935
    SetPixel m_hCkbSrcDC, 9, 11, 16711935
    SetPixel m_hCkbSrcDC, 10, 11, 8901795
    SetPixel m_hCkbSrcDC, 11, 11, 8901795
    SetPixel m_hCkbSrcDC, 12, 11, 8901795
    SetPixel m_hCkbSrcDC, 13, 11, 8901795
    SetPixel m_hCkbSrcDC, 14, 11, 8901795
    SetPixel m_hCkbSrcDC, 15, 11, 8901795
    SetPixel m_hCkbSrcDC, 16, 11, 8901795
    SetPixel m_hCkbSrcDC, 17, 11, 16711935
    SetPixel m_hCkbSrcDC, 0, 12, 16711935
    SetPixel m_hCkbSrcDC, 1, 12, 1090629
    SetPixel m_hCkbSrcDC, 2, 12, 1090629
    SetPixel m_hCkbSrcDC, 3, 12, 1090629
    SetPixel m_hCkbSrcDC, 4, 12, 1090629
    SetPixel m_hCkbSrcDC, 5, 12, 1090629
    SetPixel m_hCkbSrcDC, 6, 12, 1090629
    SetPixel m_hCkbSrcDC, 7, 12, 1090629
    SetPixel m_hCkbSrcDC, 8, 12, 16711935
    SetPixel m_hCkbSrcDC, 9, 12, 16711935
    SetPixel m_hCkbSrcDC, 10, 12, 8835488
    SetPixel m_hCkbSrcDC, 11, 12, 8835488
    SetPixel m_hCkbSrcDC, 12, 12, 8835488
    SetPixel m_hCkbSrcDC, 13, 12, 8835488
    SetPixel m_hCkbSrcDC, 14, 12, 8835488
    SetPixel m_hCkbSrcDC, 15, 12, 8835488
    SetPixel m_hCkbSrcDC, 16, 12, 8835488
    SetPixel m_hCkbSrcDC, 17, 12, 16711935
    SetPixel m_hCkbSrcDC, 0, 13, 16711935
    SetPixel m_hCkbSrcDC, 1, 13, 826944
    SetPixel m_hCkbSrcDC, 2, 13, 826944
    SetPixel m_hCkbSrcDC, 3, 13, 826944
    SetPixel m_hCkbSrcDC, 4, 13, 826944
    SetPixel m_hCkbSrcDC, 5, 13, 826944
    SetPixel m_hCkbSrcDC, 6, 13, 826944
    SetPixel m_hCkbSrcDC, 7, 13, 826944
    SetPixel m_hCkbSrcDC, 8, 13, 16711935
    SetPixel m_hCkbSrcDC, 9, 13, 16711935
    SetPixel m_hCkbSrcDC, 10, 13, 8638110
    SetPixel m_hCkbSrcDC, 11, 13, 8638110
    SetPixel m_hCkbSrcDC, 12, 13, 8638110
    SetPixel m_hCkbSrcDC, 13, 13, 8638110
    SetPixel m_hCkbSrcDC, 14, 13, 8638110
    SetPixel m_hCkbSrcDC, 15, 13, 8638110
    SetPixel m_hCkbSrcDC, 16, 13, 8638110
    SetPixel m_hCkbSrcDC, 17, 13, 16711935
    SetPixel m_hCkbSrcDC, 0, 14, 16711935
    SetPixel m_hCkbSrcDC, 1, 14, 497723
    SetPixel m_hCkbSrcDC, 2, 14, 497723
    SetPixel m_hCkbSrcDC, 3, 14, 497723
    SetPixel m_hCkbSrcDC, 4, 14, 497723
    SetPixel m_hCkbSrcDC, 5, 14, 497723
    SetPixel m_hCkbSrcDC, 6, 14, 497723
    SetPixel m_hCkbSrcDC, 7, 14, 497723
    SetPixel m_hCkbSrcDC, 8, 14, 16711935
    SetPixel m_hCkbSrcDC, 9, 14, 16711935
    SetPixel m_hCkbSrcDC, 10, 14, 8506011
    SetPixel m_hCkbSrcDC, 11, 14, 8506011
    SetPixel m_hCkbSrcDC, 12, 14, 8506011
    SetPixel m_hCkbSrcDC, 13, 14, 8506011
    SetPixel m_hCkbSrcDC, 14, 14, 8506011
    SetPixel m_hCkbSrcDC, 15, 14, 8506011
    SetPixel m_hCkbSrcDC, 16, 14, 8506011
    SetPixel m_hCkbSrcDC, 17, 14, 16711935
    SetPixel m_hCkbSrcDC, 0, 15, 16711935
    SetPixel m_hCkbSrcDC, 1, 15, 431673
    SetPixel m_hCkbSrcDC, 2, 15, 431673
    SetPixel m_hCkbSrcDC, 3, 15, 431673
    SetPixel m_hCkbSrcDC, 4, 15, 431673
    SetPixel m_hCkbSrcDC, 5, 15, 431673
    SetPixel m_hCkbSrcDC, 6, 15, 431673
    SetPixel m_hCkbSrcDC, 7, 15, 431673
    SetPixel m_hCkbSrcDC, 8, 15, 16711935
    SetPixel m_hCkbSrcDC, 9, 15, 16711935
    SetPixel m_hCkbSrcDC, 10, 15, 8440218
    SetPixel m_hCkbSrcDC, 11, 15, 8440218
    SetPixel m_hCkbSrcDC, 12, 15, 8440218
    SetPixel m_hCkbSrcDC, 13, 15, 8440218
    SetPixel m_hCkbSrcDC, 14, 15, 8440218
    SetPixel m_hCkbSrcDC, 15, 15, 8440218
    SetPixel m_hCkbSrcDC, 16, 15, 8440218
    SetPixel m_hCkbSrcDC, 17, 15, 16711935
    SetPixel m_hCkbSrcDC, 0, 16, 16711935
    SetPixel m_hCkbSrcDC, 1, 16, 16711935
    SetPixel m_hCkbSrcDC, 2, 16, 16711935
    SetPixel m_hCkbSrcDC, 3, 16, 16711935
    SetPixel m_hCkbSrcDC, 4, 16, 16711935
    SetPixel m_hCkbSrcDC, 5, 16, 16711935
    SetPixel m_hCkbSrcDC, 6, 16, 16711935
    SetPixel m_hCkbSrcDC, 7, 16, 16711935
    SetPixel m_hCkbSrcDC, 8, 16, 16711935
    SetPixel m_hCkbSrcDC, 9, 16, 16711935
    SetPixel m_hCkbSrcDC, 10, 16, 16711935
    SetPixel m_hCkbSrcDC, 11, 16, 16711935
    SetPixel m_hCkbSrcDC, 12, 16, 16711935
    SetPixel m_hCkbSrcDC, 13, 16, 16711935
    SetPixel m_hCkbSrcDC, 14, 16, 16711935
    SetPixel m_hCkbSrcDC, 15, 16, 16711935
    SetPixel m_hCkbSrcDC, 16, 16, 16711935
    SetPixel m_hCkbSrcDC, 17, 16, 16711935
    SetPixel m_hCkbSrcDC, 0, 17, 16711935
    SetPixel m_hCkbSrcDC, 1, 17, 16711935
    SetPixel m_hCkbSrcDC, 2, 17, 16711935
    SetPixel m_hCkbSrcDC, 3, 17, 16711935
    SetPixel m_hCkbSrcDC, 4, 17, 16711935
    SetPixel m_hCkbSrcDC, 5, 17, 16711935
    SetPixel m_hCkbSrcDC, 6, 17, 16711935
    SetPixel m_hCkbSrcDC, 7, 17, 16711935
    SetPixel m_hCkbSrcDC, 8, 17, 16711935
    SetPixel m_hCkbSrcDC, 9, 17, 16711935
    SetPixel m_hCkbSrcDC, 10, 17, 16711935
    SetPixel m_hCkbSrcDC, 11, 17, 16711935
    SetPixel m_hCkbSrcDC, 12, 17, 16711935
    SetPixel m_hCkbSrcDC, 13, 17, 16711935
    SetPixel m_hCkbSrcDC, 14, 17, 16711935
    SetPixel m_hCkbSrcDC, 15, 17, 16711935
    SetPixel m_hCkbSrcDC, 16, 17, 16711935
    SetPixel m_hCkbSrcDC, 17, 17, 16711935
    SetPixel m_hCkbSrcDC, 0, 18, 16711935
    SetPixel m_hCkbSrcDC, 1, 18, 16711935
    SetPixel m_hCkbSrcDC, 2, 18, 16711935
    SetPixel m_hCkbSrcDC, 3, 18, 6343815
    SetPixel m_hCkbSrcDC, 4, 18, 2669148
    SetPixel m_hCkbSrcDC, 5, 18, 6343815
    SetPixel m_hCkbSrcDC, 6, 18, 16711935
    SetPixel m_hCkbSrcDC, 7, 18, 16711935
    SetPixel m_hCkbSrcDC, 8, 18, 16711935
    SetPixel m_hCkbSrcDC, 9, 18, 16711935
    SetPixel m_hCkbSrcDC, 10, 18, 16711935
    SetPixel m_hCkbSrcDC, 11, 18, 16711935
    SetPixel m_hCkbSrcDC, 12, 18, 11527619
    SetPixel m_hCkbSrcDC, 13, 18, 9690285
    SetPixel m_hCkbSrcDC, 14, 18, 11527619
    SetPixel m_hCkbSrcDC, 15, 18, 16711935
    SetPixel m_hCkbSrcDC, 16, 18, 16711935
    SetPixel m_hCkbSrcDC, 17, 18, 16711935
    SetPixel m_hCkbSrcDC, 0, 19, 16711935
    SetPixel m_hCkbSrcDC, 1, 19, 16711935
    SetPixel m_hCkbSrcDC, 2, 19, 6211715
    SetPixel m_hCkbSrcDC, 3, 19, 1683791
    SetPixel m_hCkbSrcDC, 4, 19, 1683791
    SetPixel m_hCkbSrcDC, 5, 19, 1683791
    SetPixel m_hCkbSrcDC, 6, 19, 6211715
    SetPixel m_hCkbSrcDC, 7, 19, 16711935
    SetPixel m_hCkbSrcDC, 8, 19, 16711935
    SetPixel m_hCkbSrcDC, 9, 19, 16711935
    SetPixel m_hCkbSrcDC, 10, 19, 16711935
    SetPixel m_hCkbSrcDC, 11, 19, 11461569
    SetPixel m_hCkbSrcDC, 12, 19, 9230503
    SetPixel m_hCkbSrcDC, 13, 19, 9230503
    SetPixel m_hCkbSrcDC, 14, 19, 9230503
    SetPixel m_hCkbSrcDC, 15, 19, 11461569
    SetPixel m_hCkbSrcDC, 16, 19, 16711935
    SetPixel m_hCkbSrcDC, 17, 19, 16711935
    SetPixel m_hCkbSrcDC, 0, 20, 16711935
    SetPixel m_hCkbSrcDC, 1, 20, 16711935
    SetPixel m_hCkbSrcDC, 2, 20, 1944656
    SetPixel m_hCkbSrcDC, 3, 20, 1222727
    SetPixel m_hCkbSrcDC, 4, 20, 1222727
    SetPixel m_hCkbSrcDC, 5, 20, 1222727
    SetPixel m_hCkbSrcDC, 6, 20, 1944656
    SetPixel m_hCkbSrcDC, 7, 20, 16711935
    SetPixel m_hCkbSrcDC, 8, 20, 16711935
    SetPixel m_hCkbSrcDC, 9, 20, 16711935
    SetPixel m_hCkbSrcDC, 10, 20, 16711935
    SetPixel m_hCkbSrcDC, 11, 20, 9360807
    SetPixel m_hCkbSrcDC, 12, 20, 8967075
    SetPixel m_hCkbSrcDC, 13, 20, 8967075
    SetPixel m_hCkbSrcDC, 14, 20, 8967075
    SetPixel m_hCkbSrcDC, 15, 20, 9360807
    SetPixel m_hCkbSrcDC, 16, 20, 16711935
    SetPixel m_hCkbSrcDC, 17, 20, 16711935
    SetPixel m_hCkbSrcDC, 0, 21, 16711935
    SetPixel m_hCkbSrcDC, 1, 21, 16711935
    SetPixel m_hCkbSrcDC, 2, 21, 5552760
    SetPixel m_hCkbSrcDC, 3, 21, 761151
    SetPixel m_hCkbSrcDC, 4, 21, 761151
    SetPixel m_hCkbSrcDC, 5, 21, 761151
    SetPixel m_hCkbSrcDC, 6, 21, 5552760
    SetPixel m_hCkbSrcDC, 7, 21, 16711935
    SetPixel m_hCkbSrcDC, 8, 21, 16711935
    SetPixel m_hCkbSrcDC, 9, 21, 16711935
    SetPixel m_hCkbSrcDC, 10, 21, 16711935
    SetPixel m_hCkbSrcDC, 11, 21, 11132091
    SetPixel m_hCkbSrcDC, 12, 21, 8769183
    SetPixel m_hCkbSrcDC, 13, 21, 8769183
    SetPixel m_hCkbSrcDC, 14, 21, 8769183
    SetPixel m_hCkbSrcDC, 15, 21, 11132091
    SetPixel m_hCkbSrcDC, 16, 21, 16711935
    SetPixel m_hCkbSrcDC, 17, 21, 16711935
    SetPixel m_hCkbSrcDC, 0, 22, 16711935
    SetPixel m_hCkbSrcDC, 1, 22, 16711935
    SetPixel m_hCkbSrcDC, 2, 22, 16711935
    SetPixel m_hCkbSrcDC, 3, 22, 5223539
    SetPixel m_hCkbSrcDC, 4, 22, 1087808
    SetPixel m_hCkbSrcDC, 5, 22, 5223539
    SetPixel m_hCkbSrcDC, 6, 22, 16711935
    SetPixel m_hCkbSrcDC, 7, 22, 16711935
    SetPixel m_hCkbSrcDC, 8, 22, 16711935
    SetPixel m_hCkbSrcDC, 9, 22, 16711935
    SetPixel m_hCkbSrcDC, 10, 22, 16711935
    SetPixel m_hCkbSrcDC, 11, 22, 16711935
    SetPixel m_hCkbSrcDC, 12, 22, 11000249
    SetPixel m_hCkbSrcDC, 13, 22, 8899743
    SetPixel m_hCkbSrcDC, 14, 22, 11000249
    SetPixel m_hCkbSrcDC, 15, 22, 16711935
    SetPixel m_hCkbSrcDC, 16, 22, 16711935
    SetPixel m_hCkbSrcDC, 17, 22, 16711935
    SetPixel m_hCkbSrcDC, 0, 23, 16711935
    SetPixel m_hCkbSrcDC, 1, 23, 16711935
    SetPixel m_hCkbSrcDC, 2, 23, 16711935
    SetPixel m_hCkbSrcDC, 3, 23, 16711935
    SetPixel m_hCkbSrcDC, 4, 23, 16711935
    SetPixel m_hCkbSrcDC, 5, 23, 16711935
    SetPixel m_hCkbSrcDC, 6, 23, 16711935
    SetPixel m_hCkbSrcDC, 7, 23, 16711935
    SetPixel m_hCkbSrcDC, 8, 23, 16711935
    SetPixel m_hCkbSrcDC, 9, 23, 16711935
    SetPixel m_hCkbSrcDC, 10, 23, 16711935
    SetPixel m_hCkbSrcDC, 11, 23, 16711935
    SetPixel m_hCkbSrcDC, 12, 23, 16711935
    SetPixel m_hCkbSrcDC, 13, 23, 16711935
    SetPixel m_hCkbSrcDC, 14, 23, 16711935
    SetPixel m_hCkbSrcDC, 15, 23, 16711935
    SetPixel m_hCkbSrcDC, 16, 23, 16711935
    SetPixel m_hCkbSrcDC, 17, 23, 16711935
    
    SetPixel m_hOpbSrcDC, 0, 0, 16711935
    SetPixel m_hOpbSrcDC, 1, 0, 16711935
    SetPixel m_hOpbSrcDC, 2, 0, 16711935
    SetPixel m_hOpbSrcDC, 3, 0, 15191996
    SetPixel m_hOpbSrcDC, 4, 0, 14267285
    SetPixel m_hOpbSrcDC, 5, 0, 13871238
    SetPixel m_hOpbSrcDC, 6, 0, 13805187
    SetPixel m_hOpbSrcDC, 7, 0, 13871238
    SetPixel m_hOpbSrcDC, 8, 0, 14267285
    SetPixel m_hOpbSrcDC, 9, 0, 15191996
    SetPixel m_hOpbSrcDC, 10, 0, 16711935
    SetPixel m_hOpbSrcDC, 11, 0, 16711935
    SetPixel m_hOpbSrcDC, 12, 0, 16711935
    SetPixel m_hOpbSrcDC, 0, 1, 16711935
    SetPixel m_hOpbSrcDC, 1, 1, 16711935
    SetPixel m_hOpbSrcDC, 2, 1, 14399645
    SetPixel m_hOpbSrcDC, 3, 1, 14664361
    SetPixel m_hOpbSrcDC, 4, 1, 15523793
    SetPixel m_hOpbSrcDC, 5, 1, 15787997
    SetPixel m_hOpbSrcDC, 6, 1, 15854047
    SetPixel m_hOpbSrcDC, 7, 1, 15787997
    SetPixel m_hOpbSrcDC, 8, 1, 15523793
    SetPixel m_hOpbSrcDC, 9, 1, 14664361
    SetPixel m_hOpbSrcDC, 10, 1, 14399645
    SetPixel m_hOpbSrcDC, 11, 1, 16711935
    SetPixel m_hOpbSrcDC, 12, 1, 16711935
    SetPixel m_hOpbSrcDC, 0, 2, 16711935
    SetPixel m_hOpbSrcDC, 1, 2, 14399645
    SetPixel m_hOpbSrcDC, 2, 2, 15060665
    SetPixel m_hOpbSrcDC, 3, 2, 15985891
    SetPixel m_hOpbSrcDC, 4, 2, 16051685
    SetPixel m_hOpbSrcDC, 5, 2, 16051943
    SetPixel m_hOpbSrcDC, 6, 2, 16051943
    SetPixel m_hOpbSrcDC, 7, 2, 16051943
    SetPixel m_hOpbSrcDC, 8, 2, 16051685
    SetPixel m_hOpbSrcDC, 9, 2, 15985891
    SetPixel m_hOpbSrcDC, 10, 2, 15060665
    SetPixel m_hOpbSrcDC, 11, 2, 14399645
    SetPixel m_hOpbSrcDC, 12, 2, 16711935
    SetPixel m_hOpbSrcDC, 0, 3, 15191996
    SetPixel m_hOpbSrcDC, 1, 3, 14664361
    SetPixel m_hOpbSrcDC, 2, 3, 16051685
    SetPixel m_hOpbSrcDC, 3, 3, 16249581
    SetPixel m_hOpbSrcDC, 4, 3, 16447475
    SetPixel m_hOpbSrcDC, 5, 3, 16513527
    SetPixel m_hOpbSrcDC, 6, 3, 16513527
    SetPixel m_hOpbSrcDC, 7, 3, 16513527
    SetPixel m_hOpbSrcDC, 8, 3, 16447475
    SetPixel m_hOpbSrcDC, 9, 3, 16249581
    SetPixel m_hOpbSrcDC, 10, 3, 16051685
    SetPixel m_hOpbSrcDC, 11, 3, 14664361
    SetPixel m_hOpbSrcDC, 12, 3, 15191996
    SetPixel m_hOpbSrcDC, 0, 4, 14267285
    SetPixel m_hOpbSrcDC, 1, 4, 15721174
    SetPixel m_hOpbSrcDC, 2, 4, 16315631
    SetPixel m_hOpbSrcDC, 3, 4, 16579577
    SetPixel m_hOpbSrcDC, 4, 4, 16777215
    SetPixel m_hOpbSrcDC, 5, 4, 16777215
    SetPixel m_hOpbSrcDC, 6, 4, 16777215
    SetPixel m_hOpbSrcDC, 7, 4, 16777215
    SetPixel m_hOpbSrcDC, 8, 4, 16777215
    SetPixel m_hOpbSrcDC, 9, 4, 16579577
    SetPixel m_hOpbSrcDC, 10, 4, 16315631
    SetPixel m_hOpbSrcDC, 11, 4, 15721174
    SetPixel m_hOpbSrcDC, 12, 4, 14267285
    SetPixel m_hOpbSrcDC, 0, 5, 13871238
    SetPixel m_hOpbSrcDC, 1, 5, 16183531
    SetPixel m_hOpbSrcDC, 2, 5, 16579577
    SetPixel m_hOpbSrcDC, 3, 5, 16776958
    SetPixel m_hOpbSrcDC, 4, 5, 16777215
    SetPixel m_hOpbSrcDC, 5, 5, 16777215
    SetPixel m_hOpbSrcDC, 6, 5, 16777215
    SetPixel m_hOpbSrcDC, 7, 5, 16777215
    SetPixel m_hOpbSrcDC, 8, 5, 16777215
    SetPixel m_hOpbSrcDC, 9, 5, 16776958
    SetPixel m_hOpbSrcDC, 10, 5, 16579577
    SetPixel m_hOpbSrcDC, 11, 5, 16183531
    SetPixel m_hOpbSrcDC, 12, 5, 13871238
    SetPixel m_hOpbSrcDC, 0, 6, 13805187
    SetPixel m_hOpbSrcDC, 1, 6, 16447475
    SetPixel m_hOpbSrcDC, 2, 6, 16711421
    SetPixel m_hOpbSrcDC, 3, 6, 16777215
    SetPixel m_hOpbSrcDC, 4, 6, 16777215
    SetPixel m_hOpbSrcDC, 5, 6, 16777215
    SetPixel m_hOpbSrcDC, 6, 6, 16777215
    SetPixel m_hOpbSrcDC, 7, 6, 16777215
    SetPixel m_hOpbSrcDC, 8, 6, 16777215
    SetPixel m_hOpbSrcDC, 9, 6, 16777215
    SetPixel m_hOpbSrcDC, 10, 6, 16711421
    SetPixel m_hOpbSrcDC, 11, 6, 16447475
    SetPixel m_hOpbSrcDC, 12, 6, 13805187
    SetPixel m_hOpbSrcDC, 0, 7, 13871238
    SetPixel m_hOpbSrcDC, 1, 7, 16447476
    SetPixel m_hOpbSrcDC, 2, 7, 16777215
    SetPixel m_hOpbSrcDC, 3, 7, 16777215
    SetPixel m_hOpbSrcDC, 4, 7, 16777215
    SetPixel m_hOpbSrcDC, 5, 7, 16777215
    SetPixel m_hOpbSrcDC, 6, 7, 16777215
    SetPixel m_hOpbSrcDC, 7, 7, 16777215
    SetPixel m_hOpbSrcDC, 8, 7, 16777215
    SetPixel m_hOpbSrcDC, 9, 7, 16777215
    SetPixel m_hOpbSrcDC, 10, 7, 16777215
    SetPixel m_hOpbSrcDC, 11, 7, 16447476
    SetPixel m_hOpbSrcDC, 12, 7, 13871238
    SetPixel m_hOpbSrcDC, 0, 8, 14267285
    SetPixel m_hOpbSrcDC, 1, 8, 16117222
    SetPixel m_hOpbSrcDC, 2, 8, 16777215
    SetPixel m_hOpbSrcDC, 3, 8, 16777215
    SetPixel m_hOpbSrcDC, 4, 8, 16777215
    SetPixel m_hOpbSrcDC, 5, 8, 16777215
    SetPixel m_hOpbSrcDC, 6, 8, 16777215
    SetPixel m_hOpbSrcDC, 7, 8, 16777215
    SetPixel m_hOpbSrcDC, 8, 8, 16777215
    SetPixel m_hOpbSrcDC, 9, 8, 16777215
    SetPixel m_hOpbSrcDC, 10, 8, 16777215
    SetPixel m_hOpbSrcDC, 11, 8, 16117222
    SetPixel m_hOpbSrcDC, 12, 8, 14267285
    SetPixel m_hOpbSrcDC, 0, 9, 15191996
    SetPixel m_hOpbSrcDC, 1, 9, 14928051
    SetPixel m_hOpbSrcDC, 2, 9, 16777215
    SetPixel m_hOpbSrcDC, 3, 9, 16777215
    SetPixel m_hOpbSrcDC, 4, 9, 16777215
    SetPixel m_hOpbSrcDC, 5, 9, 16777215
    SetPixel m_hOpbSrcDC, 6, 9, 16777215
    SetPixel m_hOpbSrcDC, 7, 9, 16777215
    SetPixel m_hOpbSrcDC, 8, 9, 16777215
    SetPixel m_hOpbSrcDC, 9, 9, 16777215
    SetPixel m_hOpbSrcDC, 10, 9, 16777215
    SetPixel m_hOpbSrcDC, 11, 9, 14928051
    SetPixel m_hOpbSrcDC, 12, 9, 15191996
    SetPixel m_hOpbSrcDC, 0, 10, 16711935
    SetPixel m_hOpbSrcDC, 1, 10, 14399645
    SetPixel m_hOpbSrcDC, 2, 10, 15522250
    SetPixel m_hOpbSrcDC, 3, 10, 16777215
    SetPixel m_hOpbSrcDC, 4, 10, 16777215
    SetPixel m_hOpbSrcDC, 5, 10, 16777215
    SetPixel m_hOpbSrcDC, 6, 10, 16777215
    SetPixel m_hOpbSrcDC, 7, 10, 16777215
    SetPixel m_hOpbSrcDC, 8, 10, 16777215
    SetPixel m_hOpbSrcDC, 9, 10, 16777215
    SetPixel m_hOpbSrcDC, 10, 10, 15522250
    SetPixel m_hOpbSrcDC, 11, 10, 14399645
    SetPixel m_hOpbSrcDC, 12, 10, 16711935
    SetPixel m_hOpbSrcDC, 0, 11, 16711935
    SetPixel m_hOpbSrcDC, 1, 11, 16711935
    SetPixel m_hOpbSrcDC, 2, 11, 14399645
    SetPixel m_hOpbSrcDC, 3, 11, 15059638
    SetPixel m_hOpbSrcDC, 4, 11, 16315117
    SetPixel m_hOpbSrcDC, 5, 11, 16711164
    SetPixel m_hOpbSrcDC, 6, 11, 16777215
    SetPixel m_hOpbSrcDC, 7, 11, 16711164
    SetPixel m_hOpbSrcDC, 8, 11, 16315117
    SetPixel m_hOpbSrcDC, 9, 11, 15059638
    SetPixel m_hOpbSrcDC, 10, 11, 14399645
    SetPixel m_hOpbSrcDC, 11, 11, 16711935
    SetPixel m_hOpbSrcDC, 12, 11, 16711935
    SetPixel m_hOpbSrcDC, 0, 12, 16711935
    SetPixel m_hOpbSrcDC, 1, 12, 16711935
    SetPixel m_hOpbSrcDC, 2, 12, 16711935
    SetPixel m_hOpbSrcDC, 3, 12, 15191996
    SetPixel m_hOpbSrcDC, 4, 12, 14267285
    SetPixel m_hOpbSrcDC, 5, 12, 13871238
    SetPixel m_hOpbSrcDC, 6, 12, 13805187
    SetPixel m_hOpbSrcDC, 7, 12, 13871238
    SetPixel m_hOpbSrcDC, 8, 12, 14267285
    SetPixel m_hOpbSrcDC, 9, 12, 15191996
    SetPixel m_hOpbSrcDC, 10, 12, 16711935
    SetPixel m_hOpbSrcDC, 11, 12, 16711935
    SetPixel m_hOpbSrcDC, 12, 12, 16711935
    SetPixel m_hOpbSrcDC, 0, 13, 16711935
    SetPixel m_hOpbSrcDC, 1, 13, 16711935
    SetPixel m_hOpbSrcDC, 2, 13, 16711935
    SetPixel m_hOpbSrcDC, 3, 13, 15191996
    SetPixel m_hOpbSrcDC, 4, 13, 14267285
    SetPixel m_hOpbSrcDC, 5, 13, 13871238
    SetPixel m_hOpbSrcDC, 6, 13, 13805187
    SetPixel m_hOpbSrcDC, 7, 13, 13871238
    SetPixel m_hOpbSrcDC, 8, 13, 14267285
    SetPixel m_hOpbSrcDC, 9, 13, 15191996
    SetPixel m_hOpbSrcDC, 10, 13, 16711935
    SetPixel m_hOpbSrcDC, 11, 13, 16711935
    SetPixel m_hOpbSrcDC, 12, 13, 16711935
    SetPixel m_hOpbSrcDC, 0, 14, 16711935
    SetPixel m_hOpbSrcDC, 1, 14, 16711935
    SetPixel m_hOpbSrcDC, 2, 14, 14399645
    SetPixel m_hOpbSrcDC, 3, 14, 14664361
    SetPixel m_hOpbSrcDC, 4, 14, 16313027
    SetPixel m_hOpbSrcDC, 5, 14, 16313027
    SetPixel m_hOpbSrcDC, 6, 14, 16313027
    SetPixel m_hOpbSrcDC, 7, 14, 16313284
    SetPixel m_hOpbSrcDC, 8, 14, 16313285
    SetPixel m_hOpbSrcDC, 9, 14, 14664361
    SetPixel m_hOpbSrcDC, 10, 14, 14399645
    SetPixel m_hOpbSrcDC, 11, 14, 16711935
    SetPixel m_hOpbSrcDC, 12, 14, 16711935
    SetPixel m_hOpbSrcDC, 0, 15, 16711935
    SetPixel m_hOpbSrcDC, 1, 15, 14399645
    SetPixel m_hOpbSrcDC, 2, 15, 15060665
    SetPixel m_hOpbSrcDC, 3, 15, 16313027
    SetPixel m_hOpbSrcDC, 4, 15, 16313027
    SetPixel m_hOpbSrcDC, 5, 15, 16313027
    SetPixel m_hOpbSrcDC, 6, 15, 16313284
    SetPixel m_hOpbSrcDC, 7, 15, 16378821
    SetPixel m_hOpbSrcDC, 8, 15, 16379080
    SetPixel m_hOpbSrcDC, 9, 15, 16379338
    SetPixel m_hOpbSrcDC, 10, 15, 15060665
    SetPixel m_hOpbSrcDC, 11, 15, 14399645
    SetPixel m_hOpbSrcDC, 12, 15, 16711935
    SetPixel m_hOpbSrcDC, 0, 16, 15191996
    SetPixel m_hOpbSrcDC, 1, 16, 14664361
    SetPixel m_hOpbSrcDC, 2, 16, 16313027
    SetPixel m_hOpbSrcDC, 3, 16, 16313027
    SetPixel m_hOpbSrcDC, 4, 16, 16313027
    SetPixel m_hOpbSrcDC, 5, 16, 16313028
    SetPixel m_hOpbSrcDC, 6, 16, 16313286
    SetPixel m_hOpbSrcDC, 7, 16, 16379080
    SetPixel m_hOpbSrcDC, 8, 16, 16379339
    SetPixel m_hOpbSrcDC, 9, 16, 16445133
    SetPixel m_hOpbSrcDC, 10, 16, 16445391
    SetPixel m_hOpbSrcDC, 11, 16, 14664361
    SetPixel m_hOpbSrcDC, 12, 16, 15191996
    SetPixel m_hOpbSrcDC, 0, 17, 14267285
    SetPixel m_hOpbSrcDC, 1, 17, 16313027
    SetPixel m_hOpbSrcDC, 2, 17, 16313027
    SetPixel m_hOpbSrcDC, 3, 17, 16313028
    SetPixel m_hOpbSrcDC, 4, 17, 16313028
    SetPixel m_hOpbSrcDC, 5, 17, 16379078
    SetPixel m_hOpbSrcDC, 6, 17, 16379081
    SetPixel m_hOpbSrcDC, 7, 17, 16379339
    SetPixel m_hOpbSrcDC, 8, 17, 16445133
    SetPixel m_hOpbSrcDC, 9, 17, 16445392
    SetPixel m_hOpbSrcDC, 10, 17, 16511187
    SetPixel m_hOpbSrcDC, 11, 17, 16445910
    SetPixel m_hOpbSrcDC, 12, 17, 14267285
    SetPixel m_hOpbSrcDC, 0, 18, 13871238
    SetPixel m_hOpbSrcDC, 1, 18, 16313027
    SetPixel m_hOpbSrcDC, 2, 18, 16313027
    SetPixel m_hOpbSrcDC, 3, 18, 16313285
    SetPixel m_hOpbSrcDC, 4, 18, 16313543
    SetPixel m_hOpbSrcDC, 5, 18, 16379337
    SetPixel m_hOpbSrcDC, 6, 18, 16379339
    SetPixel m_hOpbSrcDC, 7, 18, 16445134
    SetPixel m_hOpbSrcDC, 8, 18, 16445648
    SetPixel m_hOpbSrcDC, 9, 18, 16511187
    SetPixel m_hOpbSrcDC, 10, 18, 16511446
    SetPixel m_hOpbSrcDC, 11, 18, 16511705
    SetPixel m_hOpbSrcDC, 12, 18, 13871238
    SetPixel m_hOpbSrcDC, 0, 19, 13805187
    SetPixel m_hOpbSrcDC, 1, 19, 16313028
    SetPixel m_hOpbSrcDC, 2, 19, 16313285
    SetPixel m_hOpbSrcDC, 3, 19, 16379079
    SetPixel m_hOpbSrcDC, 4, 19, 16379337
    SetPixel m_hOpbSrcDC, 5, 19, 16445132
    SetPixel m_hOpbSrcDC, 6, 19, 16445135
    SetPixel m_hOpbSrcDC, 7, 19, 16445648
    SetPixel m_hOpbSrcDC, 8, 19, 16445908
    SetPixel m_hOpbSrcDC, 9, 19, 16511446
    SetPixel m_hOpbSrcDC, 10, 19, 16511705
    SetPixel m_hOpbSrcDC, 11, 19, 16577755
    SetPixel m_hOpbSrcDC, 12, 19, 13805187
    SetPixel m_hOpbSrcDC, 0, 20, 13871238
    SetPixel m_hOpbSrcDC, 1, 20, 16313285
    SetPixel m_hOpbSrcDC, 2, 20, 16313287
    SetPixel m_hOpbSrcDC, 3, 20, 16379081
    SetPixel m_hOpbSrcDC, 4, 20, 16445132
    SetPixel m_hOpbSrcDC, 5, 20, 16445390
    SetPixel m_hOpbSrcDC, 6, 20, 16511185
    SetPixel m_hOpbSrcDC, 7, 20, 16511188
    SetPixel m_hOpbSrcDC, 8, 20, 16576983
    SetPixel m_hOpbSrcDC, 9, 20, 16577497
    SetPixel m_hOpbSrcDC, 10, 20, 16577500
    SetPixel m_hOpbSrcDC, 11, 20, 16577758
    SetPixel m_hOpbSrcDC, 12, 20, 13871238
    SetPixel m_hOpbSrcDC, 0, 21, 14267285
    SetPixel m_hOpbSrcDC, 1, 21, 16379079
    SetPixel m_hOpbSrcDC, 2, 21, 16379338
    SetPixel m_hOpbSrcDC, 3, 21, 16379596
    SetPixel m_hOpbSrcDC, 4, 21, 16445391
    SetPixel m_hOpbSrcDC, 5, 21, 16445394
    SetPixel m_hOpbSrcDC, 6, 21, 16511444
    SetPixel m_hOpbSrcDC, 7, 21, 16511704
    SetPixel m_hOpbSrcDC, 8, 21, 16511962
    SetPixel m_hOpbSrcDC, 9, 21, 16577756
    SetPixel m_hOpbSrcDC, 10, 21, 16577758
    SetPixel m_hOpbSrcDC, 11, 21, 16643552
    SetPixel m_hOpbSrcDC, 12, 21, 14267285
    SetPixel m_hOpbSrcDC, 0, 22, 15191996
    SetPixel m_hOpbSrcDC, 1, 22, 14928051
    SetPixel m_hOpbSrcDC, 2, 22, 16445133
    SetPixel m_hOpbSrcDC, 3, 22, 16445392
    SetPixel m_hOpbSrcDC, 4, 22, 16511186
    SetPixel m_hOpbSrcDC, 5, 22, 16511445
    SetPixel m_hOpbSrcDC, 6, 22, 16511704
    SetPixel m_hOpbSrcDC, 7, 22, 16577499
    SetPixel m_hOpbSrcDC, 8, 22, 16577756
    SetPixel m_hOpbSrcDC, 9, 22, 16643551
    SetPixel m_hOpbSrcDC, 10, 22, 16578016
    SetPixel m_hOpbSrcDC, 11, 22, 14928051
    SetPixel m_hOpbSrcDC, 12, 22, 15191996
    SetPixel m_hOpbSrcDC, 0, 23, 16711935
    SetPixel m_hOpbSrcDC, 1, 23, 14399645
    SetPixel m_hOpbSrcDC, 2, 23, 15522250
    SetPixel m_hOpbSrcDC, 3, 23, 16511187
    SetPixel m_hOpbSrcDC, 4, 23, 16511446
    SetPixel m_hOpbSrcDC, 5, 23, 16511704
    SetPixel m_hOpbSrcDC, 6, 23, 16577755
    SetPixel m_hOpbSrcDC, 7, 23, 16643293
    SetPixel m_hOpbSrcDC, 8, 23, 16643551
    SetPixel m_hOpbSrcDC, 9, 23, 16643553
    SetPixel m_hOpbSrcDC, 10, 23, 15522250
    SetPixel m_hOpbSrcDC, 11, 23, 14399645
    SetPixel m_hOpbSrcDC, 12, 23, 16711935
    SetPixel m_hOpbSrcDC, 0, 24, 16711935
    SetPixel m_hOpbSrcDC, 1, 24, 16711935
    SetPixel m_hOpbSrcDC, 2, 24, 14399645
    SetPixel m_hOpbSrcDC, 3, 24, 15059638
    SetPixel m_hOpbSrcDC, 4, 24, 16577497
    SetPixel m_hOpbSrcDC, 5, 24, 16577499
    SetPixel m_hOpbSrcDC, 6, 24, 16577757
    SetPixel m_hOpbSrcDC, 7, 24, 16578015
    SetPixel m_hOpbSrcDC, 8, 24, 16578017
    SetPixel m_hOpbSrcDC, 9, 24, 15059638
    SetPixel m_hOpbSrcDC, 10, 24, 14399645
    SetPixel m_hOpbSrcDC, 11, 24, 16711935
    SetPixel m_hOpbSrcDC, 12, 24, 16711935
    SetPixel m_hOpbSrcDC, 0, 25, 16711935
    SetPixel m_hOpbSrcDC, 1, 25, 16711935
    SetPixel m_hOpbSrcDC, 2, 25, 16711935
    SetPixel m_hOpbSrcDC, 3, 25, 15191996
    SetPixel m_hOpbSrcDC, 4, 25, 14267285
    SetPixel m_hOpbSrcDC, 5, 25, 13871238
    SetPixel m_hOpbSrcDC, 6, 25, 13805187
    SetPixel m_hOpbSrcDC, 7, 25, 13871238
    SetPixel m_hOpbSrcDC, 8, 25, 14267285
    SetPixel m_hOpbSrcDC, 9, 25, 15191996
    SetPixel m_hOpbSrcDC, 10, 25, 16711935
    SetPixel m_hOpbSrcDC, 11, 25, 16711935
    SetPixel m_hOpbSrcDC, 12, 25, 16711935
    SetPixel m_hOpbSrcDC, 0, 26, 16711935
    SetPixel m_hOpbSrcDC, 1, 26, 16711935
    SetPixel m_hOpbSrcDC, 2, 26, 16711935
    SetPixel m_hOpbSrcDC, 3, 26, 14531748
    SetPixel m_hOpbSrcDC, 4, 26, 13277558
    SetPixel m_hOpbSrcDC, 5, 26, 12749671
    SetPixel m_hOpbSrcDC, 6, 26, 12683620
    SetPixel m_hOpbSrcDC, 7, 26, 12749671
    SetPixel m_hOpbSrcDC, 8, 26, 13277558
    SetPixel m_hOpbSrcDC, 9, 26, 14531748
    SetPixel m_hOpbSrcDC, 10, 26, 16711935
    SetPixel m_hOpbSrcDC, 11, 26, 16711935
    SetPixel m_hOpbSrcDC, 12, 26, 16711935
    SetPixel m_hOpbSrcDC, 0, 27, 16711935
    SetPixel m_hOpbSrcDC, 1, 27, 16711935
    SetPixel m_hOpbSrcDC, 2, 27, 13475711
    SetPixel m_hOpbSrcDC, 3, 27, 13872268
    SetPixel m_hOpbSrcDC, 4, 27, 16313542
    SetPixel m_hOpbSrcDC, 5, 27, 16379338
    SetPixel m_hOpbSrcDC, 6, 27, 16445389
    SetPixel m_hOpbSrcDC, 7, 27, 16445393
    SetPixel m_hOpbSrcDC, 8, 27, 16642516
    SetPixel m_hOpbSrcDC, 9, 27, 14201753
    SetPixel m_hOpbSrcDC, 10, 27, 13475711
    SetPixel m_hOpbSrcDC, 11, 27, 16711935
    SetPixel m_hOpbSrcDC, 12, 27, 16711935
    SetPixel m_hOpbSrcDC, 0, 28, 16711935
    SetPixel m_hOpbSrcDC, 1, 28, 13475711
    SetPixel m_hOpbSrcDC, 2, 28, 14334880
    SetPixel m_hOpbSrcDC, 3, 28, 16312508
    SetPixel m_hOpbSrcDC, 4, 28, 16378562
    SetPixel m_hOpbSrcDC, 5, 28, 16379078
    SetPixel m_hOpbSrcDC, 6, 28, 16379338
    SetPixel m_hOpbSrcDC, 7, 28, 16445134
    SetPixel m_hOpbSrcDC, 8, 28, 16445393
    SetPixel m_hOpbSrcDC, 9, 28, 16445908
    SetPixel m_hOpbSrcDC, 10, 28, 15059637
    SetPixel m_hOpbSrcDC, 11, 28, 13475711
    SetPixel m_hOpbSrcDC, 12, 28, 16711935
    SetPixel m_hOpbSrcDC, 0, 29, 14531748
    SetPixel m_hOpbSrcDC, 1, 29, 13872268
    SetPixel m_hOpbSrcDC, 2, 29, 16246453
    SetPixel m_hOpbSrcDC, 3, 29, 16312250
    SetPixel m_hOpbSrcDC, 4, 29, 16312510
    SetPixel m_hOpbSrcDC, 5, 29, 16378562
    SetPixel m_hOpbSrcDC, 6, 29, 16379078
    SetPixel m_hOpbSrcDC, 7, 29, 16445130
    SetPixel m_hOpbSrcDC, 8, 29, 16445390
    SetPixel m_hOpbSrcDC, 9, 29, 16642515
    SetPixel m_hOpbSrcDC, 10, 29, 16642517
    SetPixel m_hOpbSrcDC, 11, 29, 14333340
    SetPixel m_hOpbSrcDC, 12, 29, 14531748
    SetPixel m_hOpbSrcDC, 0, 30, 13277558
    SetPixel m_hOpbSrcDC, 1, 30, 16114352
    SetPixel m_hOpbSrcDC, 2, 30, 16246195
    SetPixel m_hOpbSrcDC, 3, 30, 16246456
    SetPixel m_hOpbSrcDC, 4, 30, 16312250
    SetPixel m_hOpbSrcDC, 5, 30, 16313022
    SetPixel m_hOpbSrcDC, 6, 30, 16313539
    SetPixel m_hOpbSrcDC, 7, 30, 16444615
    SetPixel m_hOpbSrcDC, 8, 30, 16379596
    SetPixel m_hOpbSrcDC, 9, 30, 16445390
    SetPixel m_hOpbSrcDC, 10, 30, 16642515
    SetPixel m_hOpbSrcDC, 11, 30, 16445909
    SetPixel m_hOpbSrcDC, 12, 30, 13277558
    SetPixel m_hOpbSrcDC, 0, 31, 12749671
    SetPixel m_hOpbSrcDC, 1, 31, 16114350
    SetPixel m_hOpbSrcDC, 2, 31, 16245424
    SetPixel m_hOpbSrcDC, 3, 31, 16246195
    SetPixel m_hOpbSrcDC, 4, 31, 16246456
    SetPixel m_hOpbSrcDC, 5, 31, 16312251
    SetPixel m_hOpbSrcDC, 6, 31, 16313022
    SetPixel m_hOpbSrcDC, 7, 31, 16378563
    SetPixel m_hOpbSrcDC, 8, 31, 16379336
    SetPixel m_hOpbSrcDC, 9, 31, 16445133
    SetPixel m_hOpbSrcDC, 10, 31, 16641999
    SetPixel m_hOpbSrcDC, 11, 31, 16445907
    SetPixel m_hOpbSrcDC, 12, 31, 12749671
    SetPixel m_hOpbSrcDC, 0, 32, 12683620
    SetPixel m_hOpbSrcDC, 1, 32, 16114093
    SetPixel m_hOpbSrcDC, 2, 32, 16114350
    SetPixel m_hOpbSrcDC, 3, 32, 16114353
    SetPixel m_hOpbSrcDC, 4, 32, 16246196
    SetPixel m_hOpbSrcDC, 5, 32, 16246456
    SetPixel m_hOpbSrcDC, 6, 32, 16312252
    SetPixel m_hOpbSrcDC, 7, 32, 16378559
    SetPixel m_hOpbSrcDC, 8, 32, 16379075
    SetPixel m_hOpbSrcDC, 9, 32, 16379336
    SetPixel m_hOpbSrcDC, 10, 32, 16445389
    SetPixel m_hOpbSrcDC, 11, 32, 16445391
    SetPixel m_hOpbSrcDC, 12, 32, 12683620
    SetPixel m_hOpbSrcDC, 0, 33, 12749671
    SetPixel m_hOpbSrcDC, 1, 33, 16114093
    SetPixel m_hOpbSrcDC, 2, 33, 16114093
    SetPixel m_hOpbSrcDC, 3, 33, 16114094
    SetPixel m_hOpbSrcDC, 4, 33, 16246193
    SetPixel m_hOpbSrcDC, 5, 33, 16246452
    SetPixel m_hOpbSrcDC, 6, 33, 16312249
    SetPixel m_hOpbSrcDC, 7, 33, 16312507
    SetPixel m_hOpbSrcDC, 8, 33, 16312513
    SetPixel m_hOpbSrcDC, 9, 33, 16379077
    SetPixel m_hOpbSrcDC, 10, 33, 16379336
    SetPixel m_hOpbSrcDC, 11, 33, 16445133
    SetPixel m_hOpbSrcDC, 12, 33, 12749671
    SetPixel m_hOpbSrcDC, 0, 34, 13277558
    SetPixel m_hOpbSrcDC, 1, 34, 16114093
    SetPixel m_hOpbSrcDC, 2, 34, 16114093
    SetPixel m_hOpbSrcDC, 3, 34, 16114093
    SetPixel m_hOpbSrcDC, 4, 34, 16114094
    SetPixel m_hOpbSrcDC, 5, 34, 16115122
    SetPixel m_hOpbSrcDC, 6, 34, 16246452
    SetPixel m_hOpbSrcDC, 7, 34, 16312249
    SetPixel m_hOpbSrcDC, 8, 34, 16312508
    SetPixel m_hOpbSrcDC, 9, 34, 16378561
    SetPixel m_hOpbSrcDC, 10, 34, 16379078
    SetPixel m_hOpbSrcDC, 11, 34, 16445130
    SetPixel m_hOpbSrcDC, 12, 34, 13277558
    SetPixel m_hOpbSrcDC, 0, 35, 14531748
    SetPixel m_hOpbSrcDC, 1, 35, 13872268
    SetPixel m_hOpbSrcDC, 2, 35, 16114093
    SetPixel m_hOpbSrcDC, 3, 35, 16114093
    SetPixel m_hOpbSrcDC, 4, 35, 16114094
    SetPixel m_hOpbSrcDC, 5, 35, 16114352
    SetPixel m_hOpbSrcDC, 6, 35, 16246194
    SetPixel m_hOpbSrcDC, 7, 35, 16246196
    SetPixel m_hOpbSrcDC, 8, 35, 16246713
    SetPixel m_hOpbSrcDC, 9, 35, 16312510
    SetPixel m_hOpbSrcDC, 10, 35, 16378562
    SetPixel m_hOpbSrcDC, 11, 35, 14333340
    SetPixel m_hOpbSrcDC, 12, 35, 14531748
    SetPixel m_hOpbSrcDC, 0, 36, 16711935
    SetPixel m_hOpbSrcDC, 1, 36, 13475711
    SetPixel m_hOpbSrcDC, 2, 36, 14334880
    SetPixel m_hOpbSrcDC, 3, 36, 16114093
    SetPixel m_hOpbSrcDC, 4, 36, 16114093
    SetPixel m_hOpbSrcDC, 5, 36, 16114093
    SetPixel m_hOpbSrcDC, 6, 36, 16114352
    SetPixel m_hOpbSrcDC, 7, 36, 16114354
    SetPixel m_hOpbSrcDC, 8, 36, 16246453
    SetPixel m_hOpbSrcDC, 9, 36, 16312250
    SetPixel m_hOpbSrcDC, 10, 36, 15059637
    SetPixel m_hOpbSrcDC, 11, 36, 13475711
    SetPixel m_hOpbSrcDC, 12, 36, 16711935
    SetPixel m_hOpbSrcDC, 0, 37, 16711935
    SetPixel m_hOpbSrcDC, 1, 37, 16711935
    SetPixel m_hOpbSrcDC, 2, 37, 13475711
    SetPixel m_hOpbSrcDC, 3, 37, 13872268
    SetPixel m_hOpbSrcDC, 4, 37, 16114093
    SetPixel m_hOpbSrcDC, 5, 37, 16114093
    SetPixel m_hOpbSrcDC, 6, 37, 16114094
    SetPixel m_hOpbSrcDC, 7, 37, 16114352
    SetPixel m_hOpbSrcDC, 8, 37, 16246194
    SetPixel m_hOpbSrcDC, 9, 37, 14201753
    SetPixel m_hOpbSrcDC, 10, 37, 13475711
    SetPixel m_hOpbSrcDC, 11, 37, 16711935
    SetPixel m_hOpbSrcDC, 12, 37, 16711935
    SetPixel m_hOpbSrcDC, 0, 38, 16711935
    SetPixel m_hOpbSrcDC, 1, 38, 16711935
    SetPixel m_hOpbSrcDC, 2, 38, 16711935
    SetPixel m_hOpbSrcDC, 3, 38, 14531748
    SetPixel m_hOpbSrcDC, 4, 38, 13277558
    SetPixel m_hOpbSrcDC, 5, 38, 12749671
    SetPixel m_hOpbSrcDC, 6, 38, 12683620
    SetPixel m_hOpbSrcDC, 7, 38, 12749671
    SetPixel m_hOpbSrcDC, 8, 38, 13277558
    SetPixel m_hOpbSrcDC, 9, 38, 14531748
    SetPixel m_hOpbSrcDC, 10, 38, 16711935
    SetPixel m_hOpbSrcDC, 11, 38, 16711935
    SetPixel m_hOpbSrcDC, 12, 38, 16711935
    SetPixel m_hOpbSrcDC, 0, 39, 16711935
    SetPixel m_hOpbSrcDC, 1, 39, 16711935
    SetPixel m_hOpbSrcDC, 2, 39, 16711935
    SetPixel m_hOpbSrcDC, 3, 39, 15984605
    SetPixel m_hOpbSrcDC, 4, 39, 15522250
    SetPixel m_hOpbSrcDC, 5, 39, 15324098
    SetPixel m_hOpbSrcDC, 6, 39, 15258305
    SetPixel m_hOpbSrcDC, 7, 39, 15324098
    SetPixel m_hOpbSrcDC, 8, 39, 15522250
    SetPixel m_hOpbSrcDC, 9, 39, 15984605
    SetPixel m_hOpbSrcDC, 10, 39, 16711935
    SetPixel m_hOpbSrcDC, 11, 39, 16711935
    SetPixel m_hOpbSrcDC, 12, 39, 16711935
    SetPixel m_hOpbSrcDC, 0, 40, 16711935
    SetPixel m_hOpbSrcDC, 1, 40, 16711935
    SetPixel m_hOpbSrcDC, 2, 40, 15588302
    SetPixel m_hOpbSrcDC, 3, 40, 15720660
    SetPixel m_hOpbSrcDC, 4, 40, 16117736
    SetPixel m_hOpbSrcDC, 5, 40, 16249838
    SetPixel m_hOpbSrcDC, 6, 40, 16315631
    SetPixel m_hOpbSrcDC, 7, 40, 16249838
    SetPixel m_hOpbSrcDC, 8, 40, 16117736
    SetPixel m_hOpbSrcDC, 9, 40, 15720660
    SetPixel m_hOpbSrcDC, 10, 40, 15588302
    SetPixel m_hOpbSrcDC, 11, 40, 16711935
    SetPixel m_hOpbSrcDC, 12, 40, 16711935
    SetPixel m_hOpbSrcDC, 0, 41, 16711935
    SetPixel m_hOpbSrcDC, 1, 41, 15588302
    SetPixel m_hOpbSrcDC, 2, 41, 15918812
    SetPixel m_hOpbSrcDC, 3, 41, 16381425
    SetPixel m_hOpbSrcDC, 4, 41, 16381682
    SetPixel m_hOpbSrcDC, 5, 41, 16381683
    SetPixel m_hOpbSrcDC, 6, 41, 16381683
    SetPixel m_hOpbSrcDC, 7, 41, 16381683
    SetPixel m_hOpbSrcDC, 8, 41, 16381682
    SetPixel m_hOpbSrcDC, 9, 41, 16381425
    SetPixel m_hOpbSrcDC, 10, 41, 15918812
    SetPixel m_hOpbSrcDC, 11, 41, 15588302
    SetPixel m_hOpbSrcDC, 12, 41, 16711935
    SetPixel m_hOpbSrcDC, 0, 42, 15984605
    SetPixel m_hOpbSrcDC, 1, 42, 15720660
    SetPixel m_hOpbSrcDC, 2, 42, 16381682
    SetPixel m_hOpbSrcDC, 3, 42, 16513270
    SetPixel m_hOpbSrcDC, 4, 42, 16579577
    SetPixel m_hOpbSrcDC, 5, 42, 16645371
    SetPixel m_hOpbSrcDC, 6, 42, 16645371
    SetPixel m_hOpbSrcDC, 7, 42, 16645371
    SetPixel m_hOpbSrcDC, 8, 42, 16579577
    SetPixel m_hOpbSrcDC, 9, 42, 16513270
    SetPixel m_hOpbSrcDC, 10, 42, 16381682
    SetPixel m_hOpbSrcDC, 11, 42, 15720660
    SetPixel m_hOpbSrcDC, 12, 42, 15984605
    SetPixel m_hOpbSrcDC, 0, 43, 15522250
    SetPixel m_hOpbSrcDC, 1, 43, 16249066
    SetPixel m_hOpbSrcDC, 2, 43, 16513527
    SetPixel m_hOpbSrcDC, 3, 43, 16645628
    SetPixel m_hOpbSrcDC, 4, 43, 16711422
    SetPixel m_hOpbSrcDC, 5, 43, 16777215
    SetPixel m_hOpbSrcDC, 6, 43, 16777215
    SetPixel m_hOpbSrcDC, 7, 43, 16777215
    SetPixel m_hOpbSrcDC, 8, 43, 16711422
    SetPixel m_hOpbSrcDC, 9, 43, 16645628
    SetPixel m_hOpbSrcDC, 10, 43, 16513527
    SetPixel m_hOpbSrcDC, 11, 43, 16249066
    SetPixel m_hOpbSrcDC, 12, 43, 15522250
    SetPixel m_hOpbSrcDC, 0, 44, 15324098
    SetPixel m_hOpbSrcDC, 1, 44, 16447477
    SetPixel m_hOpbSrcDC, 2, 44, 16645628
    SetPixel m_hOpbSrcDC, 3, 44, 16776958
    SetPixel m_hOpbSrcDC, 4, 44, 16777215
    SetPixel m_hOpbSrcDC, 5, 44, 16777215
    SetPixel m_hOpbSrcDC, 6, 44, 16777215
    SetPixel m_hOpbSrcDC, 7, 44, 16777215
    SetPixel m_hOpbSrcDC, 8, 44, 16777215
    SetPixel m_hOpbSrcDC, 9, 44, 16776958
    SetPixel m_hOpbSrcDC, 10, 44, 16645628
    SetPixel m_hOpbSrcDC, 11, 44, 16447477
    SetPixel m_hOpbSrcDC, 12, 44, 15324098
    SetPixel m_hOpbSrcDC, 0, 45, 15258305
    SetPixel m_hOpbSrcDC, 1, 45, 16579577
    SetPixel m_hOpbSrcDC, 2, 45, 16711422
    SetPixel m_hOpbSrcDC, 3, 45, 16777215
    SetPixel m_hOpbSrcDC, 4, 45, 16777215
    SetPixel m_hOpbSrcDC, 5, 45, 16777215
    SetPixel m_hOpbSrcDC, 6, 45, 16777215
    SetPixel m_hOpbSrcDC, 7, 45, 16777215
    SetPixel m_hOpbSrcDC, 8, 45, 16777215
    SetPixel m_hOpbSrcDC, 9, 45, 16777215
    SetPixel m_hOpbSrcDC, 10, 45, 16711422
    SetPixel m_hOpbSrcDC, 11, 45, 16579577
    SetPixel m_hOpbSrcDC, 12, 45, 15258305
    SetPixel m_hOpbSrcDC, 0, 46, 15324098
    SetPixel m_hOpbSrcDC, 1, 46, 16579577
    SetPixel m_hOpbSrcDC, 2, 46, 16777215
    SetPixel m_hOpbSrcDC, 3, 46, 16777215
    SetPixel m_hOpbSrcDC, 4, 46, 16777215
    SetPixel m_hOpbSrcDC, 5, 46, 16777215
    SetPixel m_hOpbSrcDC, 6, 46, 16777215
    SetPixel m_hOpbSrcDC, 7, 46, 16777215
    SetPixel m_hOpbSrcDC, 8, 46, 16777215
    SetPixel m_hOpbSrcDC, 9, 46, 16777215
    SetPixel m_hOpbSrcDC, 10, 46, 16777215
    SetPixel m_hOpbSrcDC, 11, 46, 16579577
    SetPixel m_hOpbSrcDC, 12, 46, 15324098
    SetPixel m_hOpbSrcDC, 0, 47, 15522250
    SetPixel m_hOpbSrcDC, 1, 47, 16447218
    SetPixel m_hOpbSrcDC, 2, 47, 16777215
    SetPixel m_hOpbSrcDC, 3, 47, 16777215
    SetPixel m_hOpbSrcDC, 4, 47, 16777215
    SetPixel m_hOpbSrcDC, 5, 47, 16777215
    SetPixel m_hOpbSrcDC, 6, 47, 16777215
    SetPixel m_hOpbSrcDC, 7, 47, 16777215
    SetPixel m_hOpbSrcDC, 8, 47, 16777215
    SetPixel m_hOpbSrcDC, 9, 47, 16777215
    SetPixel m_hOpbSrcDC, 10, 47, 16777215
    SetPixel m_hOpbSrcDC, 11, 47, 16447218
    SetPixel m_hOpbSrcDC, 12, 47, 15522250
    SetPixel m_hOpbSrcDC, 0, 48, 15984605
    SetPixel m_hOpbSrcDC, 1, 48, 15852505
    SetPixel m_hOpbSrcDC, 2, 48, 16777215
    SetPixel m_hOpbSrcDC, 3, 48, 16777215
    SetPixel m_hOpbSrcDC, 4, 48, 16777215
    SetPixel m_hOpbSrcDC, 5, 48, 16777215
    SetPixel m_hOpbSrcDC, 6, 48, 16777215
    SetPixel m_hOpbSrcDC, 7, 48, 16777215
    SetPixel m_hOpbSrcDC, 8, 48, 16777215
    SetPixel m_hOpbSrcDC, 9, 48, 16777215
    SetPixel m_hOpbSrcDC, 10, 48, 16777215
    SetPixel m_hOpbSrcDC, 11, 48, 15852505
    SetPixel m_hOpbSrcDC, 12, 48, 15984605
    SetPixel m_hOpbSrcDC, 0, 49, 16711935
    SetPixel m_hOpbSrcDC, 1, 49, 15588302
    SetPixel m_hOpbSrcDC, 2, 49, 16116964
    SetPixel m_hOpbSrcDC, 3, 49, 16777215
    SetPixel m_hOpbSrcDC, 4, 49, 16777215
    SetPixel m_hOpbSrcDC, 5, 49, 16777215
    SetPixel m_hOpbSrcDC, 6, 49, 16777215
    SetPixel m_hOpbSrcDC, 7, 49, 16777215
    SetPixel m_hOpbSrcDC, 8, 49, 16777215
    SetPixel m_hOpbSrcDC, 9, 49, 16777215
    SetPixel m_hOpbSrcDC, 10, 49, 16116964
    SetPixel m_hOpbSrcDC, 11, 49, 15588302
    SetPixel m_hOpbSrcDC, 12, 49, 16711935
    SetPixel m_hOpbSrcDC, 0, 50, 16711935
    SetPixel m_hOpbSrcDC, 1, 50, 16711935
    SetPixel m_hOpbSrcDC, 2, 50, 15588302
    SetPixel m_hOpbSrcDC, 3, 50, 15918298
    SetPixel m_hOpbSrcDC, 4, 50, 16513270
    SetPixel m_hOpbSrcDC, 5, 50, 16711421
    SetPixel m_hOpbSrcDC, 6, 50, 16777215
    SetPixel m_hOpbSrcDC, 7, 50, 16711421
    SetPixel m_hOpbSrcDC, 8, 50, 16513270
    SetPixel m_hOpbSrcDC, 9, 50, 15918298
    SetPixel m_hOpbSrcDC, 10, 50, 15588302
    SetPixel m_hOpbSrcDC, 11, 50, 16711935
    SetPixel m_hOpbSrcDC, 12, 50, 16711935
    SetPixel m_hOpbSrcDC, 0, 51, 16711935
    SetPixel m_hOpbSrcDC, 1, 51, 16711935
    SetPixel m_hOpbSrcDC, 2, 51, 16711935
    SetPixel m_hOpbSrcDC, 3, 51, 15984605
    SetPixel m_hOpbSrcDC, 4, 51, 15522250
    SetPixel m_hOpbSrcDC, 5, 51, 15324098
    SetPixel m_hOpbSrcDC, 6, 51, 15258305
    SetPixel m_hOpbSrcDC, 7, 51, 15324098
    SetPixel m_hOpbSrcDC, 8, 51, 15522250
    SetPixel m_hOpbSrcDC, 9, 51, 15984605
    SetPixel m_hOpbSrcDC, 10, 51, 16711935
    SetPixel m_hOpbSrcDC, 11, 51, 16711935
    SetPixel m_hOpbSrcDC, 12, 51, 16711935
    
    SetPixel m_hHdbSrcDC, 0, 0, 16709098
    SetPixel m_hHdbSrcDC, 1, 0, 16709098
    SetPixel m_hHdbSrcDC, 2, 0, 16709098
    SetPixel m_hHdbSrcDC, 3, 0, 16639677
    SetPixel m_hHdbSrcDC, 4, 0, 16707283
    SetPixel m_hHdbSrcDC, 5, 0, 16639677
    SetPixel m_hHdbSrcDC, 6, 0, 16504727
    SetPixel m_hHdbSrcDC, 7, 0, 16504727
    SetPixel m_hHdbSrcDC, 8, 0, 16504727
    SetPixel m_hHdbSrcDC, 0, 1, 16708841
    SetPixel m_hHdbSrcDC, 1, 1, 16708841
    SetPixel m_hHdbSrcDC, 2, 1, 16708841
    SetPixel m_hHdbSrcDC, 3, 1, 16572071
    SetPixel m_hHdbSrcDC, 4, 1, 16773078
    SetPixel m_hHdbSrcDC, 5, 1, 16572071
    SetPixel m_hHdbSrcDC, 6, 1, 16436862
    SetPixel m_hHdbSrcDC, 7, 1, 16768159
    SetPixel m_hHdbSrcDC, 8, 1, 16436862
    SetPixel m_hHdbSrcDC, 0, 2, 16643305
    SetPixel m_hHdbSrcDC, 1, 2, 16643048
    SetPixel m_hHdbSrcDC, 2, 2, 16445408
    SetPixel m_hHdbSrcDC, 3, 2, 16569487
    SetPixel m_hHdbSrcDC, 4, 2, 16772819
    SetPixel m_hHdbSrcDC, 5, 2, 16372107
    SetPixel m_hHdbSrcDC, 6, 2, 16434534
    SetPixel m_hHdbSrcDC, 7, 2, 16768416
    SetPixel m_hHdbSrcDC, 8, 2, 16237412
    SetPixel m_hHdbSrcDC, 0, 3, 16643049
    SetPixel m_hHdbSrcDC, 1, 3, 16642790
    SetPixel m_hHdbSrcDC, 2, 3, 16246998
    SetPixel m_hHdbSrcDC, 3, 3, 16501623
    SetPixel m_hHdbSrcDC, 4, 3, 16772560
    SetPixel m_hHdbSrcDC, 5, 3, 16106865
    SetPixel m_hHdbSrcDC, 6, 3, 16433758
    SetPixel m_hHdbSrcDC, 7, 3, 16768674
    SetPixel m_hHdbSrcDC, 8, 3, 16039259
    SetPixel m_hHdbSrcDC, 0, 4, 16643563
    SetPixel m_hHdbSrcDC, 1, 4, 16576996
    SetPixel m_hHdbSrcDC, 2, 4, 15718844
    SetPixel m_hHdbSrcDC, 3, 4, 16434533
    SetPixel m_hHdbSrcDC, 4, 4, 16772300
    SetPixel m_hHdbSrcDC, 5, 4, 15579227
    SetPixel m_hHdbSrcDC, 6, 4, 16433758
    SetPixel m_hHdbSrcDC, 7, 4, 16768932
    SetPixel m_hHdbSrcDC, 8, 4, 15578454
    SetPixel m_hHdbSrcDC, 0, 5, 16643564
    SetPixel m_hHdbSrcDC, 1, 5, 16576738
    SetPixel m_hHdbSrcDC, 2, 5, 15454638
    SetPixel m_hHdbSrcDC, 3, 5, 16433758
    SetPixel m_hHdbSrcDC, 4, 5, 16706504
    SetPixel m_hHdbSrcDC, 5, 5, 15380820
    SetPixel m_hHdbSrcDC, 6, 5, 16433758
    SetPixel m_hHdbSrcDC, 7, 5, 16769190
    SetPixel m_hHdbSrcDC, 8, 5, 15380820
    SetPixel m_hHdbSrcDC, 0, 6, 16643821
    SetPixel m_hHdbSrcDC, 1, 6, 16510944
    SetPixel m_hHdbSrcDC, 2, 6, 15124383
    SetPixel m_hHdbSrcDC, 3, 6, 16433758
    SetPixel m_hHdbSrcDC, 4, 6, 16705988
    SetPixel m_hHdbSrcDC, 5, 6, 15051857
    SetPixel m_hHdbSrcDC, 6, 6, 16433758
    SetPixel m_hHdbSrcDC, 7, 6, 16769704
    SetPixel m_hHdbSrcDC, 8, 6, 15051857
    SetPixel m_hHdbSrcDC, 0, 7, 16643823
    SetPixel m_hHdbSrcDC, 1, 7, 16510431
    SetPixel m_hHdbSrcDC, 2, 7, 14860180
    SetPixel m_hHdbSrcDC, 3, 7, 16433758
    SetPixel m_hHdbSrcDC, 4, 7, 16705729
    SetPixel m_hHdbSrcDC, 5, 7, 14788687
    SetPixel m_hHdbSrcDC, 6, 7, 16433758
    SetPixel m_hHdbSrcDC, 7, 7, 16769963
    SetPixel m_hHdbSrcDC, 8, 7, 14788687
    SetPixel m_hHdbSrcDC, 0, 8, 16644081
    SetPixel m_hHdbSrcDC, 1, 8, 16444637
    SetPixel m_hHdbSrcDC, 2, 8, 14464386
    SetPixel m_hHdbSrcDC, 3, 8, 16433758
    SetPixel m_hHdbSrcDC, 4, 8, 16705469
    SetPixel m_hHdbSrcDC, 5, 8, 14459724
    SetPixel m_hHdbSrcDC, 6, 8, 16433758
    SetPixel m_hHdbSrcDC, 7, 8, 16770477
    SetPixel m_hHdbSrcDC, 8, 8, 14459724
    SetPixel m_hHdbSrcDC, 0, 9, 16512493
    SetPixel m_hHdbSrcDC, 1, 9, 16180691
    SetPixel m_hHdbSrcDC, 2, 9, 14266490
    SetPixel m_hHdbSrcDC, 3, 9, 16433758
    SetPixel m_hHdbSrcDC, 4, 9, 16638383
    SetPixel m_hHdbSrcDC, 5, 9, 14393931
    SetPixel m_hHdbSrcDC, 6, 9, 16433758
    SetPixel m_hHdbSrcDC, 7, 9, 16770992
    SetPixel m_hHdbSrcDC, 8, 9, 14393931
    SetPixel m_hHdbSrcDC, 0, 10, 16512237
    SetPixel m_hHdbSrcDC, 1, 10, 16180691
    SetPixel m_hHdbSrcDC, 2, 10, 14266490
    SetPixel m_hHdbSrcDC, 3, 10, 16433758
    SetPixel m_hHdbSrcDC, 4, 10, 16638126
    SetPixel m_hHdbSrcDC, 5, 10, 14393931
    SetPixel m_hHdbSrcDC, 6, 10, 16433758
    SetPixel m_hHdbSrcDC, 7, 10, 16771251
    SetPixel m_hHdbSrcDC, 8, 10, 14393931
    SetPixel m_hHdbSrcDC, 0, 11, 16511979
    SetPixel m_hHdbSrcDC, 1, 11, 16180691
    SetPixel m_hHdbSrcDC, 2, 11, 14332541
    SetPixel m_hHdbSrcDC, 3, 11, 16433758
    SetPixel m_hHdbSrcDC, 4, 11, 16638382
    SetPixel m_hHdbSrcDC, 5, 11, 14459724
    SetPixel m_hHdbSrcDC, 6, 11, 16433758
    SetPixel m_hHdbSrcDC, 7, 11, 16771766
    SetPixel m_hHdbSrcDC, 8, 11, 14459724
    SetPixel m_hHdbSrcDC, 0, 12, 16445669
    SetPixel m_hHdbSrcDC, 1, 12, 16180691
    SetPixel m_hHdbSrcDC, 2, 12, 14662284
    SetPixel m_hHdbSrcDC, 3, 12, 16433758
    SetPixel m_hHdbSrcDC, 4, 12, 16638383
    SetPixel m_hHdbSrcDC, 5, 12, 14788686
    SetPixel m_hHdbSrcDC, 6, 12, 16433758
    SetPixel m_hHdbSrcDC, 7, 12, 16772281
    SetPixel m_hHdbSrcDC, 8, 12, 14788686
    SetPixel m_hHdbSrcDC, 0, 13, 16379618
    SetPixel m_hHdbSrcDC, 1, 13, 16180691
    SetPixel m_hHdbSrcDC, 2, 13, 14926231
    SetPixel m_hHdbSrcDC, 3, 13, 16433758
    SetPixel m_hHdbSrcDC, 4, 13, 16638383
    SetPixel m_hHdbSrcDC, 5, 13, 15051857
    SetPixel m_hHdbSrcDC, 6, 13, 16434533
    SetPixel m_hHdbSrcDC, 7, 13, 16772540
    SetPixel m_hHdbSrcDC, 8, 13, 15052117
    SetPixel m_hHdbSrcDC, 0, 14, 16313566
    SetPixel m_hHdbSrcDC, 1, 14, 16180691
    SetPixel m_hHdbSrcDC, 2, 14, 15124643
    SetPixel m_hHdbSrcDC, 3, 14, 16433758
    SetPixel m_hHdbSrcDC, 4, 14, 16638383
    SetPixel m_hHdbSrcDC, 5, 14, 15315284
    SetPixel m_hHdbSrcDC, 6, 14, 16501359
    SetPixel m_hHdbSrcDC, 7, 14, 16773054
    SetPixel m_hHdbSrcDC, 8, 14, 15316575
    SetPixel m_hHdbSrcDC, 0, 15, 16312794
    SetPixel m_hHdbSrcDC, 1, 15, 16180691
    SetPixel m_hHdbSrcDC, 2, 15, 15388591
    SetPixel m_hHdbSrcDC, 3, 15, 16433758
    SetPixel m_hHdbSrcDC, 4, 15, 16638383
    SetPixel m_hHdbSrcDC, 5, 15, 15578454
    SetPixel m_hHdbSrcDC, 6, 15, 16502908
    SetPixel m_hHdbSrcDC, 7, 15, 16773829
    SetPixel m_hHdbSrcDC, 8, 15, 15646573
    SetPixel m_hHdbSrcDC, 0, 16, 16180949
    SetPixel m_hHdbSrcDC, 1, 16, 16180691
    SetPixel m_hHdbSrcDC, 2, 16, 15850692
    SetPixel m_hHdbSrcDC, 3, 16, 16433758
    SetPixel m_hHdbSrcDC, 4, 16, 16638383
    SetPixel m_hHdbSrcDC, 5, 16, 16039258
    SetPixel m_hHdbSrcDC, 6, 16, 16569995
    SetPixel m_hHdbSrcDC, 7, 16, 16773832
    SetPixel m_hHdbSrcDC, 8, 16, 16174979
    SetPixel m_hHdbSrcDC, 0, 17, 16180691
    SetPixel m_hHdbSrcDC, 1, 17, 16180691
    SetPixel m_hHdbSrcDC, 2, 17, 15983052
    SetPixel m_hHdbSrcDC, 3, 17, 16433758
    SetPixel m_hHdbSrcDC, 4, 17, 16638124
    SetPixel m_hHdbSrcDC, 5, 17, 16236636
    SetPixel m_hHdbSrcDC, 6, 17, 16637337
    SetPixel m_hHdbSrcDC, 7, 17, 16773832
    SetPixel m_hHdbSrcDC, 8, 17, 16439700
    SetPixel m_hHdbSrcDC, 0, 18, 16180691
    SetPixel m_hHdbSrcDC, 1, 18, 16180691
    SetPixel m_hHdbSrcDC, 2, 18, 16180691
    SetPixel m_hHdbSrcDC, 3, 18, 16501875
    SetPixel m_hHdbSrcDC, 4, 18, 16571816
    SetPixel m_hHdbSrcDC, 5, 18, 16501875
    SetPixel m_hHdbSrcDC, 6, 18, 16638888
    SetPixel m_hHdbSrcDC, 7, 18, 16773832
    SetPixel m_hHdbSrcDC, 8, 18, 16638888
    SetPixel m_hHdbSrcDC, 0, 19, 15915712
    SetPixel m_hHdbSrcDC, 1, 19, 15915712
    SetPixel m_hHdbSrcDC, 2, 19, 15915712
    SetPixel m_hHdbSrcDC, 3, 19, 15915712
    SetPixel m_hHdbSrcDC, 4, 19, 15915712
    SetPixel m_hHdbSrcDC, 5, 19, 15915712
    SetPixel m_hHdbSrcDC, 6, 19, 15915712
    SetPixel m_hHdbSrcDC, 7, 19, 15915712
    SetPixel m_hHdbSrcDC, 8, 19, 15915712
End Sub
Private Sub pInit2()
    m_hBtnSrcDC = pCreateDC(7, 84)
    m_hCbbSrcDC = pCreateDC(15, 18)
    m_hCkbSrcDC = pCreateDC(18, 24)
    m_hOpbSrcDC = pCreateDC(13, 52)
    m_hHdbSrcDC = pCreateDC(9, 20)
    
    SetPixel m_hBtnSrcDC, 0, 0, 12575225
    SetPixel m_hBtnSrcDC, 1, 0, 7058403
    SetPixel m_hBtnSrcDC, 2, 0, 2922456
    SetPixel m_hBtnSrcDC, 3, 0, 1806293
    SetPixel m_hBtnSrcDC, 4, 0, 2922456
    SetPixel m_hBtnSrcDC, 5, 0, 7058403
    SetPixel m_hBtnSrcDC, 6, 0, 12575225
    SetPixel m_hBtnSrcDC, 0, 1, 6927074
    SetPixel m_hBtnSrcDC, 1, 1, 8306919
    SetPixel m_hBtnSrcDC, 2, 1, 15201275
    SetPixel m_hBtnSrcDC, 3, 1, 16777215
    SetPixel m_hBtnSrcDC, 4, 1, 15201275
    SetPixel m_hBtnSrcDC, 5, 1, 8306919
    SetPixel m_hBtnSrcDC, 6, 1, 6927074
    SetPixel m_hBtnSrcDC, 0, 2, 2725335
    SetPixel m_hBtnSrcDC, 1, 2, 15529724
    SetPixel m_hBtnSrcDC, 2, 2, 16777215
    SetPixel m_hBtnSrcDC, 3, 2, 16777215
    SetPixel m_hBtnSrcDC, 4, 2, 16777215
    SetPixel m_hBtnSrcDC, 5, 2, 15529724
    SetPixel m_hBtnSrcDC, 6, 2, 2725335
    SetPixel m_hBtnSrcDC, 0, 3, 1806293
    SetPixel m_hBtnSrcDC, 1, 3, 16120573
    SetPixel m_hBtnSrcDC, 2, 3, 15857916
    SetPixel m_hBtnSrcDC, 3, 3, 15857916
    SetPixel m_hBtnSrcDC, 4, 3, 15857916
    SetPixel m_hBtnSrcDC, 5, 3, 16120573
    SetPixel m_hBtnSrcDC, 6, 3, 1806293
    SetPixel m_hBtnSrcDC, 0, 4, 1806293
    SetPixel m_hBtnSrcDC, 1, 4, 15989244
    SetPixel m_hBtnSrcDC, 2, 4, 15660795
    SetPixel m_hBtnSrcDC, 3, 4, 15660795
    SetPixel m_hBtnSrcDC, 4, 4, 15660795
    SetPixel m_hBtnSrcDC, 5, 4, 15989244
    SetPixel m_hBtnSrcDC, 6, 4, 1806293
    SetPixel m_hBtnSrcDC, 0, 5, 1806293
    SetPixel m_hBtnSrcDC, 1, 5, 15792124
    SetPixel m_hBtnSrcDC, 2, 5, 15398139
    SetPixel m_hBtnSrcDC, 3, 5, 15398139
    SetPixel m_hBtnSrcDC, 4, 5, 15398139
    SetPixel m_hBtnSrcDC, 5, 5, 15792124
    SetPixel m_hBtnSrcDC, 6, 5, 1806293
    SetPixel m_hBtnSrcDC, 0, 6, 1806293
    SetPixel m_hBtnSrcDC, 1, 6, 15660796
    SetPixel m_hBtnSrcDC, 2, 6, 15135483
    SetPixel m_hBtnSrcDC, 3, 6, 15135483
    SetPixel m_hBtnSrcDC, 4, 6, 15135483
    SetPixel m_hBtnSrcDC, 5, 6, 15660796
    SetPixel m_hBtnSrcDC, 6, 6, 1806293
    SetPixel m_hBtnSrcDC, 0, 7, 1806293
    SetPixel m_hBtnSrcDC, 1, 7, 15463932
    SetPixel m_hBtnSrcDC, 2, 7, 14872826
    SetPixel m_hBtnSrcDC, 3, 7, 14872826
    SetPixel m_hBtnSrcDC, 4, 7, 14872826
    SetPixel m_hBtnSrcDC, 5, 7, 15463932
    SetPixel m_hBtnSrcDC, 6, 7, 1806293
    SetPixel m_hBtnSrcDC, 0, 8, 1806293
    SetPixel m_hBtnSrcDC, 1, 8, 15266812
    SetPixel m_hBtnSrcDC, 2, 8, 14610170
    SetPixel m_hBtnSrcDC, 3, 8, 14610170
    SetPixel m_hBtnSrcDC, 4, 8, 14610170
    SetPixel m_hBtnSrcDC, 5, 8, 15266812
    SetPixel m_hBtnSrcDC, 6, 8, 1806293
    SetPixel m_hBtnSrcDC, 0, 9, 1806293
    SetPixel m_hBtnSrcDC, 1, 9, 15069947
    SetPixel m_hBtnSrcDC, 2, 9, 14347513
    SetPixel m_hBtnSrcDC, 3, 9, 14347513
    SetPixel m_hBtnSrcDC, 4, 9, 14347513
    SetPixel m_hBtnSrcDC, 5, 9, 15069947
    SetPixel m_hBtnSrcDC, 6, 9, 1806293
    SetPixel m_hBtnSrcDC, 0, 10, 1806293
    SetPixel m_hBtnSrcDC, 1, 10, 14938363
    SetPixel m_hBtnSrcDC, 2, 10, 14150393
    SetPixel m_hBtnSrcDC, 3, 10, 14150393
    SetPixel m_hBtnSrcDC, 4, 10, 14150393
    SetPixel m_hBtnSrcDC, 5, 10, 14938363
    SetPixel m_hBtnSrcDC, 6, 10, 1806293
    SetPixel m_hBtnSrcDC, 0, 11, 1806293
    SetPixel m_hBtnSrcDC, 1, 11, 14412792
    SetPixel m_hBtnSrcDC, 2, 11, 13427445
    SetPixel m_hBtnSrcDC, 3, 11, 13427445
    SetPixel m_hBtnSrcDC, 4, 11, 13427445
    SetPixel m_hBtnSrcDC, 5, 11, 14412792
    SetPixel m_hBtnSrcDC, 6, 11, 1806293
    SetPixel m_hBtnSrcDC, 0, 12, 1806293
    SetPixel m_hBtnSrcDC, 1, 12, 14281207
    SetPixel m_hBtnSrcDC, 2, 12, 13164531
    SetPixel m_hBtnSrcDC, 3, 12, 13164531
    SetPixel m_hBtnSrcDC, 4, 12, 13164531
    SetPixel m_hBtnSrcDC, 5, 12, 14281207
    SetPixel m_hBtnSrcDC, 6, 12, 1806293
    SetPixel m_hBtnSrcDC, 0, 13, 1806293
    SetPixel m_hBtnSrcDC, 1, 13, 14018549
    SetPixel m_hBtnSrcDC, 2, 13, 12836081
    SetPixel m_hBtnSrcDC, 3, 13, 12836081
    SetPixel m_hBtnSrcDC, 4, 13, 12836081
    SetPixel m_hBtnSrcDC, 5, 13, 14018549
    SetPixel m_hBtnSrcDC, 6, 13, 1806293
    SetPixel m_hBtnSrcDC, 0, 14, 1806293
    SetPixel m_hBtnSrcDC, 1, 14, 13821172
    SetPixel m_hBtnSrcDC, 2, 14, 12507375
    SetPixel m_hBtnSrcDC, 3, 14, 12507375
    SetPixel m_hBtnSrcDC, 4, 14, 12507375
    SetPixel m_hBtnSrcDC, 5, 14, 13821172
    SetPixel m_hBtnSrcDC, 6, 14, 1806293
    SetPixel m_hBtnSrcDC, 0, 15, 1806293
    SetPixel m_hBtnSrcDC, 1, 15, 13492978
    SetPixel m_hBtnSrcDC, 2, 15, 12113389
    SetPixel m_hBtnSrcDC, 3, 15, 12113389
    SetPixel m_hBtnSrcDC, 4, 15, 12113389
    SetPixel m_hBtnSrcDC, 5, 15, 13492978
    SetPixel m_hBtnSrcDC, 6, 15, 1806293
    SetPixel m_hBtnSrcDC, 0, 16, 1806293
    SetPixel m_hBtnSrcDC, 1, 16, 13361393
    SetPixel m_hBtnSrcDC, 2, 16, 11850475
    SetPixel m_hBtnSrcDC, 3, 16, 11850475
    SetPixel m_hBtnSrcDC, 4, 16, 11850475
    SetPixel m_hBtnSrcDC, 5, 16, 13361393
    SetPixel m_hBtnSrcDC, 6, 16, 1806293
    SetPixel m_hBtnSrcDC, 0, 17, 1806293
    SetPixel m_hBtnSrcDC, 1, 17, 13229808
    SetPixel m_hBtnSrcDC, 2, 17, 11653354
    SetPixel m_hBtnSrcDC, 3, 17, 11653354
    SetPixel m_hBtnSrcDC, 4, 17, 11653354
    SetPixel m_hBtnSrcDC, 5, 17, 13229808
    SetPixel m_hBtnSrcDC, 6, 17, 1806293
    SetPixel m_hBtnSrcDC, 0, 18, 2659799
    SetPixel m_hBtnSrcDC, 1, 18, 12244974
    SetPixel m_hBtnSrcDC, 2, 18, 12310252
    SetPixel m_hBtnSrcDC, 3, 18, 11587561
    SetPixel m_hBtnSrcDC, 4, 18, 12309996
    SetPixel m_hBtnSrcDC, 5, 18, 12244974
    SetPixel m_hBtnSrcDC, 6, 18, 2659799
    SetPixel m_hBtnSrcDC, 0, 19, 6598882
    SetPixel m_hBtnSrcDC, 1, 19, 6861537
    SetPixel m_hBtnSrcDC, 2, 19, 12310766
    SetPixel m_hBtnSrcDC, 3, 19, 13426929
    SetPixel m_hBtnSrcDC, 4, 19, 12310766
    SetPixel m_hBtnSrcDC, 5, 19, 6861537
    SetPixel m_hBtnSrcDC, 6, 19, 6467553
    SetPixel m_hBtnSrcDC, 0, 20, 13297659
    SetPixel m_hBtnSrcDC, 1, 20, 6336225
    SetPixel m_hBtnSrcDC, 2, 20, 2791128
    SetPixel m_hBtnSrcDC, 3, 20, 1806293
    SetPixel m_hBtnSrcDC, 4, 20, 2791128
    SetPixel m_hBtnSrcDC, 5, 20, 6336225
    SetPixel m_hBtnSrcDC, 6, 20, 13297659
    SetPixel m_hBtnSrcDC, 0, 21, 12575225
    SetPixel m_hBtnSrcDC, 1, 21, 7058403
    SetPixel m_hBtnSrcDC, 2, 21, 2922456
    SetPixel m_hBtnSrcDC, 3, 21, 1806293
    SetPixel m_hBtnSrcDC, 4, 21, 2922456
    SetPixel m_hBtnSrcDC, 5, 21, 7058403
    SetPixel m_hBtnSrcDC, 6, 21, 12575225
    SetPixel m_hBtnSrcDC, 0, 22, 6927074
    SetPixel m_hBtnSrcDC, 1, 22, 2663398
    SetPixel m_hBtnSrcDC, 2, 22, 3980280
    SetPixel m_hBtnSrcDC, 3, 22, 5030907
    SetPixel m_hBtnSrcDC, 4, 22, 3980280
    SetPixel m_hBtnSrcDC, 5, 22, 2663398
    SetPixel m_hBtnSrcDC, 6, 22, 6927074
    SetPixel m_hBtnSrcDC, 0, 23, 2725335
    SetPixel m_hBtnSrcDC, 1, 23, 4046073
    SetPixel m_hBtnSrcDC, 2, 23, 9296638
    SetPixel m_hBtnSrcDC, 3, 23, 13037055
    SetPixel m_hBtnSrcDC, 4, 23, 9296638
    SetPixel m_hBtnSrcDC, 5, 23, 4046073
    SetPixel m_hBtnSrcDC, 6, 23, 2725335
    SetPixel m_hBtnSrcDC, 0, 24, 1806293
    SetPixel m_hBtnSrcDC, 1, 24, 4637179
    SetPixel m_hBtnSrcDC, 2, 24, 11461629
    SetPixel m_hBtnSrcDC, 3, 24, 13823740
    SetPixel m_hBtnSrcDC, 4, 24, 11592958
    SetPixel m_hBtnSrcDC, 5, 24, 4637179
    SetPixel m_hBtnSrcDC, 6, 24, 1806293
    SetPixel m_hBtnSrcDC, 0, 25, 1806293
    SetPixel m_hBtnSrcDC, 1, 25, 4899579
    SetPixel m_hBtnSrcDC, 2, 25, 12249085
    SetPixel m_hBtnSrcDC, 3, 25, 14676732
    SetPixel m_hBtnSrcDC, 4, 25, 12314621
    SetPixel m_hBtnSrcDC, 5, 25, 4899579
    SetPixel m_hBtnSrcDC, 6, 25, 1806293
    SetPixel m_hBtnSrcDC, 0, 26, 1806293
    SetPixel m_hBtnSrcDC, 1, 26, 5030907
    SetPixel m_hBtnSrcDC, 2, 26, 12642557
    SetPixel m_hBtnSrcDC, 3, 26, 15004412
    SetPixel m_hBtnSrcDC, 4, 26, 12642557
    SetPixel m_hBtnSrcDC, 5, 26, 5030907
    SetPixel m_hBtnSrcDC, 6, 26, 1806293
    SetPixel m_hBtnSrcDC, 0, 27, 1806293
    SetPixel m_hBtnSrcDC, 1, 27, 5030907
    SetPixel m_hBtnSrcDC, 2, 27, 12642301
    SetPixel m_hBtnSrcDC, 3, 27, 14873340
    SetPixel m_hBtnSrcDC, 4, 27, 12642301
    SetPixel m_hBtnSrcDC, 5, 27, 5030907
    SetPixel m_hBtnSrcDC, 6, 27, 1806293
    SetPixel m_hBtnSrcDC, 0, 28, 1806293
    SetPixel m_hBtnSrcDC, 1, 28, 5030907
    SetPixel m_hBtnSrcDC, 2, 28, 12576764
    SetPixel m_hBtnSrcDC, 3, 28, 14742011
    SetPixel m_hBtnSrcDC, 4, 28, 12576764
    SetPixel m_hBtnSrcDC, 5, 28, 5030907
    SetPixel m_hBtnSrcDC, 6, 28, 1806293
    SetPixel m_hBtnSrcDC, 0, 29, 1806293
    SetPixel m_hBtnSrcDC, 1, 29, 4965115
    SetPixel m_hBtnSrcDC, 2, 29, 12379900
    SetPixel m_hBtnSrcDC, 3, 29, 14545147
    SetPixel m_hBtnSrcDC, 4, 29, 12379900
    SetPixel m_hBtnSrcDC, 5, 29, 4965115
    SetPixel m_hBtnSrcDC, 6, 29, 1806293
    SetPixel m_hBtnSrcDC, 0, 30, 1806293
    SetPixel m_hBtnSrcDC, 1, 30, 4899579
    SetPixel m_hBtnSrcDC, 2, 30, 12248572
    SetPixel m_hBtnSrcDC, 3, 30, 14413563
    SetPixel m_hBtnSrcDC, 4, 30, 12248572
    SetPixel m_hBtnSrcDC, 5, 30, 4899579
    SetPixel m_hBtnSrcDC, 6, 30, 1806293
    SetPixel m_hBtnSrcDC, 0, 31, 1806293
    SetPixel m_hBtnSrcDC, 1, 31, 4899579
    SetPixel m_hBtnSrcDC, 2, 31, 12182780
    SetPixel m_hBtnSrcDC, 3, 31, 14282235
    SetPixel m_hBtnSrcDC, 4, 31, 12182780
    SetPixel m_hBtnSrcDC, 5, 31, 4899579
    SetPixel m_hBtnSrcDC, 6, 31, 1806293
    SetPixel m_hBtnSrcDC, 0, 32, 1806293
    SetPixel m_hBtnSrcDC, 1, 32, 4636666
    SetPixel m_hBtnSrcDC, 2, 32, 11460089
    SetPixel m_hBtnSrcDC, 3, 32, 13362423
    SetPixel m_hBtnSrcDC, 4, 32, 11460089
    SetPixel m_hBtnSrcDC, 5, 32, 4636666
    SetPixel m_hBtnSrcDC, 6, 32, 1806293
    SetPixel m_hBtnSrcDC, 0, 33, 1806293
    SetPixel m_hBtnSrcDC, 1, 33, 4571130
    SetPixel m_hBtnSrcDC, 2, 33, 11197175
    SetPixel m_hBtnSrcDC, 3, 33, 13099509
    SetPixel m_hBtnSrcDC, 4, 33, 11197175
    SetPixel m_hBtnSrcDC, 5, 33, 4571130
    SetPixel m_hBtnSrcDC, 6, 33, 1806293
    SetPixel m_hBtnSrcDC, 0, 34, 1806293
    SetPixel m_hBtnSrcDC, 1, 34, 4505337
    SetPixel m_hBtnSrcDC, 2, 34, 11000054
    SetPixel m_hBtnSrcDC, 3, 34, 12771059
    SetPixel m_hBtnSrcDC, 4, 34, 11000054
    SetPixel m_hBtnSrcDC, 5, 34, 4505337
    SetPixel m_hBtnSrcDC, 6, 34, 1806293
    SetPixel m_hBtnSrcDC, 0, 35, 1806293
    SetPixel m_hBtnSrcDC, 1, 35, 4374008
    SetPixel m_hBtnSrcDC, 2, 35, 10671860
    SetPixel m_hBtnSrcDC, 3, 35, 12377074
    SetPixel m_hBtnSrcDC, 4, 35, 10671860
    SetPixel m_hBtnSrcDC, 5, 35, 4374008
    SetPixel m_hBtnSrcDC, 6, 35, 1806293
    SetPixel m_hBtnSrcDC, 0, 36, 1806293
    SetPixel m_hBtnSrcDC, 1, 36, 4308216
    SetPixel m_hBtnSrcDC, 2, 36, 10343155
    SetPixel m_hBtnSrcDC, 3, 36, 12048369
    SetPixel m_hBtnSrcDC, 4, 36, 10343155
    SetPixel m_hBtnSrcDC, 5, 36, 4308216
    SetPixel m_hBtnSrcDC, 6, 36, 1806293
    SetPixel m_hBtnSrcDC, 0, 37, 1806293
    SetPixel m_hBtnSrcDC, 1, 37, 4045816
    SetPixel m_hBtnSrcDC, 2, 37, 9883891
    SetPixel m_hBtnSrcDC, 3, 37, 11523567
    SetPixel m_hBtnSrcDC, 4, 37, 9752819
    SetPixel m_hBtnSrcDC, 5, 37, 4111352
    SetPixel m_hBtnSrcDC, 6, 37, 1806293
    SetPixel m_hBtnSrcDC, 0, 38, 1806293
    SetPixel m_hBtnSrcDC, 1, 38, 3783416
    SetPixel m_hBtnSrcDC, 2, 38, 9425141
    SetPixel m_hBtnSrcDC, 3, 38, 10999026
    SetPixel m_hBtnSrcDC, 4, 38, 8965363
    SetPixel m_hBtnSrcDC, 5, 38, 3783416
    SetPixel m_hBtnSrcDC, 6, 38, 1806293
    SetPixel m_hBtnSrcDC, 0, 39, 2659799
    SetPixel m_hBtnSrcDC, 1, 39, 3389174
    SetPixel m_hBtnSrcDC, 2, 39, 7391735
    SetPixel m_hBtnSrcDC, 3, 39, 10081011
    SetPixel m_hBtnSrcDC, 4, 39, 7457527
    SetPixel m_hBtnSrcDC, 5, 39, 3520246
    SetPixel m_hBtnSrcDC, 6, 39, 2659799
    SetPixel m_hBtnSrcDC, 0, 40, 6598882
    SetPixel m_hBtnSrcDC, 1, 40, 2466277
    SetPixel m_hBtnSrcDC, 2, 40, 3388918
    SetPixel m_hBtnSrcDC, 3, 40, 4176888
    SetPixel m_hBtnSrcDC, 4, 40, 3389174
    SetPixel m_hBtnSrcDC, 5, 40, 2400741
    SetPixel m_hBtnSrcDC, 6, 40, 6467553
    SetPixel m_hBtnSrcDC, 0, 41, 13297659
    SetPixel m_hBtnSrcDC, 1, 41, 6336225
    SetPixel m_hBtnSrcDC, 2, 41, 2791128
    SetPixel m_hBtnSrcDC, 3, 41, 1806293
    SetPixel m_hBtnSrcDC, 4, 41, 2791128
    SetPixel m_hBtnSrcDC, 5, 41, 6336225
    SetPixel m_hBtnSrcDC, 6, 41, 13297659
    SetPixel m_hBtnSrcDC, 0, 42, 12575225
    SetPixel m_hBtnSrcDC, 1, 42, 7058403
    SetPixel m_hBtnSrcDC, 2, 42, 2922456
    SetPixel m_hBtnSrcDC, 3, 42, 1806293
    SetPixel m_hBtnSrcDC, 4, 42, 2922456
    SetPixel m_hBtnSrcDC, 5, 42, 7058403
    SetPixel m_hBtnSrcDC, 6, 42, 12575225
    SetPixel m_hBtnSrcDC, 0, 43, 6927074
    SetPixel m_hBtnSrcDC, 1, 43, 2269670
    SetPixel m_hBtnSrcDC, 2, 43, 3192312
    SetPixel m_hBtnSrcDC, 3, 43, 3783931
    SetPixel m_hBtnSrcDC, 4, 43, 3126775
    SetPixel m_hBtnSrcDC, 5, 43, 2335206
    SetPixel m_hBtnSrcDC, 6, 43, 6927074
    SetPixel m_hBtnSrcDC, 0, 44, 2725335
    SetPixel m_hBtnSrcDC, 1, 44, 3061496
    SetPixel m_hBtnSrcDC, 2, 44, 7326200
    SetPixel m_hBtnSrcDC, 3, 44, 8834291
    SetPixel m_hBtnSrcDC, 4, 44, 7194872
    SetPixel m_hBtnSrcDC, 5, 44, 2930167
    SetPixel m_hBtnSrcDC, 6, 44, 2725335
    SetPixel m_hBtnSrcDC, 0, 45, 1806293
    SetPixel m_hBtnSrcDC, 1, 45, 3258361
    SetPixel m_hBtnSrcDC, 2, 45, 8637170
    SetPixel m_hBtnSrcDC, 3, 45, 10474739
    SetPixel m_hBtnSrcDC, 4, 45, 9096948
    SetPixel m_hBtnSrcDC, 5, 45, 3258361
    SetPixel m_hBtnSrcDC, 6, 45, 1806293
    SetPixel m_hBtnSrcDC, 0, 46, 1806293
    SetPixel m_hBtnSrcDC, 1, 46, 3520761
    SetPixel m_hBtnSrcDC, 2, 46, 9293042
    SetPixel m_hBtnSrcDC, 3, 46, 11261939
    SetPixel m_hBtnSrcDC, 4, 46, 9424370
    SetPixel m_hBtnSrcDC, 5, 46, 3520761
    SetPixel m_hBtnSrcDC, 6, 46, 1806293
    SetPixel m_hBtnSrcDC, 0, 47, 1806293
    SetPixel m_hBtnSrcDC, 1, 47, 3652089
    SetPixel m_hBtnSrcDC, 2, 47, 9817841
    SetPixel m_hBtnSrcDC, 3, 47, 11786739
    SetPixel m_hBtnSrcDC, 4, 47, 9817842
    SetPixel m_hBtnSrcDC, 5, 47, 3652089
    SetPixel m_hBtnSrcDC, 6, 47, 1806293
    SetPixel m_hBtnSrcDC, 0, 48, 1806293
    SetPixel m_hBtnSrcDC, 1, 48, 3848697
    SetPixel m_hBtnSrcDC, 2, 48, 10146034
    SetPixel m_hBtnSrcDC, 3, 48, 12049395
    SetPixel m_hBtnSrcDC, 4, 48, 10146034
    SetPixel m_hBtnSrcDC, 5, 48, 3848697
    SetPixel m_hBtnSrcDC, 6, 48, 1806293
    SetPixel m_hBtnSrcDC, 0, 49, 1806293
    SetPixel m_hBtnSrcDC, 1, 49, 3914489
    SetPixel m_hBtnSrcDC, 2, 49, 10474483
    SetPixel m_hBtnSrcDC, 3, 49, 12377589
    SetPixel m_hBtnSrcDC, 4, 49, 10474483
    SetPixel m_hBtnSrcDC, 5, 49, 3914489
    SetPixel m_hBtnSrcDC, 6, 49, 1806293
    SetPixel m_hBtnSrcDC, 0, 50, 1806293
    SetPixel m_hBtnSrcDC, 1, 50, 3980282
    SetPixel m_hBtnSrcDC, 2, 50, 10737397
    SetPixel m_hBtnSrcDC, 3, 50, 12575222
    SetPixel m_hBtnSrcDC, 4, 50, 10737397
    SetPixel m_hBtnSrcDC, 5, 50, 3980282
    SetPixel m_hBtnSrcDC, 6, 50, 1806293
    SetPixel m_hBtnSrcDC, 0, 51, 1806293
    SetPixel m_hBtnSrcDC, 1, 51, 4046074
    SetPixel m_hBtnSrcDC, 2, 51, 11000054
    SetPixel m_hBtnSrcDC, 3, 51, 12837624
    SetPixel m_hBtnSrcDC, 4, 51, 11000054
    SetPixel m_hBtnSrcDC, 5, 51, 4046074
    SetPixel m_hBtnSrcDC, 6, 51, 1806293
    SetPixel m_hBtnSrcDC, 0, 52, 1806293
    SetPixel m_hBtnSrcDC, 1, 52, 4177403
    SetPixel m_hBtnSrcDC, 2, 52, 11197176
    SetPixel m_hBtnSrcDC, 3, 52, 12969209
    SetPixel m_hBtnSrcDC, 4, 52, 11197176
    SetPixel m_hBtnSrcDC, 5, 52, 4177403
    SetPixel m_hBtnSrcDC, 6, 52, 1806293
    SetPixel m_hBtnSrcDC, 0, 53, 1806293
    SetPixel m_hBtnSrcDC, 1, 53, 4374524
    SetPixel m_hBtnSrcDC, 2, 53, 11920123
    SetPixel m_hBtnSrcDC, 3, 53, 13626363
    SetPixel m_hBtnSrcDC, 4, 53, 11920123
    SetPixel m_hBtnSrcDC, 5, 53, 4374524
    SetPixel m_hBtnSrcDC, 6, 53, 1806293
    SetPixel m_hBtnSrcDC, 0, 54, 1806293
    SetPixel m_hBtnSrcDC, 1, 54, 4440316
    SetPixel m_hBtnSrcDC, 2, 54, 12051451
    SetPixel m_hBtnSrcDC, 3, 54, 13757691
    SetPixel m_hBtnSrcDC, 4, 54, 12051451
    SetPixel m_hBtnSrcDC, 5, 54, 4440316
    SetPixel m_hBtnSrcDC, 6, 54, 1806293
    SetPixel m_hBtnSrcDC, 0, 55, 1806293
    SetPixel m_hBtnSrcDC, 1, 55, 4505852
    SetPixel m_hBtnSrcDC, 2, 55, 12182780
    SetPixel m_hBtnSrcDC, 3, 55, 13954556
    SetPixel m_hBtnSrcDC, 4, 55, 12182780
    SetPixel m_hBtnSrcDC, 5, 55, 4505852
    SetPixel m_hBtnSrcDC, 6, 55, 1806293
    SetPixel m_hBtnSrcDC, 0, 56, 1806293
    SetPixel m_hBtnSrcDC, 1, 56, 4571644
    SetPixel m_hBtnSrcDC, 2, 56, 12379644
    SetPixel m_hBtnSrcDC, 3, 56, 14085884
    SetPixel m_hBtnSrcDC, 4, 56, 12379644
    SetPixel m_hBtnSrcDC, 5, 56, 4571644
    SetPixel m_hBtnSrcDC, 6, 56, 1806293
    SetPixel m_hBtnSrcDC, 0, 57, 1806293
    SetPixel m_hBtnSrcDC, 1, 57, 4506108
    SetPixel m_hBtnSrcDC, 2, 57, 12445436
    SetPixel m_hBtnSrcDC, 3, 57, 14086140
    SetPixel m_hBtnSrcDC, 4, 57, 12445436
    SetPixel m_hBtnSrcDC, 5, 57, 4506108
    SetPixel m_hBtnSrcDC, 6, 57, 1806293
    SetPixel m_hBtnSrcDC, 0, 58, 1806293
    SetPixel m_hBtnSrcDC, 1, 58, 4374780
    SetPixel m_hBtnSrcDC, 2, 58, 12117756
    SetPixel m_hBtnSrcDC, 3, 58, 13758204
    SetPixel m_hBtnSrcDC, 4, 58, 12117756
    SetPixel m_hBtnSrcDC, 5, 58, 4374780
    SetPixel m_hBtnSrcDC, 6, 58, 1806293
    SetPixel m_hBtnSrcDC, 0, 59, 1806293
    SetPixel m_hBtnSrcDC, 1, 59, 4112380
    SetPixel m_hBtnSrcDC, 2, 59, 11461886
    SetPixel m_hBtnSrcDC, 3, 59, 12839677
    SetPixel m_hBtnSrcDC, 4, 59, 11330301
    SetPixel m_hBtnSrcDC, 5, 59, 4112380
    SetPixel m_hBtnSrcDC, 6, 59, 1806293
    SetPixel m_hBtnSrcDC, 0, 60, 2659799
    SetPixel m_hBtnSrcDC, 1, 60, 3586554
    SetPixel m_hBtnSrcDC, 2, 60, 8837118
    SetPixel m_hBtnSrcDC, 3, 60, 11461885
    SetPixel m_hBtnSrcDC, 4, 60, 8968190
    SetPixel m_hBtnSrcDC, 5, 60, 3586810
    SetPixel m_hBtnSrcDC, 6, 60, 2659799
    SetPixel m_hBtnSrcDC, 0, 61, 6598882
    SetPixel m_hBtnSrcDC, 1, 61, 2466790
    SetPixel m_hBtnSrcDC, 2, 61, 3652345
    SetPixel m_hBtnSrcDC, 3, 61, 4243965
    SetPixel m_hBtnSrcDC, 4, 61, 3652345
    SetPixel m_hBtnSrcDC, 5, 61, 2466790
    SetPixel m_hBtnSrcDC, 6, 61, 6467553
    SetPixel m_hBtnSrcDC, 0, 62, 13297659
    SetPixel m_hBtnSrcDC, 1, 62, 6336225
    SetPixel m_hBtnSrcDC, 2, 62, 2791128
    SetPixel m_hBtnSrcDC, 3, 62, 1806293
    SetPixel m_hBtnSrcDC, 4, 62, 2791128
    SetPixel m_hBtnSrcDC, 5, 62, 6336225
    SetPixel m_hBtnSrcDC, 6, 62, 13297659
    SetPixel m_hBtnSrcDC, 0, 63, 15923195
    SetPixel m_hBtnSrcDC, 1, 63, 11523567
    SetPixel m_hBtnSrcDC, 2, 63, 9422826
    SetPixel m_hBtnSrcDC, 3, 63, 8897512
    SetPixel m_hBtnSrcDC, 4, 63, 9422826
    SetPixel m_hBtnSrcDC, 5, 63, 11523567
    SetPixel m_hBtnSrcDC, 6, 63, 15923195
    SetPixel m_hBtnSrcDC, 0, 64, 11458031
    SetPixel m_hBtnSrcDC, 1, 64, 12115185
    SetPixel m_hBtnSrcDC, 2, 64, 15595003
    SetPixel m_hBtnSrcDC, 3, 64, 16317437
    SetPixel m_hBtnSrcDC, 4, 64, 15595003
    SetPixel m_hBtnSrcDC, 5, 64, 12115185
    SetPixel m_hBtnSrcDC, 6, 64, 11458031
    SetPixel m_hBtnSrcDC, 0, 65, 9357033
    SetPixel m_hBtnSrcDC, 1, 65, 15726587
    SetPixel m_hBtnSrcDC, 2, 65, 16317437
    SetPixel m_hBtnSrcDC, 3, 65, 16317437
    SetPixel m_hBtnSrcDC, 4, 65, 16317437
    SetPixel m_hBtnSrcDC, 5, 65, 15726587
    SetPixel m_hBtnSrcDC, 6, 65, 9357033
    SetPixel m_hBtnSrcDC, 0, 66, 8897512
    SetPixel m_hBtnSrcDC, 1, 66, 15988988
    SetPixel m_hBtnSrcDC, 2, 66, 15923195
    SetPixel m_hBtnSrcDC, 3, 66, 15923195
    SetPixel m_hBtnSrcDC, 4, 66, 15923195
    SetPixel m_hBtnSrcDC, 5, 66, 15988988
    SetPixel m_hBtnSrcDC, 6, 66, 8897512
    SetPixel m_hBtnSrcDC, 0, 67, 8897512
    SetPixel m_hBtnSrcDC, 1, 67, 15923451
    SetPixel m_hBtnSrcDC, 2, 67, 15792123
    SetPixel m_hBtnSrcDC, 3, 67, 15792123
    SetPixel m_hBtnSrcDC, 4, 67, 15792123
    SetPixel m_hBtnSrcDC, 5, 67, 15923451
    SetPixel m_hBtnSrcDC, 6, 67, 8897512
    SetPixel m_hBtnSrcDC, 0, 68, 8897512
    SetPixel m_hBtnSrcDC, 1, 68, 15857659
    SetPixel m_hBtnSrcDC, 2, 68, 15660795
    SetPixel m_hBtnSrcDC, 3, 68, 15660795
    SetPixel m_hBtnSrcDC, 4, 68, 15660795
    SetPixel m_hBtnSrcDC, 5, 68, 15857659
    SetPixel m_hBtnSrcDC, 6, 68, 8897512
    SetPixel m_hBtnSrcDC, 0, 69, 8897512
    SetPixel m_hBtnSrcDC, 1, 69, 15792123
    SetPixel m_hBtnSrcDC, 2, 69, 15529467
    SetPixel m_hBtnSrcDC, 3, 69, 15529467
    SetPixel m_hBtnSrcDC, 4, 69, 15529467
    SetPixel m_hBtnSrcDC, 5, 69, 15792123
    SetPixel m_hBtnSrcDC, 6, 69, 8897512
    SetPixel m_hBtnSrcDC, 0, 70, 8897512
    SetPixel m_hBtnSrcDC, 1, 70, 15726331
    SetPixel m_hBtnSrcDC, 2, 70, 15398139
    SetPixel m_hBtnSrcDC, 3, 70, 15398139
    SetPixel m_hBtnSrcDC, 4, 70, 15398139
    SetPixel m_hBtnSrcDC, 5, 70, 15726331
    SetPixel m_hBtnSrcDC, 6, 70, 8897512
    SetPixel m_hBtnSrcDC, 0, 71, 8897512
    SetPixel m_hBtnSrcDC, 1, 71, 15595003
    SetPixel m_hBtnSrcDC, 2, 71, 15266811
    SetPixel m_hBtnSrcDC, 3, 71, 15266811
    SetPixel m_hBtnSrcDC, 4, 71, 15266811
    SetPixel m_hBtnSrcDC, 5, 71, 15595003
    SetPixel m_hBtnSrcDC, 6, 71, 8897512
    SetPixel m_hBtnSrcDC, 0, 72, 8897512
    SetPixel m_hBtnSrcDC, 1, 72, 15529467
    SetPixel m_hBtnSrcDC, 2, 72, 15135482
    SetPixel m_hBtnSrcDC, 3, 72, 15135482
    SetPixel m_hBtnSrcDC, 4, 72, 15135482
    SetPixel m_hBtnSrcDC, 5, 72, 15529467
    SetPixel m_hBtnSrcDC, 6, 72, 8897512
    SetPixel m_hBtnSrcDC, 0, 73, 8897512
    SetPixel m_hBtnSrcDC, 1, 73, 15463675
    SetPixel m_hBtnSrcDC, 2, 73, 15069690
    SetPixel m_hBtnSrcDC, 3, 73, 15069690
    SetPixel m_hBtnSrcDC, 4, 73, 15069690
    SetPixel m_hBtnSrcDC, 5, 73, 15463675
    SetPixel m_hBtnSrcDC, 6, 73, 8897512
    SetPixel m_hBtnSrcDC, 0, 74, 8897512
    SetPixel m_hBtnSrcDC, 1, 74, 15200762
    SetPixel m_hBtnSrcDC, 2, 74, 14675448
    SetPixel m_hBtnSrcDC, 3, 74, 14675448
    SetPixel m_hBtnSrcDC, 4, 74, 14675448
    SetPixel m_hBtnSrcDC, 5, 74, 15200762
    SetPixel m_hBtnSrcDC, 6, 74, 8897512
    SetPixel m_hBtnSrcDC, 0, 75, 8897512
    SetPixel m_hBtnSrcDC, 1, 75, 15134969
    SetPixel m_hBtnSrcDC, 2, 75, 14543863
    SetPixel m_hBtnSrcDC, 3, 75, 14543863
    SetPixel m_hBtnSrcDC, 4, 75, 14543863
    SetPixel m_hBtnSrcDC, 5, 75, 15134969
    SetPixel m_hBtnSrcDC, 6, 75, 8897512
    SetPixel m_hBtnSrcDC, 0, 76, 8897512
    SetPixel m_hBtnSrcDC, 1, 76, 15003640
    SetPixel m_hBtnSrcDC, 2, 76, 14412534
    SetPixel m_hBtnSrcDC, 3, 76, 14412534
    SetPixel m_hBtnSrcDC, 4, 76, 14412534
    SetPixel m_hBtnSrcDC, 5, 76, 15003640
    SetPixel m_hBtnSrcDC, 6, 76, 8897512
    SetPixel m_hBtnSrcDC, 0, 77, 8897512
    SetPixel m_hBtnSrcDC, 1, 77, 14872056
    SetPixel m_hBtnSrcDC, 2, 77, 14215157
    SetPixel m_hBtnSrcDC, 3, 77, 14215157
    SetPixel m_hBtnSrcDC, 4, 77, 14215157
    SetPixel m_hBtnSrcDC, 5, 77, 14872056
    SetPixel m_hBtnSrcDC, 6, 77, 8897512
    SetPixel m_hBtnSrcDC, 0, 78, 8897512
    SetPixel m_hBtnSrcDC, 1, 78, 14740727
    SetPixel m_hBtnSrcDC, 2, 78, 14018292
    SetPixel m_hBtnSrcDC, 3, 78, 14018292
    SetPixel m_hBtnSrcDC, 4, 78, 14018292
    SetPixel m_hBtnSrcDC, 5, 78, 14740727
    SetPixel m_hBtnSrcDC, 6, 78, 8897512
    SetPixel m_hBtnSrcDC, 0, 79, 8897512
    SetPixel m_hBtnSrcDC, 1, 79, 14674934
    SetPixel m_hBtnSrcDC, 2, 79, 13886707
    SetPixel m_hBtnSrcDC, 3, 79, 13886707
    SetPixel m_hBtnSrcDC, 4, 79, 13886707
    SetPixel m_hBtnSrcDC, 5, 79, 14674934
    SetPixel m_hBtnSrcDC, 6, 79, 8897512
    SetPixel m_hBtnSrcDC, 0, 80, 8897512
    SetPixel m_hBtnSrcDC, 1, 80, 14609142
    SetPixel m_hBtnSrcDC, 2, 80, 13820915
    SetPixel m_hBtnSrcDC, 3, 80, 13820915
    SetPixel m_hBtnSrcDC, 4, 80, 13820915
    SetPixel m_hBtnSrcDC, 5, 80, 14609142
    SetPixel m_hBtnSrcDC, 6, 80, 8897512
    SetPixel m_hBtnSrcDC, 0, 81, 9291497
    SetPixel m_hBtnSrcDC, 1, 81, 14084085
    SetPixel m_hBtnSrcDC, 2, 81, 14149364
    SetPixel m_hBtnSrcDC, 3, 81, 13755378
    SetPixel m_hBtnSrcDC, 4, 81, 14149364
    SetPixel m_hBtnSrcDC, 5, 81, 14084085
    SetPixel m_hBtnSrcDC, 6, 81, 9291497
    SetPixel m_hBtnSrcDC, 0, 82, 11260911
    SetPixel m_hBtnSrcDC, 1, 82, 11392238
    SetPixel m_hBtnSrcDC, 2, 82, 14149621
    SetPixel m_hBtnSrcDC, 3, 82, 14674934
    SetPixel m_hBtnSrcDC, 4, 82, 14149621
    SetPixel m_hBtnSrcDC, 5, 82, 11392238
    SetPixel m_hBtnSrcDC, 6, 82, 11195374
    SetPixel m_hBtnSrcDC, 0, 83, 15922939
    SetPixel m_hBtnSrcDC, 1, 83, 11129582
    SetPixel m_hBtnSrcDC, 2, 83, 9357034
    SetPixel m_hBtnSrcDC, 3, 83, 8897512
    SetPixel m_hBtnSrcDC, 4, 83, 9357034
    SetPixel m_hBtnSrcDC, 5, 83, 11129582
    SetPixel m_hBtnSrcDC, 6, 83, 15922939
    SetPixel m_hCbbSrcDC, 0, 0, 5152977
    SetPixel m_hCbbSrcDC, 1, 0, 15268351
    SetPixel m_hCbbSrcDC, 2, 0, 15268351
    SetPixel m_hCbbSrcDC, 3, 0, 15268351
    SetPixel m_hCbbSrcDC, 4, 0, 5152977
    SetPixel m_hCbbSrcDC, 5, 0, 12509683
    SetPixel m_hCbbSrcDC, 6, 0, 12509683
    SetPixel m_hCbbSrcDC, 7, 0, 12509683
    SetPixel m_hCbbSrcDC, 8, 0, 1660255
    SetPixel m_hCbbSrcDC, 9, 0, 1660255
    SetPixel m_hCbbSrcDC, 10, 0, 1660255
    SetPixel m_hCbbSrcDC, 11, 0, 1660255
    SetPixel m_hCbbSrcDC, 12, 0, 1660255
    SetPixel m_hCbbSrcDC, 13, 0, 1660255
    SetPixel m_hCbbSrcDC, 14, 0, 1660255
    SetPixel m_hCbbSrcDC, 0, 1, 5152977
    SetPixel m_hCbbSrcDC, 1, 1, 15137022
    SetPixel m_hCbbSrcDC, 2, 1, 12709629
    SetPixel m_hCbbSrcDC, 3, 1, 15137022
    SetPixel m_hCbbSrcDC, 4, 1, 5152977
    SetPixel m_hCbbSrcDC, 5, 1, 12640755
    SetPixel m_hCbbSrcDC, 6, 1, 6403554
    SetPixel m_hCbbSrcDC, 7, 1, 12640755
    SetPixel m_hCbbSrcDC, 8, 1, 16711935
    SetPixel m_hCbbSrcDC, 9, 1, 1660255
    SetPixel m_hCbbSrcDC, 10, 1, 1660255
    SetPixel m_hCbbSrcDC, 11, 1, 1660255
    SetPixel m_hCbbSrcDC, 12, 1, 1660255
    SetPixel m_hCbbSrcDC, 13, 1, 1660255
    SetPixel m_hCbbSrcDC, 14, 1, 16711935
    SetPixel m_hCbbSrcDC, 0, 2, 5152977
    SetPixel m_hCbbSrcDC, 1, 2, 15005694
    SetPixel m_hCbbSrcDC, 2, 2, 12381180
    SetPixel m_hCbbSrcDC, 3, 2, 15005694
    SetPixel m_hCbbSrcDC, 4, 2, 5152977
    SetPixel m_hCbbSrcDC, 5, 2, 12706804
    SetPixel m_hCbbSrcDC, 6, 2, 6666467
    SetPixel m_hCbbSrcDC, 7, 2, 12706804
    SetPixel m_hCbbSrcDC, 8, 2, 16711935
    SetPixel m_hCbbSrcDC, 9, 2, 16711935
    SetPixel m_hCbbSrcDC, 10, 2, 1660255
    SetPixel m_hCbbSrcDC, 11, 2, 1660255
    SetPixel m_hCbbSrcDC, 12, 2, 1660255
    SetPixel m_hCbbSrcDC, 13, 2, 16711935
    SetPixel m_hCbbSrcDC, 14, 2, 16711935
    SetPixel m_hCbbSrcDC, 0, 3, 5152977
    SetPixel m_hCbbSrcDC, 1, 3, 14874365
    SetPixel m_hCbbSrcDC, 2, 3, 12052730
    SetPixel m_hCbbSrcDC, 3, 3, 14874365
    SetPixel m_hCbbSrcDC, 4, 3, 5152977
    SetPixel m_hCbbSrcDC, 5, 3, 12838133
    SetPixel m_hCbbSrcDC, 6, 3, 6994917
    SetPixel m_hCbbSrcDC, 7, 3, 12838133
    SetPixel m_hCbbSrcDC, 8, 3, 16711935
    SetPixel m_hCbbSrcDC, 9, 3, 16711935
    SetPixel m_hCbbSrcDC, 10, 3, 16711935
    SetPixel m_hCbbSrcDC, 11, 3, 1660255
    SetPixel m_hCbbSrcDC, 12, 3, 16711935
    SetPixel m_hCbbSrcDC, 13, 3, 16711935
    SetPixel m_hCbbSrcDC, 14, 3, 16711935
    SetPixel m_hCbbSrcDC, 0, 4, 5152977
    SetPixel m_hCbbSrcDC, 1, 4, 14743037
    SetPixel m_hCbbSrcDC, 2, 4, 11724281
    SetPixel m_hCbbSrcDC, 3, 4, 14743037
    SetPixel m_hCbbSrcDC, 4, 4, 5152977
    SetPixel m_hCbbSrcDC, 5, 4, 13034997
    SetPixel m_hCbbSrcDC, 6, 4, 7388902
    SetPixel m_hCbbSrcDC, 7, 4, 13034997
    SetPixel m_hCbbSrcDC, 8, 4, 8693164
    SetPixel m_hCbbSrcDC, 9, 4, 8693164
    SetPixel m_hCbbSrcDC, 10, 4, 8693164
    SetPixel m_hCbbSrcDC, 11, 4, 8693164
    SetPixel m_hCbbSrcDC, 12, 4, 8693164
    SetPixel m_hCbbSrcDC, 13, 4, 8693164
    SetPixel m_hCbbSrcDC, 14, 4, 8693164
    SetPixel m_hCbbSrcDC, 0, 5, 5152977
    SetPixel m_hCbbSrcDC, 1, 5, 14611452
    SetPixel m_hCbbSrcDC, 2, 5, 11330039
    SetPixel m_hCbbSrcDC, 3, 5, 14611452
    SetPixel m_hCbbSrcDC, 4, 5, 5152977
    SetPixel m_hCbbSrcDC, 5, 5, 13166582
    SetPixel m_hCbbSrcDC, 6, 5, 7783144
    SetPixel m_hCbbSrcDC, 7, 5, 13166582
    SetPixel m_hCbbSrcDC, 8, 5, 16711935
    SetPixel m_hCbbSrcDC, 9, 5, 8693164
    SetPixel m_hCbbSrcDC, 10, 5, 8693164
    SetPixel m_hCbbSrcDC, 11, 5, 8693164
    SetPixel m_hCbbSrcDC, 12, 5, 8693164
    SetPixel m_hCbbSrcDC, 13, 5, 8693164
    SetPixel m_hCbbSrcDC, 14, 5, 16711935
    SetPixel m_hCbbSrcDC, 0, 6, 5152977
    SetPixel m_hCbbSrcDC, 1, 6, 14414331
    SetPixel m_hCbbSrcDC, 2, 6, 10870261
    SetPixel m_hCbbSrcDC, 3, 6, 14414331
    SetPixel m_hCbbSrcDC, 4, 6, 5152977
    SetPixel m_hCbbSrcDC, 5, 6, 13363447
    SetPixel m_hCbbSrcDC, 6, 6, 8177130
    SetPixel m_hCbbSrcDC, 7, 6, 13363447
    SetPixel m_hCbbSrcDC, 8, 6, 16711935
    SetPixel m_hCbbSrcDC, 9, 6, 16711935
    SetPixel m_hCbbSrcDC, 10, 6, 8693164
    SetPixel m_hCbbSrcDC, 11, 6, 8693164
    SetPixel m_hCbbSrcDC, 12, 6, 8693164
    SetPixel m_hCbbSrcDC, 13, 6, 16711935
    SetPixel m_hCbbSrcDC, 14, 6, 16711935
    SetPixel m_hCbbSrcDC, 0, 7, 5152977
    SetPixel m_hCbbSrcDC, 1, 7, 14283002
    SetPixel m_hCbbSrcDC, 2, 7, 10476019
    SetPixel m_hCbbSrcDC, 3, 7, 14283002
    SetPixel m_hCbbSrcDC, 4, 7, 5152977
    SetPixel m_hCbbSrcDC, 5, 7, 13494775
    SetPixel m_hCbbSrcDC, 6, 7, 8636908
    SetPixel m_hCbbSrcDC, 7, 7, 13494775
    SetPixel m_hCbbSrcDC, 8, 7, 16711935
    SetPixel m_hCbbSrcDC, 9, 7, 16711935
    SetPixel m_hCbbSrcDC, 10, 7, 16711935
    SetPixel m_hCbbSrcDC, 11, 7, 8693164
    SetPixel m_hCbbSrcDC, 12, 7, 16711935
    SetPixel m_hCbbSrcDC, 13, 7, 16711935
    SetPixel m_hCbbSrcDC, 14, 7, 16711935
    SetPixel m_hCbbSrcDC, 0, 8, 5152977
    SetPixel m_hCbbSrcDC, 1, 8, 14085881
    SetPixel m_hCbbSrcDC, 2, 8, 10016241
    SetPixel m_hCbbSrcDC, 3, 8, 14085881
    SetPixel m_hCbbSrcDC, 4, 8, 5152977
    SetPixel m_hCbbSrcDC, 5, 8, 13691896
    SetPixel m_hCbbSrcDC, 6, 8, 9096686
    SetPixel m_hCbbSrcDC, 7, 8, 13691896
    SetPixel m_hCbbSrcDC, 8, 8, 16711935
    SetPixel m_hCbbSrcDC, 9, 8, 16711935
    SetPixel m_hCbbSrcDC, 10, 8, 16711935
    SetPixel m_hCbbSrcDC, 11, 8, 16711935
    SetPixel m_hCbbSrcDC, 12, 8, 16711935
    SetPixel m_hCbbSrcDC, 13, 8, 16711935
    SetPixel m_hCbbSrcDC, 14, 8, 16711935
    SetPixel m_hCbbSrcDC, 0, 9, 5152977
    SetPixel m_hCbbSrcDC, 1, 9, 13889017
    SetPixel m_hCbbSrcDC, 2, 9, 9556463
    SetPixel m_hCbbSrcDC, 3, 9, 13889017
    SetPixel m_hCbbSrcDC, 4, 9, 5152977
    SetPixel m_hCbbSrcDC, 5, 9, 13889017
    SetPixel m_hCbbSrcDC, 6, 9, 9556463
    SetPixel m_hCbbSrcDC, 7, 9, 13889017
    SetPixel m_hCbbSrcDC, 8, 9, 16711935
    SetPixel m_hCbbSrcDC, 9, 9, 16711935
    SetPixel m_hCbbSrcDC, 10, 9, 16711935
    SetPixel m_hCbbSrcDC, 11, 9, 16711935
    SetPixel m_hCbbSrcDC, 12, 9, 16711935
    SetPixel m_hCbbSrcDC, 13, 9, 16711935
    SetPixel m_hCbbSrcDC, 14, 9, 16711935
    SetPixel m_hCbbSrcDC, 0, 10, 5152977
    SetPixel m_hCbbSrcDC, 1, 10, 13691896
    SetPixel m_hCbbSrcDC, 2, 10, 9096686
    SetPixel m_hCbbSrcDC, 3, 10, 13691896
    SetPixel m_hCbbSrcDC, 4, 10, 5152977
    SetPixel m_hCbbSrcDC, 5, 10, 14085881
    SetPixel m_hCbbSrcDC, 6, 10, 10016241
    SetPixel m_hCbbSrcDC, 7, 10, 14085881
    SetPixel m_hCbbSrcDC, 8, 10, 16711935
    SetPixel m_hCbbSrcDC, 9, 10, 16711935
    SetPixel m_hCbbSrcDC, 10, 10, 16711935
    SetPixel m_hCbbSrcDC, 11, 10, 16711935
    SetPixel m_hCbbSrcDC, 12, 10, 16711935
    SetPixel m_hCbbSrcDC, 13, 10, 16711935
    SetPixel m_hCbbSrcDC, 14, 10, 16711935
    SetPixel m_hCbbSrcDC, 0, 11, 5152977
    SetPixel m_hCbbSrcDC, 1, 11, 13494775
    SetPixel m_hCbbSrcDC, 2, 11, 8636908
    SetPixel m_hCbbSrcDC, 3, 11, 13494775
    SetPixel m_hCbbSrcDC, 4, 11, 5152977
    SetPixel m_hCbbSrcDC, 5, 11, 14283002
    SetPixel m_hCbbSrcDC, 6, 11, 10476019
    SetPixel m_hCbbSrcDC, 7, 11, 14283002
    SetPixel m_hCbbSrcDC, 8, 11, 16711935
    SetPixel m_hCbbSrcDC, 9, 11, 16711935
    SetPixel m_hCbbSrcDC, 10, 11, 16711935
    SetPixel m_hCbbSrcDC, 11, 11, 16711935
    SetPixel m_hCbbSrcDC, 12, 11, 16711935
    SetPixel m_hCbbSrcDC, 13, 11, 16711935
    SetPixel m_hCbbSrcDC, 14, 11, 16711935
    SetPixel m_hCbbSrcDC, 0, 12, 5152977
    SetPixel m_hCbbSrcDC, 1, 12, 13363447
    SetPixel m_hCbbSrcDC, 2, 12, 8177130
    SetPixel m_hCbbSrcDC, 3, 12, 13363447
    SetPixel m_hCbbSrcDC, 4, 12, 5152977
    SetPixel m_hCbbSrcDC, 5, 12, 14414331
    SetPixel m_hCbbSrcDC, 6, 12, 10870261
    SetPixel m_hCbbSrcDC, 7, 12, 14414331
    SetPixel m_hCbbSrcDC, 8, 12, 16711935
    SetPixel m_hCbbSrcDC, 9, 12, 16711935
    SetPixel m_hCbbSrcDC, 10, 12, 16711935
    SetPixel m_hCbbSrcDC, 11, 12, 16711935
    SetPixel m_hCbbSrcDC, 12, 12, 16711935
    SetPixel m_hCbbSrcDC, 13, 12, 16711935
    SetPixel m_hCbbSrcDC, 14, 12, 16711935
    SetPixel m_hCbbSrcDC, 0, 13, 5152977
    SetPixel m_hCbbSrcDC, 1, 13, 13166582
    SetPixel m_hCbbSrcDC, 2, 13, 7783144
    SetPixel m_hCbbSrcDC, 3, 13, 13166582
    SetPixel m_hCbbSrcDC, 4, 13, 5152977
    SetPixel m_hCbbSrcDC, 5, 13, 14611452
    SetPixel m_hCbbSrcDC, 6, 13, 11330039
    SetPixel m_hCbbSrcDC, 7, 13, 14611452
    SetPixel m_hCbbSrcDC, 8, 13, 16711935
    SetPixel m_hCbbSrcDC, 9, 13, 16711935
    SetPixel m_hCbbSrcDC, 10, 13, 16711935
    SetPixel m_hCbbSrcDC, 11, 13, 16711935
    SetPixel m_hCbbSrcDC, 12, 13, 16711935
    SetPixel m_hCbbSrcDC, 13, 13, 16711935
    SetPixel m_hCbbSrcDC, 14, 13, 16711935
    SetPixel m_hCbbSrcDC, 0, 14, 5152977
    SetPixel m_hCbbSrcDC, 1, 14, 13034997
    SetPixel m_hCbbSrcDC, 2, 14, 7388902
    SetPixel m_hCbbSrcDC, 3, 14, 13034997
    SetPixel m_hCbbSrcDC, 4, 14, 5152977
    SetPixel m_hCbbSrcDC, 5, 14, 14743037
    SetPixel m_hCbbSrcDC, 6, 14, 11724281
    SetPixel m_hCbbSrcDC, 7, 14, 14743037
    SetPixel m_hCbbSrcDC, 8, 14, 16711935
    SetPixel m_hCbbSrcDC, 9, 14, 16711935
    SetPixel m_hCbbSrcDC, 10, 14, 16711935
    SetPixel m_hCbbSrcDC, 11, 14, 16711935
    SetPixel m_hCbbSrcDC, 12, 14, 16711935
    SetPixel m_hCbbSrcDC, 13, 14, 16711935
    SetPixel m_hCbbSrcDC, 14, 14, 16711935
    SetPixel m_hCbbSrcDC, 0, 15, 5152977
    SetPixel m_hCbbSrcDC, 1, 15, 12838133
    SetPixel m_hCbbSrcDC, 2, 15, 6994917
    SetPixel m_hCbbSrcDC, 3, 15, 12838133
    SetPixel m_hCbbSrcDC, 4, 15, 5152977
    SetPixel m_hCbbSrcDC, 5, 15, 14874365
    SetPixel m_hCbbSrcDC, 6, 15, 12052730
    SetPixel m_hCbbSrcDC, 7, 15, 14874365
    SetPixel m_hCbbSrcDC, 8, 15, 16711935
    SetPixel m_hCbbSrcDC, 9, 15, 16711935
    SetPixel m_hCbbSrcDC, 10, 15, 16711935
    SetPixel m_hCbbSrcDC, 11, 15, 16711935
    SetPixel m_hCbbSrcDC, 12, 15, 16711935
    SetPixel m_hCbbSrcDC, 13, 15, 16711935
    SetPixel m_hCbbSrcDC, 14, 15, 16711935
    SetPixel m_hCbbSrcDC, 0, 16, 5152977
    SetPixel m_hCbbSrcDC, 1, 16, 12706804
    SetPixel m_hCbbSrcDC, 2, 16, 6666467
    SetPixel m_hCbbSrcDC, 3, 16, 12706804
    SetPixel m_hCbbSrcDC, 4, 16, 5152977
    SetPixel m_hCbbSrcDC, 5, 16, 15005694
    SetPixel m_hCbbSrcDC, 6, 16, 12381180
    SetPixel m_hCbbSrcDC, 7, 16, 15005694
    SetPixel m_hCbbSrcDC, 8, 16, 16711935
    SetPixel m_hCbbSrcDC, 9, 16, 16711935
    SetPixel m_hCbbSrcDC, 10, 16, 16711935
    SetPixel m_hCbbSrcDC, 11, 16, 16711935
    SetPixel m_hCbbSrcDC, 12, 16, 16711935
    SetPixel m_hCbbSrcDC, 13, 16, 16711935
    SetPixel m_hCbbSrcDC, 14, 16, 16711935
    SetPixel m_hCbbSrcDC, 0, 17, 5152977
    SetPixel m_hCbbSrcDC, 1, 17, 12640755
    SetPixel m_hCbbSrcDC, 2, 17, 12640755
    SetPixel m_hCbbSrcDC, 3, 17, 12640755
    SetPixel m_hCbbSrcDC, 4, 17, 5152977
    SetPixel m_hCbbSrcDC, 5, 17, 15137022
    SetPixel m_hCbbSrcDC, 6, 17, 15137022
    SetPixel m_hCbbSrcDC, 7, 17, 15137022
    SetPixel m_hCbbSrcDC, 8, 17, 16711935
    SetPixel m_hCbbSrcDC, 9, 17, 16711935
    SetPixel m_hCbbSrcDC, 10, 17, 16711935
    SetPixel m_hCbbSrcDC, 11, 17, 16711935
    SetPixel m_hCbbSrcDC, 12, 17, 16711935
    SetPixel m_hCbbSrcDC, 13, 17, 16711935
    SetPixel m_hCbbSrcDC, 14, 17, 16711935
    SetPixel m_hCkbSrcDC, 0, 0, 16711935
    SetPixel m_hCkbSrcDC, 1, 0, 16711935
    SetPixel m_hCkbSrcDC, 2, 0, 16711935
    SetPixel m_hCkbSrcDC, 3, 0, 16711935
    SetPixel m_hCkbSrcDC, 4, 0, 16711935
    SetPixel m_hCkbSrcDC, 5, 0, 16711935
    SetPixel m_hCkbSrcDC, 6, 0, 16711935
    SetPixel m_hCkbSrcDC, 7, 0, 16711935
    SetPixel m_hCkbSrcDC, 8, 0, 5551901
    SetPixel m_hCkbSrcDC, 9, 0, 16711935
    SetPixel m_hCkbSrcDC, 10, 0, 16711935
    SetPixel m_hCkbSrcDC, 11, 0, 16711935
    SetPixel m_hCkbSrcDC, 12, 0, 16711935
    SetPixel m_hCkbSrcDC, 13, 0, 16711935
    SetPixel m_hCkbSrcDC, 14, 0, 16711935
    SetPixel m_hCkbSrcDC, 15, 0, 16711935
    SetPixel m_hCkbSrcDC, 16, 0, 16711935
    SetPixel m_hCkbSrcDC, 17, 0, 11065996
    SetPixel m_hCkbSrcDC, 0, 1, 16711935
    SetPixel m_hCkbSrcDC, 1, 1, 16711935
    SetPixel m_hCkbSrcDC, 2, 1, 16711935
    SetPixel m_hCkbSrcDC, 3, 1, 16711935
    SetPixel m_hCkbSrcDC, 4, 1, 16711935
    SetPixel m_hCkbSrcDC, 5, 1, 16711935
    SetPixel m_hCkbSrcDC, 6, 1, 16711935
    SetPixel m_hCkbSrcDC, 7, 1, 5354523
    SetPixel m_hCkbSrcDC, 8, 1, 5354523
    SetPixel m_hCkbSrcDC, 9, 1, 16711935
    SetPixel m_hCkbSrcDC, 10, 1, 16711935
    SetPixel m_hCkbSrcDC, 11, 1, 16711935
    SetPixel m_hCkbSrcDC, 12, 1, 16711935
    SetPixel m_hCkbSrcDC, 13, 1, 16711935
    SetPixel m_hCkbSrcDC, 14, 1, 16711935
    SetPixel m_hCkbSrcDC, 15, 1, 16711935
    SetPixel m_hCkbSrcDC, 16, 1, 10934667
    SetPixel m_hCkbSrcDC, 17, 1, 10934667
    SetPixel m_hCkbSrcDC, 0, 2, 16711935
    SetPixel m_hCkbSrcDC, 1, 2, 16711935
    SetPixel m_hCkbSrcDC, 2, 2, 16711935
    SetPixel m_hCkbSrcDC, 3, 2, 16711935
    SetPixel m_hCkbSrcDC, 4, 2, 16711935
    SetPixel m_hCkbSrcDC, 5, 2, 16711935
    SetPixel m_hCkbSrcDC, 6, 2, 5091095
    SetPixel m_hCkbSrcDC, 7, 2, 5091095
    SetPixel m_hCkbSrcDC, 8, 2, 5091095
    SetPixel m_hCkbSrcDC, 9, 2, 16711935
    SetPixel m_hCkbSrcDC, 10, 2, 16711935
    SetPixel m_hCkbSrcDC, 11, 2, 16711935
    SetPixel m_hCkbSrcDC, 12, 2, 16711935
    SetPixel m_hCkbSrcDC, 13, 2, 16711935
    SetPixel m_hCkbSrcDC, 14, 2, 16711935
    SetPixel m_hCkbSrcDC, 15, 2, 10802825
    SetPixel m_hCkbSrcDC, 16, 2, 10802825
    SetPixel m_hCkbSrcDC, 17, 2, 10802825
    SetPixel m_hCkbSrcDC, 0, 3, 4761875
    SetPixel m_hCkbSrcDC, 1, 3, 4761875
    SetPixel m_hCkbSrcDC, 2, 3, 16711935
    SetPixel m_hCkbSrcDC, 3, 3, 16711935
    SetPixel m_hCkbSrcDC, 4, 3, 16711935
    SetPixel m_hCkbSrcDC, 5, 3, 4761875
    SetPixel m_hCkbSrcDC, 6, 3, 4761875
    SetPixel m_hCkbSrcDC, 7, 3, 4761875
    SetPixel m_hCkbSrcDC, 8, 3, 16711935
    SetPixel m_hCkbSrcDC, 9, 3, 10670983
    SetPixel m_hCkbSrcDC, 10, 3, 10670983
    SetPixel m_hCkbSrcDC, 11, 3, 16711935
    SetPixel m_hCkbSrcDC, 12, 3, 16711935
    SetPixel m_hCkbSrcDC, 13, 3, 16711935
    SetPixel m_hCkbSrcDC, 14, 3, 10670983
    SetPixel m_hCkbSrcDC, 15, 3, 10670983
    SetPixel m_hCkbSrcDC, 16, 3, 10670983
    SetPixel m_hCkbSrcDC, 17, 3, 16711935
    SetPixel m_hCkbSrcDC, 0, 4, 4432399
    SetPixel m_hCkbSrcDC, 1, 4, 4432399
    SetPixel m_hCkbSrcDC, 2, 4, 4432399
    SetPixel m_hCkbSrcDC, 3, 4, 16711935
    SetPixel m_hCkbSrcDC, 4, 4, 4432399
    SetPixel m_hCkbSrcDC, 5, 4, 4432399
    SetPixel m_hCkbSrcDC, 6, 4, 4432399
    SetPixel m_hCkbSrcDC, 7, 4, 16711935
    SetPixel m_hCkbSrcDC, 8, 4, 16711935
    SetPixel m_hCkbSrcDC, 9, 4, 10473605
    SetPixel m_hCkbSrcDC, 10, 4, 10473605
    SetPixel m_hCkbSrcDC, 11, 4, 10473605
    SetPixel m_hCkbSrcDC, 12, 4, 16711935
    SetPixel m_hCkbSrcDC, 13, 4, 10473605
    SetPixel m_hCkbSrcDC, 14, 4, 10473605
    SetPixel m_hCkbSrcDC, 15, 4, 10473605
    SetPixel m_hCkbSrcDC, 16, 4, 16711935
    SetPixel m_hCkbSrcDC, 17, 4, 16711935
    SetPixel m_hCkbSrcDC, 0, 5, 16711935
    SetPixel m_hCkbSrcDC, 1, 5, 4103178
    SetPixel m_hCkbSrcDC, 2, 5, 4103178
    SetPixel m_hCkbSrcDC, 3, 5, 4103178
    SetPixel m_hCkbSrcDC, 4, 5, 4103178
    SetPixel m_hCkbSrcDC, 5, 5, 4103178
    SetPixel m_hCkbSrcDC, 6, 5, 16711935
    SetPixel m_hCkbSrcDC, 7, 5, 16711935
    SetPixel m_hCkbSrcDC, 8, 5, 16711935
    SetPixel m_hCkbSrcDC, 9, 5, 16711935
    SetPixel m_hCkbSrcDC, 10, 5, 10341762
    SetPixel m_hCkbSrcDC, 11, 5, 10341762
    SetPixel m_hCkbSrcDC, 12, 5, 10341762
    SetPixel m_hCkbSrcDC, 13, 5, 10341762
    SetPixel m_hCkbSrcDC, 14, 5, 10341762
    SetPixel m_hCkbSrcDC, 15, 5, 16711935
    SetPixel m_hCkbSrcDC, 16, 5, 16711935
    SetPixel m_hCkbSrcDC, 17, 5, 16711935
    SetPixel m_hCkbSrcDC, 0, 6, 16711935
    SetPixel m_hCkbSrcDC, 1, 6, 16711935
    SetPixel m_hCkbSrcDC, 2, 6, 3773958
    SetPixel m_hCkbSrcDC, 3, 6, 3773958
    SetPixel m_hCkbSrcDC, 4, 6, 3773958
    SetPixel m_hCkbSrcDC, 5, 6, 16711935
    SetPixel m_hCkbSrcDC, 6, 6, 16711935
    SetPixel m_hCkbSrcDC, 7, 6, 16711935
    SetPixel m_hCkbSrcDC, 8, 6, 16711935
    SetPixel m_hCkbSrcDC, 9, 6, 16711935
    SetPixel m_hCkbSrcDC, 10, 6, 16711935
    SetPixel m_hCkbSrcDC, 11, 6, 10144128
    SetPixel m_hCkbSrcDC, 12, 6, 10144128
    SetPixel m_hCkbSrcDC, 13, 6, 10144128
    SetPixel m_hCkbSrcDC, 14, 6, 16711935
    SetPixel m_hCkbSrcDC, 15, 6, 16711935
    SetPixel m_hCkbSrcDC, 16, 6, 16711935
    SetPixel m_hCkbSrcDC, 17, 6, 16711935
    SetPixel m_hCkbSrcDC, 0, 7, 16711935
    SetPixel m_hCkbSrcDC, 1, 7, 16711935
    SetPixel m_hCkbSrcDC, 2, 7, 16711935
    SetPixel m_hCkbSrcDC, 3, 7, 3510530
    SetPixel m_hCkbSrcDC, 4, 7, 16711935
    SetPixel m_hCkbSrcDC, 5, 7, 16711935
    SetPixel m_hCkbSrcDC, 6, 7, 16711935
    SetPixel m_hCkbSrcDC, 7, 7, 16711935
    SetPixel m_hCkbSrcDC, 8, 7, 16711935
    SetPixel m_hCkbSrcDC, 9, 7, 16711935
    SetPixel m_hCkbSrcDC, 10, 7, 16711935
    SetPixel m_hCkbSrcDC, 11, 7, 16711935
    SetPixel m_hCkbSrcDC, 12, 7, 10012542
    SetPixel m_hCkbSrcDC, 13, 7, 16711935
    SetPixel m_hCkbSrcDC, 14, 7, 16711935
    SetPixel m_hCkbSrcDC, 15, 7, 16711935
    SetPixel m_hCkbSrcDC, 16, 7, 16711935
    SetPixel m_hCkbSrcDC, 17, 7, 16711935
    SetPixel m_hCkbSrcDC, 0, 8, 16711935
    SetPixel m_hCkbSrcDC, 1, 8, 16711935
    SetPixel m_hCkbSrcDC, 2, 8, 16711935
    SetPixel m_hCkbSrcDC, 3, 8, 16711935
    SetPixel m_hCkbSrcDC, 4, 8, 16711935
    SetPixel m_hCkbSrcDC, 5, 8, 16711935
    SetPixel m_hCkbSrcDC, 6, 8, 16711935
    SetPixel m_hCkbSrcDC, 7, 8, 16711935
    SetPixel m_hCkbSrcDC, 8, 8, 16711935
    SetPixel m_hCkbSrcDC, 9, 8, 16711935
    SetPixel m_hCkbSrcDC, 10, 8, 16711935
    SetPixel m_hCkbSrcDC, 11, 8, 16711935
    SetPixel m_hCkbSrcDC, 12, 8, 16711935
    SetPixel m_hCkbSrcDC, 13, 8, 16711935
    SetPixel m_hCkbSrcDC, 14, 8, 16711935
    SetPixel m_hCkbSrcDC, 15, 8, 16711935
    SetPixel m_hCkbSrcDC, 16, 8, 16711935
    SetPixel m_hCkbSrcDC, 17, 8, 16711935
    SetPixel m_hCkbSrcDC, 0, 9, 16711935
    SetPixel m_hCkbSrcDC, 1, 9, 5354523
    SetPixel m_hCkbSrcDC, 2, 9, 5354523
    SetPixel m_hCkbSrcDC, 3, 9, 5354523
    SetPixel m_hCkbSrcDC, 4, 9, 5354523
    SetPixel m_hCkbSrcDC, 5, 9, 5354523
    SetPixel m_hCkbSrcDC, 6, 9, 5354523
    SetPixel m_hCkbSrcDC, 7, 9, 5354523
    SetPixel m_hCkbSrcDC, 8, 9, 16711935
    SetPixel m_hCkbSrcDC, 9, 9, 16711935
    SetPixel m_hCkbSrcDC, 10, 9, 10934667
    SetPixel m_hCkbSrcDC, 11, 9, 10934667
    SetPixel m_hCkbSrcDC, 12, 9, 10934667
    SetPixel m_hCkbSrcDC, 13, 9, 10934667
    SetPixel m_hCkbSrcDC, 14, 9, 10934667
    SetPixel m_hCkbSrcDC, 15, 9, 10934667
    SetPixel m_hCkbSrcDC, 16, 9, 10934667
    SetPixel m_hCkbSrcDC, 17, 9, 16711935
    SetPixel m_hCkbSrcDC, 0, 10, 16711935
    SetPixel m_hCkbSrcDC, 1, 10, 5156888
    SetPixel m_hCkbSrcDC, 2, 10, 5156888
    SetPixel m_hCkbSrcDC, 3, 10, 5156888
    SetPixel m_hCkbSrcDC, 4, 10, 5156888
    SetPixel m_hCkbSrcDC, 5, 10, 5156888
    SetPixel m_hCkbSrcDC, 6, 10, 5156888
    SetPixel m_hCkbSrcDC, 7, 10, 5156888
    SetPixel m_hCkbSrcDC, 8, 10, 16711935
    SetPixel m_hCkbSrcDC, 9, 10, 16711935
    SetPixel m_hCkbSrcDC, 10, 10, 10802825
    SetPixel m_hCkbSrcDC, 11, 10, 10802825
    SetPixel m_hCkbSrcDC, 12, 10, 10802825
    SetPixel m_hCkbSrcDC, 13, 10, 10802825
    SetPixel m_hCkbSrcDC, 14, 10, 10802825
    SetPixel m_hCkbSrcDC, 15, 10, 10802825
    SetPixel m_hCkbSrcDC, 16, 10, 10802825
    SetPixel m_hCkbSrcDC, 17, 10, 16711935
    SetPixel m_hCkbSrcDC, 0, 11, 16711935
    SetPixel m_hCkbSrcDC, 1, 11, 4827924
    SetPixel m_hCkbSrcDC, 2, 11, 4827924
    SetPixel m_hCkbSrcDC, 3, 11, 4827924
    SetPixel m_hCkbSrcDC, 4, 11, 4827924
    SetPixel m_hCkbSrcDC, 5, 11, 4827924
    SetPixel m_hCkbSrcDC, 6, 11, 4827924
    SetPixel m_hCkbSrcDC, 7, 11, 4827924
    SetPixel m_hCkbSrcDC, 8, 11, 16711935
    SetPixel m_hCkbSrcDC, 9, 11, 16711935
    SetPixel m_hCkbSrcDC, 10, 11, 10736775
    SetPixel m_hCkbSrcDC, 11, 11, 10736775
    SetPixel m_hCkbSrcDC, 12, 11, 10736775
    SetPixel m_hCkbSrcDC, 13, 11, 10736775
    SetPixel m_hCkbSrcDC, 14, 11, 10736775
    SetPixel m_hCkbSrcDC, 15, 11, 10736775
    SetPixel m_hCkbSrcDC, 16, 11, 10736775
    SetPixel m_hCkbSrcDC, 17, 11, 16711935
    SetPixel m_hCkbSrcDC, 0, 12, 16711935
    SetPixel m_hCkbSrcDC, 1, 12, 4563984
    SetPixel m_hCkbSrcDC, 2, 12, 4563984
    SetPixel m_hCkbSrcDC, 3, 12, 4563984
    SetPixel m_hCkbSrcDC, 4, 12, 4563984
    SetPixel m_hCkbSrcDC, 5, 12, 4563984
    SetPixel m_hCkbSrcDC, 6, 12, 4563984
    SetPixel m_hCkbSrcDC, 7, 12, 4563984
    SetPixel m_hCkbSrcDC, 8, 12, 16711935
    SetPixel m_hCkbSrcDC, 9, 12, 16711935
    SetPixel m_hCkbSrcDC, 10, 12, 10539398
    SetPixel m_hCkbSrcDC, 11, 12, 10539398
    SetPixel m_hCkbSrcDC, 12, 12, 10539398
    SetPixel m_hCkbSrcDC, 13, 12, 10539398
    SetPixel m_hCkbSrcDC, 14, 12, 10539398
    SetPixel m_hCkbSrcDC, 15, 12, 10539398
    SetPixel m_hCkbSrcDC, 16, 12, 10539398
    SetPixel m_hCkbSrcDC, 17, 12, 16711935
    SetPixel m_hCkbSrcDC, 0, 13, 16711935
    SetPixel m_hCkbSrcDC, 1, 13, 4234764
    SetPixel m_hCkbSrcDC, 2, 13, 4234764
    SetPixel m_hCkbSrcDC, 3, 13, 4234764
    SetPixel m_hCkbSrcDC, 4, 13, 4234764
    SetPixel m_hCkbSrcDC, 5, 13, 4234764
    SetPixel m_hCkbSrcDC, 6, 13, 4234764
    SetPixel m_hCkbSrcDC, 7, 13, 4234764
    SetPixel m_hCkbSrcDC, 8, 13, 16711935
    SetPixel m_hCkbSrcDC, 9, 13, 16711935
    SetPixel m_hCkbSrcDC, 10, 13, 10407555
    SetPixel m_hCkbSrcDC, 11, 13, 10407555
    SetPixel m_hCkbSrcDC, 12, 13, 10407555
    SetPixel m_hCkbSrcDC, 13, 13, 10407555
    SetPixel m_hCkbSrcDC, 14, 13, 10407555
    SetPixel m_hCkbSrcDC, 15, 13, 10407555
    SetPixel m_hCkbSrcDC, 16, 13, 10407555
    SetPixel m_hCkbSrcDC, 17, 13, 16711935
    SetPixel m_hCkbSrcDC, 0, 14, 16711935
    SetPixel m_hCkbSrcDC, 1, 14, 3905543
    SetPixel m_hCkbSrcDC, 2, 14, 3905543
    SetPixel m_hCkbSrcDC, 3, 14, 3905543
    SetPixel m_hCkbSrcDC, 4, 14, 3905543
    SetPixel m_hCkbSrcDC, 5, 14, 3905543
    SetPixel m_hCkbSrcDC, 6, 14, 3905543
    SetPixel m_hCkbSrcDC, 7, 14, 3905543
    SetPixel m_hCkbSrcDC, 8, 14, 16711935
    SetPixel m_hCkbSrcDC, 9, 14, 16711935
    SetPixel m_hCkbSrcDC, 10, 14, 10209921
    SetPixel m_hCkbSrcDC, 11, 14, 10209921
    SetPixel m_hCkbSrcDC, 12, 14, 10209921
    SetPixel m_hCkbSrcDC, 13, 14, 10209921
    SetPixel m_hCkbSrcDC, 14, 14, 10209921
    SetPixel m_hCkbSrcDC, 15, 14, 10209921
    SetPixel m_hCkbSrcDC, 16, 14, 10209921
    SetPixel m_hCkbSrcDC, 17, 14, 16711935
    SetPixel m_hCkbSrcDC, 0, 15, 16711935
    SetPixel m_hCkbSrcDC, 1, 15, 3773958
    SetPixel m_hCkbSrcDC, 2, 15, 3773958
    SetPixel m_hCkbSrcDC, 3, 15, 3773958
    SetPixel m_hCkbSrcDC, 4, 15, 3773958
    SetPixel m_hCkbSrcDC, 5, 15, 3773958
    SetPixel m_hCkbSrcDC, 6, 15, 3773958
    SetPixel m_hCkbSrcDC, 7, 15, 3773958
    SetPixel m_hCkbSrcDC, 8, 15, 16711935
    SetPixel m_hCkbSrcDC, 9, 15, 16711935
    SetPixel m_hCkbSrcDC, 10, 15, 10144128
    SetPixel m_hCkbSrcDC, 11, 15, 10144128
    SetPixel m_hCkbSrcDC, 12, 15, 10144128
    SetPixel m_hCkbSrcDC, 13, 15, 10144128
    SetPixel m_hCkbSrcDC, 14, 15, 10144128
    SetPixel m_hCkbSrcDC, 15, 15, 10144128
    SetPixel m_hCkbSrcDC, 16, 15, 10144128
    SetPixel m_hCkbSrcDC, 17, 15, 16711935
    SetPixel m_hCkbSrcDC, 0, 16, 16711935
    SetPixel m_hCkbSrcDC, 1, 16, 16711935
    SetPixel m_hCkbSrcDC, 2, 16, 16711935
    SetPixel m_hCkbSrcDC, 3, 16, 16711935
    SetPixel m_hCkbSrcDC, 4, 16, 16711935
    SetPixel m_hCkbSrcDC, 5, 16, 16711935
    SetPixel m_hCkbSrcDC, 6, 16, 16711935
    SetPixel m_hCkbSrcDC, 7, 16, 16711935
    SetPixel m_hCkbSrcDC, 8, 16, 16711935
    SetPixel m_hCkbSrcDC, 9, 16, 16711935
    SetPixel m_hCkbSrcDC, 10, 16, 16711935
    SetPixel m_hCkbSrcDC, 11, 16, 16711935
    SetPixel m_hCkbSrcDC, 12, 16, 16711935
    SetPixel m_hCkbSrcDC, 13, 16, 16711935
    SetPixel m_hCkbSrcDC, 14, 16, 16711935
    SetPixel m_hCkbSrcDC, 15, 16, 16711935
    SetPixel m_hCkbSrcDC, 16, 16, 16711935
    SetPixel m_hCkbSrcDC, 17, 16, 16711935
    SetPixel m_hCkbSrcDC, 0, 17, 16711935
    SetPixel m_hCkbSrcDC, 1, 17, 16711935
    SetPixel m_hCkbSrcDC, 2, 17, 16711935
    SetPixel m_hCkbSrcDC, 3, 17, 16711935
    SetPixel m_hCkbSrcDC, 4, 17, 16711935
    SetPixel m_hCkbSrcDC, 5, 17, 16711935
    SetPixel m_hCkbSrcDC, 6, 17, 16711935
    SetPixel m_hCkbSrcDC, 7, 17, 16711935
    SetPixel m_hCkbSrcDC, 8, 17, 16711935
    SetPixel m_hCkbSrcDC, 9, 17, 16711935
    SetPixel m_hCkbSrcDC, 10, 17, 16711935
    SetPixel m_hCkbSrcDC, 11, 17, 16711935
    SetPixel m_hCkbSrcDC, 12, 17, 16711935
    SetPixel m_hCkbSrcDC, 13, 17, 16711935
    SetPixel m_hCkbSrcDC, 14, 17, 16711935
    SetPixel m_hCkbSrcDC, 15, 17, 16711935
    SetPixel m_hCkbSrcDC, 16, 17, 16711935
    SetPixel m_hCkbSrcDC, 17, 17, 16711935
    SetPixel m_hCkbSrcDC, 0, 18, 16711935
    SetPixel m_hCkbSrcDC, 1, 18, 16711935
    SetPixel m_hCkbSrcDC, 2, 18, 16711935
    SetPixel m_hCkbSrcDC, 3, 18, 8899680
    SetPixel m_hCkbSrcDC, 4, 18, 6076968
    SetPixel m_hCkbSrcDC, 5, 18, 8899680
    SetPixel m_hCkbSrcDC, 6, 18, 16711935
    SetPixel m_hCkbSrcDC, 7, 18, 16711935
    SetPixel m_hCkbSrcDC, 8, 18, 16711935
    SetPixel m_hCkbSrcDC, 9, 18, 16711935
    SetPixel m_hCkbSrcDC, 10, 18, 16711935
    SetPixel m_hCkbSrcDC, 11, 18, 16711935
    SetPixel m_hCkbSrcDC, 12, 18, 12838319
    SetPixel m_hCkbSrcDC, 13, 18, 11394195
    SetPixel m_hCkbSrcDC, 14, 18, 12838319
    SetPixel m_hCkbSrcDC, 15, 18, 16711935
    SetPixel m_hCkbSrcDC, 16, 18, 16711935
    SetPixel m_hCkbSrcDC, 17, 18, 16711935
    SetPixel m_hCkbSrcDC, 0, 19, 16711935
    SetPixel m_hCkbSrcDC, 1, 19, 16711935
    SetPixel m_hCkbSrcDC, 2, 19, 8636510
    SetPixel m_hCkbSrcDC, 3, 19, 5222681
    SetPixel m_hCkbSrcDC, 4, 19, 5222681
    SetPixel m_hCkbSrcDC, 5, 19, 5222681
    SetPixel m_hCkbSrcDC, 6, 19, 8636510
    SetPixel m_hCkbSrcDC, 7, 19, 16711935
    SetPixel m_hCkbSrcDC, 8, 19, 16711935
    SetPixel m_hCkbSrcDC, 9, 19, 16711935
    SetPixel m_hCkbSrcDC, 10, 19, 16711935
    SetPixel m_hCkbSrcDC, 11, 19, 12706734
    SetPixel m_hCkbSrcDC, 12, 19, 10999948
    SetPixel m_hCkbSrcDC, 13, 19, 10999948
    SetPixel m_hCkbSrcDC, 14, 19, 10999948
    SetPixel m_hCkbSrcDC, 15, 19, 12706734
    SetPixel m_hCkbSrcDC, 16, 19, 16711935
    SetPixel m_hCkbSrcDC, 17, 19, 16711935
    SetPixel m_hCkbSrcDC, 0, 20, 16711935
    SetPixel m_hCkbSrcDC, 1, 20, 16711935
    SetPixel m_hCkbSrcDC, 2, 20, 5286941
    SetPixel m_hCkbSrcDC, 3, 20, 4696082
    SetPixel m_hCkbSrcDC, 4, 20, 4696082
    SetPixel m_hCkbSrcDC, 5, 20, 4696082
    SetPixel m_hCkbSrcDC, 6, 20, 5286941
    SetPixel m_hCkbSrcDC, 7, 20, 16711935
    SetPixel m_hCkbSrcDC, 8, 20, 16711935
    SetPixel m_hCkbSrcDC, 9, 20, 16711935
    SetPixel m_hCkbSrcDC, 10, 20, 16711935
    SetPixel m_hCkbSrcDC, 11, 20, 10999182
    SetPixel m_hCkbSrcDC, 12, 20, 10736520
    SetPixel m_hCkbSrcDC, 13, 20, 10736520
    SetPixel m_hCkbSrcDC, 14, 20, 10736520
    SetPixel m_hCkbSrcDC, 15, 20, 10999182
    SetPixel m_hCkbSrcDC, 16, 20, 16711935
    SetPixel m_hCkbSrcDC, 17, 20, 16711935
    SetPixel m_hCkbSrcDC, 0, 21, 16711935
    SetPixel m_hCkbSrcDC, 1, 21, 16711935
    SetPixel m_hCkbSrcDC, 2, 21, 7912020
    SetPixel m_hCkbSrcDC, 3, 21, 4168971
    SetPixel m_hCkbSrcDC, 4, 21, 4168971
    SetPixel m_hCkbSrcDC, 5, 21, 4168971
    SetPixel m_hCkbSrcDC, 6, 21, 7912020
    SetPixel m_hCkbSrcDC, 7, 21, 16711935
    SetPixel m_hCkbSrcDC, 8, 21, 16711935
    SetPixel m_hCkbSrcDC, 9, 21, 16711935
    SetPixel m_hCkbSrcDC, 10, 21, 16711935
    SetPixel m_hCkbSrcDC, 11, 21, 12311721
    SetPixel m_hCkbSrcDC, 12, 21, 10473093
    SetPixel m_hCkbSrcDC, 13, 21, 10473093
    SetPixel m_hCkbSrcDC, 14, 21, 10473093
    SetPixel m_hCkbSrcDC, 15, 21, 12311721
    SetPixel m_hCkbSrcDC, 16, 21, 16711935
    SetPixel m_hCkbSrcDC, 17, 21, 16711935
    SetPixel m_hCkbSrcDC, 0, 22, 16711935
    SetPixel m_hCkbSrcDC, 1, 22, 16711935
    SetPixel m_hCkbSrcDC, 2, 22, 16711935
    SetPixel m_hCkbSrcDC, 3, 22, 7582799
    SetPixel m_hCkbSrcDC, 4, 22, 4233488
    SetPixel m_hCkbSrcDC, 5, 22, 7582799
    SetPixel m_hCkbSrcDC, 6, 22, 16711935
    SetPixel m_hCkbSrcDC, 7, 22, 16711935
    SetPixel m_hCkbSrcDC, 8, 22, 16711935
    SetPixel m_hCkbSrcDC, 9, 22, 16711935
    SetPixel m_hCkbSrcDC, 10, 22, 16711935
    SetPixel m_hCkbSrcDC, 11, 22, 16711935
    SetPixel m_hCkbSrcDC, 12, 22, 12179879
    SetPixel m_hCkbSrcDC, 13, 22, 10472583
    SetPixel m_hCkbSrcDC, 14, 22, 12179879
    SetPixel m_hCkbSrcDC, 15, 22, 16711935
    SetPixel m_hCkbSrcDC, 16, 22, 16711935
    SetPixel m_hCkbSrcDC, 17, 22, 16711935
    SetPixel m_hCkbSrcDC, 0, 23, 16711935
    SetPixel m_hCkbSrcDC, 1, 23, 16711935
    SetPixel m_hCkbSrcDC, 2, 23, 16711935
    SetPixel m_hCkbSrcDC, 3, 23, 16711935
    SetPixel m_hCkbSrcDC, 4, 23, 16711935
    SetPixel m_hCkbSrcDC, 5, 23, 16711935
    SetPixel m_hCkbSrcDC, 6, 23, 16711935
    SetPixel m_hCkbSrcDC, 7, 23, 16711935
    SetPixel m_hCkbSrcDC, 8, 23, 16711935
    SetPixel m_hCkbSrcDC, 9, 23, 16711935
    SetPixel m_hCkbSrcDC, 10, 23, 16711935
    SetPixel m_hCkbSrcDC, 11, 23, 16711935
    SetPixel m_hCkbSrcDC, 12, 23, 16711935
    SetPixel m_hCkbSrcDC, 13, 23, 16711935
    SetPixel m_hCkbSrcDC, 14, 23, 16711935
    SetPixel m_hCkbSrcDC, 15, 23, 16711935
    SetPixel m_hCkbSrcDC, 16, 23, 16711935
    SetPixel m_hCkbSrcDC, 17, 23, 16711935
    SetPixel m_hOpbSrcDC, 0, 0, 16711935
    SetPixel m_hOpbSrcDC, 1, 0, 16711935
    SetPixel m_hOpbSrcDC, 2, 0, 16711935
    SetPixel m_hOpbSrcDC, 3, 0, 12373991
    SetPixel m_hOpbSrcDC, 4, 0, 9810905
    SetPixel m_hOpbSrcDC, 5, 0, 8825043
    SetPixel m_hOpbSrcDC, 6, 0, 8627922
    SetPixel m_hOpbSrcDC, 7, 0, 8825043
    SetPixel m_hOpbSrcDC, 8, 0, 9810905
    SetPixel m_hOpbSrcDC, 9, 0, 12373991
    SetPixel m_hOpbSrcDC, 10, 0, 16711935
    SetPixel m_hOpbSrcDC, 11, 0, 16711935
    SetPixel m_hOpbSrcDC, 12, 0, 16711935
    SetPixel m_hOpbSrcDC, 0, 1, 16711935
    SetPixel m_hOpbSrcDC, 1, 1, 16711935
    SetPixel m_hOpbSrcDC, 2, 1, 10336475
    SetPixel m_hOpbSrcDC, 3, 1, 11125471
    SetPixel m_hOpbSrcDC, 4, 1, 13754348
    SetPixel m_hOpbSrcDC, 5, 1, 14542832
    SetPixel m_hOpbSrcDC, 6, 1, 14674417
    SetPixel m_hOpbSrcDC, 7, 1, 14542832
    SetPixel m_hOpbSrcDC, 8, 1, 13754348
    SetPixel m_hOpbSrcDC, 9, 1, 11125471
    SetPixel m_hOpbSrcDC, 10, 1, 10336475
    SetPixel m_hOpbSrcDC, 11, 1, 16711935
    SetPixel m_hOpbSrcDC, 12, 1, 16711935
    SetPixel m_hOpbSrcDC, 0, 2, 16711935
    SetPixel m_hOpbSrcDC, 1, 2, 10336475
    SetPixel m_hOpbSrcDC, 2, 2, 12177125
    SetPixel m_hOpbSrcDC, 3, 2, 14937331
    SetPixel m_hOpbSrcDC, 4, 2, 15068660
    SetPixel m_hOpbSrcDC, 5, 2, 15199988
    SetPixel m_hOpbSrcDC, 6, 2, 15199988
    SetPixel m_hOpbSrcDC, 7, 2, 15199988
    SetPixel m_hOpbSrcDC, 8, 2, 15068660
    SetPixel m_hOpbSrcDC, 9, 2, 14937331
    SetPixel m_hOpbSrcDC, 10, 2, 12177125
    SetPixel m_hOpbSrcDC, 11, 2, 10336475
    SetPixel m_hOpbSrcDC, 12, 2, 16711935
    SetPixel m_hOpbSrcDC, 0, 3, 12373991
    SetPixel m_hOpbSrcDC, 1, 3, 11125471
    SetPixel m_hOpbSrcDC, 2, 3, 15068660
    SetPixel m_hOpbSrcDC, 3, 3, 15594231
    SetPixel m_hOpbSrcDC, 4, 3, 15988730
    SetPixel m_hOpbSrcDC, 5, 3, 16251387
    SetPixel m_hOpbSrcDC, 6, 3, 16251387
    SetPixel m_hOpbSrcDC, 7, 3, 16251387
    SetPixel m_hOpbSrcDC, 8, 3, 15988730
    SetPixel m_hOpbSrcDC, 9, 3, 15594231
    SetPixel m_hOpbSrcDC, 10, 3, 15068660
    SetPixel m_hOpbSrcDC, 11, 3, 11125471
    SetPixel m_hOpbSrcDC, 12, 3, 12373991
    SetPixel m_hOpbSrcDC, 0, 4, 9810905
    SetPixel m_hOpbSrcDC, 1, 4, 14082799
    SetPixel m_hOpbSrcDC, 2, 4, 15725816
    SetPixel m_hOpbSrcDC, 3, 4, 16382972
    SetPixel m_hOpbSrcDC, 4, 4, 16777215
    SetPixel m_hOpbSrcDC, 5, 4, 16777215
    SetPixel m_hOpbSrcDC, 6, 4, 16777215
    SetPixel m_hOpbSrcDC, 7, 4, 16777215
    SetPixel m_hOpbSrcDC, 8, 4, 16777215
    SetPixel m_hOpbSrcDC, 9, 4, 16382972
    SetPixel m_hOpbSrcDC, 10, 4, 15725816
    SetPixel m_hOpbSrcDC, 11, 4, 14082799
    SetPixel m_hOpbSrcDC, 12, 4, 9810905
    SetPixel m_hOpbSrcDC, 0, 5, 8825043
    SetPixel m_hOpbSrcDC, 1, 5, 15462646
    SetPixel m_hOpbSrcDC, 2, 5, 16382972
    SetPixel m_hOpbSrcDC, 3, 5, 16711423
    SetPixel m_hOpbSrcDC, 4, 5, 16777215
    SetPixel m_hOpbSrcDC, 5, 5, 16777215
    SetPixel m_hOpbSrcDC, 6, 5, 16777215
    SetPixel m_hOpbSrcDC, 7, 5, 16777215
    SetPixel m_hOpbSrcDC, 8, 5, 16777215
    SetPixel m_hOpbSrcDC, 9, 5, 16711423
    SetPixel m_hOpbSrcDC, 10, 5, 16382972
    SetPixel m_hOpbSrcDC, 11, 5, 15462646
    SetPixel m_hOpbSrcDC, 12, 5, 8825043
    SetPixel m_hOpbSrcDC, 0, 6, 8627922
    SetPixel m_hOpbSrcDC, 1, 6, 15988730
    SetPixel m_hOpbSrcDC, 2, 6, 16645886
    SetPixel m_hOpbSrcDC, 3, 6, 16777215
    SetPixel m_hOpbSrcDC, 4, 6, 16777215
    SetPixel m_hOpbSrcDC, 5, 6, 16777215
    SetPixel m_hOpbSrcDC, 6, 6, 16777215
    SetPixel m_hOpbSrcDC, 7, 6, 16777215
    SetPixel m_hOpbSrcDC, 8, 6, 16777215
    SetPixel m_hOpbSrcDC, 9, 6, 16777215
    SetPixel m_hOpbSrcDC, 10, 6, 16645886
    SetPixel m_hOpbSrcDC, 11, 6, 15988730
    SetPixel m_hOpbSrcDC, 12, 6, 8627922
    SetPixel m_hOpbSrcDC, 0, 7, 8825043
    SetPixel m_hOpbSrcDC, 1, 7, 16054266
    SetPixel m_hOpbSrcDC, 2, 7, 16777215
    SetPixel m_hOpbSrcDC, 3, 7, 16777215
    SetPixel m_hOpbSrcDC, 4, 7, 16777215
    SetPixel m_hOpbSrcDC, 5, 7, 16777215
    SetPixel m_hOpbSrcDC, 6, 7, 16777215
    SetPixel m_hOpbSrcDC, 7, 7, 16777215
    SetPixel m_hOpbSrcDC, 8, 7, 16777215
    SetPixel m_hOpbSrcDC, 9, 7, 16777215
    SetPixel m_hOpbSrcDC, 10, 7, 16777215
    SetPixel m_hOpbSrcDC, 11, 7, 16054266
    SetPixel m_hOpbSrcDC, 12, 7, 8825043
    SetPixel m_hOpbSrcDC, 0, 8, 9810905
    SetPixel m_hOpbSrcDC, 1, 8, 15134197
    SetPixel m_hOpbSrcDC, 2, 8, 16777215
    SetPixel m_hOpbSrcDC, 3, 8, 16777215
    SetPixel m_hOpbSrcDC, 4, 8, 16777215
    SetPixel m_hOpbSrcDC, 5, 8, 16777215
    SetPixel m_hOpbSrcDC, 6, 8, 16777215
    SetPixel m_hOpbSrcDC, 7, 8, 16777215
    SetPixel m_hOpbSrcDC, 8, 8, 16777215
    SetPixel m_hOpbSrcDC, 9, 8, 16777215
    SetPixel m_hOpbSrcDC, 10, 8, 16777215
    SetPixel m_hOpbSrcDC, 11, 8, 15134197
    SetPixel m_hOpbSrcDC, 12, 8, 9810905
    SetPixel m_hOpbSrcDC, 0, 9, 12373991
    SetPixel m_hOpbSrcDC, 1, 9, 11782371
    SetPixel m_hOpbSrcDC, 2, 9, 16777215
    SetPixel m_hOpbSrcDC, 3, 9, 16777215
    SetPixel m_hOpbSrcDC, 4, 9, 16777215
    SetPixel m_hOpbSrcDC, 5, 9, 16777215
    SetPixel m_hOpbSrcDC, 6, 9, 16777215
    SetPixel m_hOpbSrcDC, 7, 9, 16777215
    SetPixel m_hOpbSrcDC, 8, 9, 16777215
    SetPixel m_hOpbSrcDC, 9, 9, 16777215
    SetPixel m_hOpbSrcDC, 10, 9, 16777215
    SetPixel m_hOpbSrcDC, 11, 9, 11782371
    SetPixel m_hOpbSrcDC, 12, 9, 12373991
    SetPixel m_hOpbSrcDC, 0, 10, 16711935
    SetPixel m_hOpbSrcDC, 1, 10, 10336475
    SetPixel m_hOpbSrcDC, 2, 10, 13294060
    SetPixel m_hOpbSrcDC, 3, 10, 16777215
    SetPixel m_hOpbSrcDC, 4, 10, 16777215
    SetPixel m_hOpbSrcDC, 5, 10, 16777215
    SetPixel m_hOpbSrcDC, 6, 10, 16777215
    SetPixel m_hOpbSrcDC, 7, 10, 16777215
    SetPixel m_hOpbSrcDC, 8, 10, 16777215
    SetPixel m_hOpbSrcDC, 9, 10, 16777215
    SetPixel m_hOpbSrcDC, 10, 10, 13294060
    SetPixel m_hOpbSrcDC, 11, 10, 10336475
    SetPixel m_hOpbSrcDC, 12, 10, 16711935
    SetPixel m_hOpbSrcDC, 0, 11, 16711935
    SetPixel m_hOpbSrcDC, 1, 11, 16711935
    SetPixel m_hOpbSrcDC, 2, 11, 10336475
    SetPixel m_hOpbSrcDC, 3, 11, 11979493
    SetPixel m_hOpbSrcDC, 4, 11, 15594232
    SetPixel m_hOpbSrcDC, 5, 11, 16580094
    SetPixel m_hOpbSrcDC, 6, 11, 16777215
    SetPixel m_hOpbSrcDC, 7, 11, 16580094
    SetPixel m_hOpbSrcDC, 8, 11, 15594232
    SetPixel m_hOpbSrcDC, 9, 11, 11979493
    SetPixel m_hOpbSrcDC, 10, 11, 10336475
    SetPixel m_hOpbSrcDC, 11, 11, 16711935
    SetPixel m_hOpbSrcDC, 12, 11, 16711935
    SetPixel m_hOpbSrcDC, 0, 12, 16711935
    SetPixel m_hOpbSrcDC, 1, 12, 16711935
    SetPixel m_hOpbSrcDC, 2, 12, 16711935
    SetPixel m_hOpbSrcDC, 3, 12, 12373991
    SetPixel m_hOpbSrcDC, 4, 12, 9810905
    SetPixel m_hOpbSrcDC, 5, 12, 8825043
    SetPixel m_hOpbSrcDC, 6, 12, 8627922
    SetPixel m_hOpbSrcDC, 7, 12, 8825043
    SetPixel m_hOpbSrcDC, 8, 12, 9810905
    SetPixel m_hOpbSrcDC, 9, 12, 12373991
    SetPixel m_hOpbSrcDC, 10, 12, 16711935
    SetPixel m_hOpbSrcDC, 11, 12, 16711935
    SetPixel m_hOpbSrcDC, 12, 12, 16711935
    SetPixel m_hOpbSrcDC, 0, 13, 16711935
    SetPixel m_hOpbSrcDC, 1, 13, 16711935
    SetPixel m_hOpbSrcDC, 2, 13, 16711935
    SetPixel m_hOpbSrcDC, 3, 13, 12373991
    SetPixel m_hOpbSrcDC, 4, 13, 9810905
    SetPixel m_hOpbSrcDC, 5, 13, 8825043
    SetPixel m_hOpbSrcDC, 6, 13, 8627922
    SetPixel m_hOpbSrcDC, 7, 13, 8825043
    SetPixel m_hOpbSrcDC, 8, 13, 9810905
    SetPixel m_hOpbSrcDC, 9, 13, 12373991
    SetPixel m_hOpbSrcDC, 10, 13, 16711935
    SetPixel m_hOpbSrcDC, 11, 13, 16711935
    SetPixel m_hOpbSrcDC, 12, 13, 16711935
    SetPixel m_hOpbSrcDC, 0, 14, 16711935
    SetPixel m_hOpbSrcDC, 1, 14, 16711935
    SetPixel m_hOpbSrcDC, 2, 14, 10336475
    SetPixel m_hOpbSrcDC, 3, 14, 11125471
    SetPixel m_hOpbSrcDC, 4, 14, 12839672
    SetPixel m_hOpbSrcDC, 5, 14, 12839672
    SetPixel m_hOpbSrcDC, 6, 14, 12839672
    SetPixel m_hOpbSrcDC, 7, 14, 12905464
    SetPixel m_hOpbSrcDC, 8, 14, 12971000
    SetPixel m_hOpbSrcDC, 9, 14, 11125471
    SetPixel m_hOpbSrcDC, 10, 14, 10336475
    SetPixel m_hOpbSrcDC, 11, 14, 16711935
    SetPixel m_hOpbSrcDC, 12, 14, 16711935
    SetPixel m_hOpbSrcDC, 0, 15, 16711935
    SetPixel m_hOpbSrcDC, 1, 15, 10336475
    SetPixel m_hOpbSrcDC, 2, 15, 12177125
    SetPixel m_hOpbSrcDC, 3, 15, 12839672
    SetPixel m_hOpbSrcDC, 4, 15, 12839672
    SetPixel m_hOpbSrcDC, 5, 15, 12839672
    SetPixel m_hOpbSrcDC, 6, 15, 12905464
    SetPixel m_hOpbSrcDC, 7, 15, 12971001
    SetPixel m_hOpbSrcDC, 8, 15, 13167865
    SetPixel m_hOpbSrcDC, 9, 15, 13299193
    SetPixel m_hOpbSrcDC, 10, 15, 12177125
    SetPixel m_hOpbSrcDC, 11, 15, 10336475
    SetPixel m_hOpbSrcDC, 12, 15, 16711935
    SetPixel m_hOpbSrcDC, 0, 16, 12373991
    SetPixel m_hOpbSrcDC, 1, 16, 11125471
    SetPixel m_hOpbSrcDC, 2, 16, 12839672
    SetPixel m_hOpbSrcDC, 3, 16, 12839672
    SetPixel m_hOpbSrcDC, 4, 16, 12839672
    SetPixel m_hOpbSrcDC, 5, 16, 12905208
    SetPixel m_hOpbSrcDC, 6, 16, 13036536
    SetPixel m_hOpbSrcDC, 7, 16, 13167865
    SetPixel m_hOpbSrcDC, 8, 16, 13364729
    SetPixel m_hOpbSrcDC, 9, 16, 13496058
    SetPixel m_hOpbSrcDC, 10, 16, 13627386
    SetPixel m_hOpbSrcDC, 11, 16, 11125471
    SetPixel m_hOpbSrcDC, 12, 16, 12373991
    SetPixel m_hOpbSrcDC, 0, 17, 9810905
    SetPixel m_hOpbSrcDC, 1, 17, 12839672
    SetPixel m_hOpbSrcDC, 2, 17, 12839672
    SetPixel m_hOpbSrcDC, 3, 17, 12905208
    SetPixel m_hOpbSrcDC, 4, 17, 12905208
    SetPixel m_hOpbSrcDC, 5, 17, 13036793
    SetPixel m_hOpbSrcDC, 6, 17, 13233401
    SetPixel m_hOpbSrcDC, 7, 17, 13364729
    SetPixel m_hOpbSrcDC, 8, 17, 13496058
    SetPixel m_hOpbSrcDC, 9, 17, 13692922
    SetPixel m_hOpbSrcDC, 10, 17, 13889787
    SetPixel m_hOpbSrcDC, 11, 17, 14086650
    SetPixel m_hOpbSrcDC, 12, 17, 9810905
    SetPixel m_hOpbSrcDC, 0, 18, 8825043
    SetPixel m_hOpbSrcDC, 1, 18, 12839672
    SetPixel m_hOpbSrcDC, 2, 18, 12839672
    SetPixel m_hOpbSrcDC, 3, 18, 12971000
    SetPixel m_hOpbSrcDC, 4, 18, 13102328
    SetPixel m_hOpbSrcDC, 5, 18, 13233657
    SetPixel m_hOpbSrcDC, 6, 18, 13364729
    SetPixel m_hOpbSrcDC, 7, 18, 13561594
    SetPixel m_hOpbSrcDC, 8, 18, 13693178
    SetPixel m_hOpbSrcDC, 9, 18, 13889787
    SetPixel m_hOpbSrcDC, 10, 18, 14086651
    SetPixel m_hOpbSrcDC, 11, 18, 14283515
    SetPixel m_hOpbSrcDC, 12, 18, 8825043
    SetPixel m_hOpbSrcDC, 0, 19, 8627922
    SetPixel m_hOpbSrcDC, 1, 19, 12905208
    SetPixel m_hOpbSrcDC, 2, 19, 12971000
    SetPixel m_hOpbSrcDC, 3, 19, 13102329
    SetPixel m_hOpbSrcDC, 4, 19, 13233657
    SetPixel m_hOpbSrcDC, 5, 19, 13430522
    SetPixel m_hOpbSrcDC, 6, 19, 13627130
    SetPixel m_hOpbSrcDC, 7, 19, 13693178
    SetPixel m_hOpbSrcDC, 8, 19, 13955578
    SetPixel m_hOpbSrcDC, 9, 19, 14086651
    SetPixel m_hOpbSrcDC, 10, 19, 14283515
    SetPixel m_hOpbSrcDC, 11, 19, 14415100
    SetPixel m_hOpbSrcDC, 12, 19, 8627922
    SetPixel m_hOpbSrcDC, 0, 20, 8825043
    SetPixel m_hOpbSrcDC, 1, 20, 12971000
    SetPixel m_hOpbSrcDC, 2, 20, 13102072
    SetPixel m_hOpbSrcDC, 3, 20, 13233401
    SetPixel m_hOpbSrcDC, 4, 20, 13430522
    SetPixel m_hOpbSrcDC, 5, 20, 13561850
    SetPixel m_hOpbSrcDC, 6, 20, 13758715
    SetPixel m_hOpbSrcDC, 7, 20, 13955323
    SetPixel m_hOpbSrcDC, 8, 20, 14152188
    SetPixel m_hOpbSrcDC, 9, 20, 14283772
    SetPixel m_hOpbSrcDC, 10, 20, 14480380
    SetPixel m_hOpbSrcDC, 11, 20, 14611708
    SetPixel m_hOpbSrcDC, 12, 20, 8825043
    SetPixel m_hOpbSrcDC, 0, 21, 9810905
    SetPixel m_hOpbSrcDC, 1, 21, 13102329
    SetPixel m_hOpbSrcDC, 2, 21, 13299193
    SetPixel m_hOpbSrcDC, 3, 21, 13430521
    SetPixel m_hOpbSrcDC, 4, 21, 13627386
    SetPixel m_hOpbSrcDC, 5, 21, 13823994
    SetPixel m_hOpbSrcDC, 6, 21, 13955579
    SetPixel m_hOpbSrcDC, 7, 21, 14217979
    SetPixel m_hOpbSrcDC, 8, 21, 14349307
    SetPixel m_hOpbSrcDC, 9, 21, 14480636
    SetPixel m_hOpbSrcDC, 10, 21, 14611708
    SetPixel m_hOpbSrcDC, 11, 21, 14743037
    SetPixel m_hOpbSrcDC, 12, 21, 9810905
    SetPixel m_hOpbSrcDC, 0, 22, 12373991
    SetPixel m_hOpbSrcDC, 1, 22, 11782371
    SetPixel m_hOpbSrcDC, 2, 22, 13496058
    SetPixel m_hOpbSrcDC, 3, 22, 13692922
    SetPixel m_hOpbSrcDC, 4, 22, 13824251
    SetPixel m_hOpbSrcDC, 5, 22, 14021115
    SetPixel m_hOpbSrcDC, 6, 22, 14217979
    SetPixel m_hOpbSrcDC, 7, 22, 14414844
    SetPixel m_hOpbSrcDC, 8, 22, 14480636
    SetPixel m_hOpbSrcDC, 9, 22, 14677501
    SetPixel m_hOpbSrcDC, 10, 22, 14743036
    SetPixel m_hOpbSrcDC, 11, 22, 11782371
    SetPixel m_hOpbSrcDC, 12, 22, 12373991
    SetPixel m_hOpbSrcDC, 0, 23, 16711935
    SetPixel m_hOpbSrcDC, 1, 23, 10336475
    SetPixel m_hOpbSrcDC, 2, 23, 13294060
    SetPixel m_hOpbSrcDC, 3, 23, 13889787
    SetPixel m_hOpbSrcDC, 4, 23, 14086651
    SetPixel m_hOpbSrcDC, 5, 23, 14217979
    SetPixel m_hOpbSrcDC, 6, 23, 14415100
    SetPixel m_hOpbSrcDC, 7, 23, 14546173
    SetPixel m_hOpbSrcDC, 8, 23, 14677501
    SetPixel m_hOpbSrcDC, 9, 23, 14808573
    SetPixel m_hOpbSrcDC, 10, 23, 13294060
    SetPixel m_hOpbSrcDC, 11, 23, 10336475
    SetPixel m_hOpbSrcDC, 12, 23, 16711935
    SetPixel m_hOpbSrcDC, 0, 24, 16711935
    SetPixel m_hOpbSrcDC, 1, 24, 16711935
    SetPixel m_hOpbSrcDC, 2, 24, 10336475
    SetPixel m_hOpbSrcDC, 3, 24, 11979493
    SetPixel m_hOpbSrcDC, 4, 24, 14283772
    SetPixel m_hOpbSrcDC, 5, 24, 14414844
    SetPixel m_hOpbSrcDC, 6, 24, 14546172
    SetPixel m_hOpbSrcDC, 7, 24, 14677500
    SetPixel m_hOpbSrcDC, 8, 24, 14808572
    SetPixel m_hOpbSrcDC, 9, 24, 11979493
    SetPixel m_hOpbSrcDC, 10, 24, 10336475
    SetPixel m_hOpbSrcDC, 11, 24, 16711935
    SetPixel m_hOpbSrcDC, 12, 24, 16711935
    SetPixel m_hOpbSrcDC, 0, 25, 16711935
    SetPixel m_hOpbSrcDC, 1, 25, 16711935
    SetPixel m_hOpbSrcDC, 2, 25, 16711935
    SetPixel m_hOpbSrcDC, 3, 25, 12373991
    SetPixel m_hOpbSrcDC, 4, 25, 9810905
    SetPixel m_hOpbSrcDC, 5, 25, 8825043
    SetPixel m_hOpbSrcDC, 6, 25, 8627922
    SetPixel m_hOpbSrcDC, 7, 25, 8825043
    SetPixel m_hOpbSrcDC, 8, 25, 9810905
    SetPixel m_hOpbSrcDC, 9, 25, 12373991
    SetPixel m_hOpbSrcDC, 10, 25, 16711935
    SetPixel m_hOpbSrcDC, 11, 25, 16711935
    SetPixel m_hOpbSrcDC, 12, 25, 16711935
    SetPixel m_hOpbSrcDC, 0, 26, 16711935
    SetPixel m_hOpbSrcDC, 1, 26, 16711935
    SetPixel m_hOpbSrcDC, 2, 26, 16711935
    SetPixel m_hOpbSrcDC, 3, 26, 10796253
    SetPixel m_hOpbSrcDC, 4, 26, 7772618
    SetPixel m_hOpbSrcDC, 5, 26, 6785986
    SetPixel m_hOpbSrcDC, 6, 26, 6588865
    SetPixel m_hOpbSrcDC, 7, 26, 6785986
    SetPixel m_hOpbSrcDC, 8, 26, 7772618
    SetPixel m_hOpbSrcDC, 9, 26, 10796253
    SetPixel m_hOpbSrcDC, 10, 26, 16711935
    SetPixel m_hOpbSrcDC, 11, 26, 16711935
    SetPixel m_hOpbSrcDC, 12, 26, 16711935
    SetPixel m_hOpbSrcDC, 0, 27, 16711935
    SetPixel m_hOpbSrcDC, 1, 27, 16711935
    SetPixel m_hOpbSrcDC, 2, 27, 8363981
    SetPixel m_hOpbSrcDC, 3, 27, 9219283
    SetPixel m_hOpbSrcDC, 4, 27, 13036792
    SetPixel m_hOpbSrcDC, 5, 27, 13299193
    SetPixel m_hOpbSrcDC, 6, 27, 13496314
    SetPixel m_hOpbSrcDC, 7, 27, 13758458
    SetPixel m_hOpbSrcDC, 8, 27, 13955581
    SetPixel m_hOpbSrcDC, 9, 27, 10073048
    SetPixel m_hOpbSrcDC, 10, 27, 8363981
    SetPixel m_hOpbSrcDC, 11, 27, 16711935
    SetPixel m_hOpbSrcDC, 12, 27, 16711935
    SetPixel m_hOpbSrcDC, 0, 28, 16711935
    SetPixel m_hOpbSrcDC, 1, 28, 8363981
    SetPixel m_hOpbSrcDC, 2, 28, 10533850
    SetPixel m_hOpbSrcDC, 3, 28, 12380408
    SetPixel m_hOpbSrcDC, 4, 28, 12774137
    SetPixel m_hOpbSrcDC, 5, 28, 13036793
    SetPixel m_hOpbSrcDC, 6, 28, 13299193
    SetPixel m_hOpbSrcDC, 7, 28, 13561594
    SetPixel m_hOpbSrcDC, 8, 28, 13758458
    SetPixel m_hOpbSrcDC, 9, 28, 13955578
    SetPixel m_hOpbSrcDC, 10, 28, 11913957
    SetPixel m_hOpbSrcDC, 11, 28, 8363981
    SetPixel m_hOpbSrcDC, 12, 28, 16711935
    SetPixel m_hOpbSrcDC, 0, 29, 10796253
    SetPixel m_hOpbSrcDC, 1, 29, 9219283
    SetPixel m_hOpbSrcDC, 2, 29, 11921143
    SetPixel m_hOpbSrcDC, 3, 29, 12249080
    SetPixel m_hOpbSrcDC, 4, 29, 12511480
    SetPixel m_hOpbSrcDC, 5, 29, 12774137
    SetPixel m_hOpbSrcDC, 6, 29, 13036793
    SetPixel m_hOpbSrcDC, 7, 29, 13299450
    SetPixel m_hOpbSrcDC, 8, 29, 13561850
    SetPixel m_hOpbSrcDC, 9, 29, 13890045
    SetPixel m_hOpbSrcDC, 10, 29, 14021117
    SetPixel m_hOpbSrcDC, 11, 29, 10270170
    SetPixel m_hOpbSrcDC, 12, 29, 10796253
    SetPixel m_hOpbSrcDC, 0, 30, 7772618
    SetPixel m_hOpbSrcDC, 1, 30, 11592437
    SetPixel m_hOpbSrcDC, 2, 30, 11789815
    SetPixel m_hOpbSrcDC, 3, 30, 12117751
    SetPixel m_hOpbSrcDC, 4, 30, 12249080
    SetPixel m_hOpbSrcDC, 5, 30, 12511992
    SetPixel m_hOpbSrcDC, 6, 30, 12840184
    SetPixel m_hOpbSrcDC, 7, 30, 13102330
    SetPixel m_hOpbSrcDC, 8, 30, 13430521
    SetPixel m_hOpbSrcDC, 9, 30, 13561850
    SetPixel m_hOpbSrcDC, 10, 30, 13890045
    SetPixel m_hOpbSrcDC, 11, 30, 14021114
    SetPixel m_hOpbSrcDC, 12, 30, 7772618
    SetPixel m_hOpbSrcDC, 0, 31, 6785986
    SetPixel m_hOpbSrcDC, 1, 31, 11461365
    SetPixel m_hOpbSrcDC, 2, 31, 11592439
    SetPixel m_hOpbSrcDC, 3, 31, 11789815
    SetPixel m_hOpbSrcDC, 4, 31, 12117751
    SetPixel m_hOpbSrcDC, 5, 31, 12314616
    SetPixel m_hOpbSrcDC, 6, 31, 12511992
    SetPixel m_hOpbSrcDC, 7, 31, 12839673
    SetPixel m_hOpbSrcDC, 8, 31, 13168121
    SetPixel m_hOpbSrcDC, 9, 31, 13496058
    SetPixel m_hOpbSrcDC, 10, 31, 13627389
    SetPixel m_hOpbSrcDC, 11, 31, 13890042
    SetPixel m_hOpbSrcDC, 12, 31, 6785986
    SetPixel m_hOpbSrcDC, 0, 32, 6588865
    SetPixel m_hOpbSrcDC, 1, 32, 11395573
    SetPixel m_hOpbSrcDC, 2, 32, 11461365
    SetPixel m_hOpbSrcDC, 3, 32, 11657973
    SetPixel m_hOpbSrcDC, 4, 32, 11855351
    SetPixel m_hOpbSrcDC, 5, 32, 12117751
    SetPixel m_hOpbSrcDC, 6, 32, 12380152
    SetPixel m_hOpbSrcDC, 7, 32, 12577529
    SetPixel m_hOpbSrcDC, 8, 32, 12840185
    SetPixel m_hOpbSrcDC, 9, 32, 13168121
    SetPixel m_hOpbSrcDC, 10, 32, 13496314
    SetPixel m_hOpbSrcDC, 11, 32, 13627386
    SetPixel m_hOpbSrcDC, 12, 32, 6588865
    SetPixel m_hOpbSrcDC, 0, 33, 6785986
    SetPixel m_hOpbSrcDC, 1, 33, 11395573
    SetPixel m_hOpbSrcDC, 2, 33, 11395573
    SetPixel m_hOpbSrcDC, 3, 33, 11461109
    SetPixel m_hOpbSrcDC, 4, 33, 11658743
    SetPixel m_hOpbSrcDC, 5, 33, 11855607
    SetPixel m_hOpbSrcDC, 6, 33, 12183544
    SetPixel m_hOpbSrcDC, 7, 33, 12314872
    SetPixel m_hOpbSrcDC, 8, 33, 12708088
    SetPixel m_hOpbSrcDC, 9, 33, 12971257
    SetPixel m_hOpbSrcDC, 10, 33, 13168121
    SetPixel m_hOpbSrcDC, 11, 33, 13496058
    SetPixel m_hOpbSrcDC, 12, 33, 6785986
    SetPixel m_hOpbSrcDC, 0, 34, 7772618
    SetPixel m_hOpbSrcDC, 1, 34, 11395573
    SetPixel m_hOpbSrcDC, 2, 34, 11395573
    SetPixel m_hOpbSrcDC, 3, 34, 11395573
    SetPixel m_hOpbSrcDC, 4, 34, 11461109
    SetPixel m_hOpbSrcDC, 5, 34, 11724277
    SetPixel m_hOpbSrcDC, 6, 34, 11855607
    SetPixel m_hOpbSrcDC, 7, 34, 12183544
    SetPixel m_hOpbSrcDC, 8, 34, 12380408
    SetPixel m_hOpbSrcDC, 9, 34, 12708601
    SetPixel m_hOpbSrcDC, 10, 34, 13036793
    SetPixel m_hOpbSrcDC, 11, 34, 13299450
    SetPixel m_hOpbSrcDC, 12, 34, 7772618
    SetPixel m_hOpbSrcDC, 0, 35, 10796253
    SetPixel m_hOpbSrcDC, 1, 35, 9219283
    SetPixel m_hOpbSrcDC, 2, 35, 11395573
    SetPixel m_hOpbSrcDC, 3, 35, 11395573
    SetPixel m_hOpbSrcDC, 4, 35, 11461109
    SetPixel m_hOpbSrcDC, 5, 35, 11592437
    SetPixel m_hOpbSrcDC, 6, 35, 11724279
    SetPixel m_hOpbSrcDC, 7, 35, 11855351
    SetPixel m_hOpbSrcDC, 8, 35, 12183543
    SetPixel m_hOpbSrcDC, 9, 35, 12511480
    SetPixel m_hOpbSrcDC, 10, 35, 12774137
    SetPixel m_hOpbSrcDC, 11, 35, 10270170
    SetPixel m_hOpbSrcDC, 12, 35, 10796253
    SetPixel m_hOpbSrcDC, 0, 36, 16711935
    SetPixel m_hOpbSrcDC, 1, 36, 8363981
    SetPixel m_hOpbSrcDC, 2, 36, 10533850
    SetPixel m_hOpbSrcDC, 3, 36, 11395573
    SetPixel m_hOpbSrcDC, 4, 36, 11395573
    SetPixel m_hOpbSrcDC, 5, 36, 11395573
    SetPixel m_hOpbSrcDC, 6, 36, 11592437
    SetPixel m_hOpbSrcDC, 7, 36, 11723509
    SetPixel m_hOpbSrcDC, 8, 36, 11921143
    SetPixel m_hOpbSrcDC, 9, 36, 12249080
    SetPixel m_hOpbSrcDC, 10, 36, 11913957
    SetPixel m_hOpbSrcDC, 11, 36, 8363981
    SetPixel m_hOpbSrcDC, 12, 36, 16711935
    SetPixel m_hOpbSrcDC, 0, 37, 16711935
    SetPixel m_hOpbSrcDC, 1, 37, 16711935
    SetPixel m_hOpbSrcDC, 2, 37, 8363981
    SetPixel m_hOpbSrcDC, 3, 37, 9219283
    SetPixel m_hOpbSrcDC, 4, 37, 11395573
    SetPixel m_hOpbSrcDC, 5, 37, 11395573
    SetPixel m_hOpbSrcDC, 6, 37, 11461109
    SetPixel m_hOpbSrcDC, 7, 37, 11592437
    SetPixel m_hOpbSrcDC, 8, 37, 11724279
    SetPixel m_hOpbSrcDC, 9, 37, 10073048
    SetPixel m_hOpbSrcDC, 10, 37, 8363981
    SetPixel m_hOpbSrcDC, 11, 37, 16711935
    SetPixel m_hOpbSrcDC, 12, 37, 16711935
    SetPixel m_hOpbSrcDC, 0, 38, 16711935
    SetPixel m_hOpbSrcDC, 1, 38, 16711935
    SetPixel m_hOpbSrcDC, 2, 38, 16711935
    SetPixel m_hOpbSrcDC, 3, 38, 10796253
    SetPixel m_hOpbSrcDC, 4, 38, 7772618
    SetPixel m_hOpbSrcDC, 5, 38, 6785986
    SetPixel m_hOpbSrcDC, 6, 38, 6588865
    SetPixel m_hOpbSrcDC, 7, 38, 6785986
    SetPixel m_hOpbSrcDC, 8, 38, 7772618
    SetPixel m_hOpbSrcDC, 9, 38, 10796253
    SetPixel m_hOpbSrcDC, 10, 38, 16711935
    SetPixel m_hOpbSrcDC, 11, 38, 16711935
    SetPixel m_hOpbSrcDC, 12, 38, 16711935
    SetPixel m_hOpbSrcDC, 0, 39, 16711935
    SetPixel m_hOpbSrcDC, 1, 39, 16711935
    SetPixel m_hOpbSrcDC, 2, 39, 16711935
    SetPixel m_hOpbSrcDC, 3, 39, 14542835
    SetPixel m_hOpbSrcDC, 4, 39, 13294060
    SetPixel m_hOpbSrcDC, 5, 39, 12768233
    SetPixel m_hOpbSrcDC, 6, 39, 12702440
    SetPixel m_hOpbSrcDC, 7, 39, 12768233
    SetPixel m_hOpbSrcDC, 8, 39, 13294060
    SetPixel m_hOpbSrcDC, 9, 39, 14542835
    SetPixel m_hOpbSrcDC, 10, 39, 16711935
    SetPixel m_hOpbSrcDC, 11, 39, 16711935
    SetPixel m_hOpbSrcDC, 12, 39, 16711935
    SetPixel m_hOpbSrcDC, 0, 40, 16711935
    SetPixel m_hOpbSrcDC, 1, 40, 16711935
    SetPixel m_hOpbSrcDC, 2, 40, 13556717
    SetPixel m_hOpbSrcDC, 3, 40, 13951215
    SetPixel m_hOpbSrcDC, 4, 40, 15265781
    SetPixel m_hOpbSrcDC, 5, 40, 15660023
    SetPixel m_hOpbSrcDC, 6, 40, 15725816
    SetPixel m_hOpbSrcDC, 7, 40, 15660023
    SetPixel m_hOpbSrcDC, 8, 40, 15265781
    SetPixel m_hOpbSrcDC, 9, 40, 13951215
    SetPixel m_hOpbSrcDC, 10, 40, 13556717
    SetPixel m_hOpbSrcDC, 11, 40, 16711935
    SetPixel m_hOpbSrcDC, 12, 40, 16711935
    SetPixel m_hOpbSrcDC, 0, 41, 16711935
    SetPixel m_hOpbSrcDC, 1, 41, 13556717
    SetPixel m_hOpbSrcDC, 2, 41, 14477042
    SetPixel m_hOpbSrcDC, 3, 41, 15857145
    SetPixel m_hOpbSrcDC, 4, 41, 15922937
    SetPixel m_hOpbSrcDC, 5, 41, 15988473
    SetPixel m_hOpbSrcDC, 6, 41, 15988473
    SetPixel m_hOpbSrcDC, 7, 41, 15988473
    SetPixel m_hOpbSrcDC, 8, 41, 15922937
    SetPixel m_hOpbSrcDC, 9, 41, 15857145
    SetPixel m_hOpbSrcDC, 10, 41, 14477042
    SetPixel m_hOpbSrcDC, 11, 41, 13556717
    SetPixel m_hOpbSrcDC, 12, 41, 16711935
    SetPixel m_hOpbSrcDC, 0, 42, 14542835
    SetPixel m_hOpbSrcDC, 1, 42, 13951215
    SetPixel m_hOpbSrcDC, 2, 42, 15922937
    SetPixel m_hOpbSrcDC, 3, 42, 16185595
    SetPixel m_hOpbSrcDC, 4, 42, 16382972
    SetPixel m_hOpbSrcDC, 5, 42, 16514301
    SetPixel m_hOpbSrcDC, 6, 42, 16514301
    SetPixel m_hOpbSrcDC, 7, 42, 16514301
    SetPixel m_hOpbSrcDC, 8, 42, 16382972
    SetPixel m_hOpbSrcDC, 9, 42, 16185595
    SetPixel m_hOpbSrcDC, 10, 42, 15922937
    SetPixel m_hOpbSrcDC, 11, 42, 13951215
    SetPixel m_hOpbSrcDC, 12, 42, 14542835
    SetPixel m_hOpbSrcDC, 0, 43, 13294060
    SetPixel m_hOpbSrcDC, 1, 43, 15397111
    SetPixel m_hOpbSrcDC, 2, 43, 16251387
    SetPixel m_hOpbSrcDC, 3, 43, 16580093
    SetPixel m_hOpbSrcDC, 4, 43, 16711422
    SetPixel m_hOpbSrcDC, 5, 43, 16777215
    SetPixel m_hOpbSrcDC, 6, 43, 16777215
    SetPixel m_hOpbSrcDC, 7, 43, 16777215
    SetPixel m_hOpbSrcDC, 8, 43, 16711422
    SetPixel m_hOpbSrcDC, 9, 43, 16580093
    SetPixel m_hOpbSrcDC, 10, 43, 16251387
    SetPixel m_hOpbSrcDC, 11, 43, 15397111
    SetPixel m_hOpbSrcDC, 12, 43, 13294060
    SetPixel m_hOpbSrcDC, 0, 44, 12768233
    SetPixel m_hOpbSrcDC, 1, 44, 16119802
    SetPixel m_hOpbSrcDC, 2, 44, 16580093
    SetPixel m_hOpbSrcDC, 3, 44, 16711423
    SetPixel m_hOpbSrcDC, 4, 44, 16777215
    SetPixel m_hOpbSrcDC, 5, 44, 16777215
    SetPixel m_hOpbSrcDC, 6, 44, 16777215
    SetPixel m_hOpbSrcDC, 7, 44, 16777215
    SetPixel m_hOpbSrcDC, 8, 44, 16777215
    SetPixel m_hOpbSrcDC, 9, 44, 16711423
    SetPixel m_hOpbSrcDC, 10, 44, 16580093
    SetPixel m_hOpbSrcDC, 11, 44, 16119802
    SetPixel m_hOpbSrcDC, 12, 44, 12768233
    SetPixel m_hOpbSrcDC, 0, 45, 12702440
    SetPixel m_hOpbSrcDC, 1, 45, 16382972
    SetPixel m_hOpbSrcDC, 2, 45, 16711422
    SetPixel m_hOpbSrcDC, 3, 45, 16777215
    SetPixel m_hOpbSrcDC, 4, 45, 16777215
    SetPixel m_hOpbSrcDC, 5, 45, 16777215
    SetPixel m_hOpbSrcDC, 6, 45, 16777215
    SetPixel m_hOpbSrcDC, 7, 45, 16777215
    SetPixel m_hOpbSrcDC, 8, 45, 16777215
    SetPixel m_hOpbSrcDC, 9, 45, 16777215
    SetPixel m_hOpbSrcDC, 10, 45, 16711422
    SetPixel m_hOpbSrcDC, 11, 45, 16382972
    SetPixel m_hOpbSrcDC, 12, 45, 12702440
    SetPixel m_hOpbSrcDC, 0, 46, 12768233
    SetPixel m_hOpbSrcDC, 1, 46, 16382972
    SetPixel m_hOpbSrcDC, 2, 46, 16777215
    SetPixel m_hOpbSrcDC, 3, 46, 16777215
    SetPixel m_hOpbSrcDC, 4, 46, 16777215
    SetPixel m_hOpbSrcDC, 5, 46, 16777215
    SetPixel m_hOpbSrcDC, 6, 46, 16777215
    SetPixel m_hOpbSrcDC, 7, 46, 16777215
    SetPixel m_hOpbSrcDC, 8, 46, 16777215
    SetPixel m_hOpbSrcDC, 9, 46, 16777215
    SetPixel m_hOpbSrcDC, 10, 46, 16777215
    SetPixel m_hOpbSrcDC, 11, 46, 16382972
    SetPixel m_hOpbSrcDC, 12, 46, 12768233
    SetPixel m_hOpbSrcDC, 0, 47, 13294060
    SetPixel m_hOpbSrcDC, 1, 47, 15922938
    SetPixel m_hOpbSrcDC, 2, 47, 16777215
    SetPixel m_hOpbSrcDC, 3, 47, 16777215
    SetPixel m_hOpbSrcDC, 4, 47, 16777215
    SetPixel m_hOpbSrcDC, 5, 47, 16777215
    SetPixel m_hOpbSrcDC, 6, 47, 16777215
    SetPixel m_hOpbSrcDC, 7, 47, 16777215
    SetPixel m_hOpbSrcDC, 8, 47, 16777215
    SetPixel m_hOpbSrcDC, 9, 47, 16777215
    SetPixel m_hOpbSrcDC, 10, 47, 16777215
    SetPixel m_hOpbSrcDC, 11, 47, 15922938
    SetPixel m_hOpbSrcDC, 12, 47, 13294060
    SetPixel m_hOpbSrcDC, 0, 48, 14542835
    SetPixel m_hOpbSrcDC, 1, 48, 14279665
    SetPixel m_hOpbSrcDC, 2, 48, 16777215
    SetPixel m_hOpbSrcDC, 3, 48, 16777215
    SetPixel m_hOpbSrcDC, 4, 48, 16777215
    SetPixel m_hOpbSrcDC, 5, 48, 16777215
    SetPixel m_hOpbSrcDC, 6, 48, 16777215
    SetPixel m_hOpbSrcDC, 7, 48, 16777215
    SetPixel m_hOpbSrcDC, 8, 48, 16777215
    SetPixel m_hOpbSrcDC, 9, 48, 16777215
    SetPixel m_hOpbSrcDC, 10, 48, 16777215
    SetPixel m_hOpbSrcDC, 11, 48, 14279665
    SetPixel m_hOpbSrcDC, 12, 48, 14542835
    SetPixel m_hOpbSrcDC, 0, 49, 16711935
    SetPixel m_hOpbSrcDC, 1, 49, 13556717
    SetPixel m_hOpbSrcDC, 2, 49, 15002869
    SetPixel m_hOpbSrcDC, 3, 49, 16777215
    SetPixel m_hOpbSrcDC, 4, 49, 16777215
    SetPixel m_hOpbSrcDC, 5, 49, 16777215
    SetPixel m_hOpbSrcDC, 6, 49, 16777215
    SetPixel m_hOpbSrcDC, 7, 49, 16777215
    SetPixel m_hOpbSrcDC, 8, 49, 16777215
    SetPixel m_hOpbSrcDC, 9, 49, 16777215
    SetPixel m_hOpbSrcDC, 10, 49, 15002869
    SetPixel m_hOpbSrcDC, 11, 49, 13556717
    SetPixel m_hOpbSrcDC, 12, 49, 16711935
    SetPixel m_hOpbSrcDC, 0, 50, 16711935
    SetPixel m_hOpbSrcDC, 1, 50, 16711935
    SetPixel m_hOpbSrcDC, 2, 50, 13556717
    SetPixel m_hOpbSrcDC, 3, 50, 14345458
    SetPixel m_hOpbSrcDC, 4, 50, 16185595
    SetPixel m_hOpbSrcDC, 5, 50, 16645886
    SetPixel m_hOpbSrcDC, 6, 50, 16777215
    SetPixel m_hOpbSrcDC, 7, 50, 16645886
    SetPixel m_hOpbSrcDC, 8, 50, 16185595
    SetPixel m_hOpbSrcDC, 9, 50, 14345458
    SetPixel m_hOpbSrcDC, 10, 50, 13556717
    SetPixel m_hOpbSrcDC, 11, 50, 16711935
    SetPixel m_hOpbSrcDC, 12, 50, 16711935
    SetPixel m_hOpbSrcDC, 0, 51, 16711935
    SetPixel m_hOpbSrcDC, 1, 51, 16711935
    SetPixel m_hOpbSrcDC, 2, 51, 16711935
    SetPixel m_hOpbSrcDC, 3, 51, 14542835
    SetPixel m_hOpbSrcDC, 4, 51, 13294060
    SetPixel m_hOpbSrcDC, 5, 51, 12768233
    SetPixel m_hOpbSrcDC, 6, 51, 12702440
    SetPixel m_hOpbSrcDC, 7, 51, 12768233
    SetPixel m_hOpbSrcDC, 8, 51, 13294060
    SetPixel m_hOpbSrcDC, 9, 51, 14542835
    SetPixel m_hOpbSrcDC, 10, 51, 16711935
    SetPixel m_hOpbSrcDC, 11, 51, 16711935
    SetPixel m_hOpbSrcDC, 12, 51, 16711935
    SetPixel m_hHdbSrcDC, 0, 0, 15398398
    SetPixel m_hHdbSrcDC, 1, 0, 15398398
    SetPixel m_hHdbSrcDC, 2, 0, 15398398
    SetPixel m_hHdbSrcDC, 3, 0, 12445437
    SetPixel m_hHdbSrcDC, 4, 0, 13889278
    SetPixel m_hHdbSrcDC, 5, 0, 12445437
    SetPixel m_hHdbSrcDC, 6, 0, 9951227
    SetPixel m_hHdbSrcDC, 7, 0, 9951227
    SetPixel m_hHdbSrcDC, 8, 0, 9951227
    SetPixel m_hHdbSrcDC, 0, 1, 15332606
    SetPixel m_hHdbSrcDC, 1, 1, 15332606
    SetPixel m_hHdbSrcDC, 2, 1, 15332606
    SetPixel m_hHdbSrcDC, 3, 1, 11001596
    SetPixel m_hHdbSrcDC, 4, 1, 14086143
    SetPixel m_hHdbSrcDC, 5, 1, 11001596
    SetPixel m_hHdbSrcDC, 6, 1, 8310522
    SetPixel m_hHdbSrcDC, 7, 1, 10476799
    SetPixel m_hHdbSrcDC, 8, 1, 8310522
    SetPixel m_hHdbSrcDC, 0, 2, 15332605
    SetPixel m_hHdbSrcDC, 1, 2, 15266813
    SetPixel m_hHdbSrcDC, 2, 2, 14741498
    SetPixel m_hHdbSrcDC, 3, 2, 9426172
    SetPixel m_hHdbSrcDC, 4, 2, 13889279
    SetPixel m_hHdbSrcDC, 5, 2, 9163257
    SetPixel m_hHdbSrcDC, 6, 2, 6735354
    SetPixel m_hHdbSrcDC, 7, 2, 10542591
    SetPixel m_hHdbSrcDC, 8, 2, 6603767
    SetPixel m_hHdbSrcDC, 0, 3, 15332349
    SetPixel m_hHdbSrcDC, 1, 3, 15135485
    SetPixel m_hHdbSrcDC, 2, 3, 14084343
    SetPixel m_hHdbSrcDC, 3, 3, 7851003
    SetPixel m_hHdbSrcDC, 4, 3, 13692415
    SetPixel m_hHdbSrcDC, 5, 3, 7456245
    SetPixel m_hHdbSrcDC, 6, 3, 6210298
    SetPixel m_hHdbSrcDC, 7, 3, 10673919
    SetPixel m_hHdbSrcDC, 8, 3, 6012404
    SetPixel m_hHdbSrcDC, 0, 4, 15463933
    SetPixel m_hHdbSrcDC, 1, 4, 15004156
    SetPixel m_hHdbSrcDC, 2, 4, 12376559
    SetPixel m_hHdbSrcDC, 3, 4, 6669818
    SetPixel m_hHdbSrcDC, 4, 4, 13430015
    SetPixel m_hHdbSrcDC, 5, 4, 6011117
    SetPixel m_hHdbSrcDC, 6, 4, 6210298
    SetPixel m_hHdbSrcDC, 7, 4, 10805247
    SetPixel m_hHdbSrcDC, 8, 4, 5682669
    SetPixel m_hHdbSrcDC, 0, 5, 15529469
    SetPixel m_hHdbSrcDC, 1, 5, 14872828
    SetPixel m_hHdbSrcDC, 2, 5, 11457003
    SetPixel m_hHdbSrcDC, 3, 5, 6210298
    SetPixel m_hHdbSrcDC, 4, 5, 13167614
    SetPixel m_hHdbSrcDC, 5, 5, 5550570
    SetPixel m_hHdbSrcDC, 6, 5, 6210298
    SetPixel m_hHdbSrcDC, 7, 5, 10936575
    SetPixel m_hHdbSrcDC, 8, 5, 5550570
    SetPixel m_hHdbSrcDC, 0, 6, 15595261
    SetPixel m_hHdbSrcDC, 1, 6, 14741499
    SetPixel m_hHdbSrcDC, 2, 6, 10471398
    SetPixel m_hHdbSrcDC, 3, 6, 6210298
    SetPixel m_hHdbSrcDC, 4, 6, 12904958
    SetPixel m_hHdbSrcDC, 5, 6, 5352677
    SetPixel m_hHdbSrcDC, 6, 6, 6210298
    SetPixel m_hHdbSrcDC, 7, 6, 11068159
    SetPixel m_hHdbSrcDC, 8, 6, 5352677
    SetPixel m_hHdbSrcDC, 0, 7, 15726333
    SetPixel m_hHdbSrcDC, 1, 7, 14675451
    SetPixel m_hHdbSrcDC, 2, 7, 9748450
    SetPixel m_hHdbSrcDC, 3, 7, 6210298
    SetPixel m_hHdbSrcDC, 4, 7, 12708094
    SetPixel m_hHdbSrcDC, 5, 7, 5220577
    SetPixel m_hHdbSrcDC, 6, 7, 6210298
    SetPixel m_hHdbSrcDC, 7, 7, 11265023
    SetPixel m_hHdbSrcDC, 8, 7, 5220577
    SetPixel m_hHdbSrcDC, 0, 8, 15857661
    SetPixel m_hHdbSrcDC, 1, 8, 14544122
    SetPixel m_hHdbSrcDC, 2, 8, 8566236
    SetPixel m_hHdbSrcDC, 3, 8, 6210298
    SetPixel m_hHdbSrcDC, 4, 8, 12445694
    SetPixel m_hHdbSrcDC, 5, 8, 5022684
    SetPixel m_hHdbSrcDC, 6, 8, 6210298
    SetPixel m_hHdbSrcDC, 7, 8, 11396607
    SetPixel m_hHdbSrcDC, 8, 8, 5022684
    SetPixel m_hHdbSrcDC, 0, 9, 15595003
    SetPixel m_hHdbSrcDC, 1, 9, 13886966
    SetPixel m_hHdbSrcDC, 2, 9, 8040665
    SetPixel m_hHdbSrcDC, 3, 9, 6210298
    SetPixel m_hHdbSrcDC, 4, 9, 11526653
    SetPixel m_hHdbSrcDC, 5, 9, 4956891
    SetPixel m_hHdbSrcDC, 6, 9, 6210298
    SetPixel m_hHdbSrcDC, 7, 9, 11593727
    SetPixel m_hHdbSrcDC, 8, 9, 4956891
    SetPixel m_hHdbSrcDC, 0, 10, 15594747
    SetPixel m_hHdbSrcDC, 1, 10, 13886966
    SetPixel m_hHdbSrcDC, 2, 10, 8040665
    SetPixel m_hHdbSrcDC, 3, 10, 6210298
    SetPixel m_hHdbSrcDC, 4, 10, 11460861
    SetPixel m_hHdbSrcDC, 5, 10, 4956891
    SetPixel m_hHdbSrcDC, 6, 10, 6210298
    SetPixel m_hHdbSrcDC, 7, 10, 11790591
    SetPixel m_hHdbSrcDC, 8, 10, 4956891
    SetPixel m_hHdbSrcDC, 0, 11, 15463419
    SetPixel m_hHdbSrcDC, 1, 11, 13886966
    SetPixel m_hHdbSrcDC, 2, 11, 8237786
    SetPixel m_hHdbSrcDC, 3, 11, 6210298
    SetPixel m_hHdbSrcDC, 4, 11, 11461117
    SetPixel m_hHdbSrcDC, 5, 11, 5022684
    SetPixel m_hHdbSrcDC, 6, 11, 6210298
    SetPixel m_hHdbSrcDC, 7, 11, 11987711
    SetPixel m_hHdbSrcDC, 8, 11, 5022684
    SetPixel m_hHdbSrcDC, 0, 12, 15069434
    SetPixel m_hHdbSrcDC, 1, 12, 13886966
    SetPixel m_hHdbSrcDC, 2, 12, 9222879
    SetPixel m_hHdbSrcDC, 3, 12, 6210298
    SetPixel m_hHdbSrcDC, 4, 12, 11526653
    SetPixel m_hHdbSrcDC, 5, 12, 5155041
    SetPixel m_hHdbSrcDC, 6, 12, 6210298
    SetPixel m_hHdbSrcDC, 7, 12, 12184831
    SetPixel m_hHdbSrcDC, 8, 12, 5155041
    SetPixel m_hHdbSrcDC, 0, 13, 14872313
    SetPixel m_hHdbSrcDC, 1, 13, 13886966
    SetPixel m_hHdbSrcDC, 2, 13, 9945571
    SetPixel m_hHdbSrcDC, 3, 13, 6210298
    SetPixel m_hHdbSrcDC, 4, 13, 11526653
    SetPixel m_hHdbSrcDC, 5, 13, 5352677
    SetPixel m_hHdbSrcDC, 6, 13, 6669818
    SetPixel m_hHdbSrcDC, 7, 13, 12381695
    SetPixel m_hHdbSrcDC, 8, 13, 5615077
    SetPixel m_hHdbSrcDC, 0, 14, 14609656
    SetPixel m_hHdbSrcDC, 1, 14, 13886966
    SetPixel m_hHdbSrcDC, 2, 14, 10733798
    SetPixel m_hHdbSrcDC, 3, 14, 6210298
    SetPixel m_hHdbSrcDC, 4, 14, 11526653
    SetPixel m_hHdbSrcDC, 5, 14, 5550569
    SetPixel m_hHdbSrcDC, 6, 14, 7326459
    SetPixel m_hHdbSrcDC, 7, 14, 12513279
    SetPixel m_hHdbSrcDC, 8, 14, 6272745
    SetPixel m_hHdbSrcDC, 0, 15, 14346744
    SetPixel m_hHdbSrcDC, 1, 15, 13886966
    SetPixel m_hHdbSrcDC, 2, 15, 11522026
    SetPixel m_hHdbSrcDC, 3, 15, 6210298
    SetPixel m_hHdbSrcDC, 4, 15, 11526653
    SetPixel m_hHdbSrcDC, 5, 15, 5682669
    SetPixel m_hHdbSrcDC, 6, 15, 8179963
    SetPixel m_hHdbSrcDC, 7, 15, 12972799
    SetPixel m_hHdbSrcDC, 8, 15, 7192558
    SetPixel m_hHdbSrcDC, 0, 16, 14018294
    SetPixel m_hHdbSrcDC, 1, 16, 13886966
    SetPixel m_hHdbSrcDC, 2, 16, 12901617
    SetPixel m_hHdbSrcDC, 3, 16, 6210298
    SetPixel m_hHdbSrcDC, 4, 16, 11526653
    SetPixel m_hHdbSrcDC, 5, 16, 5946868
    SetPixel m_hHdbSrcDC, 6, 16, 9164540
    SetPixel m_hHdbSrcDC, 7, 16, 13169407
    SetPixel m_hHdbSrcDC, 8, 16, 8638454
    SetPixel m_hHdbSrcDC, 0, 17, 13886966
    SetPixel m_hHdbSrcDC, 1, 17, 13886966
    SetPixel m_hHdbSrcDC, 2, 17, 13427187
    SetPixel m_hHdbSrcDC, 3, 17, 6210298
    SetPixel m_hHdbSrcDC, 4, 17, 11329789
    SetPixel m_hHdbSrcDC, 5, 17, 6078711
    SetPixel m_hHdbSrcDC, 6, 17, 10083837
    SetPixel m_hHdbSrcDC, 7, 17, 13169407
    SetPixel m_hHdbSrcDC, 8, 17, 9755130
    SetPixel m_hHdbSrcDC, 0, 18, 13886966
    SetPixel m_hHdbSrcDC, 1, 18, 13886966
    SetPixel m_hHdbSrcDC, 2, 18, 13886966
    SetPixel m_hHdbSrcDC, 3, 18, 7589115
    SetPixel m_hHdbSrcDC, 4, 18, 11066876
    SetPixel m_hHdbSrcDC, 5, 18, 7589115
    SetPixel m_hHdbSrcDC, 6, 18, 11068413
    SetPixel m_hHdbSrcDC, 7, 18, 13169407
    SetPixel m_hHdbSrcDC, 8, 18, 11068413
    SetPixel m_hHdbSrcDC, 0, 19, 12638962
    SetPixel m_hHdbSrcDC, 1, 19, 12638962
    SetPixel m_hHdbSrcDC, 2, 19, 12638962
    SetPixel m_hHdbSrcDC, 3, 19, 12638962
    SetPixel m_hHdbSrcDC, 4, 19, 12638962
    SetPixel m_hHdbSrcDC, 5, 19, 12638962
    SetPixel m_hHdbSrcDC, 6, 19, 12638962
    SetPixel m_hHdbSrcDC, 7, 19, 12638962
    SetPixel m_hHdbSrcDC, 8, 19, 12638962

End Sub
