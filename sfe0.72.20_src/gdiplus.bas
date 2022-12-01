Attribute VB_Name = "gdiplus"

Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Public Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal stream As IStream, bitmap As Long) As GpStatus
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hwnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As String, Image As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As IStream, Image As Long) As GpStatus

Public Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Public Enum GpStatus
    OK = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
End Enum

Public m_token As Long

Public Type picTYPE
    Width As Long
    Height As Long
    x As Long
    y As Long
    black As Long
    beginaddress As Long
    datalong As Long
End Type


Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any) As Long
Public Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

Public Declare Sub ReadDataFromFile Lib "sfeAdd.dll" (ByRef data As Byte, ByVal num As Long, _
ByRef count As Long, ByRef w As Long, _
ByRef h As Long, ByRef x As Long, ByRef y As Long, _
ByRef black As Long, ByRef beginaddress As Long, ByRef Length As Long)

Public Function ArrayToPicture(ImageArray() As Byte) As Long
    Dim MemoryHandle        As Long
    Dim LockMemory          As Long
    Dim GUID(0 To 3)        As Long
    Dim Size                As Long
    Dim IIStream            As IUnknown
    GUID(0) = &H7BF80980
    GUID(1) = &H101ABF32
    GUID(2) = &HAA00BB8B
    GUID(3) = &HAB0C3000
    Size = UBound(ImageArray) - LBound(ImageArray) + 1
    MemoryHandle = GlobalAlloc(&H2, Size)
    If MemoryHandle <> 0 Then
        LockMemory = GlobalLock(MemoryHandle)
        If LockMemory <> 0 Then
            CopyMemory ByVal LockMemory, ImageArray(0), Size
            GlobalUnlock MemoryHandle
            If CreateStreamOnHGlobal(ByVal MemoryHandle, 1, IIStream) = 0 Then
                GdipCreateBitmapFromStream IIStream, ArrayToPicture
            End If
        End If
    End If
End Function

