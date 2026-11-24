Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GB_Init Lib "walnut_dll.dll" (ByVal rom_path As String) As Long
Public Declare Sub GB_RunFrame Lib "walnut_dll.dll" ()
Public Declare Sub GB_SetJoypad Lib "walnut_dll.dll" ( _
    ByVal up As Long, ByVal down As Long, ByVal left As Long, ByVal right As Long, _
    ByVal a As Long, ByVal b As Long, ByVal sel As Long, ByVal start As Long)
Public Declare Function GB_GetFrameBufferPtr Lib "walnut_dll.dll" () As Long
Public Declare Sub GB_Reset Lib "walnut_dll.dll" ()
Public Declare Sub GB_Cleanup Lib "walnut_dll.dll" ()

' --- 2. 고속 화면 출력 및 실시간 키 입력을 위한 Windows API ---
Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
End Type

Public Declare Function StretchDIBits Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, _
    ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, _
    ByVal lpBits As Long, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


