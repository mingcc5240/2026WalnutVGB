VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Walnut_CGB Gameboy Emulator"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   542
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   681
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   840
      Top             =   7560
   End
   Begin VB.PictureBox Picture1 
      Height          =   6855
      Left            =   120
      ScaleHeight     =   453
      ScaleMode       =   0  '사용자
      ScaleWidth      =   11312.94
      TabIndex        =   0
      Top             =   120
      Width           =   9075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bmi As BITMAPINFO
Private pFrameBuffer As Long
Private IsRunning As Boolean

' 게임보이 정확한 프레임 레이트 (59.7275 FPS)
Private Const TARGET_FPS As Double = 59.7275

Private Sub Form_Load()
    ' VB6 IDE 실행 시 DLL 경로 인식 문제 방지
    ChDrive App.Path
    ChDir App.Path

    ' 1. PictureBox 그래픽 설정
    Picture1.ScaleMode = vbPixels
    Picture1.AutoRedraw = True
    Picture1.BorderStyle = 0
    Picture1.Width = Me.ScaleX(160 * 3, vbPixels, Me.ScaleMode)
    Picture1.Height = Me.ScaleY(144 * 3, vbPixels, Me.ScaleMode)

    ' 2. 160x144 32비트 비트맵 헤더 초기화
    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = 160
        .biHeight = -144
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
    End With

    ' 3. 롬 파일 로드
    Dim romPath As String
    romPath = App.Path & "\mariotennis.gbc"

    Dim ret As Long
    ret = GB_Init(romPath)
    If ret <> 0 Then
        MsgBox "ROM 로드 실패! 에러 코드: " & ret & vbCrLf & "경로: " & romPath, vbCritical
        Exit Sub
    End If

    pFrameBuffer = GB_GetFrameBufferPtr()

    ' 4. 메인 루프 진입 전에 폼을 화면에 표시
    Me.Show
    DoEvents

    ' 5. 정밀 60FPS 메인 루프 시작
    IsRunning = True
    MainLoop
End Sub

Private Sub MainLoop()
    Dim freq As Currency
    Dim lastTime As Currency
    Dim currentTime As Currency
    Dim ticksPerFrame As Currency
    Dim elapsedTicks As Currency
    Dim remainingMs As Double

    ' 타이머 주파수 획득 및 1프레임당 목표 틱(Tick) 계산
    QueryPerformanceFrequency freq
    ticksPerFrame = CCur(CDbl(freq) / TARGET_FPS)
    QueryPerformanceCounter lastTime

    Do While IsRunning
        QueryPerformanceCounter currentTime
        elapsedTicks = currentTime - lastTime

        If elapsedTicks >= ticksPerFrame Then
            ' 다음 프레임 목표 시각 갱신 (누적 오차 보정)
            lastTime = lastTime + ticksPerFrame
            ' 창 이동 등으로 지연이 크게 발생했을 경우 현재 시각으로 리셋
            If currentTime - lastTime > ticksPerFrame Then
                lastTime = currentTime
            End If

            ' 1프레임 실행 및 화면 출력
            ProcessFrame
        Else
            ' 남은 시간이 2ms 이상이면 Sleep으로 CPU 점유율 절약
            remainingMs = ((CDbl(ticksPerFrame - elapsedTicks) / CDbl(freq)) * 1000#)
            If remainingMs > 2# Then
                Sleep 1
            End If
        End If

        ' 윈도우 메시지(창 닫기, 이동, 클릭 등) 처리
        DoEvents
    Loop
End Sub

Private Sub ProcessFrame()
    ' 1. 키 입력 상태 전달
    Dim kUp As Long, kDown As Long, kLeft As Long, kRight As Long
    Dim kA As Long, kB As Long, kSel As Long, kStart As Long

    kUp = IIf((GetAsyncKeyState(vbKeyUp) And &H8000) <> 0, 1, 0)
    kDown = IIf((GetAsyncKeyState(vbKeyDown) And &H8000) <> 0, 1, 0)
    kLeft = IIf((GetAsyncKeyState(vbKeyLeft) And &H8000) <> 0, 1, 0)
    kRight = IIf((GetAsyncKeyState(vbKeyRight) And &H8000) <> 0, 1, 0)
    kA = IIf((GetAsyncKeyState(vbKeyZ) And &H8000) <> 0, 1, 0)
    kB = IIf((GetAsyncKeyState(vbKeyX) And &H8000) <> 0, 1, 0)
    kSel = IIf((GetAsyncKeyState(vbKeyBack) And &H8000) <> 0, 1, 0)
    kStart = IIf((GetAsyncKeyState(vbKeyReturn) And &H8000) <> 0, 1, 0)

    GB_SetJoypad kUp, kDown, kLeft, kRight, kA, kB, kSel, kStart

    ' 2. 에뮬레이터 1프레임 실행
    GB_RunFrame

    ' 3. 화면 버퍼 복사 및 갱신
    StretchDIBits Picture1.hdc, _
                  0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
                  0, 0, 160, 144, _
                  pFrameBuffer, bmi, 0, vbSrcCopy
    Picture1.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' 창 닫기(X) 버튼 클릭 시 무한 루프를 먼저 안전하게 종료
    IsRunning = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IsRunning = False
    GB_Cleanup
End Sub


