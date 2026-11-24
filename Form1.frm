VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Walnut_CGB Gameboy Emulator"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   649
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   664
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   8880
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   840
      Top             =   7560
   End
   Begin VB.PictureBox Picture1 
      Height          =   8640
      Left            =   120
      ScaleHeight     =   572
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   636
      TabIndex        =   0
      Top             =   120
      Width           =   9600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_PATHMUSTEXIST = &H800

Private bmi As BITMAPINFO
Private pFrameBuffer As Long
Private IsRunning As Boolean

' 게임보이 정확한 프레임 레이트 (59.7275 FPS)
Private Const TARGET_FPS As Double = 59.7275
'ADD Zip file
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Function ExtractRomFromZip(ByVal zipPath As String) As String
    Dim objShell As Object
    Dim objZip As Object
    Dim objTempFolder As Object
    Dim item As Object
    Dim tempDir As String
    Dim ext As String
    
    ' 1. 윈도우 임시 폴더 경로 가져오기
    tempDir = Space$(260)
    GetTempPath 260, tempDir
    tempDir = left$(tempDir, InStr(tempDir, Chr$(0)) - 1) & "VB6_GB_Temp\"
    
    ' 임시 폴더가 없으면 생성
    If Dir(tempDir, vbDirectory) = "" Then MkDir tempDir
    
    ' 2. Shell.Application으로 ZIP 내부 탐색 (Late Binding)
    Set objShell = CreateObject("Shell.Application")
    Set objZip = objShell.NameSpace((zipPath))
    Set objTempFolder = objShell.NameSpace((tempDir))
    
    If objZip Is Nothing Then Exit Function
    
    For Each item In objZip.Items
        If Not item.IsFolder Then
            ext = LCase$(right$(item.Name, 4))
            ' .gb 또는 .gbc 파일 발견 시 추출
            If right$(ext, 3) = ".gb" Or ext = ".gbc" Then
                ' 기존에 같은 이름의 임시 파일이 있다면 삭제 (덮어쓰기 팝업 방지)
                If Dir(tempDir & item.Name) <> "" Then Kill tempDir & item.Name
                
                ' 16 = 대화상자 없이 자동 덮어쓰기/진행
                objTempFolder.CopyHere item, 16
                
                ' 비동기 복사 완료 대기 (최대 3초)
                Dim t As Single: t = Timer
                Do While Dir(tempDir & item.Name) = ""
                    DoEvents
                    If Timer - t > 3 Then Exit Do
                Loop
                
                ExtractRomFromZip = tempDir & item.Name
                Exit Function
            End If
        End If
    Next
End Function



Private Sub Command1_Click()
      ' 3. 다이얼로그 박스로 롬 파일(또는 ZIP) 선택 및 로드
    Dim ofn As OPENFILENAME
    Dim selectedPath As String
    Dim romPath As String
    
    
    IsRunning = False
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = Me.hWnd
        .hInstance = App.hInstance
        ' 필터에 *.zip 추가
        .lpstrFilter = "Game Boy ROMs (*.gb;*.gbc;*.zip)" & Chr$(0) & "*.gb;*.gbc;*.zip" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = App.Path
        .lpstrTitle = "실행할 롬(ROM) 또는 ZIP 파일을 선택하세요"
        .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    End With
    
    If GetOpenFileName(ofn) <> 0 Then
        selectedPath = left$(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr$(0)) - 1)
    Else
        MsgBox "롬 파일 선택이 취소되었습니다.", vbInformation
        Exit Sub
    End If

    ' 확장자가 .zip인지 확인 후 분기 처리
    If LCase$(right$(selectedPath, 4)) = ".zip" Then
        romPath = ExtractRomFromZip(selectedPath)
        If romPath = "" Then
            MsgBox "ZIP 파일 안에 유효한 GB 또는 GBC 롬 파일이 없습니다.", vbExclamation
            Exit Sub
        End If
    Else
        romPath = selectedPath
    End If

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

Private Sub testCommand()
   IsRunning = False
   ' 3. 다이얼로그 박스로 롬 파일 선택 및 로드
    Dim ofn As OPENFILENAME
    Dim romPath As String
    
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = Me.hWnd
        .hInstance = App.hInstance
        ' 필터 설정 (GB, GBC 파일 및 모든 파일)
        .lpstrFilter = "Game Boy ROMs (*.gb;*.gbc)" & Chr$(0) & "*.gb;*.gbc" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = App.Path ' 기본 시작 폴더
        .lpstrTitle = "실행할 롬(ROM) 파일을 선택하세요"
        .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    End With
    
    ' 파일 열기 창 띄우기
    If GetOpenFileName(ofn) <> 0 Then
        ' 선택된 경로에서 불필요한 Null 문자(Chr(0)) 제거
        romPath = left$(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr$(0)) - 1)
    Else
        ' 사용자가 취소를 누른 경우 종료
        MsgBox "롬 파일 선택이 취소되었습니다.", vbInformation
        Exit Sub
    End If

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

Private Sub Form_Load()
    ' VB6 IDE 실행 시 DLL 경로 인식 문제 방지
    ChDrive App.Path
    ChDir App.Path

    ' 1. PictureBox 그래픽 설정
    Picture1.ScaleMode = vbPixels
    Picture1.AutoRedraw = True
    Picture1.BorderStyle = 0
    Picture1.Width = Me.ScaleX(160 * 4, vbPixels, Me.ScaleMode)
    Picture1.Height = Me.ScaleY(144 * 4, vbPixels, Me.ScaleMode)

    ' 2. 160x144 32비트 비트맵 헤더 초기화
    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = 160
        .biHeight = -144
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
    End With
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


