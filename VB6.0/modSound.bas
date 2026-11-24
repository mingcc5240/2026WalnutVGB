Attribute VB_Name = "modSound"
Option Explicit

Private Declare Function waveOutOpen Lib "winmm.dll" ( _
    ByRef lphWaveOut As Long, ByVal uDeviceID As Long, _
    ByRef lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, _
    ByVal dwInstance As Long, ByVal dwFlags As Long) As Long

Private Declare Function waveOutPrepareHeader Lib "winmm.dll" ( _
    ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, _
    ByVal uSize As Long) As Long

Private Declare Function waveOutWrite Lib "winmm.dll" ( _
    ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, _
    ByVal uSize As Long) As Long

Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" ( _
    ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, _
    ByVal uSize As Long) As Long

Private Declare Function waveOutReset Lib "winmm.dll" ( _
    ByVal hWaveOut As Long) As Long

Private Declare Function waveOutClose Lib "winmm.dll" ( _
    ByVal hWaveOut As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" ( _
    ByVal Destination As Long, ByVal Length As Long)

Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type

Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    reserved As Long
End Type

Private Const WAVE_FORMAT_PCM As Long = 1
Private Const WHDR_DONE As Long = &H1
Private Const WAVE_MAPPER As Long = -1

Public Const SND_SAMPLE_RATE As Long = 32768
Public Const SND_CHANNELS As Long = 2
Public Const SND_BITS As Long = 16
Public Const SND_BUFFER_SIZE As Long = 2192 '2048 ' 512 샘플 * 2채널 * 2바이트

Private Const NUM_BUFFERS As Long = 4 ' 버퍼를 6개로 늘려 넉넉한 큐 확보

Public Declare Function GBA_GetAudioSamples Lib "gdkgba.dll" ( _
    ByVal pBuffer As Long, ByVal nByteLen As Long) As Long

Private m_hWaveOut As Long
Private m_Headers(0 To NUM_BUFFERS - 1) As WAVEHDR
Private m_SoundData(0 To NUM_BUFFERS - 1, 0 To SND_BUFFER_SIZE - 1) As Byte
Private m_CurrentBufIdx As Long
Private m_bSoundInitialized As Boolean

Public Function Sound_Init() As Boolean
    Dim wfx As WAVEFORMATEX
    Dim i As Long
    Dim ret As Long

    If m_bSoundInitialized Then Call Sound_Close

    With wfx
        .wFormatTag = WAVE_FORMAT_PCM
        .nChannels = SND_CHANNELS
        .nSamplesPerSec = SND_SAMPLE_RATE
        .wBitsPerSample = SND_BITS
        .nBlockAlign = 4
        .nAvgBytesPerSec = SND_SAMPLE_RATE * 4
        .cbSize = 0
    End With

    ret = waveOutOpen(m_hWaveOut, WAVE_MAPPER, wfx, 0, 0, 0)
    If ret <> 0 Then
        Sound_Init = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        Call ZeroMemory(VarPtr(m_SoundData(i, 0)), SND_BUFFER_SIZE)
        With m_Headers(i)
            .lpData = VarPtr(m_SoundData(i, 0))
            .dwBufferLength = SND_BUFFER_SIZE
            .dwFlags = 0
            .dwLoops = 0
        End With
        Call waveOutPrepareHeader(m_hWaveOut, m_Headers(i), Len(m_Headers(i)))
    Next i

    m_CurrentBufIdx = 0
    m_bSoundInitialized = True
    Sound_Init = True
End Function

Public Sub Sound_Update()
    If Not m_bSoundInitialized Then Exit Sub

    ' 1. DLL에서 최신 오디오 데이터 가져오기 (매 프레임 무조건 호출)
    Call GBA_GetAudioSamples(VarPtr(m_SoundData(m_CurrentBufIdx, 0)), SND_BUFFER_SIZE)

    ' 2. 버퍼가 재생 중이었으면 언프리페어 후 재준비
    If (m_Headers(m_CurrentBufIdx).dwFlags And WHDR_DONE) = WHDR_DONE Then
        Call waveOutUnprepareHeader(m_hWaveOut, m_Headers(m_CurrentBufIdx), Len(m_Headers(m_CurrentBufIdx)))
        m_Headers(m_CurrentBufIdx).dwFlags = 0
        Call waveOutPrepareHeader(m_hWaveOut, m_Headers(m_CurrentBufIdx), Len(m_Headers(m_CurrentBufIdx)))
    End If

    ' 3. 디바이스로 버퍼 밀어넣기
    Call waveOutWrite(m_hWaveOut, m_Headers(m_CurrentBufIdx), Len(m_Headers(m_CurrentBufIdx)))

    ' 4. 무조건 다음 버퍼로 인덱스 전진
    m_CurrentBufIdx = (m_CurrentBufIdx + 1) Mod NUM_BUFFERS
End Sub

Public Sub Sound_Close()
    Dim i As Long
    If Not m_bSoundInitialized Then Exit Sub
    m_bSoundInitialized = False

    If m_hWaveOut <> 0 Then
        Call waveOutReset(m_hWaveOut)
        For i = 0 To NUM_BUFFERS - 1
            Call waveOutUnprepareHeader(m_hWaveOut, m_Headers(i), Len(m_Headers(i)))
        Next i
        Call waveOutClose(m_hWaveOut)
        m_hWaveOut = 0
    End If
End Sub

Public Function Sound_GetDebugInfo() As String
    If Not m_bSoundInitialized Then
        Sound_GetDebugInfo = "Sound: Not Init"
        Exit Function
    End If

    Dim i As Long
    Dim sampleVal As Integer
    Dim maxAmp As Long
    Dim nonZeroCount As Long

    maxAmp = 0
    nonZeroCount = 0

    For i = 0 To (SND_BUFFER_SIZE - 2) Step 2
        Call CopyMemory(sampleVal, m_SoundData(m_CurrentBufIdx, i), 2)
        If sampleVal <> 0 Then
            nonZeroCount = nonZeroCount + 1
            If Abs(CLng(sampleVal)) > maxAmp Then
                maxAmp = Abs(CLng(sampleVal))
            End If
        End If
    Next i

    Sound_GetDebugInfo = "Buf: " & m_CurrentBufIdx & _
                         " | Done: " & ((m_Headers(m_CurrentBufIdx).dwFlags And WHDR_DONE) = WHDR_DONE) & _
                         " | NonZero: " & nonZeroCount & "/1024" & _
                         " | MaxAmp: " & maxAmp
End Function

