Attribute VB_Name = "soundMOD"
Public DX7 As New DirectX7, DS As DirectSound
Public BD As DSBUFFERDESC, PD As DSBUFFERDESC
Public PCM As WAVEFORMATEX, PCM2 As WAVEFORMATEX

Const MaxSoundBuffers As Integer = 10
Public Primary As DirectSoundBuffer, LTB As DirectSoundBuffer
Public StaticBuffer(MaxSoundBuffers) As DirectSoundBuffer
Public VerbBuffer(MaxSoundBuffers) As DirectSoundBuffer
Public WaveSize(MaxSoundBuffers) As Long
Public NormalFreq(MaxSoundBuffers) As Long
Public VuTempBuffer() As Byte
Public Curs As DSCURSORS, MaxSample As Integer, TempMax As Integer

Public Const FX_NoFX As Integer = 0
Public Const FX_Reverb As Integer = 1
Public Const FX_DeTune As Integer = 2

Public Function Initialize_DSEngine(hWnd As Long, SamplingRate As Long) As Boolean
On Error GoTo NoDS
Set DS = DX7.DirectSoundCreate("")

DS.SetCooperativeLevel hWnd, DSSCL_EXCLUSIVE

PCM.nSize = LenB(PCM)
PCM.nFormatTag = WAVE_FORMAT_PCM
PCM.nChannels = 1
PCM.lSamplesPerSec = SamplingRate
PCM.nBitsPerSample = 16
PCM.nBlockAlign = PCM.nBitsPerSample / 8 * PCM.nChannels
PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign
BD.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_STATIC

On Error GoTo PBErr
PD.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_PRIMARYBUFFER
Set Primary = DS.CreateSoundBuffer(PD, PCM2)
Primary.SetFormat PCM
Initialize_DSEngine = True: Exit Function
NoDS: Initialize_DSEngine = False: MsgBox "Sorry, DirectSound wan unable to initialize.", vbCritical, "Error:": Exit Function
PBErr: Initialize_DSEngine = False: MsgBox "Sorry, There was an error setting while setting up the Primary Buffer.", vbCritical, "Error:"
End Function

Public Sub Terminate_DSEngine()
For I = 0 To MaxSoundBuffers
Set StaticBuffer(I) = Nothing
Next
Set Primary = Nothing
Set DS = Nothing
Set DX7 = Nothing
End Sub

Public Function LoadFile(FileName As String, BufferIndex As Integer) As Boolean: On Error GoTo BufferErr
Set StaticBuffer(BufferIndex) = DS.CreateSoundBufferFromFile(FileName, BD, PCM)
Set VerbBuffer(BufferIndex) = DS.DuplicateSoundBuffer(StaticBuffer(BufferIndex))
NormalFreq(BufferIndex) = StaticBuffer(BufferIndex).GetFrequency
WaveSize(BufferIndex) = FileLen(FileName)
LoadFile = True: Exit Function
BufferErr: LoadFile = False: MsgBox "There was an error while trying to open that file.", vbCritical, "Error:"
End Function
Public Function LoadWaveform(WaveformBuffer() As Byte, BufferIndex As Integer) As Boolean: On Error GoTo BufferErr
Set StaticBuffer(BufferIndex) = DS.CreateSoundBuffer(BD, PCM)
StaticBuffer(BufferIndex).WriteBuffer 0, 0, WaveformBuffer(0), DSBLOCK_ENTIREBUFFER
Set VerbBuffer(BufferIndex) = DS.DuplicateSoundBuffer(StaticBuffer(BufferIndex))
NormalFreq(BufferIndex) = StaticBuffer(BufferIndex).GetFrequency
WaveSize(BufferIndex) = UBound(WaveformBuffer())
LoadWaveform = True: Exit Function
BufferErr: LoadWaveform = False
End Function
Public Sub Play(BufferIndex As Integer, ResetPos As Boolean, Effect As Integer, Optional DetuneAmount As Long)
Select Case Effect
Case 0
If ResetPos Then ResetPosition BufferIndex
StaticBuffer(BufferIndex).SetFrequency NormalFreq(BufferIndex)
StaticBuffer(BufferIndex).Play DSBPLAY_DEFAULT
Case 1
If ResetPos Then ResetPosition BufferIndex
StaticBuffer(BufferIndex).SetFrequency NormalFreq(BufferIndex)
VerbBuffer(BufferIndex).SetFrequency NormalFreq(BufferIndex) + 400
StaticBuffer(BufferIndex).Play DSBPLAY_DEFAULT
VerbBuffer(BufferIndex).Play DSBPLAY_DEFAULT
Case 2
If ResetPos Then ResetPosition BufferIndex
StaticBuffer(BufferIndex).SetFrequency NormalFreq(BufferIndex)
VerbBuffer(BufferIndex).SetFrequency NormalFreq(BufferIndex) + DetuneAmount
StaticBuffer(BufferIndex).Play DSBPLAY_DEFAULT
VerbBuffer(BufferIndex).Play DSBPLAY_DEFAULT
End Select
End Sub

Public Sub ResetPosition(BufferIndex As Integer)
StaticBuffer(BufferIndex).SetCurrentPosition 0
VerbBuffer(BufferIndex).SetCurrentPosition 0
End Sub

Public Function GetPosition(BufferIndex As Integer) As Long
StaticBuffer(BufferIndex).GetCurrentPosition Curs
GetPosition = Curs.lPlay
End Function

'Loudest sample wins VU
Public Function GetVuStatus() As Integer
On Error Resume Next
MaxSample = 0
For I = 0 To MaxSoundBuffers
If StaticBuffer(I).GetStatus = DSBSTATUS_PLAYING Then
ReDim VuTempBuffer(WaveSize(I))
StaticBuffer(I).ReadBuffer 0, 0, VuTempBuffer(0), DSBLOCK_ENTIREBUFFER
TempMax = (Abs(VuTempBuffer(GetPosition(CInt(I))) - 127) / 127) * 10
If MaxSample < TempMax Then MaxSample = TempMax
End If
Next
GetVuStatus = MaxSample
End Function

Public Function GetVolume(BufferIndex As Integer) As Long
GetVolume = StaticBuffer(BufferIndex).GetVolume
End Function

Public Sub SetVolume(BufferIndex As Integer, Value As Integer)
Select Case Value
Case 0: StaticBuffer(BufferIndex).SetVolume -10000: VerbBuffer(BufferIndex).SetVolume -10000
Case 1: StaticBuffer(BufferIndex).SetVolume -2700: VerbBuffer(BufferIndex).SetVolume -2700
Case 2: StaticBuffer(BufferIndex).SetVolume -2400: VerbBuffer(BufferIndex).SetVolume -2400
Case 3: StaticBuffer(BufferIndex).SetVolume -2100: VerbBuffer(BufferIndex).SetVolume -2100
Case 4: StaticBuffer(BufferIndex).SetVolume -1800: VerbBuffer(BufferIndex).SetVolume -1800
Case 5: StaticBuffer(BufferIndex).SetVolume -1500: VerbBuffer(BufferIndex).SetVolume -1500
Case 6: StaticBuffer(BufferIndex).SetVolume -1200: VerbBuffer(BufferIndex).SetVolume -1200
Case 7: StaticBuffer(BufferIndex).SetVolume -900: VerbBuffer(BufferIndex).SetVolume -900
Case 8: StaticBuffer(BufferIndex).SetVolume -600: VerbBuffer(BufferIndex).SetVolume -600
Case 9: StaticBuffer(BufferIndex).SetVolume -300: VerbBuffer(BufferIndex).SetVolume -300
Case 10: StaticBuffer(BufferIndex).SetVolume 0: VerbBuffer(BufferIndex).SetVolume 0
End Select
End Sub
