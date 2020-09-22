Attribute VB_Name = "Module1"
Public Buffer() As Byte
Private Type WAVHEAD
RiffID As String * 4
RiffLength As Long
WavID As String * 4
FmtID As String * 4
FmtLength As Long
wavformattag As Integer
Channels As Integer
SamplesPerSec As Integer
BytesPerSec As Integer
BlockAlign As Integer
FmtSpecific As Integer
Padding As Long
DataID As String * 4
DataLength As Long
End Type: Public WH As WAVHEAD
Declare Function timeGetTime Lib "winmm.dll" () As Long

