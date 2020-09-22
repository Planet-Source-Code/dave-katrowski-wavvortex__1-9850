VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WavVortex  -  Dmkware.2000"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5040
      Top             =   2760
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2040
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Variables"
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "FM"
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Noise"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         Height          =   855
         Left            =   4080
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   53
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   24
         Top             =   240
         Width           =   1400
         Begin VB.Line Line1 
            BorderWidth     =   3
            X1              =   44
            X2              =   0
            Y1              =   52
            Y2              =   32
         End
         Begin VB.Shape Shape1 
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   600
            Shape           =   3  'Circle
            Top             =   720
            Width           =   135
         End
         Begin VB.Shape Shape2 
            FillColor       =   &H00000080&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   1170
            Shape           =   3  'Circle
            Top             =   30
            Width           =   135
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Freak"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   23
         Top             =   1440
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Clone"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   22
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HighRake"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   20
         Top             =   1200
         Width           =   1060
      End
      Begin VB.OptionButton Option1 
         Caption         =   "LowPass"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   19
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   720
         TabIndex        =   18
         Text            =   "0.4"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Text            =   "0.8"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   560
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Play"
         Height          =   495
         Left            =   4080
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Text            =   "0.6"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   480
         TabIndex        =   9
         Text            =   "70"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Text            =   "36000"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generate"
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "0.7"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Text            =   "100"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Osc2 Options:"
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Decay"
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Attack"
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F2"
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A2"
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wave Length"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F"
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sine(35) As Single, CoSn(35) As Single, ChangeTable As Integer, Val As Integer, LastVal As Integer, CurVU As Integer, SFr As Single
Dim leng As Long, N As Long, I As Long
Dim A1 As Single, F1 As Single
Dim A2 As Single, F2 As Single
Dim Atk1 As Single, Dk1 As Single, s1 As Byte
Dim Atk2 As Single, Dk2 As Single, s2 As Byte
Dim Sample1 As Byte, Sample2 As Byte
Private Sub Command1_Click(): leng = Text3: ReDim Buffer(leng)

Picture1.ScaleMode = 0: Picture1.ScaleHeight = 255: Picture1.ScaleWidth = leng: Picture1.Cls

A1 = Text1: F1 = Text2
A2 = Text4: F2 = Text5
s1 = 0: s2 = 0: Atk1 = 0: Atk2 = 0: Dk1 = A1: Dk2 = A2: SFr = F1

Command2.Enabled = False

For I = 0 To leng: DoEvents: N = I * 0.01745329251994
On Error Resume Next
If Check1.Value = 1 Then
F1 = Sin(N) * SFr
End If
If s1 = 0 Then
If Atk1 < A1 Then Atk1 = Atk1 + (Text6 / 100)
If Atk1 >= A1 And Not ((leng - I) * (Text7 / 100)) > A1 Then s1 = 1: Dk1 = A1
Sample1 = (Cos(F1 * I) * Atk1) + &H7F
Else
If Dk1 > 0 Then Dk1 = Dk1 - (Text7 / 100)
If Dk1 < 0 Then s1 = 0: Atk1 = 0
Sample1 = (Cos(F1 * I) * Dk1) + &H7F
End If
If s2 = 0 Then
If Atk2 < A2 Then Atk2 = Atk2 + (Text6 / 100)
If Atk2 >= A2 And Not ((leng - I) * (Text7 / 100)) > A2 Then s2 = 1: Dk2 = A2
Sample2 = (Cos(F2 * I) * Atk2) + &H7F
If Option1(0).Value Then
Sample2 = Sample2 / (10 * Log(1 + (F2 / (0.1 * Atk2))))
ElseIf Option1(1).Value Then
Sample2 = Sample2 / (10 * Exp(1 + (Abs(F2) / (0.1 * (Atk2))) * (2 ^ 2)))
ElseIf Option1(3).Value Then
Sample2 = Sample2 / (Atk2 * Tan(N * Atk2))
ElseIf Option1(4).Value Then
Sample2 = Rnd * Atk2
End If
Else
If Dk2 > 0 Then Dk2 = Dk2 - (Text7 / 100)
If Dk2 < 0 Then s2 = 0: Atk2 = 0
Sample2 = (Cos(F2 * I) * Dk2) + &H7F
If Option1(0).Value Then
Sample2 = Sample2 / (10 * Log(1 + (F2 / (0.1 * Dk2))))
ElseIf Option1(1).Value Then
Sample2 = Sample2 / (10 * Exp(1 + (Abs(F2) / (0.1 * (Dk2))) * (2 ^ 2)))
ElseIf Option1(3).Value Then
Sample2 = Sample2 / (Dk2 * Tan(N * Dk2))
ElseIf Option1(4).Value Then
Sample2 = Rnd * Dk2
End If
End If
    If -(I And 1) Then
        Picture1.PSet (I, Sample1), vbGreen
        Buffer(I) = Sample1
    Else
        Picture1.PSet (I, Sample2), vbYellow
        Buffer(I) = Sample2
    End If

Next

WH.RiffID = "RIFF"
WH.RiffLength = leng - 8
WH.WavID = "WAVE"
WH.FmtID = "fmt "
WH.FmtLength = 16
WH.wavformattag = 1
WH.Channels = 1
WH.SamplesPerSec = 11250
WH.BytesPerSec = 0
WH.BlockAlign = 11250
WH.FmtSpecific = 0
WH.Padding = 524289
WH.DataID = "data"
WH.DataLength = leng - 44

Open App.Path & "\temp.wav" For Binary As #1
Put #1, , WH
Put #1, , Buffer()
Close #1

Command2.Enabled = True
End Sub

Private Sub Command2_Click()
LoadFile App.Path & "\temp.wav", 1
Play 1, True, 0
End Sub

Private Sub Command3_Click()
On Error GoTo canceled
CD.CancelError = True
CD.Filter = "Wav Files (*.wav)|*.wav"
CD.ShowSave

Open CD.FileName For Binary As #1
Put #1, , WH
Put #1, , Buffer()
Close #1
canceled:
End Sub

Private Sub Form_Load()
Initialize_DSEngine Form1.hWnd, 44100
For I = 0 To 35
Sine(I) = Sin(I * (3.14 / 18))
CoSn(I) = Cos(I * (3.14 / 18))
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Terminate_DSEngine: End
End Sub

Private Sub Text1_Change()
If Not IsNumeric(Text1) Then Text1 = 100
If Text1 > 100 Then Text1 = 100
End Sub

Private Sub Text2_Change()
If Not IsNumeric(Text2) Then Text2 = 0.7
End Sub

Private Sub Text3_Change()
If Not IsNumeric(Text3) Then Text3 = 36000
End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4) Then Text4 = 100
If Text4 > 100 Then Text4 = 100
End Sub

Private Sub Text5_Change()
If Not IsNumeric(Text5) Then Text5 = 0.6
End Sub

Private Sub Text6_Change()
If Not IsNumeric(Text6) Then Text6 = 0.8
End Sub

Private Sub Text7_Change()
If Not IsNumeric(Text7) Then Text7 = 0.4
End Sub

Private Sub Timer1_Timer()
CurVU = GetVuStatus
If CurVU > LastVal Then
Val = Val + 1
ElseIf CurVU > LastVal + 2 Then
Val = Val + 2
ElseIf CurVU < LastVal Then
Val = Val - 1
ElseIf CurVU < LastVal - 2 Then
Val = Val - 2
End If
Update Val
LastVal = Val
End Sub

Sub Update(Value As Integer)
Select Case Value
Case 0: ChangeTable = 25
Shape2.FillColor = &H80&
Case 1: ChangeTable = 24
Shape2.FillColor = &H80&
Case 2: ChangeTable = 23
Shape2.FillColor = &H80&
Case 3: ChangeTable = 21
Shape2.FillColor = &H80&
Case 4: ChangeTable = 20
Shape2.FillColor = &H80&
Case 5: ChangeTable = 19
Shape2.FillColor = &H80&
Case 6: ChangeTable = 17
Shape2.FillColor = &H80&
Case 7: ChangeTable = 16
Shape2.FillColor = &H80&
Case 8: ChangeTable = 14
Shape2.FillColor = &H80&
Case 9: ChangeTable = 12
Shape2.FillColor = &HFF
Case 10: ChangeTable = 11
Shape2.FillColor = &HFF
Case Else: ChangeTable = 25
Shape2.FillColor = &H80&
End Select
Line1.X2 = Line1.X1 + (40 * Sine(ChangeTable))
Line1.Y2 = Line1.Y1 + (40 * CoSn(ChangeTable))
End Sub

