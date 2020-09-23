VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "c"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPeakVol 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FF00&
      Height          =   3000
      Left            =   90
      Picture         =   "Form1.frx":1CCA
      ScaleHeight     =   3000
      ScaleWidth      =   4995
      TabIndex        =   32
      Top             =   1020
      Width           =   5000
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   720
      Top             =   4350
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play wave"
      Height          =   400
      Left            =   4995
      TabIndex        =   28
      Top             =   4380
      Width           =   1000
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   5220
      ScaleHeight     =   3000
      ScaleWidth      =   840
      TabIndex        =   5
      Top             =   1230
      Width           =   840
      Begin VB.Label lblVolume 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   45
         TabIndex        =   29
         Top             =   420
         Width           =   720
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   5
         Left            =   15
         TabIndex        =   27
         Top             =   2220
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   4
         Left            =   15
         TabIndex        =   26
         Top             =   2340
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   3
         Left            =   15
         TabIndex        =   25
         Top             =   2460
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   2
         Left            =   15
         TabIndex        =   24
         Top             =   2580
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   1
         Left            =   15
         TabIndex        =   23
         Top             =   2700
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   0
         Left            =   15
         TabIndex        =   22
         Top             =   2820
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   6
         Left            =   15
         TabIndex        =   21
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   7
         Left            =   15
         TabIndex        =   20
         Top             =   1980
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   8
         Left            =   15
         TabIndex        =   19
         Top             =   1860
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   9
         Left            =   15
         TabIndex        =   18
         Top             =   1740
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   10
         Left            =   15
         TabIndex        =   17
         Top             =   1620
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   11
         Left            =   375
         TabIndex        =   16
         Top             =   1620
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   12
         Left            =   375
         TabIndex        =   15
         Top             =   1740
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   13
         Left            =   375
         TabIndex        =   14
         Top             =   1860
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   14
         Left            =   375
         TabIndex        =   13
         Top             =   1980
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   15
         Left            =   375
         TabIndex        =   12
         Top             =   2100
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   21
         Left            =   375
         TabIndex        =   11
         Top             =   2820
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   20
         Left            =   375
         TabIndex        =   10
         Top             =   2700
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   19
         Left            =   375
         TabIndex        =   9
         Top             =   2580
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   18
         Left            =   375
         TabIndex        =   8
         Top             =   2460
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   17
         Left            =   375
         TabIndex        =   7
         Top             =   2340
         Width           =   345
      End
      Begin VB.Label peak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   16
         Left            =   375
         TabIndex        =   6
         Top             =   2220
         Width           =   345
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar2 
      Height          =   300
      Left            =   1155
      TabIndex        =   4
      Top             =   120
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   1155
      TabIndex        =   3
      Top             =   600
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CheckBox Check1 
      Caption         =   "get input"
      Height          =   400
      Left            =   3885
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4365
      Width           =   1000
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   225
      Top             =   4350
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   765
      Left            =   2760
      TabIndex        =   31
      Top             =   900
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "The VU meter will only work if your sound card features a peak control line."
      Height          =   300
      Left            =   90
      TabIndex        =   30
      Top             =   4905
      Width           =   5940
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "output level"
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   135
      TabIndex        =   1
      Top             =   165
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "input level"
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'_____________________________________________________________________________________
' Volume Level Control
Dim hmixer As Long                  ' mixer handle
Dim inputVolCtrl As MIXERCONTROL    ' waveout volume control
Dim outputVolCtrl As MIXERCONTROL   ' microphone volume control
Dim rc As Long                      ' return code
Dim OK As Boolean                   ' boolean return code

Dim mxcd As MIXERCONTROLDETAILS         ' control info
Dim vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Dim volume As Long                      ' volume value
Dim volHmem As Long ' handle to volume memory
Dim pic As Integer
Dim pic1 As Integer
Dim lines
Dim ppos1
'_____________________________________________________________________________________
' picture peak
Dim Y As Single
Dim pPos As Long
Dim lasty As Single
'_____________________________________________________________________________________
' for playing a wave-file
Dim SoundFileName As String
Dim i As Long

Private Sub cmdPlay_Click()
'_____________________________________________________________________________________

' IS THERE A SOUNDCARD INSTALLED???
 SoundFileName$ = App.Path & "\" & "INTRO.WAV"
 i = waveOutGetNumDevs()
 If i > 0 Then
 MsgBox " found a soundcard in your PC"
 i& = sndPlaySound(SoundFileName$, Flags&)
 Else
 MsgBox "no soundcard found or installed?!"
 Beep
 End If

End Sub

Private Sub Form_Load()
Form2.Line10.Y1 = (Form2.Line6.Y1 - Form2.Line7.Y1) / 2
Form2.Line10.Y2 = (Form2.Line6.Y1 - Form2.Line7.Y1) / 2
Form2.Line11.Y1 = (Form2.Line6.Y1 - Form2.Line7.Y1) / 2
Form2.Line11.Y2 = (Form2.Line6.Y1 - Form2.Line7.Y1) / 2
lines = Form2.Line11.Y1
pic = 1
pic1 = 1
Form1.Hide
StartInput
Form2.Line1.X2 = Form2.Shape1.Left
Form2.Line2.Y1 = Form2.Shape1.Top
Form2.Show
'_____________________________________________________________________________________


   ' Open the mixer specified by DEVICEID
   rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   
   If ((MMSYSERR_NOERROR <> rc)) Then
       MsgBox "Couldn't open the mixer."
       Exit Sub
   End If
       
   ' Get the input volume meter
   OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_WAVEIN, MIXERCONTROL_CONTROLTYPE_PEAKMETER, inputVolCtrl)
   
   If (OK <> True) Then
       OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_PEAKMETER, inputVolCtrl)
   End If
   
   If (OK = True) Then
      ProgressBar1.Min = 0
      ProgressBar1.Max = inputVolCtrl.lMaximum
   Else
      ProgressBar1.Enabled = False
      MsgBox "Couldn't get wavein meter"
   End If
       
   ' Get the output volume meter
   OK = GetControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   
   If (OK = True) Then
      ProgressBar2.Min = 0
      ProgressBar2.Max = outputVolCtrl.lMaximum
   Else
      ProgressBar2.Enabled = False
      MsgBox "Couldn't get waveout meter"
   End If
   
   ' Initialize mixercontrol structure
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1

End Sub

Private Sub Check1_Click()
   If (Check1.Value = 1) Then
      StartInput  ' Start receiving audio input
   Else
      StopInput   ' Stop receiving audio input
   End If
End Sub

Private Sub Timer1_Timer()

On Error Resume Next

' Process sound buffer if recording
  If (fRecording) Then
  For i = 0 To (NUM_BUFFERS - 1)
  If inHdr(i).dwFlags And WHDR_DONE Then
  rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
  End If
  Next
  End If

' Get the current input level
  If (ProgressBar1.Enabled = True) Then
  mxcd.dwControlID = inputVolCtrl.dwControlID
  mxcd.item = inputVolCtrl.cMultipleItems
  rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
  CopyStructFromPtr volume, mxcd.paDetails, Len(volume)

If (volume < 0) Then
volume = -volume
End If
lblVolume.Caption = volume

Y = (volume / 32768) * 2000 + 500
If pic = 1 Then
Form2.Picture1.Line (pPos - 50, lasty)-(pPos, Y), RGB(volume / 200, volume / 100, volume / 25)
Else
Form2.Picture2.Line (pPos - 50, lasty)-(pPos, Y), RGB(volume / 200, volume / 100, volume / 25)
End If
If pic1 = 1 Then
Form3.Picture1.Line (ppos1 - 50, lasty)-(ppos1, Y), RGB(volume / 200, volume / 100, volume / 25)
Else
If pic1 = 2 Then
Form3.Picture2.Line (ppos1 - 50, lasty)-(ppos1, Y), RGB(volume / 200, volume / 100, volume / 25)
Else
Form3.Picture3.Line (ppos1 - 50, lasty)-(ppos1, Y), RGB(volume / 200, volume / 100, volume / 25)
End If
End If
picPeakVol.Line (pPos - 50, lasty)-(pPos, Y), vbBlue
lasty = Y

pPos = pPos + 50
ppos1 = ppos1 + 50
If ppos1 >= 11835 Then
If pic1 = 1 Then
pic1 = 2
GoTo 70
End If
If pic1 = 2 Then
pic1 = 3
GoTo 70
End If
If pic1 = 3 Then
Form3.Picture1.Cls
Form3.Picture2.Cls
Form3.Picture3.Cls
pic1 = 1
GoTo 70
End If
End If
40
If pPos >= 7215 Then               '5000 = picpeakvol.width
picPeakVol.Cls
If pic = 1 Then
pic = 2
GoTo 67
Else
Form2.Picture1.Cls
Form2.Picture2.Cls
pic = 1
End If
GoTo 67
70
ppos1 = 0
GoTo 40
67
pPos = 0

End If
ProgressBar1.Value = volume

peak(0).BackColor = vbBlack
peak(1).BackColor = vbBlack
peak(2).BackColor = vbBlack
peak(3).BackColor = vbBlack
peak(4).BackColor = vbBlack
peak(5).BackColor = vbBlack
peak(6).BackColor = vbBlack
peak(7).BackColor = vbBlack
peak(8).BackColor = vbBlack
peak(9).BackColor = vbBlack
peak(10).BackColor = vbBlack
'---------------------------
peak(21).BackColor = vbBlack
peak(20).BackColor = vbBlack
peak(19).BackColor = vbBlack
peak(18).BackColor = vbBlack
peak(17).BackColor = vbBlack
peak(16).BackColor = vbBlack
peak(15).BackColor = vbBlack
peak(14).BackColor = vbBlack
peak(13).BackColor = vbBlack
peak(12).BackColor = vbBlack
peak(11).BackColor = vbBlack

If volume > 0 Then peak(0).BackColor = RGB(0, 150, 0)
If volume > 0 Then peak(21).BackColor = RGB(0, 20, 100)

If volume > 3200 Then peak(1).BackColor = RGB(0, 200, 0)
If volume > 3200 Then peak(20).BackColor = RGB(5, 40, 110)

If volume > 6400 Then peak(2).BackColor = RGB(0, 255, 0)
If volume > 6400 Then peak(19).BackColor = RGB(10, 60, 120)

If volume > 9600 Then peak(3).BackColor = RGB(255, 255, 0) 'gelb
If volume > 9600 Then peak(18).BackColor = RGB(15, 80, 130)

If volume > 12800 Then peak(4).BackColor = RGB(255, 255, 50)
If volume > 12800 Then peak(17).BackColor = RGB(20, 100, 140)

If volume > 16000 Then peak(5).BackColor = RGB(255, 255, 100)
If volume > 16000 Then peak(16).BackColor = RGB(25, 120, 150)

If volume > 19200 Then peak(6).BackColor = RGB(255, 255, 150)
If volume > 19200 Then peak(15).BackColor = RGB(30, 140, 160)

If volume > 22400 Then peak(7).BackColor = RGB(255, 255, 200)
If volume > 22400 Then peak(14).BackColor = RGB(40, 160, 180)

If volume > 25600 Then peak(8).BackColor = RGB(255, 255, 215) 'rot
If volume > 25600 Then peak(13).BackColor = RGB(50, 180, 190)

If volume > 28800 Then peak(9).BackColor = RGB(255, 255, 230)
If volume > 28800 Then peak(12).BackColor = RGB(60, 200, 200)

If volume > 32000 Then peak(10).BackColor = RGB(255, 255, 245)
If volume > 32000 Then peak(11).BackColor = RGB(70, 220, 210)
  Form2.Line4.X2 = (volume / 10) / 4 * 3 + Form2.Line4.X1 + volume / 100
  Form2.Line4.Y2 = (volume / 10) / 4 * 3 + Form2.Line4.Y1 + volume / 100
If XSens = True Then
Form2.Line4.X2 = (volume / 5) / 4 * 3 + Form2.Line4.X1 + volume / 50
Form2.Line4.Y2 = (volume / 5) / 4 * 3 + Form2.Line4.Y1 + volume / 50
Form2.Shape1.Height = volume / 5
Form2.Shape1.Width = volume / 5
'If volume / 5 > Form2.Line7.Y1 Then Form2.Shape2.Top = Form2.Line7.Y1 - Form2.Shape2.Height
Form2.Shape2.Top = (volume / 5) + 100
Form2.Shape3.Width = volume / 2.5
GoTo 9
End If
Form2.Line4.X2 = (volume / 10) / 4 * 3 + Form2.Line4.X1 + volume / 100
Form2.Line4.Y2 = (volume / 10) / 4 * 3 + Form2.Line4.Y1 + volume / 100
Form2.Shape1.Height = volume / 10
Form2.Shape1.Width = volume / 10
Form2.Shape2.Top = (volume / 10) + 100
Form2.Shape3.Width = volume / 5
9
Form2.Line11.Y1 = (volume / 20) + lines
Form2.Line11.Y2 = (volume / 20) + lines
Form2.Line10.Y1 = ((volume / 20) * (-0.5) * 1.5) + lines
Form2.Line10.Y2 = ((volume / 20) * (-0.5) * 1.5) + lines
Form2.Line20.X1 = Form2.Line10.X1
Form2.Line20.Y1 = Form2.Line10.Y1
Form2.Line20.X2 = Form2.Line11.X2
Form2.Line20.Y2 = Form2.Line11.Y2
Form2.Line12.X1 = Form2.Line11.X1
Form2.Line12.Y1 = Form2.Line11.Y1
Form2.Line12.X2 = Form2.Line10.X2
Form2.Line12.Y2 = Form2.Line10.Y2
If Form2.Shape1.Height >= Form2.Line8.Y2 Then Form2.TooSens
Form2.Shape1.BorderColor = RGB(volume / 200, volume / 100, volume / 25)
Form2.Line1.BorderColor = Form2.Shape1.BorderColor
Form2.Line2.BorderColor = Form2.Shape1.BorderColor
Form2.Line4.BorderColor = Form2.Shape1.BorderColor
Form2.Shape2.BorderColor = Form2.Shape1.BorderColor
Form2.Line10.BorderColor = Form2.Shape1.BorderColor
Form2.Line11.BorderColor = Form2.Shape1.BorderColor
Form2.Line12.BorderColor = Form2.Shape1.BorderColor
Form2.Line20.BorderColor = Form2.Shape1.BorderColor
Form2.Shape3.BackColor = Form2.Shape1.BorderColor
  Form2.Line1.X1 = (Form2.Shape1.Width) / 2 + Form2.Line1.X2
  Form2.Line2.Y2 = (Form2.Shape1.Height) / 2 + Form2.Line2.Y1
  Form2.Line3.X2 = Form2.Shape1.Width / 4 + Form2.Line3.X1
  Form2.Line3.Y2 = Form2.Shape1.Height / 4 + Form2.Line3.Y1
  vol1 = volume
  End If

' Get the current output level
  If (ProgressBar2.Enabled = True) Then
  mxcd.dwControlID = outputVolCtrl.dwControlID
  mxcd.item = outputVolCtrl.cMultipleItems
  rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
  CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
 
  If (volume < 0) Then volume = -volume
  
  lblVolume.Caption = volume

Y = (volume / 32768) * 2000 + 500
If pic1 = 1 Then
Form3.Picture1.Line (ppos1 - 50, lasty)-(ppos1, Y), vbBlue
Else
If pic1 = 2 Then
Form3.Picture2.Line (ppos1 - 50, lasty)-(ppos1, Y), vbBlue
Else
Form3.Picture3.Line (ppos1 - 50, lasty)-(ppos1, Y), vbBlue
End If
End If

If pic = 1 Then
Form2.Picture1.Line (pPos - 50, lasty)-(pPos, Y), vbBlue
Else
Form2.Picture2.Line (pPos - 50, lasty)-(pPos, Y), vbBlue
End If
picPeakVol.Line (pPos - 50, lasty)-(pPos, Y), vbBlue
lasty = Y

pPos = pPos + 50
ppos1 = ppos1 + 50
If ppos1 >= 11835 Then
If pic1 = 1 Then
pic1 = 2
GoTo 90
End If
If pic1 = 2 Then
pic1 = 3
GoTo 90
End If
If pic1 = 3 Then
Form3.Picture1.Cls
Form3.Picture2.Cls
Form3.Picture3.Cls
pic1 = 1
GoTo 90
End If
End If
30
If pPos >= 7215 Then               '5000 = picpeakvol.width
picPeakVol.Cls
If pic = 1 Then
pic = 2
GoTo 68
Else
Form2.Picture1.Cls
Form2.Picture2.Cls
pic = 1
End If
GoTo 68
90
ppos1 = 0
GoTo 30
68
pPos = 0
End If

  ProgressBar2.Value = volume

If volume > 0 Then peak(0).BackColor = RGB(0, 150, 0)
If volume > 0 Then peak(21).BackColor = RGB(0, 20, 100)

If volume > 3200 Then peak(1).BackColor = RGB(0, 200, 0)
If volume > 3200 Then peak(20).BackColor = RGB(5, 40, 110)

If volume > 6400 Then peak(2).BackColor = RGB(0, 255, 0)
If volume > 6400 Then peak(19).BackColor = RGB(10, 60, 120)

If volume > 9600 Then peak(3).BackColor = RGB(255, 255, 0) 'gelb
If volume > 9600 Then peak(18).BackColor = RGB(15, 80, 130)

If volume > 12800 Then peak(4).BackColor = RGB(255, 255, 50)
If volume > 12800 Then peak(17).BackColor = RGB(20, 100, 140)

If volume > 16000 Then peak(5).BackColor = RGB(255, 255, 100)
If volume > 16000 Then peak(16).BackColor = RGB(25, 120, 150)

If volume > 19200 Then peak(6).BackColor = RGB(255, 255, 150)
If volume > 19200 Then peak(15).BackColor = RGB(30, 140, 160)

If volume > 22400 Then peak(7).BackColor = RGB(255, 255, 200)
If volume > 22400 Then peak(14).BackColor = RGB(40, 160, 180)

If volume > 25600 Then peak(8).BackColor = RGB(255, 255, 215) 'rot
If volume > 25600 Then peak(13).BackColor = RGB(50, 180, 190)

If volume > 28800 Then peak(9).BackColor = RGB(255, 255, 230)
If volume > 28800 Then peak(12).BackColor = RGB(60, 200, 200)

If volume > 32000 Then peak(10).BackColor = RGB(255, 255, 245)
If volume > 32000 Then peak(11).BackColor = RGB(70, 220, 210)
  
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (fRecording = True) Then
       StopInput
   End If
   GlobalFree volHmem
End Sub


Private Sub Timer2_Timer()
  If Form2.Label1.Caption = "Label 1" Then
  Max = vol1 / 1000
  Min = vol1 / 1000
  End If
  Form2.Label1.Caption = "Output Volume: " & vol1 / 1000
End Sub
