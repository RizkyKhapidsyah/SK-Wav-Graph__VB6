VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form WavForm 
   Caption         =   "WAV Player"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Height          =   975
      Left            =   2640
      ScaleHeight     =   915
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   120
      Width           =   2415
      Begin VB.PictureBox Picture2 
         Height          =   495
         Left            =   1200
         ScaleHeight     =   495
         ScaleWidth      =   30
         TabIndex        =   6
         Top             =   0
         Width           =   25
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FF00&
         X1              =   0
         X2              =   2760
         Y1              =   250
         Y2              =   250
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   1920
      Picture         =   "WavForm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   1320
      Picture         =   "WavForm.frx":037E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   720
      Picture         =   "WavForm.frx":06F9
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "WavForm.frx":0A88
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_open 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnu_Dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "WavForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type WAVEFMT
'struct WAVEFMT {
    signature As String * 4   ' must contain 'RIFF'
    RIFFsize As Long      ' size of file (in bytes) minus 8
    type As String * 4        ' must contain 'WAVE'
    fmtchunk As String * 4  ' must contain 'fmt ' (including blank)
    fmtsize As Long    ' size of format chunk, must be 16
    format As Integer       ' normally 1 (PCM)
    channels As Integer      ' number of channels, 1=mono, 2=stereo
    samplerate As Long    ' sampling frequency: 11025, 22050 or 44100
    average_bps As Long   ' average bytes per second; samplerate * channels
    align As Integer         ' 1=byte aligned, 2=word aligned
    bitspersample As Integer ' should be 8 or 16
    datchunk As String * 4   ' must contain 'data'
    samples As Long       'number of samples
 ' };
 End Type
 Dim WavDat As WAVEFMT

Private Sub Command1_Click()
 i = mciSendString("seek voice1 to start", 0&, 0, 0)

End Sub

Private Sub Command2_Click()
i = mciSendString("seek voice1 to end", 0&, 0, 0)




End Sub

Private Sub Command3_Click()
  i = mciSendString("play voice1", 0&, 0, 0)




End Sub

Private Sub Command4_Click()
  i = mciSendString("pause voice1", 0&, 0, 0)

End Sub


Private Sub Form_Load()
Picture1.BackColor = &H0&


End Sub




Private Sub Form_Unload(Cancel As Integer)
i = mciSendString("close voice1", 0&, 0, 0)
End Sub

Private Sub mnu_Exit_Click()
End
End Sub

Private Sub mnu_open_Click()
Screen.MousePointer = 11
Err.Clear
CommonDialog1.CancelError = True
On Error GoTo EH1

CommonDialog1.Filter = "Sounds (*.wav)|*.wav"
CommonDialog1.ShowOpen

Load_Wav
Screen.MousePointer = 0
Exit Sub
EH1:
Screen.MousePointer = 0
If Err = 32755 Then Err.Clear: Exit Sub
MsgBox Err.Description, vbExclamation, "ERR #" & Err
Err.Clear


End Sub

Private Sub Picture1_Click()
MsgBox Picture1.Width
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
  Dim mssg  As String * 20
   i = mciSendString("status voice1 position", mssg, 255, f)
Picture1.Left = -((Str(mssg) * DIFF) - (Picture3.Width / 2))

End Sub



Public Sub Load_Wav()
On Error Resume Next
Static X, Y As Single
Static Byt1 As Byte
Static In1 As Integer
Dim FWAV As String

FWAV = CommonDialog1.filename

'Open the WAV file as VOICE2. This will leave VOICE1 alone
'incase this WAV file is not valid
'We open the file here just to see if it is normal and vaild.
i = mciSendString("close voice2", 0&, 0, 0) 'CLOSE incase previously opened
i = mciSendString("open " & FWAV & " type waveaudio alias voice2", 0&, 0, 0)
   
   
   If i <> 0 Then
      MsgBox "This sound file appears to be corrupted or not compatible with this software.", vbCritical, "Sound file not loaded"
      Close #1
      Screen.MousePointer = 0
      i = mciSendString("close voice2", 0&, 0, 0)
      'We have to leave here since the file is not valid.
      Exit Sub
   End If
   
  i = mciSendString("set voice2 time format ms", 0&, 0, 0) ' Set it to milliseconds
  i = mciSendString("status voice2 length", mssg, 255, 0) ' Get length of sound
  If Str(mssg) <= Picture3.Width Then
   If Str(mssg) <= Picture3.Width / 4 Then
      TimeChunk = Space(0.5) 'Short sounds will need this setting
   Else
      TimeChunk = Space(100) 'Medium length sound files can use this setting
   End If
   Picture1.Width = Str(mssg)
  Else
   TimeChunk = Space(5000) 'Large souns will need a setting like this or LARGER!!
   Picture1.Width = Picture3.Width * 4
  End If

  i = mciSendString("close voice2", 0&, 0, 0) 'Close the WAV, we checked it for everyting by now.
  Screen.MousePointer = 11
  'At this point, EVERYTHING is good so far, now
  'that we determined that the WAV is VALID!
Close #1 'Close any pre-opened files, just to be safe...

X = 0

If FWAV <> "" Then
  If Dir(FWAV, vbNormal) = "" Then Exit Sub
     'Fil = FWAV
     Open FWAV For Binary As #1
     Get #1, , WavDat
     
Select Case WavDat.bitspersample

Case 8
   Picture1.Cls 'Clear the screen and start over
   Picture1.ScaleWidth = WavDat.samples / Len(TimeChunk)
   Picture1.ScaleHeight = 6000
   X = 0
   Do
      Get #1, , Byt1
      Get #1, , TimeChunk ' The Sound is opened, so we get a numerical chunk of it
      Y = Picture1.ScaleHeight / 2
      Picture1.Line -(X, (Y) - (-Byt1 * 10) - 1000), &HFF00&
      'We draw a SIMPLE line based on that CHUNK of info
      X = X + 1
      If EOF(1) Then
         Picture1.CurrentX = 0
            'Clean up after youself, don't be a pig here...
          Picture1.CurrentY = 0
         Exit Do
      End If
   Loop

Case 16
   Picture1.Cls
   Picture1.ScaleWidth = LOF(1) / Len(TimeChunk)
   Picture1.ScaleHeight = 6000
   X = 0
   Do
      Get #1, , In1
      Get #1, , TimeChunk
      Y = Picture1.ScaleHeight / 2
      Picture1.Line -(X, Y + (-In1 / 15)), &HFF00&
      X = X + 1
      If EOF(1) Then
         'Picture1.CurrentX = 0
         'Picture1.CurrentY = 0
         Picture1.Line (0, Picture1.ScaleHeight / 2)-(X + 50, Picture1.ScaleHeight / 2), &HFF00&
         Exit Do
      End If
   Loop
   
Case Else
   MsgBox "Sound file must be recorded at 8 or 16 bits persample.", vbInformation, "WAV Format not accepted"
   Close #1
   Screen.MousePointer = 0
'   Timer1.Enabled = False
'   Picture1.Visible = False
   Exit Sub
End Select

Close #1

Else 'REMEMBER, we are still in an IF statement, so study this ELSE section...
   MsgBox "This sound file appears to be invalid with this software.", vbCritical, "Sound file not loaded"
   Close #1
   Screen.MousePointer = 0
   Timer1.Enabled = False
   Picture1.Visible = False
   Exit Sub
End If


  i = mciSendString("close voice1", 0&, 0, 0)
  i = mciSendString("open " & FWAV & " type waveaudio alias voice1", 0&, 0, 0)

  i = mciSendString("set voice1 time format ms", 0&, 0, 0)
  i = mciSendString("status voice1 length", mssg, 255, 0)

'DIFF is the difference between the width of Picture1 and the length of the sound file
'We need this 'DIFF' ratio to correctly  set the TIMER1 sub for
'moving Picture1 to the left...
     DIFF = Picture1.Width / Str(mssg)


'Gets rid of annoying line on start of WAV image
Picture1.Line (0, 0)-(0, 2800), &H0&

Screen.MousePointer = 0
Timer1.Enabled = True
Picture1.Visible = True

Picture4.Cls
Picture4.Print " Channels: " & WavDat.channels
Picture4.Print " Bitspersample: " & WavDat.bitspersample
Picture4.Print " Samplerate: " & WavDat.samplerate & " Hz"

End Sub
