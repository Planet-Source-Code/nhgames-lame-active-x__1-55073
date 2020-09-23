VERSION 5.00
Object = "{F1253762-C467-11D3-8290-0080C605ADA4}#1.0#0"; "LameEncoderX.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NHGames Encoder Active X"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Output"
      Height          =   375
      Left            =   3360
      TabIndex        =   27
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Input"
      Height          =   375
      Left            =   2280
      TabIndex        =   26
      Top             =   120
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
      Max             =   109
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Set Tag Info"
      Height          =   375
      Left            =   1200
      TabIndex        =   24
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   960
      TabIndex        =   23
      Text            =   "None"
      Top             =   4440
      Width           =   4575
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   960
      TabIndex        =   22
      Text            =   "None"
      Top             =   4080
      Width           =   4575
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   960
      TabIndex        =   21
      Text            =   "None"
      Top             =   3720
      Width           =   4575
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   960
      TabIndex        =   20
      Text            =   "None"
      Top             =   3360
      Width           =   4575
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   960
      TabIndex        =   19
      Text            =   "None"
      Top             =   3000
      Width           =   4575
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   960
      TabIndex        =   18
      Text            =   "None"
      Top             =   2640
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Tag Info"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "128"
      Top             =   720
      Width           =   495
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Voice Mode"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Mono"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Fast mode"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encode"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   480
      Top             =   5040
   End
   Begin LAMEENCODERXLib.LameEncoderX LameEncoderX1 
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tag Info "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label Label8 
      Caption         =   "Year:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Genre:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Comment:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Artist:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Album:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Bit rate:"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   750
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartEncode As Boolean
Dim Done As Boolean
Dim Press As Boolean

Private Sub Check1_Click()

'Use fast mode
If Check1.Value = 0 Then
LameEncoderX1.UseFastMode = False
End If

If Check1.Value = 1 Then
LameEncoderX1.UseFastMode = True
End If

End Sub

Private Sub Check2_Click()

'Mono
If Check2.Value = 0 Then
LameEncoderX1.DownmixToMono = False
End If

If Check2.Value = 1 Then
LameEncoderX1.DownmixToMono = True
End If

End Sub

Private Sub Check3_Click()

'Voice mode
If Check3.Value = 0 Then
LameEncoderX1.VoiceMode = False
End If

If Check3.Value = 1 Then
LameEncoderX1.VoiceMode = True
End If

End Sub

Private Sub Command1_Click()

'Check if got files
If LameEncoderX1.InputFile = "" Then
MsgBox "No input", vbCritical, "Error"
Exit Sub
End If

If LameEncoderX1.OutputFile = "" Then
MsgBox "No output", vbCritical, "Error"
Exit Sub
End If

'Start Setting up
Timer1.Enabled = True
StartEncode = True
Press = True
Label1.Visible = True

'Disable stuff
Command1.Enabled = False
Command2.Enabled = False
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Text3.Enabled = False
Command4.Enabled = False
Command3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Command5.Enabled = False
Command6.Enabled = False

'Check status
If Check1.Value = 0 Then
LameEncoderX1.UseFastMode = False
End If

If Check1.Value = 1 Then
LameEncoderX1.UseFastMode = True
End If

If Check2.Value = 0 Then
LameEncoderX1.DownmixToMono = False
End If

If Check2.Value = 1 Then
LameEncoderX1.DownmixToMono = True
End If

If Check3.Value = 0 Then
LameEncoderX1.VoiceMode = False
End If

If Check3.Value = 1 Then
LameEncoderX1.VoiceMode = True
End If

'Start Encoding
LameEncoderX1.StartEncode

End Sub

Private Sub Command2_Click()

'Show About
LameEncoderX1.AboutBox

End Sub

Private Sub Command3_Click()

'Set Form sizing for text boxs
Me.Height = 5310
Me.Width = 5745
Command4.Visible = True
Label9.Visible = True

End Sub

Private Sub Command4_Click()

'Set Form sizing for text boxs
Me.Height = 2955
Me.Width = 5745
Command4.Visible = False
Label9.Visible = False

End Sub

Private Sub Command5_Click()

'Open file to convert
CommonDialog1.DialogTitle = "Input File"
CommonDialog1.Filter = "Wave Files (*.wav)|*.wav"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
LameEncoderX1.InputFile = CommonDialog1.FileName

End Sub

Private Sub Command6_Click()

'Save Converted file
CommonDialog1.DialogTitle = "Output File"
CommonDialog1.Filter = "Mp3 File (*.mp3)|*.mp3"
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then Exit Sub
LameEncoderX1.OutputFile = CommonDialog1.FileName
Text1.Text = LameEncoderX1.OutputFile

End Sub

Private Sub Form_Load()

'Set Tag Info
LameEncoderX1.Album = "NHGames"
LameEncoderX1.Artist = "NHGames"
LameEncoderX1.Comment = "NHGames"
LameEncoderX1.Genre = "NHGames"
LameEncoderX1.Title = "NHGames"
LameEncoderX1.Year = "NHGames"

'Options
LameEncoderX1.MarkAsCopyrighted = False
LameEncoderX1.MarkAsNonOriginal = False
LameEncoderX1.Bitrate = "128"
LameEncoderX1.DisableVBRInfoTag = False
LameEncoderX1.AllowChanDifBlockTypes = False
LameEncoderX1.EncodingPriority = "3"
LameEncoderX1.ErrorProtection = True
LameEncoderX1.DisableSFBCutoff = False
LameEncoderX1.UseVBR = False
LameEncoderX1.VBRQuality = "4"
LameEncoderX1.ForceByteSwab = False
LameEncoderX1.InputIsRawPCM = False
LameEncoderX1.NoShortBlocks = False
LameEncoderX1.OnlyATHForMasking = False
LameEncoderX1.ResampleFreq = "-1"
LameEncoderX1.SamplingFreq = "-1"
LameEncoderX1.MaximumBitrate = "160"
LameEncoderX1.MinimumBitrate = "128"
LameEncoderX1.UseFastMode = False
LameEncoderX1.VoiceMode = False
LameEncoderX1.Mode = "0"
LameEncoderX1.DownmixToMono = False
LameEncoderX1.EnoderXVersion = "1"

'Input File and Output File
LameEncoderX1.InputFile = ""
LameEncoderX1.OutputFile = ""

' Get Command Line
Text1.Text = LameEncoderX1.GetCurrentCommandString

'Change Text on start
If StartEncode = True Then
Text2.Text = LameEncoderX1.GetTotalEncodingTime & " Encoding..."
End If

' Change Text on Stop
If StartEncode = False Then
Text2.Text = LameEncoderX1.GetTotalEncodingTime
End If

'Get output file
Text1.Text = LameEncoderX1.OutputFile
   
'Set Encoding to false
StartEncode = False
Command4.Visible = False
Press = False
Label1.Visible = False
Label9.Visible = False
Text2.Text = ""

'Set Sizing now
Me.Height = 2955
Me.Width = 5745

If Press = True Then
'Check Texts

'Set Album
LameEncoderX1.Album = Text4.Text

'Set Artist
LameEncoderX1.Artist = Text5.Text

'Set Comment
LameEncoderX1.Comment = Text6.Text

'Set Genre
LameEncoderX1.Genre = Text7.Text

'Set Title
LameEncoderX1.Title = Text8.Text

'Set Year
LameEncoderX1.Year = Text9.Text
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Idk why i put this here
LameEncoderX1.SetAllDefaultParams

End Sub

Private Sub LameEncoderX1_PercentComplete(ByVal lPercent As Long, ByVal lSample As Long)

'Set Progress
Label1.Caption = Label1.Caption + 1

End Sub

Private Sub Text3_Change()

'Set bitrates
If Text3.Text = "" Then Text3.Text = "128"
If Text3.Text = " " Then Text3.Text = "128"
If Text3.Text = "  " Then Text3.Text = "128"
LameEncoderX1.Bitrate = Text3.Text

End Sub

Private Sub Text4_Change()

'Set Album
LameEncoderX1.Album = Text4.Text

End Sub

Private Sub Text5_Change()

'Set Artist
LameEncoderX1.Artist = Text5.Text

End Sub

Private Sub Text6_Change()

'Set Comment
LameEncoderX1.Comment = Text6.Text

End Sub

Private Sub Text7_Change()

'Set Genre
LameEncoderX1.Genre = Text7.Text

End Sub

Private Sub Text8_Change()

'Set Title
LameEncoderX1.Title = Text8.Text

End Sub

Private Sub Text9_Change()

'Set Year
LameEncoderX1.Year = Text9.Text

End Sub

Private Sub Timer1_Timer()

'Get info
Text1.Text = LameEncoderX1.GetCurrentCommandString
Text2.Text = LameEncoderX1.GetTotalEncodingTime & " Encoding..."

'Done encoding
If ProgressBar1.Value = 100 Then
Timer1.Enabled = False
StartEncode = False
ProgressBar1.Value = 0
Label1.Caption = 0
Text1.Text = " "
Text2.Text = " "
MsgBox "Done", vbInformation, "Encoding Done"
Label1.Visible = False

'Active stuff
Command1.Enabled = True
Command2.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Text3.Enabled = True
Command4.Enabled = True
Command3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Command5.Enabled = True
Command6.Enabled = True

End If

'Make progress work
ProgressBar1.Value = Label1.Caption

End Sub
