VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form SoundChooser 
   BorderStyle     =   0  'None
   Caption         =   "SoundChooser"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "SoundChooser.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5040
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2"
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      MouseIcon       =   "SoundChooser.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "SoundChooser.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      MouseIcon       =   "SoundChooser.frx":2378
      MousePointer    =   99  'Custom
      Picture         =   "SoundChooser.frx":2682
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Play 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      MouseIcon       =   "SoundChooser.frx":43E6
      MousePointer    =   99  'Custom
      Picture         =   "SoundChooser.frx":46F0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFC0&
      Height          =   2235
      Left            =   3240
      MouseIcon       =   "SoundChooser.frx":6454
      MousePointer    =   99  'Custom
      MultiSelect     =   2  'Extended
      Pattern         =   "*.mp3; *.mid; *.wav; *.mp4; *.wma"
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      MouseIcon       =   "SoundChooser.frx":675E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H0080FF80&
      Height          =   2340
      Left            =   120
      MouseIcon       =   "SoundChooser.frx":6A68
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1680
      Width           =   3015
   End
   Begin MCI.MMControl muzica 
      Height          =   330
      Left            =   1200
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label NOWPLAYING 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Now playing..."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Music player™"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   6375
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5055
      Left            =   0
      Picture         =   "SoundChooser.frx":6D72
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "SoundChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim opnd As Boolean

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
If Drive1.Drive <> "" Then
Dir1.Path = Drive1.Drive
End If
End Sub

Private Sub File1_Click()
If Len(File1.FileName) = 0 Then
    Exit Sub
End If
opnd = False
Play.Caption = "Play"
ccc = File1.FileName
End Sub

Private Sub File1_DblClick()
If Len(File1.FileName) = 0 Then
    Exit Sub
End If
ccc = File1.FileName
muzica.Command = "close"
muzica.FileName = Dir1.Path & "\" & ccc
muzica.Command = "open"
muzica.Command = "prev"
muzica.Command = "play"
Label2.Visible = True
NOWPLAYING.Caption = ccc
opnd = True
Play.Caption = "Pause"
End Sub

Private Sub Form_activate()
If Language = 1 Then
    Label2.Caption = "Acum cântã..."
ElseIf Language = 0 Then
    Label2.Caption = "Now playing..."
End If
Me.Tag = 1
End Sub

Private Sub OK_Click()
Me.Tag = 2
Me.Hide
End Sub

Private Sub play_Click()
If Play.Caption = "Play" And opnd = False Then
    muzica.Command = "close"
    muzica.FileName = Dir1.Path & "\" & ccc
    muzica.Command = "open"
    muzica.Command = "prev"
    muzica.Command = "play"
    Label2.Visible = True
    NOWPLAYING.Caption = ccc
    Play.Caption = "Pause"
    opnd = True
ElseIf Play.Caption = "Play" And opnd = True Then
    muzica.Command = "play"
    Play.Caption = "Pause"
ElseIf Play.Caption = "Pause" Then
    muzica.Command = "pause"
    Play.Caption = "Play"
End If
End Sub

Private Sub Stop_Click()
opnd = False
muzica.Command = "close"
Label2.Visible = False
NOWPLAYING.Caption = ""
End Sub
