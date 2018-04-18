VERSION 5.00
Begin VB.Form Game 
   BackColor       =   &H00004000&
   BorderStyle     =   0  'None
   Caption         =   "6a5p1 GAMES"
   ClientHeight    =   8025
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   9585
   DrawStyle       =   5  'Transparent
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Game chooser"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Game.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "Game.frx":0BD4
   ScaleHeight     =   8025
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   480
      Top             =   840
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8535
      MouseIcon       =   "Game.frx":1E75
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   75
      Width           =   255
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   255
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8888
      MouseIcon       =   "Game.frx":217F
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   60
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   255
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   60
      Width           =   255
   End
   Begin VB.Label ImpasseLabel 
      BackStyle       =   0  'Transparent
      Height          =   1695
      Left            =   480
      MouseIcon       =   "Game.frx":2489
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Impasse"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   735
      Left            =   720
      MouseIcon       =   "Game.frx":2793
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   6240
      Width           =   3015
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000004&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      FillColor       =   &H00E0E0E0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1695
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   6120
      MouseIcon       =   "Game.frx":2A9D
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label SudokuLabel 
      BackStyle       =   0  'Transparent
      Height          =   1695
      Left            =   480
      MouseIcon       =   "Game.frx":2DA7
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label TicTacToeLabel 
      BackStyle       =   0  'Transparent
      Height          =   1695
      Left            =   480
      MouseIcon       =   "Game.frx":30B1
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Image Romanian 
      Height          =   975
      Left            =   7560
      MouseIcon       =   "Game.frx":33BB
      MousePointer    =   99  'Custom
      Picture         =   "Game.frx":36C5
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Image English 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   5280
      MouseIcon       =   "Game.frx":88BB
      MousePointer    =   99  'Custom
      Picture         =   "Game.frx":8BC5
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6a5p1 GAMES v1.0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MouseIcon       =   "Game.frx":DDBB
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   60
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9240
      MouseIcon       =   "Game.frx":E0C5
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   60
      Width           =   255
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   255
      Left            =   9240
      Shape           =   3  'Circle
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Games"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   6960
      TabIndex        =   8
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6a5p1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   855
      Left            =   6120
      TabIndex        =   4
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sudoku"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      FillColor       =   &H00E0E0E0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1695
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   4
      FillColor       =   &H00E0E0E0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   975
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tic Tac Toe"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   765
      Left            =   495
      TabIndex        =   0
      Top             =   1920
      Width           =   3315
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      FillColor       =   &H00E0E0E0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1695
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   2295
      Left            =   4320
      TabIndex        =   6
      Top             =   600
      Width           =   4695
   End
   Begin VB.Image Image2 
      Height          =   7695
      Left            =   0
      Picture         =   "Game.frx":E3CF
      Stretch         =   -1  'True
      Top             =   360
      Width           =   9615
   End
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub English_Click()
If Language = 1 Then
    Language = 0
    English.BorderStyle = 1
    Romanian.BorderStyle = 0
    Label5.Caption = "Choose a game"
    Label3.Caption = "Exit"
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label9.FontSize <> 9 Then Label9.FontSize = 9
If Shape8.BackColor = &HFF& Then Shape8.BackColor = &HC0&
If Shape5.BackColor = &HFF& Then Shape5.BackColor = &HC0&
If Shape12.BackColor = &HFF& Then Shape12.BackColor = &HC0&
End Sub

Private Sub Label11_Click()
WindowState = 1
Unload GameAbout
End Sub

Private Sub Label12_Click()
If SoundChooser.Tag = 2 Then
    SoundChooser.Tag = 1
    SoundChooser.Show
ElseIf SoundChooser.Tag = 1 Then
    SoundChooser.Tag = 2
    SoundChooser.Hide
End If
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label9.FontSize = 9 Then Label9.FontSize = 11
End Sub

Private Sub Romanian_Click()
If Language = 0 Then
    Language = 1
    English.BorderStyle = 0
    Romanian.BorderStyle = 1
    Label5.Caption = "Alege un joc"
    Label3.Caption = "Ieºire"
End If
End Sub

Private Sub Form_activate()
Unload Sudoku
Unload TicTacToe
Unload Impasse
TicTacToe.thelevel = 2
If Language = 1 Then Language = 0: Romanian_Click
End Sub

Private Sub Label4_Click()
End
End Sub

Private Sub Label8_Click()
End
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape8.BackColor = &HC0& Then Shape8.BackColor = &HFF&
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape5.BackColor = &HC0& Then Shape5.BackColor = &HFF&
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape12.BackColor = &HC0& Then Shape12.BackColor = &HFF&
End Sub

Private Sub Label9_Click()
WhichAbout = 0: GameAbout.Show
End Sub

Private Sub TicTacToeLabel_Click()
Unload Me
TicTacToe.Show
End Sub
Private Sub SudokuLabel_Click()
Unload Me
Sudoku.Show
End Sub
Private Sub ImpasseLabel_Click()
Unload Me
Impasse.Show
End Sub

Private Sub tictactoelabel_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.ForeColor = &H808080 Then Label1.ForeColor = &H80000012
End Sub
Private Sub sudokulabel_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label2.ForeColor = &H808080 Then Label2.ForeColor = &H80000012
End Sub
Private Sub impasselabel_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label10.ForeColor = &H808080 Then Label10.ForeColor = &H80000012
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.ForeColor = &H808080 Then Label3.ForeColor = &H80000012
End Sub
Private Sub image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.ForeColor = &H80000012 Then Label1.ForeColor = &H808080
If Label2.ForeColor = &H80000012 = True Then Label2.ForeColor = &H808080
If Label3.ForeColor = &H80000012 = True Then Label3.ForeColor = &H808080
If Label10.ForeColor = &H80000012 = True Then Label10.ForeColor = &H808080
If Shape8.BackColor = &HFF& Then Shape8.BackColor = &HC0&
If Shape5.BackColor = &HFF& Then Shape5.BackColor = &HC0&
If Shape12.BackColor = &HFF& Then Shape12.BackColor = &HC0&
If Label9.FontSize <> 9 Then Label9.FontSize = 9
End Sub

Private Sub Timer1_Timer()
Unload Welcome
Welcome.music.Command = "stop"
Timer1.Enabled = False
End Sub
