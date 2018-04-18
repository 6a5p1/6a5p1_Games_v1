VERSION 5.00
Begin VB.Form GameWell 
   BorderStyle     =   0  'None
   Caption         =   "Dialog Caption"
   ClientHeight    =   3930
   ClientLeft      =   4935
   ClientTop       =   5250
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "GameWell.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   3930
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      MouseIcon       =   "GameWell.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Draw !!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Not Bad"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1335
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404080&
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      Height          =   615
      Left            =   1920
      Shape           =   2  'Oval
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3975
      Left            =   0
      Picture         =   "GameWell.frx":0614
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "GameWell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
If TicTacToe.thelevel <> 4 Then
    If Language = 1 And WhichWell = 0 Then
        Label1(0).Caption = "Felicitãri"
        Label1(1).Caption = "Ai câºtigat!!!"
    ElseIf Language = 1 And WhichWell = 1 Then
        Label1(0).Caption = "Nu-i Rãu"
        Label1(1).Caption = "Egalitate!!!"
    ElseIf Language = 1 And WhichWell = 2 Then
        Label1(0).Caption = "Rãu Jucat"
        Label1(1).Caption = "Ai pierdut!!!"
    ElseIf Language = 1 And WhichWell = 3 Then
        Label1(0).Caption = "Felicitãri"
        Label1(1).Caption = "Puzzle comlet!!!"
    ElseIf Language = 0 And WhichWell = 0 Then
        Label1(0).Caption = "Very Well"
        Label1(1).Caption = "You won!!!"
    ElseIf Language = 0 And WhichWell = 2 Then
        Label1(0).Caption = "Too Bad"
        Label1(1).Caption = "You lost!!!"
    ElseIf Language = 0 And WhichWell = 3 Then
        Label1(0).Caption = "Congratulations"
        Label1(1).Caption = "Puzzle comlete!!!"
    End If
ElseIf TicTacToe.thelevel = 4 Then
    If Language = 1 And WhichWell = 1 Then
        Label1(0).Caption = "Nu-i Rãu"
        Label1(1).Caption = "Egalitate!!!"
    ElseIf Language = 0 And WhichWell = 1 Then
        Label1(0).Caption = "Not Bad"
        Label1(1).Caption = "Draw!!!"
    ElseIf Language = 0 And TicTacToe.xor0 = False Then
        Label1(0).Caption = "Very Well"
        Label1(1).Caption = "X won!!!"
    ElseIf Language = 0 And TicTacToe.xor0 = True Then
        Label1(0).Caption = "Very Well"
        Label1(1).Caption = "O won!!!"
    ElseIf Language = 1 And TicTacToe.xor0 = False Then
        Label1(0).Caption = "Felicitãri"
        Label1(1).Caption = "X câºtigã!!!"
    ElseIf Language = 0 And TicTacToe.xor0 = True Then
        Label1(0).Caption = "Felicitãri"
        Label1(1).Caption = "O câºtigã!!!"
    End If
End If
End Sub

Private Sub Label2_Click()
Unload Me
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape1.BackColor = &H404080 Then Shape1.BackColor = &H40C0&
End Sub
Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape1.BackColor = &H40C0& Then Shape1.BackColor = &H404080
End Sub
