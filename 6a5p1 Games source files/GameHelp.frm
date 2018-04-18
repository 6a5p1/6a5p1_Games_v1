VERSION 5.00
Begin VB.Form GameHelp 
   BackColor       =   &H80000011&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   5805
   ClientLeft      =   4230
   ClientTop       =   3000
   ClientWidth     =   6870
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "GameHelp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "GameHelp.frx":0442
   MousePointer    =   99  'Custom
   ScaleHeight     =   4006.714
   ScaleMode       =   0  'User
   ScaleWidth      =   6451.286
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"GameHelp.frx":074C
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      TabIndex        =   5
      Top             =   3120
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fill the grid with digits between 1 and 9 so that each digit appears only once in each row, column or in each 3x3 block."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help and Hints"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6615
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   6535.801
      Y1              =   3975.654
      Y2              =   3975.654
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6a5p1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   1440
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Left            =   4320
      MouseIcon       =   "GameHelp.frx":07DD
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404080&
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      Height          =   615
      Left            =   4320
      Shape           =   2  'Oval
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   6535.801
      Y1              =   20.707
      Y2              =   20.707
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sudoku"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   585
      Left            =   720
      TabIndex        =   0
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   0
      Picture         =   "GameHelp.frx":0AE7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "GameHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If Language = 1 And WhichHelp = 0 Then
  lblTitle.Caption = "Ajutor ºi Sfaturi"
  Label1.Caption = "Completeazã grila cu cifre între 1 ºi 9 astfel încât fiecare cifrã sã aparã o singurã datã în fiecare linie, coloanã sau casuþã 3x3."
  Label2.Caption = "Începe cu linia, coloana sau casuþa 3x3 cu cel mai mic numãr de cifre lipsã ºi încearcã sã gãseºti poziþia celor care lipsesc."
ElseIf Language = 1 And WhichHelp = 1 Then
  lblDisclaimer.Caption = "Tic Tac Toe"
  lblTitle.Caption = "Ajutor ºi Sfaturi"
  Label1.Caption = "Încearcã sã baþi adversarul completând spaþiile libere, in ordinea corectã. Fii primul care face 3 într-o linie, coloanã sau diagonalã."
  Label2.Caption = "Începe cu linia, coloana sau diagonala care iþi dã o mai mare ºansã de a câºtiga."
ElseIf Language = 0 And WhichHelp = 1 Then
  lblDisclaimer.Caption = "Tic Tac Toe"
  Label1.Caption = "Try to beat the opponent by filling the empty squares, in the right order. Be first who makes 3 in a line, column or diagonal."
  Label2.Caption = "Start with the line, the column or the diagonal which gives you a better chance to win."
ElseIf Language = 0 And WhichHelp = 2 Then
  lblDisclaimer.Caption = "Impasse"
  Label1.Caption = "Initially 9 pawns are situated in the south-west corner of the 6x6 table. Each player moves one pawn in the own direction, however free cells"
  Label2.Caption = "desires. The player moves the pieces from the bottom to upper side, and the computer clockwise. First player that can't move anymore looses the match."
ElseIf Language = 1 And WhichHelp = 2 Then
  lblDisclaimer.Caption = "Impasse"
  lblTitle.Caption = "Ajutor ºi Sfaturi"
  Label1.Caption = "Iniþial sunt plasaþi 9 pioni in colþul de sud-vest al tablei de 6x6. Pe rând, fiecare jucãtor deplaseazã un pion in direcþia proprie, oricâte cãsuþe"
  Label2.Caption = "libere doreºte. Jucãtorul mutã piesele de jos în sus, iar calculatorul de la stânga la dreapta. Primul jucãtor care nu mai poate muta pierde partida."
End If
End Sub

Private Sub Label3_Click()
Unload Me
End Sub
Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape1.BackColor = &H40C0& Then Shape1.BackColor = &H404080
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape1.BackColor = &H404080 Then Shape1.BackColor = &H40C0&
End Sub
