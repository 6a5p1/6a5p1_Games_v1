VERSION 5.00
Begin VB.Form GameQuit 
   BorderStyle     =   0  'None
   Caption         =   "Exit Confirmation"
   ClientHeight    =   5205
   ClientLeft      =   6555
   ClientTop       =   5025
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "GameQuit.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5205
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure you want to quit ?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1335
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No"
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
      Left            =   3360
      MouseIcon       =   "GameQuit.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
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
      Left            =   600
      MouseIcon       =   "GameQuit.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   6120
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   0
      X1              =   -360
      X2              =   6120
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label GameQuitTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sudoku"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   44.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404080&
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      Height          =   615
      Left            =   600
      Shape           =   2  'Oval
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404080&
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      Height          =   615
      Left            =   3360
      Shape           =   2  'Oval
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   0
      Picture         =   "GameQuit.frx":091E
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   6135
   End
End
Attribute VB_Name = "GameQuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
If WhichQuit = 1 Then GameQuitTitle.Caption = "Tic Tac Toe"
If WhichQuit = 2 Then GameQuitTitle.Caption = "Impasse"
If Language = 1 Then
    Label2.Caption = "Eºti sigur cã vrei sã ieºi?"
    Label3.Caption = "Da"
    Label4.Caption = "Nu"
End If
End Sub

Private Sub Label3_Click()
Unload Sudoku
Unload GameAbout
Unload GameHelp
Unload GameWell
Unload Me
Game.Show
End Sub

Private Sub Label4_Click()
Unload Me
End Sub
Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape1.BackColor = &H40C0& Then Shape1.BackColor = &H404080
If Shape2.BackColor = &H40C0& Then Shape2.BackColor = &H404080
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape1.BackColor = &H404080 Then Shape1.BackColor = &H40C0&
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape2.BackColor = &H404080 Then Shape2.BackColor = &H40C0&
End Sub

