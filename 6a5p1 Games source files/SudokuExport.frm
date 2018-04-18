VERSION 5.00
Begin VB.Form SudokuExport 
   BorderStyle     =   0  'None
   Caption         =   "Import"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "SudokuExport.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   8141.538
   ScaleMode       =   0  'User
   ScaleWidth      =   8194.615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      HideSelection   =   0   'False
      Left            =   960
      TabIndex        =   5
      Top             =   4440
      Width           =   4455
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
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
      Left            =   4920
      MouseIcon       =   "SudokuExport.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "SudokuExport.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
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
      Left            =   3600
      MouseIcon       =   "SudokuExport.frx":2378
      MousePointer    =   99  'Custom
      Picture         =   "SudokuExport.frx":2682
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFC0&
      Height          =   2235
      Left            =   3240
      MouseIcon       =   "SudokuExport.frx":43E6
      MousePointer    =   99  'Custom
      MultiSelect     =   2  'Extended
      Pattern         =   "*.sdk"
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H0080FF80&
      Height          =   2340
      Left            =   120
      MouseIcon       =   "SudokuExport.frx":46F0
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      MouseIcon       =   "SudokuExport.frx":49FA
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sudoku puzzle files (*.sdk)"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4080
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Export file"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   6615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5055
      Left            =   0
      Picture         =   "SudokuExport.frx":4D04
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "SudokuExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
Unload Me
End Sub

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
NomFicConf = File1.FileName
End Sub

'Private Sub File1_DblClick()
'If Len(File1.FileName) = 0 Then
    'Exit Sub
'End If
'NomFicConf = Dir1.Path & "\" & File1.FileName
'Sudoku.Export2
'End Sub
Private Sub Form_activate()
If Language = 1 Then
    Label1.Caption = "Export� un fi�ier"
    Label2.Caption = "Fi�iere puzzle Sudoku (*.sdk)"
    Cancel.Caption = "Anulare"
ElseIf Language = 0 Then
    Label1.Caption = "Export file"
    Label2.Caption = "Sudoku puzzle files (*.sdk)"
    Cancel.Caption = "Cancel"
End If
End Sub

Private Sub OK_Click()
If Text1.Text <> "" Then
    NomFicConf = Dir1.Path & "\" & Text1.Text & ".sdk"
ElseIf Text1.Text = "" Then
    NomFicConf = Dir1.Path & "\" & File1.FileName
End If
Sudoku.Export2
Unload Me
End Sub

'Private Sub Text1_Change()
'NomFicConf = Dir1.Path & "\" & Text1.Text
'End Sub
