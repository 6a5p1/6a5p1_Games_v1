VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Sudoku 
   BackColor       =   &H00004000&
   BorderStyle     =   0  'None
   Caption         =   "6a5p1 SUDOKU"
   ClientHeight    =   7725
   ClientLeft      =   2115
   ClientTop       =   3750
   ClientWidth     =   8040
   Icon            =   "Sudoku.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Sudoku.frx":030A
   MousePointer    =   99  'Custom
   Picture         =   "Sudoku.frx":0614
   ScaleHeight     =   7725
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7440
      Top             =   3600
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   7560
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "\"
      Orientation     =   2
   End
   Begin VB.Label Label13 
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
      Left            =   6968
      MouseIcon       =   "Sudoku.frx":18B5
      MousePointer    =   99  'Custom
      TabIndex        =   111
      Top             =   75
      Width           =   255
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   255
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label5 
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
      Left            =   7328
      MouseIcon       =   "Sudoku.frx":1BBF
      MousePointer    =   99  'Custom
      TabIndex        =   110
      Top             =   60
      Width           =   255
   End
   Begin VB.Label exitgame 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   7200
      MouseIcon       =   "Sudoku.frx":1EC9
      MousePointer    =   99  'Custom
      TabIndex        =   109
      Top             =   480
      Width           =   735
   End
   Begin VB.Label newgame 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   120
      MouseIcon       =   "Sudoku.frx":21D3
      MousePointer    =   99  'Custom
      TabIndex        =   108
      Top             =   480
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   8040
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Label Label4 
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
      Left            =   7680
      MouseIcon       =   "Sudoku.frx":24DD
      MousePointer    =   99  'Custom
      TabIndex        =   106
      Top             =   60
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   255
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Mute 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5160
      MouseIcon       =   "Sudoku.frx":27E7
      MousePointer    =   99  'Custom
      TabIndex        =   105
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   11
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      MouseIcon       =   "Sudoku.frx":2AF1
      MousePointer    =   99  'Custom
      TabIndex        =   104
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   10
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   2
      X1              =   960
      X2              =   7050
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Wait 
      Alignment       =   2  'Center
      Caption         =   "Creating puzzle... Please wait..."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   103
      Top             =   6120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Index           =   9
      Left            =   2760
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label Numbers 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      Caption         =   "D e l e t e"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   2760
      MouseIcon       =   "Sudoku.frx":2DFB
      MousePointer    =   99  'Custom
      TabIndex        =   91
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6a5p1                                                Sudoku"
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
      Height          =   495
      Left            =   0
      TabIndex        =   102
      Top             =   7200
      Width           =   8055
   End
   Begin VB.Label HelpGame 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":3105
      MousePointer    =   99  'Custom
      TabIndex        =   99
      Top             =   480
      Width           =   855
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7080
      TabIndex        =   98
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Index           =   8
      Left            =   7080
      Top             =   6600
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Index           =   7
      Left            =   6240
      Top             =   6600
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Index           =   6
      Left            =   5400
      Top             =   6600
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Index           =   5
      Left            =   4440
      Top             =   6600
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Index           =   4
      Left            =   3600
      Top             =   6600
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Index           =   3
      Left            =   2760
      Top             =   6600
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Index           =   2
      Left            =   1800
      Top             =   6600
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Index           =   1
      Left            =   960
      Top             =   6600
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Numbers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      MouseIcon       =   "Sudoku.frx":340F
      MousePointer    =   99  'Custom
      TabIndex        =   97
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Numbers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7080
      MouseIcon       =   "Sudoku.frx":3719
      MousePointer    =   99  'Custom
      TabIndex        =   96
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Numbers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":3A23
      MousePointer    =   99  'Custom
      TabIndex        =   95
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Numbers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5400
      MouseIcon       =   "Sudoku.frx":3D2D
      MousePointer    =   99  'Custom
      TabIndex        =   94
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Numbers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4440
      MouseIcon       =   "Sudoku.frx":4037
      MousePointer    =   99  'Custom
      TabIndex        =   93
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Numbers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3600
      MouseIcon       =   "Sudoku.frx":4341
      MousePointer    =   99  'Custom
      TabIndex        =   92
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Numbers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      MouseIcon       =   "Sudoku.frx":464B
      MousePointer    =   99  'Custom
      TabIndex        =   90
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Numbers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      MouseIcon       =   "Sudoku.frx":4955
      MousePointer    =   99  'Custom
      TabIndex        =   89
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Numbers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      MouseIcon       =   "Sudoku.frx":4C5F
      MousePointer    =   99  'Custom
      TabIndex        =   88
      Top             =   6600
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   80
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   79
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   78
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   77
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   76
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   75
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   74
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   73
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   72
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   81
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":4F69
      MousePointer    =   99  'Custom
      TabIndex        =   87
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   80
      Left            =   5640
      MouseIcon       =   "Sudoku.frx":5273
      MousePointer    =   99  'Custom
      TabIndex        =   86
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   79
      Left            =   5040
      MouseIcon       =   "Sudoku.frx":557D
      MousePointer    =   99  'Custom
      TabIndex        =   85
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   78
      Left            =   4320
      MouseIcon       =   "Sudoku.frx":5887
      MousePointer    =   99  'Custom
      TabIndex        =   84
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   77
      Left            =   3720
      MouseIcon       =   "Sudoku.frx":5B91
      MousePointer    =   99  'Custom
      TabIndex        =   83
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   76
      Left            =   3120
      MouseIcon       =   "Sudoku.frx":5E9B
      MousePointer    =   99  'Custom
      TabIndex        =   82
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   75
      Left            =   2400
      MouseIcon       =   "Sudoku.frx":61A5
      MousePointer    =   99  'Custom
      TabIndex        =   81
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   74
      Left            =   1800
      MouseIcon       =   "Sudoku.frx":64AF
      MousePointer    =   99  'Custom
      TabIndex        =   80
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   73
      Left            =   1200
      MouseIcon       =   "Sudoku.frx":67B9
      MousePointer    =   99  'Custom
      TabIndex        =   79
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   71
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   70
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   69
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   68
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   67
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   66
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   65
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   64
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   63
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   72
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":6AC3
      MousePointer    =   99  'Custom
      TabIndex        =   78
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   71
      Left            =   5640
      MouseIcon       =   "Sudoku.frx":6DCD
      MousePointer    =   99  'Custom
      TabIndex        =   77
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   70
      Left            =   5040
      MouseIcon       =   "Sudoku.frx":70D7
      MousePointer    =   99  'Custom
      TabIndex        =   76
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   69
      Left            =   4320
      MouseIcon       =   "Sudoku.frx":73E1
      MousePointer    =   99  'Custom
      TabIndex        =   75
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   68
      Left            =   3720
      MouseIcon       =   "Sudoku.frx":76EB
      MousePointer    =   99  'Custom
      TabIndex        =   74
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   67
      Left            =   3120
      MouseIcon       =   "Sudoku.frx":79F5
      MousePointer    =   99  'Custom
      TabIndex        =   73
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   66
      Left            =   2400
      MouseIcon       =   "Sudoku.frx":7CFF
      MousePointer    =   99  'Custom
      TabIndex        =   72
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   65
      Left            =   1800
      MouseIcon       =   "Sudoku.frx":8009
      MousePointer    =   99  'Custom
      TabIndex        =   71
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   64
      Left            =   1200
      MouseIcon       =   "Sudoku.frx":8313
      MousePointer    =   99  'Custom
      TabIndex        =   70
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   62
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   61
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   60
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   59
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   58
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   57
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   56
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   55
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   54
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   63
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":861D
      MousePointer    =   99  'Custom
      TabIndex        =   69
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   62
      Left            =   5640
      MouseIcon       =   "Sudoku.frx":8927
      MousePointer    =   99  'Custom
      TabIndex        =   68
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   61
      Left            =   5040
      MouseIcon       =   "Sudoku.frx":8C31
      MousePointer    =   99  'Custom
      TabIndex        =   67
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   60
      Left            =   4320
      MouseIcon       =   "Sudoku.frx":8F3B
      MousePointer    =   99  'Custom
      TabIndex        =   66
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   59
      Left            =   3720
      MouseIcon       =   "Sudoku.frx":9245
      MousePointer    =   99  'Custom
      TabIndex        =   65
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   58
      Left            =   3120
      MouseIcon       =   "Sudoku.frx":954F
      MousePointer    =   99  'Custom
      TabIndex        =   64
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   57
      Left            =   2400
      MouseIcon       =   "Sudoku.frx":9859
      MousePointer    =   99  'Custom
      TabIndex        =   63
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   56
      Left            =   1800
      MouseIcon       =   "Sudoku.frx":9B63
      MousePointer    =   99  'Custom
      TabIndex        =   62
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   55
      Left            =   1200
      MouseIcon       =   "Sudoku.frx":9E6D
      MousePointer    =   99  'Custom
      TabIndex        =   61
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   53
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   52
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   51
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   50
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   49
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   48
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   47
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   46
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   45
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   54
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":A177
      MousePointer    =   99  'Custom
      TabIndex        =   60
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   53
      Left            =   5640
      MouseIcon       =   "Sudoku.frx":A481
      MousePointer    =   99  'Custom
      TabIndex        =   59
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   52
      Left            =   5040
      MouseIcon       =   "Sudoku.frx":A78B
      MousePointer    =   99  'Custom
      TabIndex        =   58
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   51
      Left            =   4320
      MouseIcon       =   "Sudoku.frx":AA95
      MousePointer    =   99  'Custom
      TabIndex        =   57
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   50
      Left            =   3720
      MouseIcon       =   "Sudoku.frx":AD9F
      MousePointer    =   99  'Custom
      TabIndex        =   56
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   49
      Left            =   3120
      MouseIcon       =   "Sudoku.frx":B0A9
      MousePointer    =   99  'Custom
      TabIndex        =   55
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   2400
      MouseIcon       =   "Sudoku.frx":B3B3
      MousePointer    =   99  'Custom
      TabIndex        =   54
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   1800
      MouseIcon       =   "Sudoku.frx":B6BD
      MousePointer    =   99  'Custom
      TabIndex        =   53
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   46
      Left            =   1200
      MouseIcon       =   "Sudoku.frx":B9C7
      MousePointer    =   99  'Custom
      TabIndex        =   52
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   44
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   43
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   42
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   41
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   40
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   39
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   38
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   37
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   36
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":BCD1
      MousePointer    =   99  'Custom
      TabIndex        =   51
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   44
      Left            =   5640
      MouseIcon       =   "Sudoku.frx":BFDB
      MousePointer    =   99  'Custom
      TabIndex        =   50
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   43
      Left            =   5040
      MouseIcon       =   "Sudoku.frx":C2E5
      MousePointer    =   99  'Custom
      TabIndex        =   49
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   42
      Left            =   4320
      MouseIcon       =   "Sudoku.frx":C5EF
      MousePointer    =   99  'Custom
      TabIndex        =   48
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   41
      Left            =   3720
      MouseIcon       =   "Sudoku.frx":C8F9
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   40
      Left            =   3120
      MouseIcon       =   "Sudoku.frx":CC03
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   39
      Left            =   2400
      MouseIcon       =   "Sudoku.frx":CF0D
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   38
      Left            =   1800
      MouseIcon       =   "Sudoku.frx":D217
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   37
      Left            =   1200
      MouseIcon       =   "Sudoku.frx":D521
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   35
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   34
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   33
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   32
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   31
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   30
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   29
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   28
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   27
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":D82B
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   5640
      MouseIcon       =   "Sudoku.frx":DB35
      MousePointer    =   99  'Custom
      TabIndex        =   41
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   5040
      MouseIcon       =   "Sudoku.frx":DE3F
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   4320
      MouseIcon       =   "Sudoku.frx":E149
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   3720
      MouseIcon       =   "Sudoku.frx":E453
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   3120
      MouseIcon       =   "Sudoku.frx":E75D
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   2400
      MouseIcon       =   "Sudoku.frx":EA67
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   1800
      MouseIcon       =   "Sudoku.frx":ED71
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   1200
      MouseIcon       =   "Sudoku.frx":F07B
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   26
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   25
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   24
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   23
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   22
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   21
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   20
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   19
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   18
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":F385
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   5640
      MouseIcon       =   "Sudoku.frx":F68F
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   5040
      MouseIcon       =   "Sudoku.frx":F999
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   4320
      MouseIcon       =   "Sudoku.frx":FCA3
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   3720
      MouseIcon       =   "Sudoku.frx":FFAD
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   3120
      MouseIcon       =   "Sudoku.frx":102B7
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   2400
      MouseIcon       =   "Sudoku.frx":105C1
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   1800
      MouseIcon       =   "Sudoku.frx":108CB
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   1200
      MouseIcon       =   "Sudoku.frx":10BD5
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   17
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   16
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   15
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   14
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   13
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   12
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   11
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   10
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   0
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":10EDF
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   5640
      MouseIcon       =   "Sudoku.frx":111E9
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   5040
      MouseIcon       =   "Sudoku.frx":114F3
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   4320
      MouseIcon       =   "Sudoku.frx":117FD
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   3720
      MouseIcon       =   "Sudoku.frx":11B07
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   3120
      MouseIcon       =   "Sudoku.frx":11E11
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   2400
      MouseIcon       =   "Sudoku.frx":1211B
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   1800
      MouseIcon       =   "Sudoku.frx":12425
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   1200
      MouseIcon       =   "Sudoku.frx":1272F
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   9
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   8
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   7
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   6
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   5
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   4
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   3
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   2
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   375
      Index           =   1
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":12A39
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5640
      MouseIcon       =   "Sudoku.frx":12D43
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5040
      MouseIcon       =   "Sudoku.frx":1304D
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4320
      MouseIcon       =   "Sudoku.frx":13357
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3720
      MouseIcon       =   "Sudoku.frx":13661
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3120
      MouseIcon       =   "Sudoku.frx":1396B
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1200
      MouseIcon       =   "Sudoku.frx":13C75
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      MouseIcon       =   "Sudoku.frx":13F7F
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2400
      MouseIcon       =   "Sudoku.frx":14289
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   8040
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label exitgame2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   7200
      TabIndex        =   15
      Top             =   720
      Width           =   735
   End
   Begin VB.Label changesound 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sound"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5160
      MouseIcon       =   "Sudoku.frx":14593
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   480
      Width           =   975
   End
   Begin VB.Label changesquares 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change theme"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4080
      MouseIcon       =   "Sudoku.frx":1489D
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.Label changebackground 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change bkground"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3000
      MouseIcon       =   "Sudoku.frx":14BA7
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Make grid"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      MouseIcon       =   "Sudoku.frx":14EB1
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label newgame2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New game"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Import 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Import"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      MouseIcon       =   "Sudoku.frx":151BB
      MousePointer    =   99  'Custom
      TabIndex        =   100
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Export 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      MouseIcon       =   "Sudoku.frx":154C5
      MousePointer    =   99  'Custom
      TabIndex        =   101
      Top             =   720
      Width           =   855
   End
   Begin VB.Label aboutgame 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6240
      MouseIcon       =   "Sudoku.frx":157CF
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   735
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   8
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   9
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   1
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   495
      Index           =   2
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   495
      Index           =   3
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   4
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   5
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   7
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   735
      Index           =   6
      Left            =   7200
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6a5p1 SUDOKU v1.0"
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
      Left            =   2640
      MouseIcon       =   "Sudoku.frx":15AD9
      MousePointer    =   99  'Custom
      TabIndex        =   107
      Top             =   60
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   0
      Picture         =   "Sudoku.frx":15DE3
      Stretch         =   -1  'True
      Top             =   360
      Width           =   8055
   End
   Begin VB.Image Image3 
      Height          =   7335
      Left            =   0
      Picture         =   "Sudoku.frx":29828
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Image Image5 
      Height          =   7335
      Left            =   0
      Picture         =   "Sudoku.frx":55F03
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Image Image4 
      Height          =   7335
      Left            =   0
      Picture         =   "Sudoku.frx":5FAE6
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Image Image2 
      Height          =   7335
      Left            =   0
      Picture         =   "Sudoku.frx":70C0D
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   255
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   60
      Width           =   255
   End
End
Attribute VB_Name = "Sudoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer
Dim lev As Integer
Dim corect As Boolean
Dim csp As Integer
Dim teren(81) As Integer
Dim era(81) As Integer
Dim puzzle As Integer
Public Init As Boolean
Public PuzzleName As String
Private whattoput As String

Private Sub changebackground_Click()
If Image1.Visible = True Then
  Image1.Visible = False
  Image2.Visible = True
Else
 If Image2.Visible = True Then
 Image2.Visible = False
 Image3.Visible = True
Else
 If Image3.Visible = True Then
 Image3.Visible = False
 Image4.Visible = True
Else
 If Image4.Visible = True Then
 Image4.Visible = False
 Image5.Visible = True
Else
 If Image5.Visible = True Then
 Image5.Visible = False
 Image1.Visible = True
End If
End If
End If
End If
End If
End Sub
Private Sub color()
For i = 1 To 81
t(i).BackColor = t(0).BackColor
Next i
End Sub
Private Sub verify()
color
corect = True
For X = 1 To 81 Step 9
  For j = X To X + 7
    For i = j + 1 To X + 8
    If teren(j) = teren(i) And teren(i) <> 0 Then
      corect = False
      t(i).BackColor = vbYellow
      t(j).BackColor = vbYellow
    End If
    Next i
  Next j
Next X
For X = 1 To 9
  For j = X To X + 63 Step 9
    For i = j + 9 To X + 72 Step 9
    If teren(j) = teren(i) And teren(i) <> 0 Then
      corect = False
      t(i).BackColor = vbYellow
      t(j).BackColor = vbYellow
    End If
    Next i
  Next j
Next X
For i = 1 To 7 Step 3
  If teren(i) = teren(i + 10) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 10).BackColor = vbYellow
  If teren(i) = teren(i + 11) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 11).BackColor = vbYellow
  If teren(i) = teren(i + 19) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
  If teren(i) = teren(i + 20) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 1) = teren(i + 9) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 9).BackColor = vbYellow
  If teren(i + 1) = teren(i + 11) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 11).BackColor = vbYellow
  If teren(i + 1) = teren(i + 18) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 1) = teren(i + 20) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 2) = teren(i + 9) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 9).BackColor = vbYellow
  If teren(i + 2) = teren(i + 10) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 10).BackColor = vbYellow
  If teren(i + 2) = teren(i + 18) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 2) = teren(i + 19) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
  If teren(i + 9) = teren(i + 19) And teren(i + 9) <> 0 Then corect = False: t(i + 9).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
  If teren(i + 9) = teren(i + 20) And teren(i + 9) <> 0 Then corect = False: t(i + 9).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 10) = teren(i + 18) And teren(i + 10) <> 0 Then corect = False: t(i + 10).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 10) = teren(i + 20) And teren(i + 10) <> 0 Then corect = False: t(i + 10).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 11) = teren(i + 18) And teren(i + 11) <> 0 Then corect = False: t(i + 11).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 11) = teren(i + 19) And teren(i + 11) <> 0 Then corect = False: t(i + 11).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
Next i
For i = 28 To 36 Step 3
  If teren(i) = teren(i + 10) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 10).BackColor = vbYellow
  If teren(i) = teren(i + 11) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 11).BackColor = vbYellow
  If teren(i) = teren(i + 19) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
  If teren(i) = teren(i + 20) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 1) = teren(i + 9) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 9).BackColor = vbYellow
  If teren(i + 1) = teren(i + 11) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 11).BackColor = vbYellow
  If teren(i + 1) = teren(i + 18) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 1) = teren(i + 20) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 2) = teren(i + 9) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 9).BackColor = vbYellow
  If teren(i + 2) = teren(i + 10) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 10).BackColor = vbYellow
  If teren(i + 2) = teren(i + 18) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 2) = teren(i + 19) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
  If teren(i + 9) = teren(i + 19) And teren(i + 9) <> 0 Then corect = False: t(i + 9).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
  If teren(i + 9) = teren(i + 20) And teren(i + 9) <> 0 Then corect = False: t(i + 9).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 10) = teren(i + 18) And teren(i + 10) <> 0 Then corect = False: t(i + 10).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 10) = teren(i + 20) And teren(i + 10) <> 0 Then corect = False: t(i + 10).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 11) = teren(i + 18) And teren(i + 11) <> 0 Then corect = False: t(i + 11).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 11) = teren(i + 19) And teren(i + 11) <> 0 Then corect = False: t(i + 11).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
Next i
For i = 55 To 61 Step 3
  If teren(i) = teren(i + 10) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 10).BackColor = vbYellow
  If teren(i) = teren(i + 11) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 11).BackColor = vbYellow
  If teren(i) = teren(i + 19) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
  If teren(i) = teren(i + 20) And teren(i) <> 0 Then corect = False: t(i).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 1) = teren(i + 9) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 9).BackColor = vbYellow
  If teren(i + 1) = teren(i + 11) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 11).BackColor = vbYellow
  If teren(i + 1) = teren(i + 18) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 1) = teren(i + 20) And teren(i + 1) <> 0 Then corect = False: t(i + 1).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 2) = teren(i + 9) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 9).BackColor = vbYellow
  If teren(i + 2) = teren(i + 10) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 10).BackColor = vbYellow
  If teren(i + 2) = teren(i + 18) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 2) = teren(i + 19) And teren(i + 2) <> 0 Then corect = False: t(i + 2).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
  If teren(i + 9) = teren(i + 19) And teren(i + 9) <> 0 Then corect = False: t(i + 9).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
  If teren(i + 9) = teren(i + 20) And teren(i + 9) <> 0 Then corect = False: t(i + 9).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 10) = teren(i + 18) And teren(i + 10) <> 0 Then corect = False: t(i + 10).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 10) = teren(i + 20) And teren(i + 10) <> 0 Then corect = False: t(i + 10).BackColor = vbYellow: t(i + 20).BackColor = vbYellow
  If teren(i + 11) = teren(i + 18) And teren(i + 11) <> 0 Then corect = False: t(i + 11).BackColor = vbYellow: t(i + 18).BackColor = vbYellow
  If teren(i + 11) = teren(i + 19) And teren(i + 11) <> 0 Then corect = False: t(i + 11).BackColor = vbYellow: t(i + 19).BackColor = vbYellow
Next i
For i = 1 To 81
If teren(i) = 0 Then corect = False
Next i
If corect = True Then WhichWell = 3: GameWell.Show
4
End Sub
Private Sub levelload()
For i = 1 To 81
era(i) = 0
Next i
For i = 0 To 80
If nil(i) > 0 And nil(i) < 10 Then t(i + 1).Caption = nil(i): t(i + 1).ForeColor = vbBlack: teren(i + 1) = nil(i): era(i + 1) = 1
If nil(i) = 0 Then t(i + 1).Caption = "": teren(i + 1) = 0
Next i
End Sub

Private Sub Export_Click()
SudokuExport.Show
'With dlgCommonDialog
'        .FileName = PuzzleName
'        .DialogTitle = "Export"
'        .CancelError = False
'        .Filter = "Sudoku puzzle files(*.sdk)|*.sdk"
'        .ShowSave
'        If Len(.FileName) = 0 Then
'            Exit Sub
'        End If
'        NomFicConf = .FileName
'    End With
'    NumFicConf = FreeFile
'    On Error Resume Next
'    Open NomFicConf For Output As NumFicConf
'    For i = 1 To 81
'        If t(i).Caption <> "" Then
'            whattoput = t(i).Caption & " "
'        Else
'            whattoput = "0 "
'        End If
'    Print #NumFicConf, whattoput
'    Next i
'    Close NumFicConf
End Sub
Public Sub Export2()
'With dlgCommonDialog
'        .FileName = PuzzleName
'        .DialogTitle = "Export"
'        .CancelError = False
'        .Filter = "Sudoku puzzle files(*.sdk)|*.sdk"
'        .ShowSave
'        If Len(.FileName) = 0 Then
'            Exit Sub
'        End If
'        NomFicConf = .FileName
'    End With
    NumFicConf = FreeFile
    On Error Resume Next
    Open NomFicConf For Output As NumFicConf
    For i = 1 To 81
        If t(i).Caption <> "" Then
            whattoput = t(i).Caption & " "
        Else
            whattoput = "0 "
        End If
    Print #NumFicConf, whattoput
    Next i
    Close NumFicConf
End Sub

Private Sub Form_Load()
Unload Game
WhichAbout = 2
WhichHelp = 0
WhichQuit = 0
If Language = 1 Then
  newgame2.Caption = "Joc nou"
  Import.Caption = "Importã"
  Export.Caption = "Exportã"
  Level.Caption = "Grilã nouã"
  Label1.Caption = "Uºor"
  changebackground.Caption = "Schimbã fundal"
  changesquares.Caption = "Schimbã temã"
  changesound.Caption = "Sunet"
  Mute.Caption = "Oprit"
  HelpGame.Caption = "Ajutor"
  aboutgame.Caption = "Despre"
  exitgame2.Caption = " Ieºire"
  Wait.Caption = "Se creeazã puzzle-ul. Asteptaþi..."
  Numbers(9).Caption = "ª t e r g e"
End If
NomFicConf = ""
PuzzleName = ""
csp = 9
lev = 1
k = 0
Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.FontSize <> 9 Then Label3.FontSize = 9
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape5.BackColor = &HFF& Then Shape5.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
End Sub

Private Sub HelpGame_Click()
GameHelp.Show
End Sub

Private Sub Import_Click()
'    With dlgCommonDialog
'        .DialogTitle = "Import"
'        .CancelError = False
'        .Filter = "Sudoku puzzle files(*.sdk)|*.sdk"
'        .ShowOpen
'        If Len(.FileName) = 0 Then
'            Exit Sub
'        End If
'        NomFicConf = .FileName
'    End With
'NumFicConf = FreeFile
'k = 0
'On Error Resume Next
'Open NomFicConf For Input As NumFicConf
'While Not EOF(NumFicConf)
'  Input #NumFicConf, nil(k)
'  k = k + 1
'Wend
'Close NumFicConf
'levelload
'color
SudokuImport.Show
End Sub
Public Sub Import2()
NumFicConf = FreeFile
k = 0
On Error Resume Next
Open NomFicConf For Input As NumFicConf
While Not EOF(NumFicConf)
  Input #NumFicConf, nil(k)
  k = k + 1
Wend
Close NumFicConf
levelload
color
End Sub


Private Sub Label1_Click()
If lev = 1 Then
    lev = 2
    If Language = 0 Then
        Label1.Caption = "Medium"
    Else
        Label1.Caption = "Mediu"
    End If
ElseIf lev = 2 Then
    lev = 3
    If Language = 0 Then
        Label1.Caption = "Hard"
    Else
        Label1.Caption = "Greu"
    End If
ElseIf lev = 3 Then
    lev = 1
    If Language = 0 Then
        Label1.Caption = "Easy"
    Else
        Label1.Caption = "Usor"
    End If
End If
End Sub

Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape5.BackColor = &HFF& Then Shape5.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub
Private Sub image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape5.BackColor = &HFF& Then Shape5.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape5.BackColor = &HFF& Then Shape5.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub
Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape5.BackColor = &HFF& Then Shape5.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub
Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape5.BackColor = &HFF& Then Shape5.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub

Private Sub Label13_Click()
If SoundChooser.Tag = 2 Then
    SoundChooser.Tag = 1
    SoundChooser.Show
ElseIf SoundChooser.Tag = 1 Then
    SoundChooser.Tag = 2
    SoundChooser.Hide
End If
End Sub

Private Sub Label3_Click()
GameAbout.Show
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.FontSize = 9 Then Label3.FontSize = 11
End Sub

Private Sub Label4_Click()
GameQuit.Show
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HC0& Then Shape4.BackColor = &HFF&
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape5.BackColor = &HC0& Then Shape5.BackColor = &HFF&
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape13.BackColor = &HC0& Then Shape13.BackColor = &HFF&
End Sub

Private Sub Label5_Click()
WindowState = 1
Unload GameWell
Unload GameAbout
Unload GameHelp
Unload GameQuit
End Sub

Public Sub Level_Click()
Wait.Visible = True
CreerGrille (lev)
levelload
color
Wait.Visible = False
End Sub
Private Sub changesound_Click()
SoundChooser.Show
End Sub

Private Sub changesquares_Click()
If t(0).BackColor = &HC0FFC0 Then
  For i = 0 To 81
  t(i).BackColor = &HFFFFC0
  Next i
  For i = 0 To 9
  Numbers(i).BackColor = &HFFFFC0
  Next i
  For i = 0 To 11
  Shape3(i).BackColor = &HFFFF80
  Next i
  verify
Else
If t(0).BackColor = &HFFFFC0 Then
  For i = 0 To 81
  t(i).BackColor = &HFFC0FF
  Next i
  For i = 0 To 9
  Numbers(i).BackColor = &HFFC0FF
  Next i
  For i = 0 To 11
  Shape3(i).BackColor = &HFF80FF
  Next i
  verify
Else
If t(0).BackColor = &HFFC0FF Then
  For i = 0 To 81
  t(i).BackColor = &HC0E0FF
  Next i
  For i = 0 To 9
  Numbers(i).BackColor = &HC0E0FF
  Next i
  For i = 0 To 11
  Shape3(i).BackColor = &H80C0FF
  Next i
  verify
Else
If t(0).BackColor = &HC0E0FF Then
  For i = 0 To 81
  t(i).BackColor = vbWhite
  Next i
  For i = 0 To 9
  Numbers(i).BackColor = vbWhite
  Next i
  For i = 0 To 11
  Shape3(i).BackColor = &HE0E0E0
  Next i
  verify
Else
If t(0).BackColor = vbWhite Then
  For i = 0 To 81
  t(i).BackColor = &HC0FFFF
  Next i
  For i = 0 To 9
  Numbers(i).BackColor = &HC0FFFF
  Next i
  For i = 0 To 11
  Shape3(i).BackColor = &H80FFFF
  Next i
  verify
Else
If t(0).BackColor = &HC0FFFF Then
  For i = 0 To 81
  t(i).BackColor = &HC0C0FF
  Next i
  For i = 0 To 9
  Numbers(i).BackColor = &HC0C0FF
  Next i
  For i = 0 To 11
  Shape3(i).BackColor = &H8080FF
  Next i
  verify
Else
If t(0).BackColor = &HC0C0FF Then
  For i = 0 To 81
  t(i).BackColor = &HFFC0C0
  Next i
  For i = 0 To 9
  Numbers(i).BackColor = &HFFC0C0
  Next i
  For i = 0 To 11
  Shape3(i).BackColor = &HFF8080
  Next i
  verify
Else
If t(0).BackColor = &HFFC0C0 Then
  For i = 0 To 81
  t(i).BackColor = &HC0FFC0
  Next i
  For i = 0 To 9
  Numbers(i).BackColor = &HC0FFC0
  Next i
  For i = 0 To 11
  Shape3(i).BackColor = &H80FF80
  Next i
  verify
End If
End If
End If
End If
End If
End If
End If
End If
Numbers(9).BackColor = &H8080&
csp = 9
End Sub

Private Sub exitgame_Click()
GameQuit.Show
End Sub
Private Sub aboutgame_Click()
GameAbout.Show
End Sub

Private Sub Mute_Click()
Unload SoundChooser
'SoundChooser.muzica.Command = "close"
End Sub

Private Sub newgame_Click()
levelload
color
End Sub

Private Sub Numbers_Click(Index As Integer)
For i = 0 To 9
Numbers(i).BackColor = t(0).BackColor
Next i
Numbers(Index).BackColor = &H8080&
csp = Index
End Sub

Private Sub t_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
    Case vbLeftButton
        If teren(Index) <> 0 And csp = 9 And era(Index) = 0 Then teren(Index) = 0: t(Index).Caption = "": verify
        If teren(Index) = 0 And csp >= 0 And csp < 9 Then teren(Index) = csp + 1: t(Index).ForeColor = &H80&: t(Index).Caption = csp + 1: verify
    Case vbRightButton
        If teren(Index) <> 0 And era(Index) = 0 Then teren(Index) = 0: t(Index).Caption = "": verify
End Select
End Sub

Private Sub Timer1_Timer()
Level_Click
levelload
Timer1.Enabled = False
End Sub

