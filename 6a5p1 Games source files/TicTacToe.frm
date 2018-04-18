VERSION 5.00
Begin VB.Form TicTacToe 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "6a5p1 TIC TAC TOE"
   ClientHeight    =   5670
   ClientLeft      =   4185
   ClientTop       =   1905
   ClientWidth     =   7200
   Icon            =   "TicTacToe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "TicTacToe.frx":030A
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   OLEDropMode     =   1  'Manual
   PaletteMode     =   1  'UseZOrder
   Picture         =   "TicTacToe.frx":0614
   ScaleHeight     =   5670
   ScaleMode       =   0  'User
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer whichlevel 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6840
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6840
      Top             =   2640
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
      Left            =   6128
      MouseIcon       =   "TicTacToe.frx":18B5
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   75
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      Index           =   2
      X1              =   960
      X2              =   6240
      Y1              =   1080
      Y2              =   1080
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
      Height          =   225
      Left            =   5280
      MouseIcon       =   "TicTacToe.frx":1BBF
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Level 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   960
      MouseIcon       =   "TicTacToe.frx":1EC9
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   480
      Width           =   975
   End
   Begin VB.Label newgame 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   120
      MouseIcon       =   "TicTacToe.frx":21D3
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   480
      Width           =   735
   End
   Begin VB.Label exitgame 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   6360
      MouseIcon       =   "TicTacToe.frx":24DD
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   480
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   7440
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
      Left            =   6840
      MouseIcon       =   "TicTacToe.frx":27E7
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   60
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   255
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6a5p1 TIC TAC TOE v1.0"
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
      Left            =   2040
      MouseIcon       =   "TicTacToe.frx":2AF1
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   60
      Width           =   3135
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
      Left            =   6360
      TabIndex        =   15
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   735
      Index           =   4
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   735
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
      Left            =   4200
      MouseIcon       =   "TicTacToe.frx":2DFB
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   8
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      Index           =   2
      X1              =   4560
      X2              =   5160
      Y1              =   5160
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      Index           =   0
      X1              =   2640
      X2              =   4560
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      Index           =   1
      X1              =   2040
      X2              =   2640
      Y1              =   4680
      Y2              =   5160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0    -    0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      FillStyle       =   7  'Diagonal Cross
      Height          =   495
      Index           =   2
      Left            =   3360
      Shape           =   2  'Oval
      Top             =   1800
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      FillStyle       =   7  'Diagonal Cross
      Height          =   735
      Index           =   1
      Left            =   2040
      Shape           =   2  'Oval
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   615
      Index           =   2
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   615
      Index           =   1
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      Height          =   615
      Index           =   0
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   615
      Index           =   5
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   615
      Index           =   4
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   615
      Index           =   3
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   615
      Index           =   8
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   615
      Index           =   7
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   5
      Height          =   615
      Index           =   6
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      Index           =   1
      X1              =   -120
      X2              =   7560
      Y1              =   5625
      Y2              =   5625
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   735
      Index           =   9
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
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
      Height          =   225
      Left            =   5280
      MouseIcon       =   "TicTacToe.frx":3105
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   7
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Level2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level medium"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   17
      Top             =   480
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
      Left            =   4200
      MouseIcon       =   "TicTacToe.frx":340F
      MousePointer    =   99  'Custom
      TabIndex        =   16
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
      Left            =   3120
      MouseIcon       =   "TicTacToe.frx":3719
      MousePointer    =   99  'Custom
      TabIndex        =   13
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
      Left            =   2040
      MouseIcon       =   "TicTacToe.frx":3A23
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      FillStyle       =   7  'Diagonal Cross
      Height          =   735
      Index           =   0
      Left            =   3960
      Shape           =   2  'Oval
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6a5p1                                  Tic Tac Toe"
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
      TabIndex        =   10
      Top             =   5160
      Width           =   7215
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   4080
      MouseIcon       =   "TicTacToe.frx":3D2D
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   3240
      MouseIcon       =   "TicTacToe.frx":4037
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   2400
      MouseIcon       =   "TicTacToe.frx":4341
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   4080
      MouseIcon       =   "TicTacToe.frx":464B
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3240
      MouseIcon       =   "TicTacToe.frx":4955
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2400
      MouseIcon       =   "TicTacToe.frx":4C5F
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   4080
      MouseIcon       =   "TicTacToe.frx":4F69
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3240
      MouseIcon       =   "TicTacToe.frx":5273
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2640
      Width           =   735
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
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.Label t 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2400
      MouseIcon       =   "TicTacToe.frx":557D
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2640
      Width           =   735
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
      Height          =   495
      Index           =   1
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   3
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   255
      Index           =   6
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   495
      Index           =   5
      Left            =   3120
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
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   5295
      Left            =   0
      Picture         =   "TicTacToe.frx":5887
      Stretch         =   -1  'True
      Top             =   360
      Width           =   7215
   End
   Begin VB.Image Image5 
      Height          =   5295
      Left            =   0
      Picture         =   "TicTacToe.frx":17C5F
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Image Image4 
      Height          =   5295
      Left            =   0
      Picture         =   "TicTacToe.frx":4AE1D
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Image Image3 
      Height          =   5295
      Left            =   0
      Picture         =   "TicTacToe.frx":5BCEA
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Image Image2 
      Height          =   5295
      Left            =   0
      Picture         =   "TicTacToe.frx":837A6
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
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
      Left            =   6488
      MouseIcon       =   "TicTacToe.frx":9CB20
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   60
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   255
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   60
      Width           =   255
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   255
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   60
      Width           =   255
   End
End
Attribute VB_Name = "TicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim teren(9) As Integer
Dim daca As Integer
Dim scoreu, scorel As Integer
Dim cm As Integer
Public thelevel As Integer
Public xor0 As Boolean

Private Sub aboutgame_Click()
WhichAbout = 1: GameAbout.Show
End Sub

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

Private Sub changesound_Click()
SoundChooser.Show
End Sub
Private Sub changesquares_Click()
If t(0).BackColor = &HC0FFC0 Then
  For i = 0 To 8
  t(i).BackColor = &HFFFFC0
  Shape3(i).BackColor = &HFFFF80
  Next i
  For i = 0 To 2
  Shape2(i).BackColor = &HFFFFC0
  Next i
Else
If t(0).BackColor = &HFFFFC0 Then
  For i = 0 To 8
  t(i).BackColor = &HFFC0FF
  Shape3(i).BackColor = &HFF80FF
  Next i
  For i = 0 To 2
  Shape2(i).BackColor = &HFFC0FF
  Next i
Else
If t(0).BackColor = &HFFC0FF Then
  For i = 0 To 8
  t(i).BackColor = &HC0E0FF
  Shape3(i).BackColor = &H80C0FF
  Next i
  For i = 0 To 2
  Shape2(i).BackColor = &HC0E0FF
  Next i
Else
If t(0).BackColor = &HC0E0FF Then
  For i = 0 To 8
  t(i).BackColor = vbWhite
  Shape3(i).BackColor = &HE0E0E0
  Next i
  For i = 0 To 2
  Shape2(i).BackColor = vbWhite
  Next i
Else
If t(0).BackColor = vbWhite Then
  For i = 0 To 8
  t(i).BackColor = &HC0FFFF
  Shape3(i).BackColor = &H80FFFF
  Next i
  For i = 0 To 2
  Shape2(i).BackColor = &HC0FFFF
  Next i
Else
If t(0).BackColor = &HC0FFFF Then
  For i = 0 To 8
  t(i).BackColor = &HC0C0FF
  Shape3(i).BackColor = &H8080FF
  Next i
  For i = 0 To 2
  Shape2(i).BackColor = &HC0C0FF
  Next i
Else
If t(0).BackColor = &HC0C0FF Then
  For i = 0 To 8
  t(i).BackColor = &HFFC0C0
  Shape3(i).BackColor = &HFF8080
  Next i
  For i = 0 To 2
  Shape2(i).BackColor = &HFFC0C0
  Next i
Else
If t(0).BackColor = &HFFC0C0 Then
  For i = 0 To 8
  t(i).BackColor = &HC0FFC0
  Shape3(i).BackColor = &H80FF80
  Next i
  For i = 0 To 2
  Shape2(i).BackColor = &HC0FFC0
  Next i
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub exitgame_Click()
WhichQuit = 1: GameQuit.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then WhichQuit = 1: GameQuit.Show
End Sub

Private Sub Form_Load()
Unload Game
If Language = 1 Then
    newgame2.Caption = "Joc nou"
    Level2.Caption = "Nivel mediu"
    changebackground.Caption = "Schimbã fundal"
    changesquares.Caption = "Schimbã temã"
    changesound.Caption = "Sunet"
    Mute.Caption = "Oprit"
    aboutgame.Caption = "Despre"
    HelpGame.Caption = "Ajutor"
    exitgame2.Caption = "Ieºire"
End If
Randomize
thelevel = 2
scoreu = 0
scorel = 0
ganduri = 0
daca = 0
For i = 0 To 9
teren(i) = 0
Next i
cm = Int(Rnd * 2)
If cm = 1 Then whichlevel_timer
End Sub

Private Sub finished()
If scoreu <= 10 And scorel <= 10 Then Label1.Caption = scoreu & "    -    " & scorel: GoTo 1
If scoreu <= 10 And scorel >= 10 Then Label1.Caption = scoreu & "     -   " & scorel: GoTo 1
If scoreu >= 10 And scorel <= 10 Then Label1.Caption = scoreu & "   -     " & scorel: GoTo 1
1 Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape4.BackColor = &HFF& Then Shape4.BackColor = &HC0&
If Shape5.BackColor = &HFF& Then Shape5.BackColor = &HC0&
If Shape13.BackColor = &HFF& Then Shape13.BackColor = &HC0&
If Label3.FontSize <> 9 Then Label3.FontSize = 9
End Sub

Private Sub HelpGame_Click()
WhichHelp = 1: GameHelp.Show
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
WhichAbout = 1: GameAbout.Show
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.FontSize = 9 Then Label3.FontSize = 11
End Sub

Private Sub Label4_Click()
WhichQuit = 1: GameQuit.Show
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

Private Sub Mute_Click()
SoundChooser.muzica.Command = "close"
End Sub

Private Sub whichlevel_timer()
If thelevel = 1 Then randcomp1
If thelevel = 2 Then randcomp2
If thelevel = 3 Then randcomp3
If thelevel = 4 Then whichlevel.Enabled = False: Exit Sub
whichlevel.Enabled = False
End Sub
Private Sub newgame_Click()
scoreu = 0
scorel = 0
Label1.Caption = scoreu & "    -    " & scorel
For i = 0 To 8
    teren(i) = 0
    t(i).Caption = ""
Next i
If thelevel <> 4 Then
    If cm = 1 Then
        cm = 0
        whichlevel.Enabled = True
    Else
        cm = 1
    End If
End If
End Sub

Private Sub Level_Click()
If thelevel = 1 Then
    thelevel = 2
    If Language = 1 Then
        Level2.Caption = "Nivel mediu"
    Else
        Level2.Caption = "Level medium"
    End If
    newgame_Click
ElseIf thelevel = 2 Then
    thelevel = 3
    If Language = 1 Then
        Level2.Caption = "Nivel greu"
    Else
        Level2.Caption = "Level hard"
    End If
    newgame_Click
ElseIf thelevel = 3 Then
    thelevel = 4
    If Language = 1 Then
        Level2.Caption = "Doi jucãtori"
    Else
        Level2.Caption = "Two players"
    End If
    newgame_Click
ElseIf thelevel = 4 Then
    thelevel = 1
    If Language = 1 Then
        Level2.Caption = "Nivel uºor"
    Else
        Level2.Caption = "Level easy"
    End If
    newgame_Click
End If
End Sub

Private Sub verify()
If teren(0) = teren(1) And teren(1) = teren(2) And teren(2) = 1 Then daca = 1: WhichWell = 0: GameWell.Show: scoreu = scoreu + 1: Exit Sub
If teren(3) = teren(4) And teren(4) = teren(5) And teren(5) = 1 Then daca = 1: WhichWell = 0: GameWell.Show: scoreu = scoreu + 1: Exit Sub
If teren(6) = teren(7) And teren(7) = teren(8) And teren(8) = 1 Then daca = 1: WhichWell = 0: GameWell.Show: scoreu = scoreu + 1: Exit Sub
If teren(0) = teren(3) And teren(3) = teren(6) And teren(6) = 1 Then daca = 1: WhichWell = 0: GameWell.Show: scoreu = scoreu + 1: Exit Sub
If teren(1) = teren(4) And teren(4) = teren(7) And teren(7) = 1 Then daca = 1: WhichWell = 0: GameWell.Show: scoreu = scoreu + 1: Exit Sub
If teren(2) = teren(5) And teren(5) = teren(8) And teren(8) = 1 Then daca = 1: WhichWell = 0: GameWell.Show: scoreu = scoreu + 1: Exit Sub
If teren(0) = teren(4) And teren(4) = teren(8) And teren(8) = 1 Then daca = 1: WhichWell = 0: GameWell.Show: scoreu = scoreu + 1: Exit Sub
If teren(2) = teren(4) And teren(4) = teren(6) And teren(6) = 1 Then daca = 1: WhichWell = 0: GameWell.Show: scoreu = scoreu + 1: Exit Sub

If teren(0) = teren(1) And teren(1) = teren(2) And teren(2) = 2 Then daca = 1: WhichWell = 2: GameWell.Show: scorel = scorel + 1: Exit Sub
If teren(3) = teren(4) And teren(4) = teren(5) And teren(5) = 2 Then daca = 1: WhichWell = 2: GameWell.Show: scorel = scorel + 1: Exit Sub
If teren(6) = teren(7) And teren(7) = teren(8) And teren(8) = 2 Then daca = 1: WhichWell = 2: GameWell.Show: scorel = scorel + 1: Exit Sub
If teren(0) = teren(3) And teren(3) = teren(6) And teren(6) = 2 Then daca = 1: WhichWell = 2: GameWell.Show: scorel = scorel + 1: Exit Sub
If teren(1) = teren(4) And teren(4) = teren(7) And teren(7) = 2 Then daca = 1: WhichWell = 2: GameWell.Show: scorel = scorel + 1: Exit Sub
If teren(2) = teren(5) And teren(5) = teren(8) And teren(8) = 2 Then daca = 1: WhichWell = 2: GameWell.Show: scorel = scorel + 1: Exit Sub
If teren(0) = teren(4) And teren(4) = teren(8) And teren(8) = 2 Then daca = 1: WhichWell = 2: GameWell.Show: scorel = scorel + 1: Exit Sub
If teren(2) = teren(4) And teren(4) = teren(6) And teren(6) = 2 Then daca = 1: WhichWell = 2: GameWell.Show: scorel = scorel + 1: Exit Sub

If teren(0) > 0 And teren(1) > 0 And teren(2) > 0 And teren(3) > 0 And teren(4) > 0 And teren(5) > 0 And teren(6) > 0 And teren(7) > 0 And teren(8) > 0 And daca = 0 Then daca = 1: WhichWell = 1: GameWell.Show: Exit Sub
End Sub

Private Sub t_Click(Index As Integer)
If scoreu <= 10 And scorel <= 10 Then
    Label1.Caption = scoreu & "    -    " & scorel
ElseIf scoreu <= 10 And scorel >= 10 Then
    Label1.Caption = scoreu & "     -   " & scorel
ElseIf scoreu >= 10 And scorel <= 10 Then
    Label1.Caption = scoreu & "   -     " & scorel
End If
If whichlevel.Enabled = False And Timer1.Enabled = False And teren(Index) = 0 And thelevel <> 4 Then
    teren(Index) = 1
    t(Index).Caption = "X"
    verify
    If daca = 1 Then
        finished
    Else
        whichlevel.Enabled = True
    End If
ElseIf teren(Index) = 0 And thelevel = 4 And Timer1.Enabled = False Then
    If xor0 = False Then
        teren(Index) = 2
        t(Index).Caption = "O"
        xor0 = True
    ElseIf xor0 = True Then
        teren(Index) = 1
        t(Index).Caption = "X"
        xor0 = False
    End If
    verify
    If daca = 1 Then finished
End If
End Sub
Private Sub randcomp1()
Randomize
1 Where = Int(Rnd * 9)
2  If teren(Where) = 1 Or teren(Where) = 2 Then GoTo 1
3  If teren(Where) = 0 Then
   teren(Where) = 2
   t(Where).Caption = "O"
   verify
   If daca = 1 Then finished
End If
End Sub
Private Sub randcomp2()
Randomize
Where = Int(Rnd * 9)
If teren(0) = 1 Or teren(2) = 1 Or teren(6) = 1 Or teren(8) = 1 Then Where = 4
If teren(0) = 2 And teren(1) = 2 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 2 And teren(2) = 2 And teren(1) = 0 Then Where = 1: GoTo 3
If teren(1) = 2 And teren(2) = 2 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(3) = 2 And teren(4) = 2 And teren(5) = 0 Then Where = 5: GoTo 3
If teren(3) = 2 And teren(5) = 2 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 2 And teren(5) = 2 And teren(3) = 0 Then Where = 3: GoTo 3
If teren(6) = 2 And teren(7) = 2 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(6) = 2 And teren(8) = 2 And teren(7) = 0 Then Where = 7: GoTo 3
If teren(7) = 2 And teren(8) = 2 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(0) = 2 And teren(3) = 2 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(0) = 2 And teren(6) = 2 And teren(3) = 0 Then Where = 3: GoTo 3
If teren(3) = 2 And teren(6) = 2 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(1) = 2 And teren(4) = 2 And teren(7) = 0 Then Where = 7: GoTo 3
If teren(1) = 2 And teren(7) = 2 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 2 And teren(7) = 2 And teren(1) = 0 Then Where = 1: GoTo 3
If teren(2) = 2 And teren(5) = 2 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(2) = 2 And teren(8) = 2 And teren(5) = 0 Then Where = 5: GoTo 3
If teren(5) = 2 And teren(8) = 2 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 2 And teren(4) = 2 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(0) = 2 And teren(8) = 2 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 2 And teren(8) = 2 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(2) = 2 And teren(4) = 2 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(2) = 2 And teren(6) = 2 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 2 And teren(6) = 2 And teren(2) = 0 Then Where = 2: GoTo 3

If teren(0) = 1 And teren(1) = 1 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 1 And teren(2) = 1 And teren(1) = 0 Then Where = 1: GoTo 3
If teren(1) = 1 And teren(2) = 1 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(3) = 1 And teren(4) = 1 And teren(5) = 0 Then Where = 5: GoTo 3
If teren(3) = 1 And teren(5) = 1 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 1 And teren(5) = 1 And teren(3) = 0 Then Where = 3: GoTo 3
If teren(6) = 1 And teren(7) = 1 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(6) = 1 And teren(8) = 1 And teren(7) = 0 Then Where = 7: GoTo 3
If teren(7) = 1 And teren(8) = 1 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(0) = 1 And teren(3) = 1 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(0) = 1 And teren(6) = 1 And teren(3) = 0 Then Where = 3: GoTo 3
If teren(3) = 1 And teren(6) = 1 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(1) = 1 And teren(4) = 1 And teren(7) = 0 Then Where = 7: GoTo 3
If teren(1) = 1 And teren(7) = 1 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 1 And teren(7) = 1 And teren(1) = 0 Then Where = 1: GoTo 3
If teren(2) = 1 And teren(5) = 1 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(2) = 1 And teren(8) = 1 And teren(5) = 0 Then Where = 5: GoTo 3
If teren(5) = 1 And teren(8) = 1 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 1 And teren(4) = 1 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(0) = 1 And teren(8) = 1 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 1 And teren(8) = 1 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(2) = 1 And teren(4) = 1 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(2) = 1 And teren(6) = 1 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 1 And teren(6) = 1 And teren(2) = 0 Then Where = 2: GoTo 3
GoTo 2
1 Where = Int(Rnd * 9)
2  If teren(Where) = 1 Or teren(Where) = 2 Then GoTo 1
3  If teren(Where) = 0 Then
   teren(Where) = 2
   t(Where).Caption = "O"
   verify
   If daca = 1 Then finished
End If
End Sub
Private Sub randcomp3()
Randomize
Where = 2 * Int(Rnd * 5)
If teren(0) = 2 And teren(1) = 2 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 2 And teren(2) = 2 And teren(1) = 0 Then Where = 1: GoTo 3
If teren(1) = 2 And teren(2) = 2 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(3) = 2 And teren(4) = 2 And teren(5) = 0 Then Where = 5: GoTo 3
If teren(3) = 2 And teren(5) = 2 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 2 And teren(5) = 2 And teren(3) = 0 Then Where = 3: GoTo 3
If teren(6) = 2 And teren(7) = 2 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(6) = 2 And teren(8) = 2 And teren(7) = 0 Then Where = 7: GoTo 3
If teren(7) = 2 And teren(8) = 2 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(0) = 2 And teren(3) = 2 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(0) = 2 And teren(6) = 2 And teren(3) = 0 Then Where = 3: GoTo 3
If teren(3) = 2 And teren(6) = 2 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(1) = 2 And teren(4) = 2 And teren(7) = 0 Then Where = 7: GoTo 3
If teren(1) = 2 And teren(7) = 2 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 2 And teren(7) = 2 And teren(1) = 0 Then Where = 1: GoTo 3
If teren(2) = 2 And teren(5) = 2 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(2) = 2 And teren(8) = 2 And teren(5) = 0 Then Where = 5: GoTo 3
If teren(5) = 2 And teren(8) = 2 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 2 And teren(4) = 2 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(0) = 2 And teren(8) = 2 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 2 And teren(8) = 2 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(2) = 2 And teren(4) = 2 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(2) = 2 And teren(6) = 2 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 2 And teren(6) = 2 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 1 And teren(1) = 1 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 1 And teren(2) = 1 And teren(1) = 0 Then Where = 1: GoTo 3
If teren(1) = 1 And teren(2) = 1 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(3) = 1 And teren(4) = 1 And teren(5) = 0 Then Where = 5: GoTo 3
If teren(3) = 1 And teren(5) = 1 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 1 And teren(5) = 1 And teren(3) = 0 Then Where = 3: GoTo 3
If teren(6) = 1 And teren(7) = 1 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(6) = 1 And teren(8) = 1 And teren(7) = 0 Then Where = 7: GoTo 3
If teren(7) = 1 And teren(8) = 1 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(0) = 1 And teren(3) = 1 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(0) = 1 And teren(6) = 1 And teren(3) = 0 Then Where = 3: GoTo 3
If teren(3) = 1 And teren(6) = 1 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(1) = 1 And teren(4) = 1 And teren(7) = 0 Then Where = 7: GoTo 3
If teren(1) = 1 And teren(7) = 1 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 1 And teren(7) = 1 And teren(1) = 0 Then Where = 1: GoTo 3
If teren(2) = 1 And teren(5) = 1 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(2) = 1 And teren(8) = 1 And teren(5) = 0 Then Where = 5: GoTo 3
If teren(5) = 1 And teren(8) = 1 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 1 And teren(4) = 1 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(0) = 1 And teren(8) = 1 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 1 And teren(8) = 1 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(2) = 1 And teren(4) = 1 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(2) = 1 And teren(6) = 1 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(4) = 1 And teren(6) = 1 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 1 And teren(8) = 1 And teren(4) = 0 Then Where = 4: GoTo 3
If teren(2) = 1 And teren(6) = 1 And teren(4) = 0 Then Where = 4: GoTo 3

If teren(0) = 0 And teren(1) = 1 And teren(2) = 0 And teren(3) = 1 And teren(4) = 2 And teren(5) = 0 And teren(6) = 0 And teren(7) = 0 And teren(8) = 0 Then Where = 0: GoTo 3
If teren(0) = 0 And teren(1) = 1 And teren(2) = 0 And teren(3) = 0 And teren(4) = 2 And teren(5) = 1 And teren(6) = 0 And teren(7) = 0 And teren(8) = 0 Then Where = 2: GoTo 3
If teren(0) = 0 And teren(1) = 0 And teren(2) = 0 And teren(3) = 0 And teren(4) = 2 And teren(5) = 1 And teren(6) = 0 And teren(7) = 1 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(0) = 0 And teren(1) = 0 And teren(2) = 0 And teren(3) = 1 And teren(4) = 2 And teren(5) = 0 And teren(6) = 0 And teren(7) = 1 And teren(8) = 0 Then Where = 6: GoTo 3
If teren(0) = 0 And teren(1) = 1 And teren(2) = 0 And teren(3) = 0 And teren(4) = 0 And teren(5) = 0 And teren(6) = 0 And teren(7) = 0 And teren(8) = 0 Then Where = 4: GoTo 3
If teren(0) = 0 And teren(1) = 0 And teren(2) = 0 And teren(3) = 1 And teren(4) = 0 And teren(5) = 0 And teren(6) = 0 And teren(7) = 0 And teren(8) = 0 Then Where = 4: GoTo 3
If teren(0) = 0 And teren(1) = 0 And teren(2) = 0 And teren(3) = 0 And teren(4) = 0 And teren(5) = 1 And teren(6) = 0 And teren(7) = 0 And teren(8) = 0 Then Where = 4: GoTo 3
If teren(0) = 0 And teren(1) = 0 And teren(2) = 0 And teren(3) = 0 And teren(4) = 0 And teren(5) = 0 And teren(6) = 0 And teren(7) = 1 And teren(8) = 0 Then Where = 4: GoTo 3

If teren(0) = 1 And teren(4) = 1 And teren(8) = 2 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 1 And teren(4) = 1 And teren(8) = 2 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(0) = 2 And teren(4) = 1 And teren(8) = 1 And teren(2) = 0 Then Where = 2: GoTo 3
If teren(0) = 2 And teren(4) = 1 And teren(8) = 1 And teren(6) = 0 Then Where = 6: GoTo 3
If teren(2) = 1 And teren(4) = 1 And teren(6) = 2 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(2) = 1 And teren(4) = 1 And teren(6) = 2 And teren(8) = 0 Then Where = 8: GoTo 3
If teren(2) = 2 And teren(4) = 1 And teren(6) = 1 And teren(0) = 0 Then Where = 0: GoTo 3
If teren(2) = 2 And teren(4) = 1 And teren(6) = 1 And teren(8) = 0 Then Where = 8: GoTo 3

If teren(0) = 0 And teren(1) = 1 And teren(2) = 0 And ((teren(5) = 0 And teren(8) = 1) Or (teren(5) = 1 And teren(8) = 0)) Then Where = 2: GoTo 3
If teren(0) = 0 And teren(1) = 1 And teren(2) = 0 And ((teren(3) = 0 And teren(6) = 1) Or (teren(3) = 1 And teren(6) = 0)) Then Where = 0: GoTo 3
If teren(2) = 0 And teren(5) = 1 And teren(8) = 0 And ((teren(6) = 0 And teren(7) = 1) Or (teren(6) = 1 And teren(7) = 0)) Then Where = 8: GoTo 3
If teren(2) = 0 And teren(5) = 1 And teren(8) = 0 And ((teren(0) = 0 And teren(1) = 1) Or (teren(0) = 1 And teren(1) = 0)) Then Where = 2: GoTo 3
If teren(8) = 0 And teren(7) = 1 And teren(6) = 0 And ((teren(3) = 0 And teren(0) = 1) Or (teren(3) = 1 And teren(0) = 0)) Then Where = 6: GoTo 3
If teren(8) = 0 And teren(7) = 1 And teren(6) = 0 And ((teren(5) = 0 And teren(2) = 1) Or (teren(5) = 1 And teren(2) = 0)) Then Where = 8: GoTo 3
If teren(6) = 0 And teren(3) = 1 And teren(0) = 0 And ((teren(1) = 0 And teren(2) = 1) Or (teren(1) = 1 And teren(2) = 0)) Then Where = 0: GoTo 3
If teren(6) = 0 And teren(3) = 1 And teren(0) = 0 And ((teren(7) = 0 And teren(8) = 1) Or (teren(7) = 1 And teren(8) = 0)) Then Where = 6: GoTo 3

If teren(0) = 0 And teren(1) = 0 And teren(2) = 0 And teren(3) = 0 And teren(4) = 0 And teren(5) = 0 And teren(6) = 0 And teren(7) = 0 And teren(8) = 0 Then GoTo 2
If teren(0) = 0 And teren(1) = 0 And teren(2) = 0 And teren(3) = 0 And teren(4) = 1 And teren(5) = 0 And teren(6) = 0 And teren(7) = 0 And teren(8) = 0 Then GoTo 2

If (teren(0) = 1 Or teren(2) = 1 Or teren(6) = 1 Or teren(8) = 1) And teren(4) = 0 Then Where = 4: GoTo 3

Randomize: Where = 2 * Int(Rnd * 4) + 1
If teren(Where) = 1 Or teren(Where) = 2 Then Randomize: Where = 2 * Int(Rnd * 4) + 1
If teren(Where) = 1 Or teren(Where) = 2 Then Randomize: Where = 2 * Int(Rnd * 4) + 1
If teren(Where) = 1 Or teren(Where) = 2 Then Randomize: Where = 2 * Int(Rnd * 4) + 1
If teren(Where) = 1 Or teren(Where) = 2 Then Randomize: Where = 2 * Int(Rnd * 4) + 1
GoTo 3
1 Randomize: Where = Int((Rnd * 9) + 0)
GoTo 3
2 Randomize: Where = 2 * Int(Rnd * 5)
If Where = 4 And teren(4) > 0 Then GoTo 2
3 If teren(Where) = 1 Or teren(Where) = 2 Or Where > 8 Or Where < 0 Then GoTo 1
  If teren(Where) = 0 Then
   teren(Where) = 2
   t(Where).Caption = "O"
   verify
   If daca = 1 Then finished
End If
End Sub

Private Sub Timer1_Timer()
For i = 0 To 8
 teren(i) = 0
 t(i).Caption = ""
Next i
daca = 0
Timer1.Enabled = False
Randomize
If cm = 1 Then
  cm = 0
  whichlevel.Enabled = True
Else
  cm = 1
End If
End Sub
